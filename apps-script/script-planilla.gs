/**
 * Mapa de clientes: conexión entre la app y esta planilla.
 *
 * Estructura (se arma sola la primera vez, a partir de la hoja original):
 *  - Clientes: un resumen general por cliente (datos, zona, estado y última visita).
 *  - Visitas:  una fila por visita (fecha, hora, cliente, zona y resumen).
 *  - Mapa:     ubicación de cada cliente (hoja técnica, oculta).
 *  - Base original: la hoja que había antes, sin cambios, como respaldo.
 *
 * La app lee todo con doGet(?accion=todo) y escribe con doPost:
 *  - "alta": cliente nuevo (si no tiene código, se le asigna SC-001, SC-002…)
 *  - "lote" / "edicion": cambios de datos, ubicación, zona y estado
 *  - "registro_lote" / "visita" / "desvisita": visitas y resúmenes
 *  - "reinicio": fecha de corte de "Reiniciar visitas del mes" (las visitas anteriores
 *    dejan de contar como visitadas, pero quedan en el historial)
 *  - "cobrada": marca o desmarca una factura como cobrada (hoja Cobradas)
 *
 * Cobranzas: actualizarDeudas() entra al sistema de la empresa con el usuario y la clave
 * guardados en las propiedades del script (SISTEMA_USUARIO y SISTEMA_CLAVE), trae la cuenta
 * de cada cliente y la guarda en la hoja Deudas. Se ejecuta cada mañana (ver
 * instalarActualizacionDiaria) y también desde el botón "Actualizar" de la app.
 *
 * Productos: actualizarProductos() lee las facturas (PDF) de los últimos 12 meses de cada cliente,
 * guarda los renglones en la hoja Productos (oculta) y arma el resumen que muestra la ficha
 * ("Compró anteriormente" y "Podría interesarle"). También completa la columna Rubro cuando está
 * vacía. La primera carga se hace en tandas solas (ejecutar cargarProductosAhora una vez); después
 * se actualiza cada mañana junto con las deudas.
 */
const H_CLIENTES = 'Clientes', H_VISITAS = 'Visitas', H_MAPA = 'Mapa';
const COLS_CLIENTES = ['Código', 'Nombre', 'Razón social', 'Dirección', 'Localidad', 'Partido', 'Teléfono', 'Contacto', 'Rubro', 'Comentarios', 'Horario', 'Plazo de pago', 'Medio de pago', 'Zona', 'Estado', 'Deuda', 'Última visita', 'Último resumen'];
const CAMPO_COL = { codigo: 'codigo', nombre: 'nombre', razon: 'razon social', direccion: 'direccion', localidad: 'localidad', partido: 'partido', telefono: 'telefono', contacto: 'contacto', rubro: 'rubro', notas: 'comentarios', horario: 'horario', plazo: 'plazo de pago', medio: 'medio de pago', zona: 'zona', estado: 'estado' };
const H_DEUDAS = 'Deudas', H_COBRADAS = 'Cobradas';
const SISTEMA = 'https://tresgeventas.com.ar';
const COLS_VISITAS = ['Fecha', 'Hora', 'Código', 'Nombre', 'Dirección', 'Zona', 'Resumen'];

/* ================= Lectura ================= */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    estructura(ss);
    const props = PropertiesService.getScriptProperties();
    if (p.accion === 'todo') return out({ ok: true, clientes: leerClientes(ss), visitas: leerVisitas(ss), reinicio: Number(props.getProperty('reinicio') || 0),
      deudas: leerDeudas(ss), cobradas: leerCobradas(ss), deudasAct: Number(props.getProperty('deudasAct') || 0),
      productos: leerResumenProductos(ss), prodAct: Number(props.getProperty('prodAct') || 0), prodPend: Number(props.getProperty('prodPend') || 0) });
    if (p.accion === 'actualizarDeudas') return out(actualizarDeudas(ss));
    if (p.accion === 'registro') return out({ ok: true, filas: leerVisitas(ss) });
    return out({ ok: true, msg: 'Conectado' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function leerClientes(ss) {
  const cl = ss.getSheetByName(H_CLIENTES);
  const data = cl.getDataRange().getDisplayValues();
  const head = data[0].map(norm);
  const conocidas = new Set(Object.values(CAMPO_COL).concat(['ultima visita', 'ultimo resumen', 'deuda']));
  const mapa = {};
  ss.getSheetByName(H_MAPA).getDataRange().getDisplayValues().slice(1)
    .forEach(f => { if (f[0]) mapa[norm(f[0])] = { lat: f[1], lng: f[2] }; });
  return data.slice(1).filter(f => f.some(v => String(v).trim())).map(f => {
    const o = { extras: {} };
    head.forEach((h, i) => {
      const campo = Object.keys(CAMPO_COL).find(k => CAMPO_COL[k] === h);
      if (campo) o[campo] = f[i];
      else if (h && !conocidas.has(h) && String(f[i]).trim()) o.extras[data[0][i]] = f[i];
    });
    const m = mapa[norm(o.codigo)] || {};
    o.lat = m.lat || ''; o.lng = m.lng || '';
    return o;
  });
}

function leerVisitas(ss) {
  const v = ss.getSheetByName(H_VISITAS).getDataRange().getDisplayValues().slice(1);
  return v.slice(-4000).map(f => ({ fecha: f[0], hora: f[1], codigo: f[2], nombre: f[3], direccion: f[4], zona: f[5], resumen: f[6] }));
}

/* ================= Escritura ================= */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    estructura(ss);
    let r;
    if (d.tipo === 'alta') r = alta(ss, d);
    else if (d.tipo === 'cobrada') r = marcarCobrada(ss, d);
    else if (d.tipo === 'reinicio') { PropertiesService.getScriptProperties().setProperty('reinicio', String(Number(d.ts) || Date.now())); r = { ok: true }; }
    else if (d.tipo === 'lote') r = { ok: true, res: aplicarLote(ss, d.items || []) };
    else if (d.tipo === 'edicion') r = aplicarLote(ss, [{ orig: d.orig, campos: d.campos }])[0];
    else if (d.tipo === 'registro_lote') { (d.items || []).forEach(it => registrar(ss, it)); r = { ok: true }; }
    else if (d.tipo === 'visita' || d.tipo === 'desvisita') r = registrar(ss, d);
    else if (d.texto) r = registrar(ss, { tipo: 'visita', fecha: d.fecha, hora: '', codigo: d.codigo, nombre: d.nombre, direccion: d.direccion, resumen: d.texto }); // versiones viejas de la app
    else r = { ok: false, error: 'Pedido desconocido' };
    if (d.tipo !== 'lote' && d.tipo !== 'edicion' && d.tipo !== 'cobrada') actualizarResumenes(ss);
    return out(r);
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function idxClientes(cl) {
  const head = cl.getRange(1, 1, 1, cl.getLastColumn()).getDisplayValues()[0].map(norm);
  const idx = {};
  Object.keys(CAMPO_COL).forEach(k => { idx[k] = head.indexOf(CAMPO_COL[k]); });
  idx.ultima = head.indexOf('ultima visita');
  idx.resumen = head.indexOf('ultimo resumen');
  return idx;
}

// Busca la fila: por código; si no tiene código válido, por nombre y dirección
function findRow(data, idx, d) {
  const cod = String(d.codigo || '').trim();
  const codOk = cod && !esSinCodigo(cod);
  const dirs = [norm(d.direccion), norm([d.direccion, d.localidad].filter(Boolean).join(', '))];
  for (let r = 1; r < data.length; r++) {
    if (codOk) { if (norm(data[r][idx.codigo]) === norm(cod)) return r; }
    else if (norm(data[r][idx.nombre]) === norm(d.nombre) && dirs.indexOf(norm(data[r][idx.direccion])) >= 0) return r;
  }
  return -1;
}

function prepararCampos(c0, idx) {
  const c = Object.assign({}, c0);
  if (!(idx && idx.localidad >= 0)) {
    if (c.localidad) c.direccion = [c.direccion, c.localidad].filter(Boolean).join(', ');
    delete c.localidad;
  }
  Object.keys(c).forEach(k => { c[k] = c[k] == null ? '' : String(c[k]); });
  return c;
}

function alta(ss, d) {
  const cl = ss.getSheetByName(H_CLIENTES);
  const data = cl.getDataRange().getDisplayValues();
  const idx = idxClientes(cl);
  const existente = findRow(data, idx, d);
  if (existente >= 0) return { ok: true, existente: true, codigo: data[existente][idx.codigo] };
  const c = prepararCampos(d, idx);
  let cod = String(c.codigo || '').trim();
  if (!cod || esSinCodigo(cod)) cod = nuevoCodigoSC(data.map(f => f[idx.codigo]));
  c.codigo = cod;
  if (!c.estado) c.estado = 'Activo';
  const fila = new Array(data[0].length).fill('');
  Object.keys(idx).forEach(k => { if (idx[k] >= 0 && c[k] !== undefined) fila[idx[k]] = c[k]; });
  cl.appendRow(fila);
  if (c.lat && c.lng) guardarUbicaciones(ss, [[cod, c.lat, c.lng]]);
  return { ok: true, alta: true, codigo: cod };
}

// Aplica varios cambios de una vez, escribiendo columna por columna
function aplicarLote(ss, items) {
  const cl = ss.getSheetByName(H_CLIENTES);
  const data = cl.getDataRange().getDisplayValues();
  const idx = idxClientes(cl);
  const cambios = {}, ubic = [], renombres = [];
  const res = items.map(it => {
    const r = findRow(data, idx, it.orig || {});
    if (r < 0) return { ok: false, error: 'Cliente no encontrado en la planilla' };
    const c = prepararCampos(it.campos || {}, idx);
    const codViejo = data[r][idx.codigo];
    if (c.codigo !== undefined) {
      if (!c.codigo.trim() || esSinCodigo(c.codigo)) delete c.codigo;
      else if (norm(c.codigo) !== norm(codViejo)) renombres.push([codViejo, c.codigo]);
    }
    Object.keys(c).forEach(k => {
      const i = idx[k];
      if (i === undefined || i < 0) return;
      (cambios[i] = cambios[i] || {})[r] = c[k];
      data[r][i] = c[k];
    });
    if (c.lat !== undefined || c.lng !== undefined) ubic.push([data[r][idx.codigo], c.lat || '', c.lng || '']);
    return { ok: true };
  });
  escribirColumnas(cl, data.length, cambios);
  renombres.forEach(([a, b]) => renombrarCodigo(ss, a, b));
  if (ubic.length) guardarUbicaciones(ss, ubic);
  if (renombres.length) actualizarResumenes(ss);
  return res;
}

function escribirColumnas(sh, filas, cambios) {
  if (filas < 2) return;
  Object.keys(cambios).forEach(k => {
    const i = Number(k);
    const rango = sh.getRange(2, i + 1, filas - 1, 1);
    const valores = rango.getValues();
    Object.keys(cambios[i]).forEach(r => { valores[Number(r) - 1][0] = cambios[i][r]; });
    rango.setValues(valores);
  });
}

function guardarUbicaciones(ss, lista) {
  const mp = ss.getSheetByName(H_MAPA);
  const data = mp.getDataRange().getDisplayValues();
  const pos = {};
  data.forEach((f, r) => { if (r > 0 && f[0]) pos[norm(f[0])] = r; });
  const nuevas = [];
  lista.forEach(([cod, lat, lng]) => {
    const r = pos[norm(cod)];
    if (r !== undefined) { data[r][1] = lat; data[r][2] = lng; }
    else { nuevas.push([cod, lat, lng]); pos[norm(cod)] = -1; }
  });
  if (data.length > 1) mp.getRange(2, 1, data.length - 1, 3).setValues(data.slice(1).map(f => [f[0], f[1], f[2]]));
  if (nuevas.length) mp.getRange(mp.getLastRow() + 1, 1, nuevas.length, 3).setNumberFormat('@').setValues(nuevas);
}

function renombrarCodigo(ss, viejo, nuevo) {
  [[H_MAPA, 0], [H_VISITAS, 2]].forEach(([nombre, col]) => {
    const sh = ss.getSheetByName(nombre);
    if (sh.getLastRow() < 2) return;
    const rango = sh.getRange(2, col + 1, sh.getLastRow() - 1, 1);
    const v = rango.getDisplayValues();
    let cambio = false;
    v.forEach(f => { if (norm(f[0]) === norm(viejo)) { f[0] = nuevo; cambio = true; } });
    if (cambio) rango.setValues(v);
  });
}

/* ================= Visitas ================= */
function registrar(ss, d) {
  const vis = ss.getSheetByName(H_VISITAS);
  const v = vis.getDataRange().getDisplayValues();
  const cod = String(d.codigo || '').trim();
  const codOk = cod && !esSinCodigo(cod);
  let fila = -1;
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] !== d.fecha) continue;
    if (codOk ? norm(v[i][2]) === norm(cod) : norm(v[i][3]) === norm(d.nombre)) { fila = i; break; }
  }
  if (d.tipo === 'desvisita') {
    if (fila >= 0) vis.deleteRow(fila + 1);
    return { ok: true };
  }
  if (fila >= 0) {
    if (d.resumen) vis.getRange(fila + 1, 7).setValue(d.resumen);
    return { ok: true, actualizado: true };
  }
  vis.appendRow([d.fecha, d.hora || '', cod || 'S/C', d.nombre || '', d.direccion || '', d.zona || '', d.resumen || '']);
  return { ok: true, registrado: true };
}

// Completa "Última visita" y "Último resumen" de la hoja Clientes a partir de Visitas
function actualizarResumenes(ss) {
  const cl = ss.getSheetByName(H_CLIENTES);
  const idx = idxClientes(cl);
  if (idx.ultima < 0 || idx.resumen < 0 || cl.getLastRow() < 2) return;
  const ult = {}, res = {};
  ss.getSheetByName(H_VISITAS).getDataRange().getDisplayValues().slice(1).forEach(f => {
    const k = norm(f[2]); if (!k) return;
    const t = tsFecha(f[0], f[1]);
    if (t && (!ult[k] || t >= ult[k].t)) ult[k] = { t: t, fecha: f[0] };
    if (f[6] && (!res[k] || t >= res[k].t)) res[k] = { t: t, txt: (t ? f[0] + ' - ' : '') + f[6] };
  });
  const n = cl.getLastRow() - 1;
  const cods = cl.getRange(2, idx.codigo + 1, n, 1).getDisplayValues();
  cl.getRange(2, idx.ultima + 1, n, 1).setValues(cods.map(c => [ult[norm(c[0])] ? ult[norm(c[0])].fecha : '']));
  cl.getRange(2, idx.resumen + 1, n, 1).setValues(cods.map(c => [res[norm(c[0])] ? res[norm(c[0])].txt : '']));
}

/* ================= Estructura y migración ================= */
function estructura(ss) {
  if (!ss.getSheetByName(H_CLIENTES)) migrar(ss);
  // Si ya existe Clientes pero faltan otras hojas (por ejemplo, después de importar un archivo), se crean vacías
  if (!ss.getSheetByName(H_VISITAS)) {
    const v = ss.insertSheet(H_VISITAS, Math.min(1, ss.getSheets().length));
    v.getRange('A:G').setNumberFormat('@');
    v.getRange(1, 1, 1, COLS_VISITAS.length).setValues([COLS_VISITAS]);
    formatearEncabezado(v, COLS_VISITAS.length);
  }
  if (!ss.getSheetByName(H_MAPA)) {
    const m = ss.insertSheet(H_MAPA, ss.getSheets().length);
    m.getRange('A:C').setNumberFormat('@');
    m.getRange(1, 1, 1, 3).setValues([['Código', 'Lat', 'Lng']]);
    formatearEncabezado(m, 3);
    m.hideSheet();
  }
  asegurarHorario(ss);
  const cl = ss.getSheetByName(H_CLIENTES);
  asegurarColumna(cl, 'Razón social', 'nombre', '@');
  asegurarColumna(cl, 'Localidad', 'direccion', '@');
  asegurarColumna(cl, 'Partido', 'localidad', '@');
  asegurarColumna(cl, 'Plazo de pago', 'horario', '@');
  asegurarColumna(cl, 'Medio de pago', 'plazo de pago', '@');
  asegurarColumna(cl, 'Deuda', 'estado', '"$"#,##0.00');
}

// Agrega una columna en la hoja Clientes si falta, a continuación de otra
function asegurarColumna(cl, titulo, despuesDe, formato) {
  const n = cl.getLastColumn();
  const nh = cl.getRange(1, 1, 1, n).getDisplayValues()[0].map(norm);
  if (nh.indexOf(norm(titulo)) >= 0) return;
  const i = nh.indexOf(despuesDe);
  const pos = i >= 0 ? i + 1 : n;
  cl.insertColumnsAfter(pos, 1);
  cl.getRange(1, pos + 1, cl.getMaxRows(), 1).setNumberFormat(formato || '@');
  cl.getRange(1, pos + 1).setNumberFormat('@').setValue(titulo)
    .setFontWeight('bold').setBackground('#14313a').setFontColor('#ffffff');
}

// Separa "Comentarios / Horario" en "Comentarios" y "Horario"
function asegurarHorario(ss) {
  const cl = ss.getSheetByName(H_CLIENTES); if (!cl) return;
  const n = cl.getLastColumn();
  const nh = cl.getRange(1, 1, 1, n).getDisplayValues()[0].map(norm);
  let iCom = nh.indexOf('comentarios horario');
  if (iCom >= 0) { cl.getRange(1, iCom + 1).setValue('Comentarios'); nh[iCom] = 'comentarios'; }
  iCom = nh.indexOf('comentarios');
  if (nh.indexOf('horario') >= 0) return;
  const despues = iCom >= 0 ? iCom + 1 : n;
  cl.insertColumnsAfter(despues, 1);
  cl.getRange(1, despues + 1, cl.getMaxRows(), 1).setNumberFormat('@');
  cl.getRange(1, despues + 1, 1, 1).setValues([['Horario']])
    .setFontWeight('bold').setBackground('#14313a').setFontColor('#ffffff');
}

// También se puede ejecutar a mano desde el editor de Apps Script
function migrar(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(H_CLIENTES) && ss.getSheetByName(H_MAPA) && ss.getSheetByName(H_VISITAS)) return;
  const reservadas = ['clientes', 'visitas', 'mapa', 'registro'];
  const src = ss.getSheets().find(s => reservadas.indexOf(norm(s.getName())) < 0 && !/^base original/.test(norm(s.getName()))) || ss.getSheets()[0];
  const data = src.getDataRange().getDisplayValues();
  const head = data[0].map(norm);
  const col = al => head.findIndex(h => al.some(a => h === a || h.startsWith(a + ' ')));
  const i = {
    codigo: col(['codigo', 'cod']), nombre: col(['nombre', 'cliente', 'razon social', 'comercio']),
    direccion: col(['direccion', 'domicilio']), localidad: col(['localidad', 'barrio', 'ciudad', 'partido']),
    telefono: col(['telefono', 'tel', 'celular', 'whatsapp']), contacto: col(['contacto', 'encargado', 'responsable']),
    rubro: col(['rubro', 'categoria']), notas: col(['comentarios', 'comentario', 'notas', 'observaciones', 'obs', 'horario']),
    lat: head.indexOf('lat'), lng: head.indexOf('lng'), zona: head.indexOf('zona mapa'), estado: head.indexOf('estado mapa')
  };
  const usadas = new Set(Object.values(i).filter(x => x >= 0));
  const vcols = head.map((h, k) => /^visitas? ?\d*$/.test(h) ? { k: k, n: Number((h.match(/\d+/) || ['1'])[0]) } : null).filter(Boolean).sort((a, b) => a.n - b.n);
  vcols.forEach(v => usadas.add(v.k));
  const extras = head.map((h, k) => (h && !usadas.has(k) && !h.startsWith('zona')) ? k : -1).filter(k => k >= 0);
  const filas = data.slice(1).filter(f => f.some(v => String(v).trim()));
  const val = (f, k) => k >= 0 ? String(f[k] || '').trim() : '';

  // Códigos: los "S/C" o vacíos reciben SC-001, SC-002…
  const todos = filas.map(f => val(f, i.codigo));
  filas.forEach((f, r) => { if (!todos[r] || esSinCodigo(todos[r])) todos[r] = nuevoCodigoSC(todos); });

  // Hoja original como respaldo
  src.setName(nombreLibre(ss, 'Base original'));

  // Clientes
  const cl = ss.insertSheet(H_CLIENTES, 0);
  const encab = COLS_CLIENTES.concat(extras.map(k => data[0][k]));
  const valores = filas.map((f, r) => {
    const dir = [val(f, i.direccion), val(f, i.localidad)].filter(Boolean).join(', ');
    const est = /oculto/i.test(val(f, i.estado)) ? 'Oculto' : 'Activo';
    return [todos[r], val(f, i.nombre), '', val(f, i.direccion), val(f, i.localidad), '', val(f, i.telefono), val(f, i.contacto), val(f, i.rubro), val(f, i.notas), '', '', '',
            est === 'Oculto' ? '' : val(f, i.zona), est, '', '', ''].concat(extras.map(k => val(f, k)));
  });
  cl.getRange(1, 1, valores.length + 1, encab.length).setNumberFormat('@').setValues([encab].concat(valores));
  formatearEncabezado(cl, encab.length);

  // Mapa
  const mp = ss.insertSheet(H_MAPA);
  const ubic = filas.map((f, r) => [todos[r], val(f, i.lat), val(f, i.lng)]);
  mp.getRange(1, 1, ubic.length + 1, 3).setNumberFormat('@').setValues([['Código', 'Lat', 'Lng']].concat(ubic));
  formatearEncabezado(mp, 3);

  // Visitas: se reutiliza "Registro" si existe y se suman los resúmenes de las columnas Visita N
  let vis = ss.getSheetByName('Registro');
  if (vis) vis.setName(H_VISITAS);
  else { vis = ss.insertSheet(H_VISITAS); vis.getRange('A:G').setNumberFormat('@'); vis.appendRow(COLS_VISITAS); }
  formatearEncabezado(vis, 7);
  const previas = vis.getDataRange().getDisplayValues();
  const porNombre = {}; filas.forEach((f, r) => { porNombre[norm(val(f, i.nombre))] = todos[r]; });
  previas.forEach((f, r) => { if (r > 0 && (!f[2] || esSinCodigo(f[2])) && porNombre[norm(f[3])]) f[2] = porNombre[norm(f[3])]; });
  if (previas.length > 1) vis.getRange(2, 1, previas.length - 1, 7).setValues(previas.slice(1));
  const ya = new Set(previas.slice(1).map(f => f[0] + '|' + norm(f[2]) + '|' + norm(f[6])));
  const nuevas = [];
  filas.forEach((f, r) => {
    vcols.forEach(v => {
      const t = val(f, v.k); if (!t) return;
      const m = t.match(/^(\d{1,2}\/\d{1,2}\/\d{4})\s*-\s*([\s\S]*)$/);
      const fecha = m ? m[1] : 'Anterior', txt = m ? m[2] : t;
      const clave = fecha + '|' + norm(todos[r]) + '|' + norm(txt);
      if (ya.has(clave)) return; ya.add(clave);
      nuevas.push([fecha, '', todos[r], val(f, i.nombre), [val(f, i.direccion), val(f, i.localidad)].filter(Boolean).join(', '), val(f, i.zona), txt]);
    });
  });
  if (nuevas.length) vis.getRange(vis.getLastRow() + 1, 1, nuevas.length, 7).setNumberFormat('@').setValues(nuevas);

  // Orden de las hojas
  [cl, vis, mp].forEach((sh, pos) => { ss.setActiveSheet(sh); ss.moveActiveSheet(pos + 1); });
  ss.setActiveSheet(src); ss.moveActiveSheet(ss.getSheets().length);
  mp.hideSheet();
  ss.setActiveSheet(cl);
  actualizarResumenes(ss);
}

function formatearEncabezado(sh, n) {
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#14313a').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

/* ================= Cobranzas ================= */
function credenciales() {
  const p = PropertiesService.getScriptProperties();
  return { usuario: p.getProperty('SISTEMA_USUARIO'), clave: p.getProperty('SISTEMA_CLAVE') };
}
function guardarCookies(resp, jar) {
  const h = resp.getAllHeaders();
  let sc = h['Set-Cookie'] || h['set-cookie'];
  if (!sc) return jar;
  if (!Array.isArray(sc)) sc = [sc];
  sc.forEach(c => { const kv = String(c).split(';')[0]; const i = kv.indexOf('='); if (i > 0) jar[kv.slice(0, i).trim()] = kv.slice(i + 1).trim(); });
  return jar;
}
function textoCookies(jar) { return Object.keys(jar).map(k => k + '=' + jar[k]).join('; '); }

// Inicia sesión en el sistema y devuelve las cookies de la sesión
function loginSistema() {
  const cr = credenciales();
  if (!cr.usuario || !cr.clave) throw new Error('Falta cargar SISTEMA_USUARIO y SISTEMA_CLAVE en las propiedades del script');
  const jar = {};
  guardarCookies(UrlFetchApp.fetch(SISTEMA + '/login', { muteHttpExceptions: true, followRedirects: false }), jar);
  const r1 = UrlFetchApp.fetch(SISTEMA + '/login', {
    method: 'post', payload: { name: cr.usuario, pass: cr.clave },
    headers: { Cookie: textoCookies(jar) }, followRedirects: false, muteHttpExceptions: true
  });
  guardarCookies(r1, jar);
  const h = r1.getAllHeaders(), loc = h['Location'] || h['location'];
  if (loc) guardarCookies(UrlFetchApp.fetch(/^http/.test(loc) ? loc : SISTEMA + loc, { headers: { Cookie: textoCookies(jar) }, followRedirects: false, muteHttpExceptions: true }), jar);
  const u = UrlFetchApp.fetch(SISTEMA + '/usuario', { headers: { Cookie: textoCookies(jar), 'X-Requested-With': 'XMLHttpRequest', Accept: 'application/json' }, muteHttpExceptions: true });
  let j = null; try { j = JSON.parse(u.getContentText()); } catch (e) {}
  if (!j || !j.clients) throw new Error('No pude iniciar sesión en el sistema: revisá el usuario y la clave');
  return jar;
}

// Para probar desde el editor: Ejecutar → probarConexionSistema
function probarConexionSistema() {
  loginSistema();
  Logger.log('Conexión correcta con el sistema');
}

function fechaSistema(s) {
  const m = String(s || '').match(/(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}
function fmtFecha(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy'); }

// Trae la cuenta de cada cliente de la hoja Clientes y la guarda en la hoja Deudas
function actualizarDeudas(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  estructura(ss);
  const jar = loginSistema(), ck = textoCookies(jar);
  const cl = ss.getSheetByName(H_CLIENTES);
  const data = cl.getDataRange().getDisplayValues();
  const head = data[0].map(norm);
  const iCod = head.indexOf('codigo'), iNom = head.indexOf('nombre'), iPlazo = head.indexOf('plazo de pago'),
        iMedio = head.indexOf('medio de pago'), iDeuda = head.indexOf('deuda');
  const clientes = [];
  data.slice(1).forEach((f, r) => {
    const cod = String(f[iCod] || '').trim();
    if (/^\d+$/.test(cod)) clientes.push({ fila: r + 1, codigo: cod, nombre: f[iNom], plazo: parseInt(f[iPlazo], 10) || 30, medio: iMedio >= 0 ? f[iMedio] : '' });
  });
  const pedidos = clientes.map(c => ({ url: SISTEMA + '/clientes/' + c.codigo + '/cuenta', headers: { Cookie: ck, 'X-Requested-With': 'XMLHttpRequest', Accept: 'application/json' }, muteHttpExceptions: true }));
  const resp = [];
  for (let i = 0; i < pedidos.length; i += 25) resp.push.apply(resp, UrlFetchApp.fetchAll(pedidos.slice(i, i + 25)));

  const cobradas = {}; leerCobradas(ss).forEach(x => { cobradas[norm(x.codigo) + '|' + x.numero] = true; });
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const filas = [], saldos = {}, vigentes = new Set();
  let errores = 0, conDeuda = 0;
  clientes.forEach((c, k) => {
    let a = null; try { a = JSON.parse(resp[k].getContentText()); } catch (e) {}
    if (!Array.isArray(a)) { errores++; return; }
    if (a.length) conDeuda++;
    saldos[c.fila] = a.length ? (Number(a[0].Saldo) || a.reduce((s, x) => s + (Number(x.Importe) || 0), 0)) : 0;
    a.forEach(x => {
      const f = fechaSistema(x.Fecha); if (!f) return;
      const imp = Number(x.Importe) || 0;
      const dias = Math.round((hoy - f) / 864e5);
      const vence = new Date(f.getTime()); vence.setDate(vence.getDate() + c.plazo);
      const restan = Math.round((vence - hoy) / 864e5);
      const clave = norm(c.codigo) + '|' + x.NroComprobante;
      vigentes.add(clave);
      const estado = imp <= 0 ? 'Crédito' : cobradas[clave] ? 'Cobrada' : restan < 0 ? 'Vencida' : restan <= 3 ? 'Por vencer' : 'Al día';
      filas.push([c.codigo, c.nombre, x.TipoComprobante || '', x.NroComprobante || '', fmtFecha(f), imp, dias, imp > 0 ? fmtFecha(vence) : '', estado, c.medio]);
    });
  });

  // Hoja Deudas
  let dh = ss.getSheetByName(H_DEUDAS);
  if (!dh) { dh = ss.insertSheet(H_DEUDAS, Math.min(2, ss.getSheets().length)); }
  dh.clearContents();
  const enc = ['Código', 'Nombre', 'Tipo', 'Número', 'Fecha', 'Importe', 'Días', 'Vence', 'Estado', 'Medio de pago'];
  dh.getRange(1, 1, 1, enc.length).setValues([enc]).setFontWeight('bold').setBackground('#14313a').setFontColor('#ffffff');
  dh.setFrozenRows(1);
  if (filas.length) {
    dh.getRange(2, 1, filas.length, 5).setNumberFormat('@');
    dh.getRange(2, 8, filas.length, 3).setNumberFormat('@');
    dh.getRange(2, 1, filas.length, enc.length).setValues(filas);
    dh.getRange(2, 6, filas.length, 1).setNumberFormat('"$"#,##0.00');
  }

  // Columna Deuda de Clientes
  if (iDeuda >= 0 && data.length > 1) {
    const col = cl.getRange(2, iDeuda + 1, data.length - 1, 1);
    const v = data.slice(1).map((f, r) => [saldos[r + 1] !== undefined ? saldos[r + 1] : (String(f[iCod]).match(/^\d+$/) ? 0 : '')]);
    col.setValues(v);
  }

  // Las marcas de "cobrada" de facturas que ya no figuran se borran
  const cb = ss.getSheetByName(H_COBRADAS);
  if (cb && cb.getLastRow() > 1) {
    const v = cb.getRange(2, 1, cb.getLastRow() - 1, 3).getDisplayValues().filter(f => vigentes.has(norm(f[0]) + '|' + f[1]));
    cb.getRange(2, 1, cb.getLastRow() - 1, 3).clearContent();
    if (v.length) cb.getRange(2, 1, v.length, 3).setValues(v);
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('deudasAct', String(Date.now()));
  return { ok: true, clientes: clientes.length, conDeuda: conDeuda, comprobantes: filas.length, errores: errores };
}

// Versión para el disparador diario
function actualizarDeudasProgramado() {
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  const t0 = Date.now();
  try {
    actualizarDeudas();
    try { actualizarProductos(null, Math.max(60 * 1000, 5 * 60 * 1000 - (Date.now() - t0))); } catch (e) { Logger.log('Productos: ' + e); }
  } finally { lock.releaseLock(); }
}

// Ejecutar UNA vez desde el editor: actualiza las deudas todos los días entre las 6 y las 7
function instalarActualizacionDiaria() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'actualizarDeudasProgramado').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('actualizarDeudasProgramado').timeBased().everyDays(1).atHour(6).create();
  Logger.log('Listo: las deudas se actualizan todos los días a la mañana');
}

function leerDeudas(ss) {
  const dh = ss.getSheetByName(H_DEUDAS);
  const out = {};
  if (dh && dh.getLastRow() > 1) {
    dh.getRange(2, 1, dh.getLastRow() - 1, 6).getValues().forEach(f => {
      const k = String(f[0]).trim(); if (!k) return;
      (out[k] = out[k] || { saldo: 0, items: [] }).items.push({ tipo: String(f[2]), numero: String(f[3]), fecha: f[4] instanceof Date ? fmtFecha(f[4]) : String(f[4]), importe: Number(f[5]) || 0 });
    });
  }
  const cl = ss.getSheetByName(H_CLIENTES);
  const data = cl.getDataRange().getValues();
  const head = data[0].map(norm), iCod = head.indexOf('codigo'), iDeuda = head.indexOf('deuda');
  if (iDeuda >= 0) data.slice(1).forEach(f => { const k = String(f[iCod]).trim(); if (out[k]) out[k].saldo = Number(f[iDeuda]) || 0; });
  return out;
}

function hojaCobradas(ss) {
  let cb = ss.getSheetByName(H_COBRADAS);
  if (!cb) {
    cb = ss.insertSheet(H_COBRADAS, ss.getSheets().length);
    cb.getRange('A:C').setNumberFormat('@');
    cb.getRange(1, 1, 1, 3).setValues([['Código', 'Número', 'Marcada']]).setFontWeight('bold').setBackground('#14313a').setFontColor('#ffffff');
    cb.setFrozenRows(1);
    cb.hideSheet();
  }
  return cb;
}
function leerCobradas(ss) {
  const cb = ss.getSheetByName(H_COBRADAS);
  if (!cb || cb.getLastRow() < 2) return [];
  return cb.getRange(2, 1, cb.getLastRow() - 1, 3).getDisplayValues().filter(f => f[0]).map(f => ({ codigo: f[0], numero: f[1], marcada: f[2] }));
}
function marcarCobrada(ss, d) {
  const cb = hojaCobradas(ss);
  const v = cb.getLastRow() > 1 ? cb.getRange(2, 1, cb.getLastRow() - 1, 3).getDisplayValues() : [];
  const i = v.findIndex(f => norm(f[0]) === norm(d.codigo) && f[1] === d.numero);
  if (d.valor && i < 0) cb.appendRow([String(d.codigo), String(d.numero), d.marcada || fmtFecha(new Date())]);
  if (!d.valor && i >= 0) cb.deleteRow(i + 2);
  return { ok: true };
}

/* ================= Utilidades ================= */
function esSinCodigo(c) { return /^s\s*\/?\s*c$/i.test(String(c).trim()); }

function nuevoCodigoSC(lista) {
  let max = 0;
  lista.forEach(c => { const m = String(c || '').match(/^SC-(\d+)$/i); if (m) max = Math.max(max, Number(m[1])); });
  const cod = 'SC-' + String(max + 1).padStart(3, '0');
  lista.push(cod);
  return cod;
}

function nombreLibre(ss, base) {
  let n = base, k = 2;
  while (ss.getSheetByName(n)) n = base + ' ' + (k++);
  return n;
}

function tsFecha(fecha, hora) {
  const m = String(fecha).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return 0;
  const h = String(hora || '').match(/(\d{1,2}):(\d{2})/);
  return new Date(+m[3], +m[2] - 1, +m[1], h ? +h[1] : 0, h ? +h[2] : 0).getTime();
}

function norm(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= Productos de las facturas ================= */
// Lee las facturas (PDF) de los últimos 12 meses de cada cliente, guarda sus renglones en la hoja
// Productos y arma un resumen por cliente para la app: lo que más compró y sugerencias por rubro.
const H_PRODUCTOS = 'Productos', H_PROD_RES = 'ProductosResumen';
const COLS_PRODUCTOS = ['Código', 'Fecha', 'Comprobante', 'Número', 'Artículo', 'Descripción', 'Cantidad', 'Subtotal'];
const RUBROS_NOMBRE = [
  ['Bulonera', /bulon|\bbul[- ]?maq|bulmaq|tornillo|fijacion/],
  ['Pinturería', /pintur|pintesir|\bpint\b/],
  ['Sanitarios', /sanitar|hidraulic/],
  ['Materiales eléctricos', /electr|mat\.? ?elec|\bwatt\b|energia/],
  ['Corralón', /corralon|materiales el obelisco/],
  ['Zinguería', /zingueria/],
  ['Matafuegos y seguridad', /matafuego|seguridad/],
  ['Gomas y plásticos', /goma|plast|gom-vi/],
  ['Herramientas y maquinaria', /herramienta|maq\b|maquin|soldadura|acero/],
  ['Ferretería', /ferret|^ferr|\bferr\w*|ferro/]
];

// Pasa los bytes del PDF a texto y extrae los renglones de la factura
function renglonesFactura(bytes) {
  const n = bytes.length; let s = '';
  for (let i = 0; i < n; i += 8192) s += String.fromCharCode.apply(null, bytes.slice(i, i + 8192).map(b => b & 255));
  let txt = '';
  const re = /stream\r?\n/g; let m;
  while ((m = re.exec(s))) {
    const ini = m.index + m[0].length, fin = s.indexOf('endstream', ini);
    if (fin < 0) break;
    const dic = s.slice(Math.max(0, m.index - 250), m.index);
    if (!/FlateDecode/.test(dic) || /\/Image/.test(dic)) continue;
    try { const d = inflarZlib(bytes.slice(ini, fin)); for (let i = 0; i < d.length; i += 8192) txt += String.fromCharCode.apply(null, d.slice(i, i + 8192)); txt += '\n'; } catch (e) {}
    re.lastIndex = fin;
  }
  const els = []; const rt = /1 0 0 1 ([\d.]+) ([\d.]+) Tm\s*\(((?:\\.|[^\\)])*)\)\s*Tj/g;
  while ((m = rt.exec(txt))) els.push({ x: +m[1], y: +m[2], s: m[3].replace(/\\(.)/g, '$1').trim() });
  els.sort((a, b) => a.y - b.y);
  const filas = [];
  els.forEach(e => { const l = filas[filas.length - 1]; if (l && e.y - l.y0 <= 12) l.r.push(e); else filas.push({ y0: e.y, r: [e] }); });
  // Importes impresos al pie (Sub-Total y Sub-Total con bonificación), para controlar los renglones
  const pie = t => { const et = els.find(e => e.s.replace(/\s/g, '') === t); if (!et) return 0;
    const v = els.filter(e => /^\$/.test(e.s) && e.y < et.y && e.y > et.y - 500 && Math.abs(e.x - et.x) < 700).sort((a, b) => Math.abs(a.x - et.x) - Math.abs(b.x - et.x))[0];
    return v ? Number(v.s.replace(/[$\s]/g, '')) || 0 : 0; };
  const out = [];
  out.controles = [pie('Sub-Total+Bonif.'), pie('Sub-Total')].filter(Boolean);
  filas.forEach(f => {
    const g = (a, b) => { const e = f.r.find(e => e.x >= a && e.x < b); return e ? e.s : ''; };
    const cod = g(6500, 7900), cant = g(400, 1400), desc = g(1500, 6000), sub = g(10300, 11400);
    if (cod && /^[\d.,]+$/.test(cant) && desc && /\$/.test(sub)) out.push({ cod: cod, desc: desc, cant: Number(cant), sub: Number(sub.replace(/[$\s]/g, '')) || 0 });
  });
  return out;
}

// Control de los renglones contra el Sub-Total impreso en la factura (o, si no se encuentra, contra el
// importe de la lista, que incluye IVA y percepciones). Si no cierran, se sacan renglones repetidos y,
// si todavía no cierra, se ajustan los importes en proporción.
function ajustarRenglones(items, controles, monto) {
  if (!items.length) return items;
  controles = controles || [];
  let objetivo = controles[0] || 0, cierra;
  if (objetivo) cierra = l => { const s = l.reduce((t, i) => t + i.sub, 0); return controles.some(c => Math.abs(s - c) <= Math.max(1, c * 0.01)); };
  else if (monto && monto < 999999) { objetivo = monto / 1.21; cierra = l => { const f = monto / (l.reduce((t, i) => t + i.sub, 0) || 1); return f >= 1.14 && f <= 1.32; }; }
  else return items; // sin forma de controlar
  if (cierra(items)) return items;
  const vistos = {}, unicos = items.filter(i => { const k = i.cod + '|' + i.cant + '|' + i.sub; if (vistos[k]) return false; vistos[k] = 1; return true; });
  if (cierra(unicos)) return unicos;
  const s = unicos.reduce((t, i) => t + i.sub, 0) || 1, f = objetivo / s;
  return unicos.map(i => ({ cod: i.cod, desc: i.desc, cant: i.cant, sub: Math.round(i.sub * f * 100) / 100 }));
}

function hojaProductos(ss) {
  let sh = ss.getSheetByName(H_PRODUCTOS);
  if (!sh) {
    sh = ss.insertSheet(H_PRODUCTOS, ss.getSheets().length);
    sh.getRange('A:F').setNumberFormat('@');
    sh.getRange(1, 1, 1, COLS_PRODUCTOS.length).setValues([COLS_PRODUCTOS]).setFontWeight('bold').setBackground('#14313a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

// Ejecutar desde el editor para la primera carga (sigue sola en tandas hasta terminar)
function cargarProductosAhora() { continuarProductos(); }

function continuarProductos() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { programarContinuacion(); return; }
  try { const r = actualizarProductos(); Logger.log(JSON.stringify(r)); } finally { lock.releaseLock(); }
}

function programarContinuacion() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'continuarProductos').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('continuarProductos').timeBased().after(60 * 1000).create();
}

function actualizarProductos(ss, limiteMs) {
  const t0 = Date.now(); limiteMs = limiteMs || 4.5 * 60 * 1000;
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  estructura(ss);
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'continuarProductos').forEach(t => ScriptApp.deleteTrigger(t));
  const props = PropertiesService.getScriptProperties();
  const jar = loginSistema(), ck = textoCookies(jar);
  const hdr = { Cookie: ck, 'X-Requested-With': 'XMLHttpRequest', Accept: 'application/json' };
  const cl = ss.getSheetByName(H_CLIENTES), data = cl.getDataRange().getDisplayValues(), head = data[0].map(norm);
  const iCod = head.indexOf('codigo');
  const codigos = data.slice(1).map(f => String(f[iCod] || '').trim()).filter(c => /^\d+$/.test(c));

  // Facturas de los últimos 12 meses de cada cliente
  const desde = new Date(); desde.setFullYear(desde.getFullYear() - 1); desde.setHours(0, 0, 0, 0);
  const resp = [];
  const pedidos = codigos.map(c => ({ url: SISTEMA + '/clientes/' + c + '/comprobantes', headers: hdr, muteHttpExceptions: true }));
  for (let i = 0; i < pedidos.length; i += 25) resp.push.apply(resp, UrlFetchApp.fetchAll(pedidos.slice(i, i + 25)));
  const sh = hojaProductos(ss);
  const hechos = new Set(sh.getLastRow() > 1 ? sh.getRange(2, 3, sh.getLastRow() - 1, 1).getDisplayValues().map(f => f[0]) : []);
  let pend = [];
  codigos.forEach((c, k) => {
    let a = null; try { a = JSON.parse(resp[k].getContentText()); } catch (e) {}
    if (!Array.isArray(a)) return;
    a.forEach(x => { const f = fechaSistema(x.date);
      if (f && f >= desde && /^factura/i.test(x.type || '') && !hechos.has(String(x.id))) pend.push({ cod: c, id: String(x.id), num: x.number || '', fecha: f, monto: Number(x.amount) || 0 }); });
  });
  pend.sort((a, b) => b.fecha - a.fecha); // primero lo más reciente
  const total = pend.length;
  let leidas = 0;
  while (pend.length && Date.now() - t0 < limiteMs) {
    const tanda = pend.splice(0, 20);
    const r = UrlFetchApp.fetchAll(tanda.map(x => ({ url: SISTEMA + '/comprobantes/' + x.id, headers: { Cookie: ck }, muteHttpExceptions: true })));
    const filas = [];
    tanda.forEach((x, k) => {
      if (r[k].getResponseCode() !== 200) return;
      let items = []; try { const rr = renglonesFactura(r[k].getContent()); items = ajustarRenglones(rr, rr.controles, x.monto); } catch (e) {}
      if (!items.length) filas.push([x.cod, fmtFecha(x.fecha), x.id, x.num, '', '', 0, 0]); // marca: ya leída, sin renglones
      items.forEach(it => filas.push([x.cod, fmtFecha(x.fecha), x.id, x.num, it.cod, it.desc, it.cant, it.sub]));
      leidas++;
    });
    if (filas.length) sh.getRange(sh.getLastRow() + 1, 1, filas.length, COLS_PRODUCTOS.length).setValues(filas);
  }
  if (pend.length) programarContinuacion();
  else limpiarProductosViejos(sh);
  asignarRubros(ss);
  resumirProductos(ss);
  props.setProperty('prodAct', String(Date.now()));
  props.setProperty('prodPend', String(pend.length));
  return { ok: true, facturasNuevas: total, leidas: leidas, pendientes: pend.length };
}

// Borra renglones de más de 13 meses
function limpiarProductosViejos(sh) {
  if (sh.getLastRow() < 2) return;
  const lim = new Date(); lim.setMonth(lim.getMonth() - 13);
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, COLS_PRODUCTOS.length).getValues();
  const quedan = v.filter(f => { const d = fechaTexto(f[1]); return !d || d >= lim; });
  if (quedan.length === v.length) return;
  sh.getRange(2, 1, v.length, COLS_PRODUCTOS.length).clearContent();
  if (quedan.length) sh.getRange(2, 1, quedan.length, COLS_PRODUCTOS.length).setValues(quedan);
}
function fechaTexto(s) {
  if (s instanceof Date) return s;
  const m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? new Date(+m[3], +m[2] - 1, +m[1]) : null;
}

function leerRenglones(ss) {
  const sh = ss.getSheetByName(H_PRODUCTOS);
  if (!sh || sh.getLastRow() < 2) return [];
  const desde = new Date(); desde.setFullYear(desde.getFullYear() - 1); desde.setHours(0, 0, 0, 0);
  return sh.getRange(2, 1, sh.getLastRow() - 1, COLS_PRODUCTOS.length).getValues()
    .map(f => ({ cli: String(f[0]).trim(), fecha: fechaTexto(f[1]), art: String(f[4]).trim(), desc: String(f[5]).trim(), cant: Number(f[6]) || 0, sub: Number(f[7]) || 0 }))
    .filter(r => r.art && r.fecha && r.fecha >= desde);
}

// Rubro: primero por el nombre; si no se puede, por el parecido de lo que compra con cada rubro.
// Solo completa los clientes con el Rubro vacío.
function asignarRubros(ss) {
  const cl = ss.getSheetByName(H_CLIENTES), data = cl.getDataRange().getDisplayValues(), head = data[0].map(norm);
  const iCod = head.indexOf('codigo'), iNom = head.indexOf('nombre'), iRaz = head.indexOf('razon social'), iRub = head.indexOf('rubro');
  if (iRub < 0) return;
  const rubro = {}, vacios = [];
  data.slice(1).forEach((f, r) => {
    const cod = String(f[iCod]).trim(); if (!cod) return;
    if (String(f[iRub]).trim()) { rubro[cod] = String(f[iRub]).trim(); return; }
    const nn = String(f[iNom]).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, ''), rr = String(f[iRaz] || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
    const rb = (RUBROS_NOMBRE.find(p => p[1].test(nn)) || RUBROS_NOMBRE.find(p => p[1].test(rr)) || [''])[0];
    if (rb) { rubro[cod] = rb; vacios.push([r + 1, rb]); } else vacios.push([r + 1, '', cod]);
  });
  // Por compras: perfil de cada cliente (participación de cada artículo en lo que compró)
  const pendientes = vacios.filter(v => !v[1]);
  if (pendientes.length) {
    const perfil = {};
    leerRenglones(ss).forEach(r => { const p = perfil[r.cli] = perfil[r.cli] || {}; p[r.art] = (p[r.art] || 0) + r.sub; });
    const normal = p => { const t = Object.keys(p).reduce((s, k) => s + p[k], 0) || 1, o = {}; Object.keys(p).forEach(k => o[k] = p[k] / t); return o; };
    const centro = {}, cuantos = {};
    Object.keys(perfil).forEach(c => { const rb = rubro[c]; if (!rb) return; const p = normal(perfil[c]); const z = centro[rb] = centro[rb] || {}; cuantos[rb] = (cuantos[rb] || 0) + 1; Object.keys(p).forEach(k => z[k] = (z[k] || 0) + p[k]); });
    const coseno = (a, b) => { let s = 0, na = 0, nb = 0; Object.keys(a).forEach(k => { na += a[k] * a[k]; if (b[k]) s += a[k] * b[k]; }); Object.keys(b).forEach(k => nb += b[k] * b[k]); return s / (Math.sqrt(na * nb) || 1); };
    pendientes.forEach(v => {
      const p = perfil[v[2]]; if (!p) return;
      const pn = normal(p); let mejor = '', ms = 0.1;
      Object.keys(centro).forEach(rb => { if (cuantos[rb] < 3) return; const s = coseno(pn, centro[rb]); if (s > ms) { ms = s; mejor = rb; } });
      if (mejor) { v[1] = mejor; rubro[v[2]] = mejor; }
    });
  }
  vacios.filter(v => v[1]).forEach(v => cl.getRange(v[0] + 1, iRub + 1).setValue(v[1]));
}

// Resumen por cliente para la app: top 5 por plata (con la última compra) y 3 sugerencias de su rubro
function resumirProductos(ss) {
  const ren = leerRenglones(ss);
  const cl = ss.getSheetByName(H_CLIENTES), data = cl.getDataRange().getDisplayValues(), head = data[0].map(norm);
  const iCod = head.indexOf('codigo'), iRub = head.indexOf('rubro');
  const rubro = {}; data.slice(1).forEach(f => { const c = String(f[iCod]).trim(); if (c) rubro[c] = iRub >= 0 ? String(f[iRub]).trim() : ''; });
  const porCli = {}, nombreArt = {};
  ren.forEach(r => {
    const a = (porCli[r.cli] = porCli[r.cli] || {})[r.art] = (porCli[r.cli][r.art] || { plata: 0, cant: 0, ult: null });
    a.plata += r.sub; a.cant += r.cant; if (!a.ult || r.fecha > a.ult) { a.ult = r.fecha; nombreArt[r.art] = r.desc; }
  });
  // Qué artículos compra cada rubro (cantidad de clientes distintos)
  const rub = {}, todos = { n: 0, art: {} };
  Object.keys(porCli).forEach(c => {
    const g = rubro[c] || '';
    const grupo = rub[g] = rub[g] || { n: 0, art: {} };
    grupo.n++; todos.n++;
    Object.keys(porCli[c]).forEach(a => { grupo.art[a] = (grupo.art[a] || 0) + 1; todos.art[a] = (todos.art[a] || 0) + 1; });
  });
  const filas = [];
  Object.keys(porCli).forEach(c => {
    const arts = porCli[c];
    const top = Object.keys(arts).sort((a, b) => arts[b].plata - arts[a].plata).slice(0, 5)
      .map(a => [nombreArt[a], Math.round(arts[a].cant * 100) / 100, Utilities.formatDate(arts[a].ult, Session.getScriptTimeZone(), 'yyyy-MM-dd')]);
    let g = rub[rubro[c] || ''];
    if (!rubro[c] || !g || g.n < 5) g = todos;
    const minimo = Math.max(2, Math.ceil(g.n * 0.25));
    const sug = Object.keys(g.art).filter(a => !arts[a] && g.art[a] >= minimo)
      .sort((a, b) => g.art[b] - g.art[a]).slice(0, 3).map(a => nombreArt[a]);
    filas.push([c, JSON.stringify({ t: top, s: sug })]);
  });
  let sh = ss.getSheetByName(H_PROD_RES);
  if (!sh) { sh = ss.insertSheet(H_PROD_RES, ss.getSheets().length); sh.getRange('A:B').setNumberFormat('@'); sh.hideSheet(); }
  sh.clearContents();
  sh.getRange(1, 1, 1, 2).setValues([['Código', 'Resumen']]);
  if (filas.length) sh.getRange(2, 1, filas.length, 2).setValues(filas);
}

function leerResumenProductos(ss) {
  const sh = ss.getSheetByName(H_PROD_RES), o = {};
  if (!sh || sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(f => { try { o[f[0]] = JSON.parse(f[1]); } catch (e) {} });
  return o;
}
// Descompresor zlib/deflate (RFC 1950/1951) en JavaScript puro, para leer los PDF de las facturas
function inflarZlib(bytes) {
  const src = bytes, out = [];
  let pos = 2, bit = 0, buf = 0; // se saltean los 2 bytes de cabecera zlib
  const lim = src.length + 4;
  function bits(n) { while (bit < n) { if (pos >= lim) throw new Error('Datos incompletos'); buf |= (src[pos++] & 255) << bit; bit += 8; } const v = buf & ((1 << n) - 1); buf >>>= n; bit -= n; return v; }
  function arbol(lens, n) {
    const cnt = new Array(16).fill(0), offs = new Array(16).fill(0), sym = new Array(n);
    for (let i = 0; i < n; i++) cnt[lens[i]]++;
    cnt[0] = 0; for (let i = 1; i < 16; i++) offs[i] = offs[i - 1] + cnt[i - 1];
    for (let i = 0; i < n; i++) if (lens[i]) sym[offs[lens[i]]++] = i;
    return { cnt: cnt, sym: sym };
  }
  function dec(t) { let code = 0, first = 0, idx = 0; for (let len = 1; len < 16; len++) { code |= bits(1); const c = t.cnt[len]; if (code - c < first) return t.sym[idx + (code - first)]; idx += c; first += c; first <<= 1; code <<= 1; } throw new Error('Huffman'); }
  const LB = [3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258], LE = [0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0];
  const DB = [1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577], DE = [0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13];
  let fijoL = null, fijoD = null;
  let fin = 0;
  while (!fin) {
    fin = bits(1); const tipo = bits(2);
    if (tipo === 0) {
      buf = 0; bit = 0; const len = (src[pos] & 255) | ((src[pos + 1] & 255) << 8); pos += 4;
      for (let i = 0; i < len; i++) out.push(src[pos++] & 255);
      continue;
    }
    let lt, dt;
    if (tipo === 1) {
      if (!fijoL) { const l = new Array(288); for (let i = 0; i < 288; i++) l[i] = i < 144 ? 8 : i < 256 ? 9 : i < 280 ? 7 : 8; fijoL = arbol(l, 288); fijoD = arbol(new Array(30).fill(5), 30); }
      lt = fijoL; dt = fijoD;
    } else if (tipo === 2) {
      const hl = bits(5) + 257, hd = bits(5) + 1, hc = bits(4) + 4, ord = [16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];
      const cl = new Array(19).fill(0); for (let i = 0; i < hc; i++) cl[ord[i]] = bits(3);
      const ct = arbol(cl, 19), lens = [];
      while (lens.length < hl + hd) {
        const s = dec(ct);
        if (s < 16) lens.push(s);
        else if (s === 16) { const p = lens[lens.length - 1], r = 3 + bits(2); for (let i = 0; i < r; i++) lens.push(p); }
        else if (s === 17) { const r = 3 + bits(3); for (let i = 0; i < r; i++) lens.push(0); }
        else { const r = 11 + bits(7); for (let i = 0; i < r; i++) lens.push(0); }
      }
      lt = arbol(lens.slice(0, hl), hl); dt = arbol(lens.slice(hl), hd);
    } else throw new Error('Bloque inválido');
    for (;;) {
      const s = dec(lt);
      if (s < 256) out.push(s);
      else if (s === 256) break;
      else if (out.length > 4000000) throw new Error('Demasiado grande');
      else { const l = LB[s - 257] + bits(LE[s - 257]), ds = dec(dt), d = DB[ds] + bits(DE[ds]); const st = out.length - d; for (let i = 0; i < l; i++) out.push(out[st + i]); }
    }
  }
  return out;
}
