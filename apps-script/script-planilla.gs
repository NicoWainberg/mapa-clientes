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
 */
const H_CLIENTES = 'Clientes', H_VISITAS = 'Visitas', H_MAPA = 'Mapa';
const COLS_CLIENTES = ['Código', 'Nombre', 'Dirección', 'Teléfono', 'Contacto', 'Rubro', 'Comentarios', 'Horario', 'Zona', 'Estado', 'Última visita', 'Último resumen'];
const CAMPO_COL = { codigo: 'codigo', nombre: 'nombre', direccion: 'direccion', telefono: 'telefono', contacto: 'contacto', rubro: 'rubro', notas: 'comentarios', horario: 'horario', zona: 'zona', estado: 'estado' };
const COLS_VISITAS = ['Fecha', 'Hora', 'Código', 'Nombre', 'Dirección', 'Zona', 'Resumen'];

/* ================= Lectura ================= */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    estructura(ss);
    if (p.accion === 'todo') return out({ ok: true, clientes: leerClientes(ss), visitas: leerVisitas(ss), reinicio: Number(PropertiesService.getScriptProperties().getProperty('reinicio') || 0) });
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
  const conocidas = new Set(Object.values(CAMPO_COL).concat(['ultima visita', 'ultimo resumen']));
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
    else if (d.tipo === 'reinicio') { PropertiesService.getScriptProperties().setProperty('reinicio', String(Number(d.ts) || Date.now())); r = { ok: true }; }
    else if (d.tipo === 'lote') r = { ok: true, res: aplicarLote(ss, d.items || []) };
    else if (d.tipo === 'edicion') r = aplicarLote(ss, [{ orig: d.orig, campos: d.campos }])[0];
    else if (d.tipo === 'registro_lote') { (d.items || []).forEach(it => registrar(ss, it)); r = { ok: true }; }
    else if (d.tipo === 'visita' || d.tipo === 'desvisita') r = registrar(ss, d);
    else if (d.texto) r = registrar(ss, { tipo: 'visita', fecha: d.fecha, hora: '', codigo: d.codigo, nombre: d.nombre, direccion: d.direccion, resumen: d.texto }); // versiones viejas de la app
    else r = { ok: false, error: 'Pedido desconocido' };
    if (d.tipo !== 'lote' && d.tipo !== 'edicion') actualizarResumenes(ss);
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

function prepararCampos(c0) {
  const c = Object.assign({}, c0);
  if (c.localidad) c.direccion = [c.direccion, c.localidad].filter(Boolean).join(', ');
  delete c.localidad;
  Object.keys(c).forEach(k => { c[k] = c[k] == null ? '' : String(c[k]); });
  return c;
}

function alta(ss, d) {
  const cl = ss.getSheetByName(H_CLIENTES);
  const data = cl.getDataRange().getDisplayValues();
  const idx = idxClientes(cl);
  const existente = findRow(data, idx, d);
  if (existente >= 0) return { ok: true, existente: true, codigo: data[existente][idx.codigo] };
  const c = prepararCampos(d);
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
    const c = prepararCampos(it.campos || {});
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
  if (!(ss.getSheetByName(H_CLIENTES) && ss.getSheetByName(H_MAPA) && ss.getSheetByName(H_VISITAS))) migrar(ss);
  asegurarHorario(ss);
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
    return [todos[r], val(f, i.nombre), dir, val(f, i.telefono), val(f, i.contacto), val(f, i.rubro), val(f, i.notas), '',
            est === 'Oculto' ? '' : val(f, i.zona), est, '', ''].concat(extras.map(k => val(f, k)));
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
