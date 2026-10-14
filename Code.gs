/**
 * ============================================================
 *  CALIBRACIÓN DE BOQUILLAS – DRONES  |  Ingenio Madre Tierra
 *  Backend Google Apps Script para "BASE DE DATOS CALIBRACION DRONES"
 * ============================================================
 *  Pestañas que usa:
 *   USUARIOS            -> login, cargo y botones extras permitidos
 *   SUPERVISOR          -> lista de supervisores
 *   DESCRIPCION DRONES  -> tipo de dron, velocidad (m/s) y ancho de faja (m)
 *   CALCULO DE FLUJOS   -> parámetros del medidor de flujos (% variación máxima)
 *   LABORES             -> labor y L/ha teórico
 *   CALIBRACIONES       -> aquí se guarda cada calibración (una fila por registro)
 *
 *  Publicar: Implementar > Nueva implementación > Aplicación web
 *            Ejecutar como: Yo | Quién tiene acceso: Cualquier persona
 * ============================================================
 */

const SHEET_ID  = '1hz59q81WbmV0hS70VMETl5Zju_dvOAflBSzJcNlAABs';
const FOLDER_ID = '1Bg1AyYQQSx7xb0qOIk7fSkDioc0cRNMW';   // carpeta de fotos
const TZ        = 'America/Guatemala';
const TOKEN_DIAS = 30;                                     // días que dura la sesión
const BOQUILLAS  = 2;

/* ---------- utilidades ---------- */

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function norm_(s) {
  return String(s == null ? '' : s)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

function num_(v) {
  if (typeof v === 'number') return v;
  if (v == null || v === '') return NaN;
  let s = String(v).trim();
  const pct = s.indexOf('%') >= 0;
  s = s.replace(/[%\s]/g, '');
  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  else s = s.replace(',', '.');
  const n = parseFloat(s);
  return pct ? n / 100 : n;
}

function r_(n, d) { const f = Math.pow(10, d); return Math.sign(n) * Math.round(Math.abs(n) * f * (1 + 1e-12)) / f; }  // redondeo como Excel

function getSheet_(...claves) {
  const hojas = ss_().getSheets();
  for (const h of hojas) {
    const n = norm_(h.getName());
    if (claves.every(c => n.indexOf(c) >= 0)) return h;
  }
  return null;
}

function fechaTxt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'dd/MM/yyyy');
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) return ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + m[3];
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[3] + '/' + iso[2] + '/' + iso[1];
  return s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- entrada web ---------- */

function doGet(e) {
  const a = (e && e.parameter && e.parameter.action) || 'ping';
  if (a === 'diag') return json_(diagnostico_());
  return json_({ ok: true, app: 'Calibración de boquillas', hora: Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') });
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents || '{}');
    switch (p.action) {
      case 'ping':       return json_({ ok: true });
      case 'usuarios':   return json_({ ok: true, usuarios: listaUsuarios_() });
      case 'login':      return json_(login_(p.usuario, p.password));
      case 'catalogos':  return json_(conSesion_(p.token, u => ({ ok: true, usuario: u, catalogos: catalogos_() })));
      case 'guardar':    return json_(conSesion_(p.token, u => guardar_(u, p.datos)));
      case 'guardarVuelo': return json_(conSesion_(p.token, u => guardarVuelo_(u, p.datos)));
      case 'vuelos':     return json_(conSesion_(p.token, u => ({ ok: true, vuelos: vuelos_(60) })));
      case 'reporte':    return json_(conSesion_(p.token, u => generarReporte_(u, p.desde, p.hasta, p.enviar !== false)));
      case 'editarCalibracion': return json_(conSesion_(p.token, u => editarCalibracion_(u, p.datos)));
      case 'editarVuelo': return json_(conSesion_(p.token, u => editarVuelo_(u, p.datos)));
      case 'hoy':        return json_(conSesion_(p.token, u => ({ ok: true, registros: registros_(true) })));
      case 'dashboard':  return json_(conSesion_(p.token, u => {
                            if (!u.permisos.dashboard) return { ok: false, error: 'No tiene permiso para el dashboard.' };
                            let vs = []; try { vs = vuelos_(100000); } catch (e) {}
                            return { ok: true, registros: registros_(false), vuelos: vs };
                          }));
      case 'logout':     PropertiesService.getScriptProperties().deleteProperty('tok_' + p.token); return json_({ ok: true });
      default:           return json_({ ok: false, error: 'Acción no reconocida.' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------- usuarios y sesión ---------- */

function leerUsuarios_() {
  const sh = getSheet_('USUARIO');
  if (!sh) throw new Error('No existe la pestaña USUARIOS.');
  const v = sh.getDataRange().getValues();
  const h = v[0].map(norm_);
  const col = k => h.findIndex(x => x.indexOf(k) >= 0);
  const cU = col('USUARIO'), cP = col('CONTRASE'), cC = col('CORREO'), cB = col('BOTONES'), cG = col('CARGO');
  return v.slice(1).filter(r => String(r[cU]).trim()).map(r => {
    const cargo = cG >= 0 ? String(r[cG]).trim() : '';
    const botones = cB >= 0 ? norm_(r[cB]) : '';
    const admin = norm_(cargo).indexOf('ADMIN') >= 0;
    return {
      nombre: String(r[cU]).trim(),
      pass: cP >= 0 ? String(r[cP]).trim() : '',
      correo: cC >= 0 ? String(r[cC]).trim() : '',
      cargo: cargo,
      esAdmin: admin,
      permisos: {
        flujos: admin || botones.indexOf('MEDIDOR') >= 0 || botones.indexOf('FLUJO') >= 0,
        dashboard: admin || botones.indexOf('DASHBOARD') >= 0
      }
    };
  });
}

function publico_(u) {
  return { nombre: u.nombre, correo: u.correo, cargo: u.cargo, esAdmin: u.esAdmin, permisos: u.permisos };
}

function listaUsuarios_() { return leerUsuarios_().map(u => u.nombre); }

function login_(usuario, password) {
  const u = leerUsuarios_().find(x => norm_(x.nombre) === norm_(usuario));
  if (!u) return { ok: false, error: 'Usuario no encontrado.' };
  if (!u.pass) return { ok: false, error: 'Este usuario aún no tiene contraseña asignada. Pídala al administrador.' };
  if (String(password).trim() !== u.pass) return { ok: false, error: 'Contraseña incorrecta.' };
  const token = Utilities.getUuid();
  const props = PropertiesService.getScriptProperties();
  limpiarTokens_(props);
  props.setProperty('tok_' + token, JSON.stringify({ u: u.nombre, exp: Date.now() + TOKEN_DIAS * 864e5 }));
  return { ok: true, token: token, usuario: publico_(u), catalogos: catalogos_() };
}

function conSesion_(token, fn) {
  if (!token) return { ok: false, sesion: false, error: 'Sesión no válida.' };
  const raw = PropertiesService.getScriptProperties().getProperty('tok_' + token);
  if (!raw) return { ok: false, sesion: false, error: 'Su sesión expiró. Ingrese de nuevo.' };
  const t = JSON.parse(raw);
  if (t.exp < Date.now()) return { ok: false, sesion: false, error: 'Su sesión expiró. Ingrese de nuevo.' };
  const u = leerUsuarios_().find(x => norm_(x.nombre) === norm_(t.u));
  if (!u) return { ok: false, sesion: false, error: 'Usuario dado de baja.' };
  return fn(publico_(u));
}

function limpiarTokens_(props) {
  const all = props.getProperties(), now = Date.now();
  Object.keys(all).forEach(k => {
    if (k.indexOf('tok_') !== 0) return;
    try { if (JSON.parse(all[k]).exp < now) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); }
  });
}

/* ---------- catálogos ---------- */

function catalogos_() {
  return {
    supervisores: supervisores_(),
    drones: drones_(),
    labores: labores_(),
    flujos: paramsFlujos_()
  };
}

function supervisores_() {
  const sh = getSheet_('SUPERVISOR');
  if (!sh) return [];
  const v = sh.getDataRange().getValues();
  const out = [];
  v.forEach(r => {
    const nombre = String(r[1] || '').trim();
    if (!nombre || norm_(nombre) === 'NOMBRE') return;
    if (out.indexOf(nombre) < 0) out.push(nombre);
  });
  return out;
}

function labores_() {
  const sh = getSheet_('LABOR');
  if (!sh) return [];
  return sh.getDataRange().getValues().slice(1)
    .filter(r => String(r[0]).trim() && !isNaN(num_(r[1])))
    .map(r => ({ labor: String(r[0]).trim(), lha: num_(r[1]) }));
}

/**
 * Lee DESCRIPCION DRONES buscando automáticamente la fila de encabezados y las
 * columnas de: nombre/tipo de dron, velocidad (m/s o km/h) y ancho de faja.
 */
function drones_() {
  const sh = getSheet_('DRON', 'DESCRIP') || getSheet_('DRONES');
  if (!sh) return [];
  const v = sh.getDataRange().getValues();
  let hr = -1;
  for (let i = 0; i < Math.min(v.length, 15); i++) {
    const fila = v[i].map(norm_).join(' | ');
    if (fila.indexOf('VELOC') >= 0 && (fila.indexOf('ANCHO') >= 0 || fila.indexOf('FAJA') >= 0)) { hr = i; break; }
  }
  if (hr < 0) return [];
  const h = v[hr].map(norm_);
  const busca = (...pruebas) => { for (const p of pruebas) { const i = h.findIndex(p); if (i >= 0) return i; } return -1; };
  const cNom = busca(x => x === 'TIPO DRON' || x === 'TIPO DE DRON', x => x.indexOf('DRON') >= 0 && x.indexOf('VELOC') < 0,
                     x => /EQUIPO|CODIGO|NOMBRE|MODELO|IDENTIF/.test(x), x => x.indexOf('TIPO') >= 0);
  const cMs  = busca(x => x.indexOf('VELOC') >= 0 && x.indexOf('M/S') >= 0, x => x.indexOf('VELOC') >= 0 && x.indexOf('KM') < 0);
  const cKm  = busca(x => x.indexOf('VELOC') >= 0 && x.indexOf('KM') >= 0);
  const cAn  = busca(x => x.indexOf('ANCHO') >= 0, x => x.indexOf('FAJA') >= 0);
  const cBq  = busca(x => x.indexOf('BOQUILLA') >= 0);
  const col0 = cNom >= 0 ? cNom : 0;
  const out = [];
  for (const r of v.slice(hr + 1)) {
    const nombre = String(r[col0] || '').trim();
    if (!nombre) { if (out.length) break; else continue; }       // fin de la tabla de drones
    if (!/[A-Z]/i.test(nombre) || !/\d/.test(nombre) || /^\d+[.)]/.test(nombre)) continue;  // solo nombres tipo T40-1
    let vel = cMs >= 0 ? num_(r[cMs]) : NaN;
    if (isNaN(vel) && cKm >= 0) vel = num_(r[cKm]) / 3.6;
    const ancho = cAn >= 0 ? num_(r[cAn]) : NaN;
    if (isNaN(vel) || isNaN(ancho)) continue;
    out.push({ dron: nombre, vel: r_(vel, 3), ancho: ancho, boquillas: cBq >= 0 && !isNaN(num_(r[cBq])) ? num_(r[cBq]) : BOQUILLAS });
  }
  return out;
}

function paramsFlujos_() {
  const def = { base: 14, min: 11.5, max: 16.5, vel: 13.8, ancho: 8.5, variacion: 0.05 };
  const sh = getSheet_('FLUJO');
  if (!sh) return def;
  const v = sh.getRange(1, 1, Math.min(sh.getLastRow(), 20), Math.min(sh.getLastColumn(), 12)).getValues();
  for (let i = 0; i < v.length - 1; i++) {
    for (let j = 0; j < v[i].length; j++) {
      if (norm_(v[i][j]).indexOf('L/HA BASE') >= 0) {
        const f = v[i + 1];
        const vals = [num_(f[j]), num_(f[j + 1]), num_(f[j + 2]), num_(f[j + 3]), num_(f[j + 4]), num_(f[j + 5])];
        if (vals.every(x => !isNaN(x))) {
          let varx = vals[5]; if (varx > 1) varx = varx / 100;
          return { base: vals[0], min: vals[1], max: vals[2], vel: vals[3], ancho: vals[4], variacion: varx };
        }
      }
    }
  }
  return def;
}

/* ---------- cálculo (igual que la hoja CALIBRACIONES) ---------- */

function calcular_(d, tol) {
  const t   = d.tiempo > 0 ? d.tiempo : 1;
  const b1  = d.b1 / t, b2 = d.b2 / t;                       // L/min por boquilla
  const caudal = d.lha * d.vel * d.ancho * 60 / 10000;       // L/min total
  const qb  = caudal / BOQUILLAS;                            // L/min por boquilla
  const min = qb * (1 - tol), max = qb * (1 + tol);
  const desc = b1 + b2;
  const ok1 = b1 >= min && b1 <= max, ok2 = b2 >= min && b2 <= max;
  return {
    b1: r_(b1, 3), b2: r_(b2, 3),
    diferencia: r_(b1 - b2, 3),
    significancia: (b1 - b2) / d.lha,                        // igual a la hoja: diferencia ÷ L/ha teórico
    varB1: (b1 - qb) / qb,
    varB2: (b2 - qb) / qb,
    descargaGeneral: r_(desc, 3),
    varGeneral: (desc - caudal) / caudal,
    lhaTeorico: d.lha,
    velMs: d.vel, velKmh: r_(d.vel * 3.6, 2), ancho: d.ancho,
    caudal: r_(caudal, 3), caudalBoq: r_(qb, 3),
    min: r_(min, 3), max: r_(max, 3),
    estado: ok1 && ok2 ? 'DENTRO DE RANGO' : 'FUERA DE RANGO'
  };
}

/* ---------- hoja CALIBRACIONES ---------- */

const CAMPOS = [
  ['fecha', 'FECHA', 'dd/MM/yyyy'],
  ['supervisor', 'SUPERVISOR', '@'],
  ['dron', 'TIPO DRON', '@'],
  ['tiempo', 'TIEMPO DE CALIBRACION', '0.##'],
  ['descarga', 'NUMERO DE DESCARGA', '@'],
  ['b1', 'BOQUILLA 1', '0.000'],
  ['b2', 'BOQUILLA 2', '0.000'],
  ['diferencia', 'DIFERENCIA ENTRE', '0.000'],
  ['significancia', 'SIGNIFICANCIA', '0.00%'],
  ['varB1', 'VARIACION B1', '0.00%'],
  ['varB2', 'VARIACION B2', '0.00%'],
  ['descargaGeneral', 'DESCARGA GENERAL', '0.000'],
  ['varGeneral', 'VARIACION GENERAL', '0.00%'],
  ['lhaTeorico', 'LTS/HA TEORICO', '0.00'],
  ['velMs', 'VELOCIDAD M/S', '0.00'],
  ['velKmh', 'VELOCIDAD KM', '0.00'],
  ['ancho', 'ANCHO DE FAJA', '0.00'],
  ['caudal', 'CAUDAL LTS/MIN', '0.000'],
  ['caudalBoq', 'CAUDAL POR BOQUILLA', '0.000'],
  ['min', 'MIN', '0.000'],
  ['max', 'MAX', '0.000'],
  ['foto', 'FOTO', '@'],
  // columnas nuevas (se crean solas al final si no existen)
  ['labor', 'LABOR', '@'],
  ['estado', 'ESTADO', '@'],
  ['registradoPor', 'REGISTRADO POR', '@'],
  ['marcaTiempo', 'MARCA DE TIEMPO', 'dd/MM/yyyy HH:mm'],
  ['observaciones', 'OBSERVACIONES', '@'],
  ['parametros', 'PARAMETROS', '@'],
  ['editadoPor', 'EDITADO POR', '@'],
  ['idRegistro', 'ID REGISTRO', '@']
];

function coincide_(h, clave) {
  if (h === clave) return true;
  if (h.indexOf(clave) !== 0) return false;
  const sig = h.charAt(clave.length);
  return !/[A-Z0-9]/.test(sig);
}

function hojaCalibraciones_() {
  const sh = getSheet_('CALIBRACION');
  if (!sh) throw new Error('No existe la pestaña CALIBRACIONES.');
  return sh;
}

function mapaColumnas_(sh, crear, campos) {
  campos = campos || CAMPOS;
  let ultima = Math.max(sh.getLastColumn(), 1);
  const h = sh.getRange(1, 1, 1, ultima).getValues()[0].map(norm_);
  const mapa = {};
  campos.forEach(([k, clave, , alias]) => {
    const i = h.findIndex(x => coincide_(x, clave) || (alias && coincide_(x, alias)));
    if (i >= 0) mapa[k] = i + 1;
  });
  if (crear) {
    campos.forEach(([k, clave]) => {
      if (mapa[k]) return;
      ultima += 1;
      sh.getRange(1, ultima).setValue(clave).setFontWeight('bold');
      mapa[k] = ultima;
    });
  }
  return mapa;
}

/** Velocidad, ancho y L/ha: del catálogo, o los que escriba el administrador. */
function parametros_(u, d, dron, labor) {
  const p = { vel: dron.vel, ancho: dron.ancho, lha: labor.lha, manual: false };
  if (u.esAdmin && d) {
    [['vel', 'vel'], ['ancho', 'ancho'], ['lha', 'lha']].forEach(([k, src]) => {
      const v = num_(d[src]);
      if (v > 0 && Math.abs(v - p[k]) > 1e-9) { p[k] = v; p.manual = true; }
    });
  }
  return p;
}

function textoObs_(t) { return String(t == null ? '' : t).replace(/\s+/g, ' ').trim().slice(0, 500); }

/** Escribe en una fila solo las claves que vienen en "valores" (respeta las demás columnas). */
function escribirFila_(sh, fila, campos, mapa, valores) {
  const ancho = sh.getLastColumn();
  const rango = sh.getRange(fila, 1, 1, ancho);
  const actual = rango.getValues()[0];
  const formatos = rango.getNumberFormats()[0];
  campos.forEach(([k, , fmt]) => {
    const col = mapa[k];
    if (!col || !(k in valores)) return;
    actual[col - 1] = valores[k] == null ? '' : valores[k];
    formatos[col - 1] = fmt;
  });
  rango.setNumberFormats([formatos]);
  rango.setValues([actual]);
  SpreadsheetApp.flush();
}

/** Busca un registro por su ID (para no duplicar cuando el teléfono reintenta). Devuelve la fila o 0. */
function filaPorId_(sh, mapa, id) {
  if (!id || !mapa.idRegistro) return 0;
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, mapa.idRegistro, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

/** Quién registró: el usuario que lo guardó en el teléfono (si existe), si no el de la sesión. */
function quienRegistra_(u, d) {
  if (d && d.registradoPor) {
    const x = leerUsuarios_().find(z => norm_(z.nombre) === norm_(d.registradoPor));
    if (x) return x.nombre;
  }
  return u.nombre;
}

function editado_(u) { return u.nombre + ' · ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'); }

function editarCalibracion_(u, d) {
  if (!u.esAdmin) return { ok: false, error: 'Solo el administrador puede editar registros.' };
  if (!d || !(d.fila >= 2)) return { ok: false, error: 'Registro no válido.' };
  const cat = catalogos_();
  const dron = cat.drones.find(x => norm_(x.dron) === norm_(d.dron));
  if (!dron) return { ok: false, error: 'Dron no encontrado: ' + d.dron };
  const labor = cat.labores.find(x => norm_(x.labor) === norm_(d.labor));
  if (!labor) return { ok: false, error: 'Labor no encontrada: ' + d.labor };
  const b1 = num_(d.b1), b2 = num_(d.b2);
  if (!(b1 > 0) || !(b2 > 0)) return { ok: false, error: 'Ingrese la descarga de ambas boquillas.' };
  const par = parametros_(u, d, dron, labor);
  const c = calcular_({ b1, b2, tiempo: 1, lha: par.lha, vel: par.vel, ancho: par.ancho }, cat.flujos.variacion);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = hojaCalibraciones_();
    const mapa = mapaColumnas_(sh, true);
    if (d.fila > sh.getLastRow()) return { ok: false, error: 'El registro ya no existe.' };
    const fila = sh.getRange(d.fila, 1, 1, sh.getLastColumn()).getValues()[0];
    if (fechaTxt_(fila[mapa.fecha - 1]) !== d.fechaOrig || norm_(fila[mapa.dron - 1]) !== norm_(d.dronOrig))
      return { ok: false, error: 'La hoja cambió (alguien movió o borró filas). Recargue e intente de nuevo.' };
    const tiempo = num_(d.tiempo);
    escribirFila_(sh, d.fila, CAMPOS, mapa, Object.assign({}, c, {
      supervisor: d.supervisor || fila[mapa.supervisor - 1],
      dron: dron.dron, labor: labor.labor,
      tiempo: tiempo > 0 ? tiempo : fila[mapa.tiempo - 1],
      observaciones: textoObs_(d.observaciones),
      parametros: par.manual ? 'MANUAL (admin)' : 'AUTOMATICO',
      editadoPor: editado_(u)
    }));
    return { ok: true, calculo: c };
  } finally { lock.releaseLock(); }
}

function editarVuelo_(u, d) {
  if (!u.esAdmin) return { ok: false, error: 'Solo el administrador puede editar registros.' };
  if (!d || !(d.fila >= 2)) return { ok: false, error: 'Registro no válido.' };
  const cat = catalogos_();
  const dron = cat.drones.find(x => norm_(x.dron) === norm_(d.dron));
  if (!dron) return { ok: false, error: 'Dron no encontrado: ' + d.dron };
  const labor = cat.labores.find(x => norm_(x.labor) === norm_(d.labor));
  if (!labor) return { ok: false, error: 'Labor no encontrada: ' + d.labor };
  const n = k => num_(d[k]);
  const v = { areaPoligono: n('areaPoligono'), distanciaMax: n('distanciaMax'), carga: n('carga'),
              areaReal: n('areaReal'), aplicados: n('aplicados'), devueltos: n('devueltos') };
  for (const k of ['areaReal', 'aplicados', 'carga']) if (!(v[k] > 0)) return { ok: false, error: 'Falta un dato del vuelo (' + k + ').' };
  if (isNaN(v.devueltos) || v.devueltos < 0) return { ok: false, error: 'Ingrese los litros devueltos (0 si no hubo).' };
  const lha = n('lha') > 0 ? n('lha') : labor.lha;
  const c = calcularVuelo_(Object.assign({ lhaTeorico: lha }, v));
  ['areaPoligono', 'distanciaMax'].forEach(k => { if (isNaN(v[k])) v[k] = ''; });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = hojaVuelos_();
    const mapa = mapaColumnas_(sh, true, CAMPOS_VUELO);
    if (d.fila > sh.getLastRow()) return { ok: false, error: 'El registro ya no existe.' };
    const fila = sh.getRange(d.fila, 1, 1, sh.getLastColumn()).getValues()[0];
    if (norm_(fila[mapa.dron - 1]) !== norm_(d.dronOrig))
      return { ok: false, error: 'La hoja cambió (alguien movió o borró filas). Recargue e intente de nuevo.' };
    escribirFila_(sh, d.fila, CAMPOS_VUELO, mapa, Object.assign({}, v, c, {
      dron: dron.dron, labor: labor.labor, lhaTeorico: lha,
      supervisor: d.supervisor || fila[mapa.supervisor - 1],
      observaciones: textoObs_(d.observaciones),
      editadoPor: editado_(u)
    }));
    return { ok: true, calculo: c };
  } finally { lock.releaseLock(); }
}

function guardar_(u, d) {
  if (!d) return { ok: false, error: 'Sin datos.' };
  const cat = catalogos_();
  const dron = cat.drones.find(x => norm_(x.dron) === norm_(d.dron));
  if (!dron) return { ok: false, error: 'Dron no encontrado en DESCRIPCION DRONES: ' + d.dron };
  const labor = cat.labores.find(x => norm_(x.labor) === norm_(d.labor));
  if (!labor) return { ok: false, error: 'Labor no encontrada en LABORES: ' + d.labor };
  const b1 = num_(d.b1), b2 = num_(d.b2), tiempo = num_(d.tiempo) || 1;
  if (!(b1 > 0) || !(b2 > 0)) return { ok: false, error: 'Ingrese la descarga de ambas boquillas.' };
  if (!d.fecha || !/^\d{4}-\d{2}-\d{2}$/.test(d.fecha)) return { ok: false, error: 'Fecha inválida.' };

  const par = parametros_(u, d, dron, labor);
  const c = calcular_({ b1, b2, tiempo, lha: par.lha, vel: par.vel, ancho: par.ancho }, cat.flujos.variacion);
  const [Y, M, D] = d.fecha.split('-').map(Number);
  const fecha = new Date(Y, M - 1, D, 12, 0, 0);
  const fechaStr = fechaTxt_(d.fecha);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = hojaCalibraciones_();
    const mapa = mapaColumnas_(sh, true);

    // ¿ya se guardó antes este mismo registro? (reintento desde la cola del teléfono)
    const ya = filaPorId_(sh, mapa, d.id);
    if (ya) {
      const r = sh.getRange(ya, 1, 1, sh.getLastColumn()).getValues()[0];
      return { ok: true, duplicado: true, descarga: String(r[mapa.descarga - 1]), calculo: { estado: String(r[mapa.estado - 1] || c.estado) }, fila: ya };
    }

    // número de descarga del día para ese dron: R1, R2, R3...
    let n = 1;
    const last = sh.getLastRow();
    if (last > 1) {
      const fs = sh.getRange(2, mapa.fecha, last - 1, 1).getValues();
      const ds = sh.getRange(2, mapa.dron, last - 1, 1).getValues();
      for (let i = 0; i < fs.length; i++) {
        if (fechaTxt_(fs[i][0]) === fechaStr && norm_(ds[i][0]) === norm_(dron.dron)) n++;
      }
    }

    // foto (opcional)
    let fotoUrl = '';
    if (d.foto && d.foto.data) {
      const nombre = Utilities.formatDate(fecha, TZ, 'yyyyMMdd') + '_' + dron.dron + '_R' + n + '.jpg';
      const blob = Utilities.newBlob(Utilities.base64Decode(d.foto.data), d.foto.tipo || 'image/jpeg', nombre);
      const file = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
      fotoUrl = file.getUrl();
    }

    const valores = Object.assign({}, c, {
      fecha: fecha,
      supervisor: d.supervisor || u.nombre,
      dron: dron.dron,
      tiempo: tiempo,
      descarga: 'R' + n,
      foto: fotoUrl,
      labor: labor.labor,
      registradoPor: quienRegistra_(u, d),
      marcaTiempo: new Date(),
      idRegistro: d.id || '',
      observaciones: textoObs_(d.observaciones),
      parametros: par.manual ? 'MANUAL (admin)' : 'AUTOMATICO'
    });

    const ancho = sh.getLastColumn();
    const fila = new Array(ancho).fill('');
    const formatos = new Array(ancho).fill('General');
    CAMPOS.forEach(([k, , fmt]) => {
      const col = mapa[k];
      if (!col) return;
      fila[col - 1] = valores[k] == null ? '' : valores[k];
      formatos[col - 1] = fmt === '@' ? '@' : fmt;
    });
    const destino = sh.getRange(last + 1, 1, 1, ancho);
    destino.setNumberFormats([formatos]);
    destino.setValues([fila]);
    SpreadsheetApp.flush();

    return { ok: true, descarga: 'R' + n, foto: fotoUrl, calculo: c, fila: last + 1 };
  } finally {
    lock.releaseLock();
  }
}

function registros_(soloHoy) {
  const sh = hojaCalibraciones_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const mapa = mapaColumnas_(sh, false);
  const v = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const hoy = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy');
  const out = [];
  v.forEach((r, i) => {
    const g = k => mapa[k] ? r[mapa[k] - 1] : '';
    const fecha = fechaTxt_(g('fecha'));
    if (!fecha) return;
    if (soloHoy && fecha !== hoy) return;
    const o = { fecha: fecha, fila: i + 2 };
    ['supervisor', 'dron', 'descarga', 'foto', 'labor', 'estado', 'registradoPor', 'observaciones', 'parametros', 'editadoPor'].forEach(k => o[k] = String(g(k) || ''));
    ['tiempo', 'b1', 'b2', 'diferencia', 'significancia', 'varB1', 'varB2', 'descargaGeneral', 'varGeneral',
     'lhaTeorico', 'velMs', 'velKmh', 'ancho', 'caudal', 'caudalBoq', 'min', 'max'].forEach(k => {
      const n = num_(g(k)); o[k] = isNaN(n) ? null : n;
    });
    if (!o.estado && o.b1 != null && o.min != null) {
      o.estado = (o.b1 >= o.min && o.b1 <= o.max && o.b2 >= o.min && o.b2 <= o.max) ? 'DENTRO DE RANGO' : 'FUERA DE RANGO';
    }
    out.push(o);
  });
  return out;
}

/* ---------- hoja VUELO DEL DRONE ---------- */

const CAMPOS_VUELO = [
  ['dron', 'DRON', '@'],
  ['areaPoligono', 'AREA DEL POLIGONO', '0.00'],
  ['distanciaMax', 'DISTANCIA MAX', '0.##'],
  ['lhaTeorico', 'LTS/HA TEORICO', '0.00'],
  ['carga', 'CARGA INICIAL', '0.00'],
  ['areaReal', 'AREA REAL', '0.00'],
  ['aplicados', 'LTS APLICADOS', '0.00'],
  ['devueltos', 'LTS DEVUELTOS', '0.00', 'LST DEVUELTOS'],
  ['lhaReal', 'LTS/HA REAL', '0.00'],
  ['variacion', '% VARIACION', '0.00%'],
  ['varCaudal', '% VAR CAUDAL', '0.00%'],
  // columnas nuevas (se crean solas al final si no existen)
  ['fecha', 'FECHA', 'dd/MM/yyyy'],
  ['labor', 'LABOR', '@'],
  ['supervisor', 'SUPERVISOR', '@'],
  ['calibracion', 'CALIBRACION', '@'],
  ['consumo', 'LTS CONSUMIDOS (CARGA - DEVUELTOS)', '0.00'],
  ['registradoPor', 'REGISTRADO POR', '@'],
  ['marcaTiempo', 'MARCA DE TIEMPO', 'dd/MM/yyyy HH:mm'],
  ['observaciones', 'OBSERVACIONES', '@'],
  ['editadoPor', 'EDITADO POR', '@'],
  ['idRegistro', 'ID REGISTRO', '@']
];

function hojaVuelos_() {
  const sh = getSheet_('VUELO');
  if (!sh) throw new Error('No existe la pestaña VUELO DEL DRONE.');
  return sh;
}

function calcularVuelo_(d) {
  const lhaReal = d.aplicados / d.areaReal;
  const consumo = d.carga - d.devueltos;
  return {
    lhaReal: r_(lhaReal, 2),
    variacion: (lhaReal - d.lhaTeorico) / d.lhaTeorico,          // L/ha real vs teórico
    consumo: r_(consumo, 2),
    varCaudal: (d.aplicados - consumo) / d.carga                 // lts que reporta el dron vs (carga - devueltos), sobre la carga
  };
}

function guardarVuelo_(u, d) {
  if (!d) return { ok: false, error: 'Sin datos.' };
  const cat = catalogos_();
  const dron = cat.drones.find(x => norm_(x.dron) === norm_(d.dron));
  if (!dron) return { ok: false, error: 'Dron no encontrado: ' + d.dron };
  const labor = cat.labores.find(x => norm_(x.labor) === norm_(d.labor));
  if (!labor) return { ok: false, error: 'Labor no encontrada: ' + d.labor };
  const n = k => num_(d[k]);
  const v = { areaPoligono: n('areaPoligono'), distanciaMax: n('distanciaMax'), carga: n('carga'),
              areaReal: n('areaReal'), aplicados: n('aplicados'), devueltos: n('devueltos') };
  for (const k of ['areaReal', 'aplicados', 'carga']) if (!(v[k] > 0)) return { ok: false, error: 'Falta un dato del vuelo (' + k + ').' };
  if (isNaN(v.devueltos) || v.devueltos < 0) return { ok: false, error: 'Ingrese los litros devueltos (0 si no hubo).' };
  if (!d.fecha || !/^\d{4}-\d{2}-\d{2}$/.test(d.fecha)) return { ok: false, error: 'Fecha inválida.' };
  const lha = u.esAdmin && num_(d.lha) > 0 ? num_(d.lha) : labor.lha;
  const c = calcularVuelo_(Object.assign({ lhaTeorico: lha }, v));
  const [Y, M, D] = d.fecha.split('-').map(Number);

  const valores = Object.assign({}, v, c, {
    dron: dron.dron, lhaTeorico: lha, labor: labor.labor, observaciones: textoObs_(d.observaciones),
    fecha: new Date(Y, M - 1, D, 12, 0, 0),
    supervisor: d.supervisor || u.nombre, calibracion: d.calibracion || '',
    registradoPor: quienRegistra_(u, d), marcaTiempo: new Date(), idRegistro: d.id || ''
  });
  ['areaPoligono', 'distanciaMax'].forEach(k => { if (isNaN(valores[k])) valores[k] = ''; });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = hojaVuelos_();
    const mapa = mapaColumnas_(sh, true, CAMPOS_VUELO);
    if (filaPorId_(sh, mapa, d.id)) return { ok: true, duplicado: true, calculo: c };
    // calibración ligada (guardada desde la cola): obtener su número R
    if (!valores.calibracion && d.calId) {
      const shc = hojaCalibraciones_(), mc = mapaColumnas_(shc, false);
      const fc = filaPorId_(shc, mc, d.calId);
      if (fc) valores.calibracion = String(shc.getRange(fc, mc.descarga).getValue());
    }
    const last = sh.getLastRow(), ancho = sh.getLastColumn();
    const fila = new Array(ancho).fill(''), formatos = new Array(ancho).fill('General');
    CAMPOS_VUELO.forEach(([k, , fmt]) => {
      const col = mapa[k]; if (!col) return;
      fila[col - 1] = valores[k] == null ? '' : valores[k];
      formatos[col - 1] = fmt;
    });
    const destino = sh.getRange(last + 1, 1, 1, ancho);
    destino.setNumberFormats([formatos]);
    destino.setValues([fila]);
    SpreadsheetApp.flush();
    return { ok: true, calculo: c, fila: last + 1 };
  } finally {
    lock.releaseLock();
  }
}

function vuelos_(max) {
  const sh = hojaVuelos_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const mapa = mapaColumnas_(sh, false, CAMPOS_VUELO);
  const v = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const out = [];
  v.forEach((r, i) => {
    const g = k => mapa[k] ? r[mapa[k] - 1] : '';
    if (!String(g('dron')).trim()) return;
    const o = { fecha: fechaTxt_(g('fecha')), fila: i + 2 };
    ['dron', 'labor', 'supervisor', 'calibracion', 'registradoPor', 'observaciones', 'editadoPor'].forEach(k => o[k] = String(g(k) || ''));
    ['areaPoligono', 'distanciaMax', 'lhaTeorico', 'carga', 'areaReal', 'aplicados', 'devueltos', 'lhaReal', 'variacion', 'varCaudal'].forEach(k => {
      const n = num_(g(k)); o[k] = isNaN(n) ? null : n;
    });
    out.push(o);
  });
  return out.slice(-max).reverse();
}

/* ============================================================
 *  REPORTE DE ANÁLISIS POR DRON  (botón "Generar reportes")
 * ============================================================ */

function isoDe_(dmy) { const m = String(dmy).match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? m[3] + '-' + m[2] + '-' + m[1] : ''; }
function prom_(a) { return a.length ? a.reduce((x, y) => x + y, 0) / a.length : NaN; }
function p_(n, d) { return isNaN(n) || n == null ? '—' : (n > 0 ? '+' : '') + (n * 100).toFixed(d == null ? 2 : d) + '%'; }
function pa_(n, d) { return isNaN(n) || n == null ? '—' : (n * 100).toFixed(d == null ? 1 : d) + '%'; }
function n_(n, d) { return isNaN(n) || n == null ? '—' : Number(n).toFixed(d == null ? 3 : d); }
function h_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function dmy_(iso) { return iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : ''; }

function analizar_(desde, hasta) {
  const tol = paramsFlujos_().variacion;
  const cal = registros_(false).map(r => Object.assign(r, { iso: isoDe_(r.fecha) }))
    .filter(r => r.iso && r.iso >= desde && r.iso <= hasta && r.b1 != null && r.b2 != null);
  let vue = [];
  try { vue = vuelos_(100000).map(v => Object.assign(v, { iso: isoDe_(v.fecha) })).filter(v => v.iso && v.iso >= desde && v.iso <= hasta); } catch (e) {}

  const clave = r => r.iso + '|' + ('0000' + String(r.descarga).replace(/\D/g, '')).slice(-4);
  const nombres = Array.from(new Set(cal.map(r => r.dron).concat(vue.map(v => v.dron))))
    .sort((a, b) => a.localeCompare(b, 'es', { numeric: true }));

  const drones = nombres.map(dr => {
    const rs = cal.filter(r => r.dron === dr).sort((a, b) => clave(a) < clave(b) ? -1 : 1);
    const vs = vue.filter(v => v.dron === dr);
    const dentro = rs.filter(r => /DENTRO/i.test(r.estado)).length;
    const v1 = rs.map(r => r.varB1).filter(x => x != null), v2 = rs.map(r => r.varB2).filter(x => x != null);
    const vg = rs.map(r => r.varGeneral).filter(x => x != null);
    const desb = rs.filter(r => r.caudalBoq).map(r => Math.abs(r.b1 - r.b2) / r.caudalBoq);
    const fuera1 = rs.filter(r => r.b1 < r.min || r.b1 > r.max).length;
    const fuera2 = rs.filter(r => r.b2 < r.min || r.b2 > r.max).length;
    const mitad = Math.floor(vg.length / 2);
    const t1 = prom_(vg.slice(0, mitad).map(Math.abs)), t2 = prom_(vg.slice(vg.length - mitad).map(Math.abs));
    let tendencia = 'Sin datos suficientes';
    if (vg.length >= 4) tendencia = t2 < t1 * 0.8 ? 'Mejorando' : t2 > t1 * 1.2 ? 'Empeorando' : 'Estable';
    const ult = rs[rs.length - 1];
    const pctDentro = rs.length ? dentro / rs.length : NaN;
    const absG = prom_(vg.map(Math.abs));
    let clase = 'Sin calibraciones';
    if (rs.length) {
      if (pctDentro >= 0.8 && absG <= tol && /DENTRO/i.test(ult.estado)) clase = 'Bueno';
      else if (pctDentro < 0.5 || (!/DENTRO/i.test(ult.estado) && pctDentro < 0.8)) clase = 'Crítico';
      else clase = 'Atención';
    }
    const vVar = vs.map(v => v.variacion).filter(x => x != null);
    const vQ = vs.map(v => v.varCaudal).filter(x => x != null);
    return {
      dron: dr, n: rs.length, dentro: dentro, pctDentro: pctDentro,
      promB1: prom_(v1), promB2: prom_(v2), promAbsGen: absG,
      maxAbs: Math.max.apply(null, [0].concat(v1.map(Math.abs), v2.map(Math.abs))),
      desbalance: prom_(desb), fuera1: fuera1, fuera2: fuera2, tendencia: tendencia,
      ultimo: ult ? { fecha: ult.fecha, descarga: ult.descarga, estado: ult.estado, varB1: ult.varB1, varB2: ult.varB2 } : null,
      clase: clase,
      vuelos: vs.length, vueloVar: prom_(vVar), vueloAbsVar: prom_(vVar.map(Math.abs)), vueloVarQ: prom_(vQ),
      vuelosFuera: vVar.filter(x => Math.abs(x) > tol).length
    };
  });

  // Hallazgos automáticos
  const tot = cal.length, totDentro = cal.filter(r => /DENTRO/i.test(r.estado)).length;
  const con = drones.filter(d => d.n);
  const hall = [];
  if (tot) hall.push('Se registraron <b>' + tot + '</b> calibraciones en <b>' + con.length + '</b> drones; <b>' + pa_(totDentro / tot) + '</b> quedaron dentro de rango (±' + (tol * 100).toFixed(0) + '% del caudal por boquilla).');
  const conMin = con.filter(d => d.n >= 2);
  if (conMin.length) {
    const mejor = conMin.slice().sort((a, b) => b.pctDentro - a.pctDentro || a.promAbsGen - b.promAbsGen)[0];
    const peor = conMin.slice().sort((a, b) => a.pctDentro - b.pctDentro || b.promAbsGen - a.promAbsGen)[0];
    hall.push('Mejor desempeño: <b>' + h_(mejor.dron) + '</b> (' + pa_(mejor.pctDentro) + ' dentro de rango, variación general promedio ' + pa_(mejor.promAbsGen, 2) + ').');
    if (peor.dron !== mejor.dron) hall.push('Mayor dificultad: <b>' + h_(peor.dron) + '</b> (' + pa_(peor.pctDentro) + ' dentro de rango, variación general promedio ' + pa_(peor.promAbsGen, 2) + ').');
  }
  con.forEach(d => {
    [['Boquilla 1', d.promB1], ['Boquilla 2', d.promB2]].forEach(([b, v]) => {
      if (d.n >= 2 && Math.abs(v) > tol * 0.6) hall.push('<b>' + h_(d.dron) + '</b>: la ' + b + ' descarga en promedio ' + p_(v) + ' ' + (v < 0 ? 'por debajo' : 'por encima') + ' de lo teórico (tendencia sistemática): revisar boquilla, filtro o presión.');
    });
    if (d.n >= 2 && d.desbalance > tol) hall.push('<b>' + h_(d.dron) + '</b>: desbalance promedio entre boquillas de ' + pa_(d.desbalance) + ' del caudal por boquilla: una boquilla trabaja más que la otra.');
    if (d.tendencia === 'Empeorando') hall.push('<b>' + h_(d.dron) + '</b>: la variación va en aumento dentro del período.');
    if (d.ultimo && !/DENTRO/i.test(d.ultimo.estado)) hall.push('<b>' + h_(d.dron) + '</b>: su última calibración (' + d.ultimo.fecha + ', ' + h_(d.ultimo.descarga) + ') quedó <b>fuera de rango</b>.');
    if (d.vuelos && Math.abs(d.vueloVar) > tol) hall.push('<b>' + h_(d.dron) + '</b>: en vuelo aplicó en promedio ' + p_(d.vueloVar) + ' respecto al L/ha teórico.');
  });
  if (!tot) hall.push('No hay calibraciones registradas en el período seleccionado.');

  const rec = [];
  const crit = con.filter(d => d.clase === 'Crítico').map(d => d.dron);
  const aten = con.filter(d => d.clase === 'Atención').map(d => d.dron);
  if (crit.length) rec.push('Recalibrar y revisar boquillas, filtros y bomba de: <b>' + crit.map(h_).join(', ') + '</b> antes de su próximo vuelo.');
  if (aten.length) rec.push('Dar seguimiento en la próxima calibración a: <b>' + aten.map(h_).join(', ') + '</b>.');
  if (!crit.length && !aten.length && con.length) rec.push('Todos los drones calibrados se mantienen en buen estado. Continuar con la rutina de calibración.');

  const obs = cal.filter(r => r.observaciones).map(r => ({ fecha: r.fecha, dron: r.dron, ref: r.descarga, tipo: 'Calibración', texto: r.observaciones }))
    .concat(vue.filter(v => v.observaciones).map(v => ({ fecha: v.fecha, dron: v.dron, ref: v.calibracion || '', tipo: 'Vuelo', texto: v.observaciones })))
    .sort((x, y) => isoDe_(x.fecha) < isoDe_(y.fecha) ? -1 : 1);
  const manuales = cal.filter(r => /MANUAL/i.test(r.parametros)).length;
  if (manuales) hall.push('<b>' + manuales + '</b> calibración(es) se hicieron con parámetros ajustados manualmente por el administrador.');
  return { desde: desde, hasta: hasta, tol: tol, total: tot, dentro: totDentro, drones: drones, hallazgos: hall, recomendaciones: rec, vuelos: vue.length, observaciones: obs };
}

const VERDE = '#2e7d32', VERDE_OSC = '#1b5e20', VERDE_CLARO = '#eaf4ea', GRIS = '#5f6b73', ROJO = '#c0392b', AMBAR = '#b7791f';
const COLOR_CLASE = { 'Bueno': '#2e7d32', 'Atención': '#b7791f', 'Crítico': '#c0392b', 'Sin calibraciones': '#6b7279' };
const LOGO_B64 = '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';
function logoBlob_() { return Utilities.newBlob(Utilities.base64Decode(LOGO_B64), 'image/png', 'logo.png'); }

/** HTML del reporte. modo 'correo' usa cid:logo; 'web' incrusta el logo. */
function htmlReporte_(a, u, modo) {
  const logo = modo === 'correo' ? 'cid:logo' : 'data:image/png;base64,' + LOGO_B64;
  const th = 'style="background:' + VERDE + ';color:#fff;padding:6px;text-align:left;font-size:11px;vertical-align:bottom"';
  const td = 'style="padding:6px;border-top:1px solid #dfe2dc;font-size:11px"';
  const h3 = t => '<h3 style="margin:20px 0 8px;font-size:16px;color:' + VERDE + ';border-bottom:2px solid ' + VERDE_CLARO + ';padding-bottom:4px">' + t + '</h3>';
  const barra = v => '<div style="background:#e9ebe6;border-radius:4px;height:8px;width:70px;display:inline-block;vertical-align:middle"><div style="background:' + (v >= 0.8 ? VERDE : v >= 0.5 ? AMBAR : ROJO) + ';height:8px;border-radius:4px;width:' + Math.max(0, Math.min(100, (v || 0) * 100)).toFixed(0) + '%"></div></div>';
  const filas = a.drones.map((d, i) => '<tr style="background:' + (i % 2 ? '#f7faf6' : '#fff') + '">' +
    '<td ' + td + '><b>' + h_(d.dron) + '</b></td>' +
    '<td ' + td + '><span style="color:#fff;background:' + COLOR_CLASE[d.clase] + ';border-radius:10px;padding:2px 8px;font-weight:bold;font-size:11px">' + d.clase + '</span></td>' +
    '<td ' + td + '>' + d.n + '</td>' +
    '<td ' + td + '>' + (d.n ? barra(d.pctDentro) + ' ' + pa_(d.pctDentro, 0) : '—') + '</td>' +
    '<td ' + td + '>' + p_(d.promB1) + '</td><td ' + td + '>' + p_(d.promB2) + '</td>' +
    '<td ' + td + '>' + pa_(d.promAbsGen, 2) + '</td><td ' + td + '>' + pa_(d.maxAbs, 1) + '</td>' +
    '<td ' + td + '>' + pa_(d.desbalance, 1) + '</td>' +
    '<td ' + td + '>' + d.fuera1 + ' / ' + d.fuera2 + '</td>' +
    '<td ' + td + '>' + d.tendencia + '</td>' +
    '<td ' + td + '>' + (d.ultimo ? d.ultimo.fecha + ' ' + h_(d.ultimo.descarga) + ' · <b style="color:' + (/DENTRO/i.test(d.ultimo.estado) ? VERDE : ROJO) + '">' + (/DENTRO/i.test(d.ultimo.estado) ? 'OK' : 'FUERA') + '</b>' : '—') + '</td>' +
    '</tr>').join('');
  const conV = a.drones.filter(d => d.vuelos);
  const cv = x => 'style="padding:6px;border-top:1px solid #dfe2dc;font-size:11px;color:' + (Math.abs(x) > a.tol ? ROJO : VERDE) + ';font-weight:bold"';
  const filasV = conV.map(d => '<tr><td ' + td + '><b>' + h_(d.dron) + '</b></td><td ' + td + '>' + d.vuelos + '</td><td ' + cv(d.vueloVar) + '>' + p_(d.vueloVar) + '</td><td ' + td + '>' + pa_(d.vueloAbsVar, 2) + '</td><td ' + td + '>' + p_(d.vueloVarQ) + '</td><td ' + td + '>' + d.vuelosFuera + '</td></tr>').join('');
  const kpi = (t, v, s, c) => '<td style="padding:5px;width:25%"><div style="border:1px solid #d5e5d5;border-top:4px solid ' + c + ';border-radius:8px;padding:10px 12px;background:#fff"><div style="font-size:11px;color:' + GRIS + ';text-transform:uppercase;letter-spacing:.05em">' + t + '</div><div style="font-size:26px;font-weight:bold;color:' + c + '">' + v + '</div><div style="font-size:11px;color:' + GRIS + '">' + (s || '') + '</div></div></td>';
  const conN = a.drones.filter(d => d.n).length;
  const nBuenos = a.drones.filter(d => d.clase === 'Bueno').length, nCrit = a.drones.filter(d => d.clase === 'Crítico').length;
  const pctOk = a.total ? a.dentro / a.total : 0;
  const obs = (a.observaciones || []).map(o => '<tr><td ' + td + '>' + o.fecha + '</td><td ' + td + '><b>' + h_(o.dron) + '</b></td><td ' + td + '>' + o.tipo + ' ' + h_(o.ref) + '</td><td ' + td + '>' + h_(o.texto) + '</td></tr>').join('');
  return '<div style="font-family:Arial,Helvetica,sans-serif;color:#1b1e21;max-width:900px">' +
    '<table style="width:100%;border-collapse:collapse;border-bottom:4px solid ' + VERDE + '"><tr>' +
    '<td style="width:60px;padding:8px 10px 10px 0;vertical-align:middle"><img src="' + logo + '" alt="Ingenio Madre Tierra" width="44" style="display:block"></td>' +
    '<td style="vertical-align:middle;padding-bottom:6px"><div style="font-size:11px;letter-spacing:.18em;color:' + VERDE + ';font-weight:bold">INGENIO MADRE TIERRA · MALEZAS Y MADURANTES</div>' +
    '<div style="font-size:22px;font-weight:bold;color:' + VERDE_OSC + ';margin-top:2px">Reporte de calibración de boquillas por dron</div>' +
    '<div style="font-size:12px;color:' + GRIS + ';margin-top:3px">Período: <b style="color:#1b1e21">' + dmy_(a.desde) + ' al ' + dmy_(a.hasta) + '</b> · Generado por ' + h_(u.nombre) + ' el ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') + '</div></td></tr></table>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:12px"><tr>' +
    kpi('Calibraciones', a.total, conN + ' drones', VERDE_OSC) +
    kpi('Dentro de rango', a.total ? pa_(pctOk, 0) : '—', a.dentro + ' de ' + a.total, pctOk >= 0.8 ? VERDE : pctOk >= 0.5 ? AMBAR : ROJO) +
    kpi('Drones en buen estado', nBuenos, 'de ' + conN, VERDE) +
    kpi('Drones críticos', nCrit, 'requieren recalibración', nCrit ? ROJO : VERDE) +
    '</tr></table>' +
    h3('Hallazgos principales') + '<ul style="margin:0 0 0 18px;padding:0;font-size:13px;line-height:1.6">' + a.hallazgos.map(x => '<li>' + x + '</li>').join('') + '</ul>' +
    (a.recomendaciones.length ? h3('Recomendaciones') + '<ul style="margin:0 0 0 18px;padding:0;font-size:13px;line-height:1.6">' + a.recomendaciones.map(x => '<li>' + x + '</li>').join('') + '</ul>' : '') +
    h3('Detalle por dron — calibraciones') +
    '<div style="overflow-x:auto"><table style="border-collapse:collapse;width:100%"><tr>' +
    ['Dron', 'Estado', 'Calib.', '% dentro de rango', 'Var. prom. B1', 'Var. prom. B2', '|Var. gral.| prom.', 'Var. máx.', 'Desbalance B1–B2', 'Fuera B1 / B2', 'Tendencia', 'Última calibración'].map(x => '<th ' + th + '>' + x + '</th>').join('') +
    '</tr>' + (filas || '<tr><td ' + td + ' colspan="12">Sin datos</td></tr>') + '</table></div>' +
    (conV.length ? h3('Vuelos del drone en el período') + '<table style="border-collapse:collapse;width:100%"><tr>' +
      ['Dron', 'Vuelos', 'Var. L/ha prom.', '|Var. L/ha| prom.', 'Var. caudal prom.', 'Vuelos fuera de rango'].map(x => '<th ' + th + '>' + x + '</th>').join('') + '</tr>' + filasV + '</table>' : '') +
    (obs ? h3('Observaciones de campo') + '<table style="border-collapse:collapse;width:100%"><tr>' + ['Fecha', 'Dron', 'Registro', 'Observación'].map(x => '<th ' + th + '>' + x + '</th>').join('') + '</tr>' + obs + '</table>' : '') +
    '<p style="font-size:11px;color:' + GRIS + ';margin:16px 0 0;line-height:1.5"><b style="color:' + VERDE + '">Cómo leerlo:</b> la variación compara la descarga de cada boquilla con el caudal teórico por boquilla (L/ha × velocidad × ancho × 60 ÷ 10 000 ÷ 2). ' +
    'Desbalance = |B1 − B2| ÷ caudal por boquilla. Tendencia compara la primera y la segunda mitad del período. ' +
    'Estado: <b style="color:' + VERDE + '">Bueno</b> = 80% o más dentro de rango y última calibración OK; <b style="color:' + ROJO + '">Crítico</b> = menos de 50% dentro de rango, o última calibración fuera con menos de 80%; <b style="color:' + AMBAR + '">Atención</b> = el resto.</p>' +
    '</div>';
}

/* ---------- PDF con Google Docs: logo en la esquina de cada hoja ---------- */

function rico_(par, html) {
  let txt = ''; const negs = [];
  String(html).split(/(<b>.*?<\/b>)/).forEach(p => {
    const m = p.match(/^<b>(.*)<\/b>$/);
    const t = (m ? m[1] : p).replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"');
    if (m && t) negs.push([txt.length, txt.length + t.length - 1]);
    txt += t;
  });
  par.setText(txt || ' ');
  const e = par.editAsText(); e.setBold(false);
  negs.forEach(([x, y]) => e.setBold(x, y, true));
  return par;
}

function tablaDoc_(body, enc, filas, anchos) {
  const t = body.appendTable([enc].concat(filas.length ? filas : [enc.map((x, i) => i ? '' : 'Sin datos')]));
  t.setBorderColor('#cfdccf').setBorderWidth(0.5);
  for (let r = 0; r < t.getNumRows(); r++) {
    const row = t.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(4).setPaddingRight(4);
      const e = cell.editAsText(); e.setFontSize(8).setFontFamily('Arial');
      if (r === 0) { cell.setBackgroundColor(VERDE); e.setForegroundColor('#ffffff').setBold(true); }
      else { e.setForegroundColor('#1b1e21').setBold(false); if (r % 2 === 0) cell.setBackgroundColor('#f4f8f3'); }
    }
  }
  if (anchos) anchos.forEach((w, i) => { if (w) t.setColumnWidth(i, w); });
  return t;
}

function tituloDoc_(body, texto) {
  const p = body.appendParagraph(texto);
  p.setSpacingBefore(12).setSpacingAfter(4);
  p.editAsText().setFontSize(13).setBold(true).setForegroundColor(VERDE).setFontFamily('Arial');
  return p;
}

function pdfReporte_(a, u, nombre) {
  const doc = DocumentApp.create(nombre.replace(/\.pdf$/, ''));
  const id = doc.getId();
  try {
    const body = doc.getBody();
    body.setPageWidth(792).setPageHeight(612).setMarginTop(30).setMarginBottom(30).setMarginLeft(36).setMarginRight(36);

    // Encabezado (se repite en cada hoja): logo en la esquina + nombre
    const hd = doc.addHeader();
    const hp = hd.getNumChildren() ? hd.getChild(0).asParagraph() : hd.appendParagraph('');
    const img = hp.appendInlineImage(logoBlob_());
    const k = 30 / img.getHeight(); img.setHeight(30).setWidth(Math.round(img.getWidth() * k));
    hp.appendText('   INGENIO MADRE TIERRA · Malezas y Madurantes').setFontSize(9).setBold(true).setForegroundColor(VERDE).setFontFamily('Arial');
    const ft = doc.addFooter();
    const fp = ft.getNumChildren() ? ft.getChild(0).asParagraph() : ft.appendParagraph('');
    fp.setText('Reporte de calibración de boquillas · ' + dmy_(a.desde) + ' al ' + dmy_(a.hasta)).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    fp.editAsText().setFontSize(8).setForegroundColor(GRIS).setFontFamily('Arial');

    const t0 = body.getChild(0).asParagraph();
    t0.setText('Reporte de calibración de boquillas por dron');
    t0.editAsText().setFontSize(20).setBold(true).setForegroundColor(VERDE_OSC).setFontFamily('Arial');
    const sub = body.appendParagraph('Período: ' + dmy_(a.desde) + ' al ' + dmy_(a.hasta) + '  ·  Generado por ' + u.nombre + ' el ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'));
    sub.editAsText().setFontSize(10).setBold(false).setForegroundColor(GRIS).setFontFamily('Arial');

    // KPIs
    const conN = a.drones.filter(d => d.n).length, pctOk = a.total ? a.dentro / a.total : 0;
    const nCrit = a.drones.filter(d => d.clase === 'Crítico').length;
    const kp = [['CALIBRACIONES', String(a.total), conN + ' drones', VERDE_OSC],
                ['DENTRO DE RANGO', a.total ? pa_(pctOk, 0) : '—', a.dentro + ' de ' + a.total, pctOk >= 0.8 ? VERDE : pctOk >= 0.5 ? AMBAR : ROJO],
                ['DRONES EN BUEN ESTADO', String(a.drones.filter(d => d.clase === 'Bueno').length), 'de ' + conN, VERDE],
                ['DRONES CRÍTICOS', String(nCrit), 'requieren recalibración', nCrit ? ROJO : VERDE]];
    const kt = body.appendTable([kp.map(x => x[0])]);
    kt.setBorderColor('#ffffff').setBorderWidth(3);
    kp.forEach((x, i) => {
      const cell = kt.getRow(0).getCell(i);
      cell.setBackgroundColor(VERDE_CLARO).setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(8);
      cell.getChild(0).asParagraph().editAsText().setFontSize(8).setBold(true).setForegroundColor(GRIS).setFontFamily('Arial');
      cell.appendParagraph(x[1]).editAsText().setFontSize(22).setBold(true).setForegroundColor(x[3]).setFontFamily('Arial');
      cell.appendParagraph(x[2]).editAsText().setFontSize(8).setBold(false).setForegroundColor(GRIS).setFontFamily('Arial');
    });

    tituloDoc_(body, 'Hallazgos principales');
    a.hallazgos.forEach(hx => { const li = body.appendListItem(''); rico_(li, hx); li.setGlyphType(DocumentApp.GlyphType.BULLET); li.editAsText().setFontSize(10).setForegroundColor('#1b1e21').setFontFamily('Arial'); });
    if (a.recomendaciones.length) {
      tituloDoc_(body, 'Recomendaciones');
      a.recomendaciones.forEach(hx => { const li = body.appendListItem(''); rico_(li, hx); li.setGlyphType(DocumentApp.GlyphType.BULLET); li.editAsText().setFontSize(10).setForegroundColor('#1b1e21').setFontFamily('Arial'); });
    }

    tituloDoc_(body, 'Detalle por dron — calibraciones');
    const tD = tablaDoc_(body,
      ['Dron', 'Estado', 'Calib.', '% dentro', 'Var. B1', 'Var. B2', '|Var. gral.|', 'Var. máx.', 'Desbalance', 'Fuera B1/B2', 'Tendencia', 'Última calibración'],
      a.drones.map(d => [d.dron, d.clase, String(d.n), d.n ? pa_(d.pctDentro, 0) : '—', p_(d.promB1), p_(d.promB2), pa_(d.promAbsGen, 2), pa_(d.maxAbs, 1), pa_(d.desbalance, 1), d.fuera1 + ' / ' + d.fuera2, d.tendencia,
        d.ultimo ? d.ultimo.fecha + ' ' + d.ultimo.descarga + ' · ' + (/DENTRO/i.test(d.ultimo.estado) ? 'OK' : 'FUERA') : '—']),
      [48, 62, 36, 46, 48, 48, 52, 46, 54, 52, 66, 0]);
    a.drones.forEach((d, i) => {
      const r = tD.getRow(i + 1);
      r.getCell(0).editAsText().setBold(true);
      r.getCell(1).editAsText().setBold(true).setForegroundColor(COLOR_CLASE[d.clase]);
      r.getCell(3).editAsText().setBold(true).setForegroundColor(d.pctDentro >= 0.8 ? VERDE : d.pctDentro >= 0.5 ? AMBAR : ROJO);
      if (d.ultimo) r.getCell(11).editAsText().setForegroundColor(/DENTRO/i.test(d.ultimo.estado) ? VERDE : ROJO);
    });

    const conV = a.drones.filter(d => d.vuelos);
    if (conV.length) {
      tituloDoc_(body, 'Vuelos del drone en el período');
      const tV = tablaDoc_(body, ['Dron', 'Vuelos', 'Var. L/ha prom.', '|Var. L/ha| prom.', 'Var. caudal prom.', 'Vuelos fuera de rango'],
        conV.map(d => [d.dron, String(d.vuelos), p_(d.vueloVar), pa_(d.vueloAbsVar, 2), p_(d.vueloVarQ), String(d.vuelosFuera)]));
      conV.forEach((d, i) => { const r = tV.getRow(i + 1); r.getCell(0).editAsText().setBold(true); r.getCell(2).editAsText().setBold(true).setForegroundColor(Math.abs(d.vueloVar) > a.tol ? ROJO : VERDE); });
    }

    if ((a.observaciones || []).length) {
      tituloDoc_(body, 'Observaciones de campo');
      tablaDoc_(body, ['Fecha', 'Dron', 'Registro', 'Observación'], a.observaciones.map(o => [o.fecha, o.dron, o.tipo + ' ' + o.ref, o.texto]), [62, 56, 86, 0]);
    }

    const nota = body.appendParagraph('Cómo leerlo: la variación compara la descarga de cada boquilla con el caudal teórico por boquilla (L/ha × velocidad × ancho × 60 ÷ 10 000 ÷ 2). Desbalance = |B1 − B2| ÷ caudal por boquilla. Tendencia compara la primera y la segunda mitad del período. Estado: Bueno = 80% o más dentro de rango y última calibración OK; Crítico = menos de 50% dentro de rango, o última calibración fuera con menos de 80%; Atención = el resto.');
    nota.setSpacingBefore(12);
    nota.editAsText().setFontSize(8).setBold(false).setForegroundColor(GRIS).setFontFamily('Arial').setBold(0, 12, true).setForegroundColor(0, 12, VERDE);

    doc.saveAndClose();
    return DriveApp.getFileById(id).getAs('application/pdf').setName(nombre);
  } finally {
    try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
  }
}

function generarReporte_(u, desde, hasta, enviar) {
  if (!u.esAdmin) return { ok: false, error: 'Solo el administrador puede generar reportes.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(desde || '') || !/^\d{4}-\d{2}-\d{2}$/.test(hasta || '')) return { ok: false, error: 'Seleccione fecha inicial y final.' };
  if (desde > hasta) return { ok: false, error: 'La fecha inicial es mayor que la final.' };
  const a = analizar_(desde, hasta);
  const html = htmlReporte_(a, u, 'web');
  let enviadoA = '';
  if (enviar) {
    if (!u.correo) return { ok: false, error: 'Su usuario no tiene CORREO ELECTRONICO en la pestaña USUARIOS.' };
    const nombre = 'Reporte_calibracion_' + desde + '_a_' + hasta + '.pdf';
    const pdf = pdfReporte_(a, u, nombre);
    MailApp.sendEmail({
      to: u.correo,
      subject: 'Reporte de calibración de boquillas · ' + dmy_(desde) + ' al ' + dmy_(hasta),
      htmlBody: htmlReporte_(a, u, 'correo'),
      inlineImages: { logo: logoBlob_() },
      attachments: [pdf],
      name: 'Calibración de Boquillas · Ingenio Madre Tierra'
    });
    enviadoA = u.correo;
  }
  return { ok: true, enviadoA: enviadoA, html: html, resumen: { total: a.total, dentro: a.dentro, drones: a.drones.length } };
}

/* ---------- diagnóstico: abrir  URL_DE_LA_APP?action=diag ---------- */

function diagnostico_() {
  // USUARIOS se omite a propósito: tiene contraseñas.
  const hojas = ss_().getSheets().filter(h => norm_(h.getName()).indexOf('USUARIO') < 0).map(h => ({
    hoja: h.getName(),
    encabezados: h.getLastColumn() ? h.getRange(1, 1, Math.min(3, Math.max(h.getLastRow(), 1)), h.getLastColumn()).getDisplayValues() : []
  }));
  let drones = [], err = '';
  try { drones = drones_(); } catch (e) { err = String(e); }
  return { ok: true, hojas: hojas, dronesDetectados: drones, laboresDetectadas: labores_(), flujos: paramsFlujos_(), error: err };
}

/** Ejecútelo una vez desde el editor para autorizar permisos de Hoja y Drive. */
function autorizar() {
  Logger.log(JSON.stringify(catalogos_(), null, 2));
  Logger.log(DriveApp.getFolderById(FOLDER_ID).getName());
  Logger.log('Correos restantes hoy: ' + MailApp.getRemainingDailyQuota());
  const d = DocumentApp.create('prueba permisos'); DriveApp.getFileById(d.getId()).setTrashed(true);
}
