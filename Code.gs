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
      case 'usuarios':   return json_({ ok: true, usuarios: listaUsuarios_() });
      case 'login':      return json_(login_(p.usuario, p.password));
      case 'catalogos':  return json_(conSesion_(p.token, u => ({ ok: true, usuario: u, catalogos: catalogos_() })));
      case 'guardar':    return json_(conSesion_(p.token, u => guardar_(u, p.datos)));
      case 'guardarVuelo': return json_(conSesion_(p.token, u => guardarVuelo_(u, p.datos)));
      case 'vuelos':     return json_(conSesion_(p.token, u => ({ ok: true, vuelos: vuelos_(60) })));
      case 'hoy':        return json_(conSesion_(p.token, u => ({ ok: true, registros: registros_(true) })));
      case 'dashboard':  return json_(conSesion_(p.token, u => {
                            if (!u.permisos.dashboard) return { ok: false, error: 'No tiene permiso para el dashboard.' };
                            return { ok: true, registros: registros_(false) };
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
  ['marcaTiempo', 'MARCA DE TIEMPO', 'dd/MM/yyyy HH:mm']
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

  const c = calcular_({ b1, b2, tiempo, lha: labor.lha, vel: dron.vel, ancho: dron.ancho }, cat.flujos.variacion);
  const [Y, M, D] = d.fecha.split('-').map(Number);
  const fecha = new Date(Y, M - 1, D, 12, 0, 0);
  const fechaStr = fechaTxt_(d.fecha);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = hojaCalibraciones_();
    const mapa = mapaColumnas_(sh, true);

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
      registradoPor: u.nombre,
      marcaTiempo: new Date()
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
  v.forEach(r => {
    const g = k => mapa[k] ? r[mapa[k] - 1] : '';
    const fecha = fechaTxt_(g('fecha'));
    if (!fecha) return;
    if (soloHoy && fecha !== hoy) return;
    const o = { fecha: fecha };
    ['supervisor', 'dron', 'descarga', 'foto', 'labor', 'estado', 'registradoPor'].forEach(k => o[k] = String(g(k) || ''));
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
  ['marcaTiempo', 'MARCA DE TIEMPO', 'dd/MM/yyyy HH:mm']
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
  const c = calcularVuelo_(Object.assign({ lhaTeorico: labor.lha }, v));
  const [Y, M, D] = d.fecha.split('-').map(Number);

  const valores = Object.assign({}, v, c, {
    dron: dron.dron, lhaTeorico: labor.lha, labor: labor.labor,
    fecha: new Date(Y, M - 1, D, 12, 0, 0),
    supervisor: d.supervisor || u.nombre, calibracion: d.calibracion || '',
    registradoPor: u.nombre, marcaTiempo: new Date()
  });
  ['areaPoligono', 'distanciaMax'].forEach(k => { if (isNaN(valores[k])) valores[k] = ''; });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = hojaVuelos_();
    const mapa = mapaColumnas_(sh, true, CAMPOS_VUELO);
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
  v.forEach(r => {
    const g = k => mapa[k] ? r[mapa[k] - 1] : '';
    if (!String(g('dron')).trim()) return;
    const o = { fecha: fechaTxt_(g('fecha')) };
    ['dron', 'labor', 'supervisor', 'calibracion', 'registradoPor'].forEach(k => o[k] = String(g(k) || ''));
    ['areaPoligono', 'distanciaMax', 'lhaTeorico', 'carga', 'areaReal', 'aplicados', 'devueltos', 'lhaReal', 'variacion', 'varCaudal'].forEach(k => {
      const n = num_(g(k)); o[k] = isNaN(n) ? null : n;
    });
    out.push(o);
  });
  return out.slice(-max).reverse();
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
}
