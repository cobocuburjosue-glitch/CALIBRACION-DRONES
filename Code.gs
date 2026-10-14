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
      case 'reporte':    return json_(conSesion_(p.token, u => generarReporte_(u, p.desde, p.hasta, p.enviar !== false)));
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
  if (tot) hall.push('Se registraron <b>' + tot + '</b> calibraciones en <b>' + con.length + '</b> drones; <b>' + pa_(totDentro / tot) + '</b> quedaron dentro de la tolerancia de ±' + (tol * 100).toFixed(0) + '%.');
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

  return { desde: desde, hasta: hasta, tol: tol, total: tot, dentro: totDentro, drones: drones, hallazgos: hall, recomendaciones: rec, vuelos: vue.length };
}

function htmlReporte_(a, u) {
  const col = { 'Bueno': '#2f8a46', 'Atención': '#b7791f', 'Crítico': '#c0392b', 'Sin calibraciones': '#6b7279' };
  const th = 'style="background:#23272b;color:#fff;padding:6px 6px;text-align:left;font-size:11px;vertical-align:bottom"';
  const td = 'style="padding:6px 6px;border-top:1px solid #dfe2dc;font-size:11px"';
  const barra = v => {
    const w = Math.max(0, Math.min(100, (v || 0) * 100));
    return '<div style="background:#e9ebe6;border-radius:4px;height:8px;width:70px;display:inline-block;vertical-align:middle"><div style="background:' + (v >= 0.8 ? '#2f8a46' : v >= 0.5 ? '#b7791f' : '#c0392b') + ';height:8px;border-radius:4px;width:' + w.toFixed(0) + '%"></div></div>';
  };
  const filas = a.drones.map(d => '<tr>' +
    '<td ' + td + '><b>' + h_(d.dron) + '</b></td>' +
    '<td ' + td + '><span style="color:#fff;background:' + col[d.clase] + ';border-radius:10px;padding:2px 8px;font-weight:bold;font-size:11px">' + d.clase + '</span></td>' +
    '<td ' + td + '>' + d.n + '</td>' +
    '<td ' + td + '>' + (d.n ? barra(d.pctDentro) + ' ' + pa_(d.pctDentro, 0) : '—') + '</td>' +
    '<td ' + td + '>' + p_(d.promB1) + '</td><td ' + td + '>' + p_(d.promB2) + '</td>' +
    '<td ' + td + '>' + pa_(d.promAbsGen, 2) + '</td><td ' + td + '>' + pa_(d.maxAbs, 1) + '</td>' +
    '<td ' + td + '>' + pa_(d.desbalance, 1) + '</td>' +
    '<td ' + td + '>' + d.fuera1 + ' / ' + d.fuera2 + '</td>' +
    '<td ' + td + '>' + d.tendencia + '</td>' +
    '<td ' + td + '>' + (d.ultimo ? d.ultimo.fecha + ' ' + h_(d.ultimo.descarga) + ' · ' + (/DENTRO/i.test(d.ultimo.estado) ? 'OK' : 'FUERA') : '—') + '</td>' +
    '</tr>').join('');
  const conV = a.drones.filter(d => d.vuelos);
  const filasV = conV.map(d => '<tr><td ' + td + '><b>' + h_(d.dron) + '</b></td><td ' + td + '>' + d.vuelos + '</td><td ' + td + '>' + p_(d.vueloVar) + '</td><td ' + td + '>' + pa_(d.vueloAbsVar, 2) + '</td><td ' + td + '>' + p_(d.vueloVarQ) + '</td><td ' + td + '>' + d.vuelosFuera + '</td></tr>').join('');
  const kpi = (t, v, s) => '<td style="padding:6px"><div style="border:1px solid #dfe2dc;border-radius:10px;padding:10px 12px"><div style="font-size:11px;color:#6b7279;text-transform:uppercase;letter-spacing:.05em">' + t + '</div><div style="font-size:24px;font-weight:bold;color:#1b1e21">' + v + '</div><div style="font-size:11px;color:#6b7279">' + (s || '') + '</div></div></td>';
  const nBuenos = a.drones.filter(d => d.clase === 'Bueno').length, nCrit = a.drones.filter(d => d.clase === 'Crítico').length;
  return '<div style="font-family:Arial,Helvetica,sans-serif;color:#1b1e21;max-width:900px">' +
    '<div style="background:#141618;color:#fff;padding:18px 20px;border-radius:12px 12px 0 0">' +
    '<div style="font-size:12px;letter-spacing:.2em;color:#aeb4ba">INGENIO MADRE TIERRA · MALEZAS Y MADURANTES</div>' +
    '<div style="font-size:22px;font-weight:bold;margin-top:4px">Reporte de calibración de boquillas por dron</div>' +
    '<div style="font-size:13px;color:#aeb4ba;margin-top:4px">Período: ' + dmy_(a.desde) + ' al ' + dmy_(a.hasta) + ' · Tolerancia ±' + (a.tol * 100).toFixed(0) + '% · Generado por ' + h_(u.nombre) + ' el ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') + '</div></div>' +
    '<div style="border:1px solid #dfe2dc;border-top:0;padding:16px 14px;border-radius:0 0 12px 12px">' +
    '<table style="width:100%;border-collapse:collapse"><tr>' +
    kpi('Calibraciones', a.total, a.drones.filter(d => d.n).length + ' drones') +
    kpi('Dentro de rango', a.total ? pa_(a.dentro / a.total, 0) : '—', a.dentro + ' de ' + a.total) +
    kpi('Drones en buen estado', nBuenos, 'de ' + a.drones.filter(d => d.n).length) +
    kpi('Drones críticos', nCrit, 'requieren recalibración') +
    '</tr></table>' +
    '<h3 style="margin:18px 6px 6px;font-size:16px">Hallazgos principales</h3><ul style="margin:0 6px 0 22px;padding:0;font-size:13px;line-height:1.6">' + a.hallazgos.map(x => '<li>' + x + '</li>').join('') + '</ul>' +
    (a.recomendaciones.length ? '<h3 style="margin:18px 6px 6px;font-size:16px">Recomendaciones</h3><ul style="margin:0 6px 0 22px;padding:0;font-size:13px;line-height:1.6">' + a.recomendaciones.map(x => '<li>' + x + '</li>').join('') + '</ul>' : '') +
    '<h3 style="margin:18px 6px 8px;font-size:16px">Detalle por dron — calibraciones</h3>' +
    '<div style="overflow-x:auto"><table style="border-collapse:collapse;width:100%"><tr>' +
    ['Dron', 'Estado', 'Calib.', '% dentro de rango', 'Var. prom. B1', 'Var. prom. B2', '|Var. gral.| prom.', 'Var. máx.', 'Desbalance B1–B2', 'Fuera B1 / B2', 'Tendencia', 'Última calibración'].map(x => '<th ' + th + '>' + x + '</th>').join('') +
    '</tr>' + (filas || '<tr><td ' + td + ' colspan="12">Sin datos</td></tr>') + '</table></div>' +
    (conV.length ? '<h3 style="margin:18px 6px 8px;font-size:16px">Vuelos del drone en el período</h3><table style="border-collapse:collapse;width:100%"><tr>' +
      ['Dron', 'Vuelos', 'Var. L/ha prom.', '|Var. L/ha| prom.', 'Var. caudal prom.', 'Vuelos fuera de ±' + (a.tol * 100).toFixed(0) + '%'].map(x => '<th ' + th + '>' + x + '</th>').join('') + '</tr>' + filasV + '</table>' : '') +
    '<p style="font-size:11px;color:#6b7279;margin:16px 6px 0;line-height:1.5"><b>Cómo leerlo:</b> la variación compara la descarga de cada boquilla con el caudal teórico por boquilla (L/ha × velocidad × ancho × 60 ÷ 10 000 ÷ 2). ' +
    'Desbalance = |B1 − B2| ÷ caudal por boquilla. Tendencia compara la primera y la segunda mitad del período. ' +
    'Estado: <b>Bueno</b> = 80% o más dentro de rango, variación promedio dentro de la tolerancia y última calibración OK; <b>Crítico</b> = menos de 50% dentro de rango, o última calibración fuera con menos de 80%; <b>Atención</b> = el resto.</p>' +
    '</div></div>';
}

function generarReporte_(u, desde, hasta, enviar) {
  if (!u.esAdmin) return { ok: false, error: 'Solo el administrador puede generar reportes.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(desde || '') || !/^\d{4}-\d{2}-\d{2}$/.test(hasta || '')) return { ok: false, error: 'Seleccione fecha inicial y final.' };
  if (desde > hasta) return { ok: false, error: 'La fecha inicial es mayor que la final.' };
  const a = analizar_(desde, hasta);
  const html = htmlReporte_(a, u);
  let enviadoA = '';
  if (enviar) {
    if (!u.correo) return { ok: false, error: 'Su usuario no tiene CORREO ELECTRONICO en la pestaña USUARIOS.' };
    const nombre = 'Reporte_calibracion_' + desde + '_a_' + hasta + '.pdf';
    const pdf = Utilities.newBlob('<html><head><meta charset="utf-8"></head><body>' + html + '</body></html>', 'text/html', 'r.html').getAs('application/pdf').setName(nombre);
    MailApp.sendEmail({
      to: u.correo,
      subject: 'Reporte de calibración de boquillas · ' + dmy_(desde) + ' al ' + dmy_(hasta),
      htmlBody: html,
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
}
