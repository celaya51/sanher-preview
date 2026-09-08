/**
 * ============================================================
 *  SANHER — WebApp de Testimonios (Google Apps Script)
 * ============================================================
 *  SIRVE A:
 *   1. El BOT de WhatsApp (lee los testimonios pendientes y
 *      marca aprobado/rechazado).
 *   2. La web del carrusel (opcional: traer solo aprobados).
 *
 *  NOTA: Autodetecta las columnas por palabra clave en el encabezado,
 *  así no necesitas saber el índice exacto. Ajusta las CLAVE_* si
 *  los títulos de tu hoja cambian.
 * ============================================================
 */

// ------------------------------------------------------------------
//  CONFIGURACIÓN
// ------------------------------------------------------------------
// Si el script está VINCULADO a la hoja (creado desde la hoja),
// deja ID_HOJA vacío. Si es suelto, pega el ID de la hoja.
const ID_HOJA = '1SKpP-if4TXIAaGN8jCrOp-6joiW50eD2P79A1cSM1No';
const NOMBRE_HOJA = 'Respuestas de formulario 1'; // pestaña donde caen las respuestas
const URL_WEBHOOK_TESTIMONIOS = 'https://openclaw.veia.com.mx/hooks/sanher/testimonials';
const PROPIEDAD_SECRETO = 'SANHER_WEBHOOK_SECRET';

// Palabras clave para detectar columnas por su encabezado:
const CLAVE_MARCA      = 'marca';        // Timestamp (marca de tiempo)
const CLAVE_NOMBRE     = 'nombre';       // Nombre (opcional)
const CLAVE_COMO       = 'particip' ;    // Participó como (arrendador/inquilino)
const CLAVE_TESTIMONIO = 'gustó';        // Pregunta 6 → texto del testimonio
const CLAVE_AUTORIZA   = 'autoriza';     // Pregunta 7 → autoriza compartir
// ------------------------------------------------------------------

// ------------------------------------------------------------------
//  HELPERS
// ------------------------------------------------------------------
function obtenerHoja() {
  let ss;
  if (ID_HOJA) {
    ss = SpreadsheetApp.openById(ID_HOJA);
  } else {
    const activo = SpreadsheetApp.getActiveSpreadsheet();
    if (!activo) throw new Error('No hay hoja vinculada. Pon ID_HOJA.');
    ss = activo;
  }
  let hoja = ss.getSheetByName(NOMBRE_HOJA);
  if (!hoja) hoja = ss.getSheets()[0];
  return hoja;
}

function obtenerColumnas(hoja) {
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const idx = {};
  const buscar = (clave) => {
    if (!clave) return -1;
    for (let i = 0; i < encabezados.length; i++)
      if (String(encabezados[i]).toLowerCase().includes(clave)) return i;
    return -1;
  };
  idx.marca      = buscar(CLAVE_MARCA);
  idx.nombre     = buscar(CLAVE_NOMBRE);
  idx.como       = buscar(CLAVE_COMO);
  idx.testimonio = buscar(CLAVE_TESTIMONIO);
  idx.autoriza   = buscar(CLAVE_AUTORIZA);
  idx.estado     = buscar('estado');
  if (idx.estado === -1) {
    const nc = encabezados.length + 1;
    hoja.getRange(1, nc).setValue('Estado');
    idx.estado = nc - 1;
  }
  return idx;
}

// Devuelve testimonios pendientes: autorizado (no "No") y sin estado aún
function obtenerPendientes() {
  const hoja = obtenerHoja();
  const idx = obtenerColumnas(hoja);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  const datos = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
  const pendientes = [];
  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];
    const autoriza = idx.autoriza >= 0 ? String(fila[idx.autoriza]).toLowerCase() : 'si';
    // Rechaza solamente respuestas negativas (p. ej. "No autorizo").
    // No uses includes('no'): "Sí, con mi nombre" contiene "no" dentro de "nombre".
    const autorizado = !/^no\b/.test(autoriza.normalize('NFD').replace(/[\u0300-\u036f]/g, ''));
    const estado = idx.estado >= 0 ? String(fila[idx.estado]).trim().toLowerCase() : '';
    if (autorizado && estado === '') {
      pendientes.push({
        fila: i + 2,
        marca: idx.marca >= 0 ? fila[idx.marca] : null,
        nombre: idx.nombre >= 0 ? fila[idx.nombre] : null,
        como: idx.como >= 0 ? fila[idx.como] : null,
        testimonio: idx.testimonio >= 0 ? fila[idx.testimonio] : null,
        autoriza: fila[idx.autoriza]
      });
    }
  }
  return pendientes;
}

// Devuelve únicamente testimonios aprobados y autorizados para publicación.
// Si la persona eligió publicación anónima, nunca expone su nombre.
function obtenerAprobados() {
  const hoja = obtenerHoja();
  const idx = obtenerColumnas(hoja);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];

  const datos = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
  const aprobados = [];
  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];
    const autoriza = idx.autoriza >= 0 ? String(fila[idx.autoriza]).trim() : '';
    const autorizacionNormalizada = autoriza
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
    const estado = idx.estado >= 0 ? String(fila[idx.estado]).trim().toLowerCase() : '';
    const testimonio = idx.testimonio >= 0 ? String(fila[idx.testimonio]).trim() : '';

    if (estado !== 'aprobado' || !testimonio || /^no\b/.test(autorizacionNormalizada)) continue;

    const esAnonimo = autorizacionNormalizada.includes('anonim');
    aprobados.push({
      fila: i + 2,
      nombre: esAnonimo ? 'Cliente SANHER' : (String(fila[idx.nombre] || '').trim() || 'Cliente SANHER'),
      como: idx.como >= 0 ? String(fila[idx.como]).trim() : '',
      testimonio: testimonio
    });
  }
  return aprobados.reverse();
}

// Marca el estado de una fila
function marcarEstado(numeroFila, estado) {
  const hoja = obtenerHoja();
  const idx = obtenerColumnas(hoja);
  if (idx.estado < 0) throw new Error('No existe columna Estado');
  hoja.getRange(numeroFila, idx.estado + 1).setValue(estado);
}

function obtenerSecreto() {
  const secreto = PropertiesService.getScriptProperties().getProperty(PROPIEDAD_SECRETO);
  if (!secreto) throw new Error('Falta la propiedad de script SANHER_WEBHOOK_SECRET');
  return secreto;
}

function salidaJson(valor) {
  return ContentService.createTextOutput(JSON.stringify(valor))
    .setMimeType(ContentService.MimeType.JSON);
}

function salidaPublica(valor, callback) {
  if (!callback) return salidaJson(valor);
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return salidaJson({ ok: false, error: 'Callback inválido' });
  }
  return ContentService
    .createTextOutput('/**/' + callback + '(' + JSON.stringify(valor) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ------------------------------------------------------------------
//  WEB APP
// ------------------------------------------------------------------
function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    if (String(params.accion || '').toLowerCase() === 'aprobados') {
      return salidaPublica({ ok: true, data: obtenerAprobados() }, String(params.callback || ''));
    }
    return salidaJson({ ok: true, servicio: 'SANHER testimonios' });
  } catch (err) {
    return salidaJson({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = (e && e.postData && e.postData.contents) || '';
    let params = {};
    try { params = JSON.parse(body); } catch (err) { params = body ? Object.fromEntries(new URLSearchParams(body)) : {}; }
    if (String(params.secret || '') !== obtenerSecreto()) throw new Error('No autorizado');

    const accion = String(params.accion || 'marcar').toLowerCase();
    if (accion === 'listar') return salidaJson({ ok: true, data: obtenerPendientes() });
    if (accion !== 'marcar') throw new Error('Acción inválida');

    const fila = parseInt(params.fila, 10);
    const estado = (params.estado || '').toLowerCase();
    if (!fila || !['aprobado', 'rechazado', 'pendiente'].includes(estado))
      throw new Error('Parámetros inválidos. Se espera {fila, estado}');
    marcarEstado(fila, estado);
    return salidaJson({ ok: true, fila: fila, estado: estado });
  } catch (err) {
    return salidaJson({ ok: false, error: String(err) });
  }
}

// Crea un disparador instalable para esta función: "Al enviar formulario".
function notificarNuevoTestimonio(e) {
  const hoja = obtenerHoja();
  const fila = e && e.range ? e.range.getRow() : hoja.getLastRow();
  if (fila < 2) return;

  const respuesta = UrlFetchApp.fetch(URL_WEBHOOK_TESTIMONIOS, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'X-Sanher-Webhook-Secret': obtenerSecreto() },
    payload: JSON.stringify({ fila: fila }),
    muteHttpExceptions: true
  });

  const codigo = respuesta.getResponseCode();
  if (codigo < 200 || codigo >= 300) {
    throw new Error('Webhook SANHER respondió ' + codigo + ': ' + respuesta.getContentText());
  }
}

// Prueba manual: notifica usando la última fila de respuestas.
function probarWebhookTestimonios() {
  notificarNuevoTestimonio(null);
}

function myFunction() {}
