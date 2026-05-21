// ============================================================
//  CONFIGURACIÓN  ←  EDITA ESTOS 3 VALORES ANTES DE USAR
// ============================================================
var CONFIG = {
  KOBO_API_TOKEN:     'PON_AQUI_TU_TOKEN_DE_KOBO',   // Tu token de API de KoboToolbox
  NOTIFICATION_EMAIL: 'PON_AQUI_EL_CORREO_DESTINO',  // Correo que recibirá las alertas
  KOBO_ASSET_UID:     'aJHSDMnJqjhZ6YPUsiyEze'        // UID del formulario (ya configurado)
};

var KOBO_API_URL = 'https://kf.kobotoolbox.org/api/v2/assets/' +
                   CONFIG.KOBO_ASSET_UID + '/data/?format=json&limit=5000';

var SHEET_NAME = 'Ingresos Creamos';

// ============================================================
//  FUNCIÓN PRINCIPAL — se ejecuta con el trigger automático
// ============================================================
function checkNewSubmissions() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    setupHeaders(sheet);
  }

  var processedIds = getProcessedIds(sheet);
  var submissions  = fetchKoboData();

  if (!submissions) {
    Logger.log('No se pudo obtener datos de KoboToolbox.');
    return;
  }

  var newCount = 0;

  submissions.forEach(function(submission) {
    var id = String(submission['_id']);

    if (!processedIds[id]) {
      appendToSheet(sheet, submission);
      sendNotificationEmail(submission);
      newCount++;
    }
  });

  Logger.log(newCount > 0
    ? '✓ ' + newCount + ' nueva(s) respuesta(s) procesada(s).'
    : 'Sin respuestas nuevas.');
}

// ============================================================
//  OBTENER DATOS DE KOBOTOOLBOX
// ============================================================
function fetchKoboData() {
  try {
    var options = {
      method: 'GET',
      headers: { 'Authorization': 'Token ' + CONFIG.KOBO_API_TOKEN },
      muteHttpExceptions: true
    };

    var allResults = [];
    var url = KOBO_API_URL;

    // Maneja paginación por si hay muchas respuestas
    while (url) {
      var response = UrlFetchApp.fetch(url, options);
      var code     = response.getResponseCode();

      if (code !== 200) {
        Logger.log('Error KoboToolbox (' + code + '): ' + response.getContentText());
        return null;
      }

      var data = JSON.parse(response.getContentText());
      allResults = allResults.concat(data.results || []);
      url = data.next || null;
    }

    return allResults;

  } catch (e) {
    Logger.log('Error de conexión: ' + e.message);
    return null;
  }
}

// ============================================================
//  ENCABEZADOS DE LA HOJA
// ============================================================
function setupHeaders(sheet) {
  var headers = [
    'ID',
    'Nombre del Visitante',
    'Fecha de Visita',
    'Fecha de Envío',
    'Boletín (newsletter)',
    'Contacto de Emergencia',
    'Teléfono de Emergencia',
    'Permiso de Fotos',
    'Exención de Responsabilidad',
    'Estado'
  ];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
       .setFontWeight('bold')
       .setBackground('#4a86e8')
       .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
}

// ============================================================
//  IDs YA PROCESADOS (para no duplicar)
// ============================================================
function getProcessedIds(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var map = {};
  ids.forEach(function(row) { map[String(row[0])] = true; });
  return map;
}

// ============================================================
//  AGREGAR FILA A LA HOJA
// ============================================================
function appendToSheet(sheet, s) {
  var row = [
    s['_id']                                  || '',
    s['visitor_info/visitor_name']            || '',
    s['visitor_info/visit_date']              || '',
    s['_submission_time']                     || '',
    s['visitor_info/newsletter']              || '',
    s['emergency_contact/emergency_name']     || '',
    s['emergency_contact/emergency_phone']    || '',
    s['guidelines_section/photo_permission']  || '',
    s['liability_section/liability_waiver']   || '',
    'Nuevo'
  ];
  sheet.appendRow(row);
}

// ============================================================
//  ENVIAR CORREO DE NOTIFICACIÓN
// ============================================================
function sendNotificationEmail(s) {
  var name           = s['visitor_info/visitor_name']            || 'Sin nombre';
  var visitDate      = s['visitor_info/visit_date']              || 'No especificada';
  var submissionTime = s['_submission_time']                     || '';
  var newsletter     = s['visitor_info/newsletter']              || '';
  var emergencyName  = s['emergency_contact/emergency_name']     || '';
  var emergencyPhone = s['emergency_contact/emergency_phone']    || '';
  var photoPermit    = s['guidelines_section/photo_permission']  || '';
  var liability      = s['liability_section/liability_waiver']   || '';
  var id             = s['_id']                                  || '';

  var subject = 'Nuevo ingreso a Creamos: ' + name;

  var body =
    'Nueva respuesta recibida en el formulario de ingresos de Creamos.\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '  INFORMACIÓN DEL VISITANTE\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'Nombre:          ' + name           + '\n' +
    'Fecha de visita: ' + visitDate      + '\n' +
    'Enviado el:      ' + submissionTime + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '  CONTACTO DE EMERGENCIA\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'Nombre:   ' + emergencyName  + '\n' +
    'Teléfono: ' + emergencyPhone + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '  PREFERENCIAS Y ACUERDOS\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'Boletín informativo:       ' + newsletter  + '\n' +
    'Permiso de fotos:          ' + photoPermit + '\n' +
    'Exención de responsabilidad: ' + liability + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'ID de respuesta: ' + id + '\n' +
    'Ver todos los ingresos en Google Sheets.';

  // Versión HTML del correo (se ve mejor en Gmail)
  var htmlBody =
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">' +
    '<div style="background:#4a86e8;padding:20px;border-radius:8px 8px 0 0;">' +
    '<h2 style="color:#fff;margin:0;">&#128100; Nuevo Ingreso a Creamos</h2></div>' +
    '<div style="background:#f8f9fa;padding:20px;border:1px solid #e0e0e0;">' +

    '<h3 style="color:#4a86e8;border-bottom:2px solid #4a86e8;padding-bottom:5px;">Información del Visitante</h3>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><td style="padding:6px 0;font-weight:bold;width:40%;">Nombre:</td>' +
    '    <td style="padding:6px 0;">' + name + '</td></tr>' +
    '<tr><td style="padding:6px 0;font-weight:bold;">Fecha de visita:</td>' +
    '    <td style="padding:6px 0;">' + visitDate + '</td></tr>' +
    '<tr><td style="padding:6px 0;font-weight:bold;">Enviado el:</td>' +
    '    <td style="padding:6px 0;">' + submissionTime + '</td></tr>' +
    '</table>' +

    '<h3 style="color:#4a86e8;border-bottom:2px solid #4a86e8;padding-bottom:5px;margin-top:20px;">Contacto de Emergencia</h3>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><td style="padding:6px 0;font-weight:bold;width:40%;">Nombre:</td>' +
    '    <td style="padding:6px 0;">' + emergencyName + '</td></tr>' +
    '<tr><td style="padding:6px 0;font-weight:bold;">Teléfono:</td>' +
    '    <td style="padding:6px 0;">' + emergencyPhone + '</td></tr>' +
    '</table>' +

    '<h3 style="color:#4a86e8;border-bottom:2px solid #4a86e8;padding-bottom:5px;margin-top:20px;">Preferencias y Acuerdos</h3>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><td style="padding:6px 0;font-weight:bold;width:40%;">Boletín:</td>' +
    '    <td style="padding:6px 0;">' + newsletter + '</td></tr>' +
    '<tr><td style="padding:6px 0;font-weight:bold;">Permiso de fotos:</td>' +
    '    <td style="padding:6px 0;">' + photoPermit + '</td></tr>' +
    '<tr><td style="padding:6px 0;font-weight:bold;">Exención:</td>' +
    '    <td style="padding:6px 0;">' + liability + '</td></tr>' +
    '</table>' +

    '</div>' +
    '<div style="background:#e8e8e8;padding:10px 20px;border-radius:0 0 8px 8px;font-size:12px;color:#666;">' +
    'ID: ' + id + '</div>' +
    '</div>';

  GmailApp.sendEmail(
    CONFIG.NOTIFICATION_EMAIL,
    subject,
    body,
    { htmlBody: htmlBody }
  );
}

// ============================================================
//  CONFIGURAR TRIGGER AUTOMÁTICO (ejecutar UNA SOLA VEZ)
// ============================================================
function setupTrigger() {
  // Elimina triggers anteriores para no duplicar
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('checkNewSubmissions')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('✓ Trigger listo: checkNewSubmissions correrá cada hora.');
}

// ============================================================
//  PRUEBA MANUAL — corre esto una vez para verificar conexión
// ============================================================
function testConnection() {
  var submissions = fetchKoboData();

  if (!submissions) {
    Logger.log('ERROR: No se pudo conectar. Revisa tu KOBO_API_TOKEN.');
    return;
  }

  Logger.log('Conexión exitosa. Total de respuestas: ' + submissions.length);

  if (submissions.length > 0) {
    var last = submissions[submissions.length - 1];
    Logger.log('Última respuesta:');
    Logger.log('  ID: '     + last['_id']);
    Logger.log('  Nombre: ' + last['visitor_info/visitor_name']);
    Logger.log('  Fecha: '  + last['visitor_info/visit_date']);
  }
}
