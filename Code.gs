// ============================================================
//  CONFIGURACIÓN  ←  EDITA ESTOS 3 VALORES ANTES DE USAR
// ============================================================
var CONFIG = {
  KOBO_API_TOKEN:     '64cc018b88067397addd36b09288be8b6539cf39',   // Token de API de KoboToolbox
  NOTIFICATION_EMAIL: 'adrian@creamosguatemala.org',                 // Correo que recibirá el resumen
  KOBO_ASSET_UID:     'aJHSDMnJqjhZ6YPUsiyEze'                       // UID del formulario
};

var KOBO_API_URL = 'https://kf.kobotoolbox.org/api/v2/assets/' +
                   CONFIG.KOBO_ASSET_UID + '/data/?format=json&limit=5000';

var SHEET_NAME  = 'Ingresos Creamos';
var STATUS_COL  = 11;  // Columna "Estado" (K)

// ============================================================
//  TRIGGER 1 (cada hora) — Sincroniza datos, NO envía correo
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
      newCount++;
    }
  });

  Logger.log(newCount > 0
    ? '✓ ' + newCount + ' nueva(s) respuesta(s) guardada(s) en el Sheet.'
    : 'Sin respuestas nuevas.');
}

// ============================================================
//  TRIGGER 2 (cada lunes) — Envía resumen solo si hay nuevos
// ============================================================
function sendWeeklySummary() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('No hay datos en el Sheet todavía.');
    return;
  }

  // Primero sincroniza por si hay respuestas de fin de semana
  checkNewSubmissions();

  // Lee todas las filas con estado "Nuevo"
  var lastRow  = sheet.getLastRow();
  var allData  = sheet.getRange(2, 1, lastRow - 1, STATUS_COL).getValues();
  var newRows  = [];
  var rowNums  = [];

  allData.forEach(function(row, i) {
    if (row[STATUS_COL - 1] === 'Nuevo') {
      newRows.push(row);
      rowNums.push(i + 2);  // +2 porque empieza en fila 2
    }
  });

  if (newRows.length === 0) {
    Logger.log('No hay ingresos nuevos esta semana. No se envía correo.');
    return;
  }

  // Enviar correo resumen
  var today   = new Date();
  var options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  var fecha   = today.toLocaleDateString('es-GT', options);

  var subject = 'Resumen semanal de ingresos a Creamos — ' + fecha;
  var html    = buildWeeklySummaryHtml(newRows, fecha);
  var plain   = buildWeeklySummaryPlain(newRows, fecha);

  GmailApp.sendEmail(CONFIG.NOTIFICATION_EMAIL, subject, plain, { htmlBody: html });

  // Marcar filas como "Enviado"
  rowNums.forEach(function(rowNum) {
    sheet.getRange(rowNum, STATUS_COL).setValue('Enviado');
  });

  Logger.log('✓ Correo enviado con ' + newRows.length + ' ingreso(s). Filas marcadas como Enviado.');
}

// ============================================================
//  CONSTRUIR HTML DEL CORREO SEMANAL
// ============================================================
function buildWeeklySummaryHtml(rows, fecha) {
  var rowsHtml = rows.map(function(r, i) {
    var bg = i % 2 === 0 ? '#f8f9fa' : '#ffffff';
    return '<tr style="background:' + bg + ';">' +
      '<td style="padding:10px 12px;border-bottom:1px solid #e0e0e0;">' + (i + 1) + '</td>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #e0e0e0;font-weight:bold;">' + (r[1] || '—') + '</td>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #e0e0e0;">' + (r[2] || '—') + '</td>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #e0e0e0;">' + (r[3] || '—') + '</td>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #e0e0e0;">' + (r[6] || '—') + '</td>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #e0e0e0;">' + (r[7] || '—') + '</td>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #e0e0e0;">' + (r[4] === 'yes' ? '✔ Sí' : 'No') + '</td>' +
      '</tr>';
  }).join('');

  return '<div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;">' +

    '<div style="background:#4a86e8;padding:24px 28px;border-radius:8px 8px 0 0;">' +
    '<h2 style="color:#fff;margin:0;font-size:20px;">Resumen Semanal de Ingresos</h2>' +
    '<p style="color:#c8d8f8;margin:6px 0 0;font-size:14px;">Creamos Guatemala — ' + fecha + '</p>' +
    '</div>' +

    '<div style="background:#fff;padding:24px 28px;border:1px solid #e0e0e0;">' +
    '<p style="font-size:15px;color:#333;">Esta semana se registraron <strong>' + rows.length +
    ' ingreso(s) nuevo(s)</strong>:</p>' +

    '<div style="overflow-x:auto;">' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
    '<thead>' +
    '<tr style="background:#4a86e8;color:#fff;">' +
    '<th style="padding:10px 12px;text-align:left;">#</th>' +
    '<th style="padding:10px 12px;text-align:left;">Nombre</th>' +
    '<th style="padding:10px 12px;text-align:left;">Correo</th>' +
    '<th style="padding:10px 12px;text-align:left;">Fecha Visita</th>' +
    '<th style="padding:10px 12px;text-align:left;">Contacto Emergencia</th>' +
    '<th style="padding:10px 12px;text-align:left;">Tel. Emergencia</th>' +
    '<th style="padding:10px 12px;text-align:left;">Boletín</th>' +
    '</tr>' +
    '</thead>' +
    '<tbody>' + rowsHtml + '</tbody>' +
    '</table>' +
    '</div>' +

    '</div>' +

    '<div style="background:#f0f4ff;padding:12px 28px;border-radius:0 0 8px 8px;border:1px solid #e0e0e0;border-top:none;">' +
    '<p style="font-size:12px;color:#888;margin:0;">' +
    'Este correo se genera automáticamente cada lunes. Ver detalle completo en Google Sheets.' +
    '</p>' +
    '</div>' +

    '</div>';
}

// ============================================================
//  CONSTRUIR VERSIÓN TEXTO PLANO DEL CORREO
// ============================================================
function buildWeeklySummaryPlain(rows, fecha) {
  var lines = [
    'RESUMEN SEMANAL DE INGRESOS — CREAMOS GUATEMALA',
    fecha,
    '',
    rows.length + ' ingreso(s) nuevo(s) esta semana:',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'
  ];

  rows.forEach(function(r, i) {
    lines.push('');
    lines.push((i + 1) + '. ' + (r[1] || 'Sin nombre'));
    lines.push('   Correo:      '  + (r[2] || '—'));
    lines.push('   Visita:      '  + (r[3] || '—'));
    lines.push('   Emergencia:  '  + (r[6] || '—') + '  ' + (r[7] || ''));
    lines.push('   Boletín:     '  + (r[4] === 'yes' ? 'Sí' : 'No'));
  });

  lines.push('');
  lines.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('Ver detalle completo en Google Sheets.');

  return lines.join('\n');
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
    'ID',                          // A - col 1
    'Nombre del Visitante',        // B - col 2
    'Correo del Visitante',        // C - col 3  ← NUEVO
    'Fecha de Visita',             // D - col 4
    'Boletín (newsletter)',        // E - col 5
    'Fecha de Envío',              // F - col 6
    'Contacto de Emergencia',      // G - col 7
    'Teléfono de Emergencia',      // H - col 8
    'Permiso de Fotos',            // I - col 9
    'Exención de Responsabilidad', // J - col 10
    'Estado'                       // K - col 11
  ];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
       .setFontWeight('bold')
       .setBackground('#4a86e8')
       .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 220);
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
    s['_id']                                  || '',  // A
    s['visitor_info/visitor_name']            || '',  // B
    s['visitor_info/visitor_email']           || '',  // C ← campo nuevo del formulario
    s['visitor_info/visit_date']              || '',  // D
    s['visitor_info/newsletter']              || '',  // E
    s['_submission_time']                     || '',  // F
    s['emergency_contact/emergency_name']     || '',  // G
    s['emergency_contact/emergency_phone']    || '',  // H
    s['guidelines_section/photo_permission']  || '',  // I
    s['liability_section/liability_waiver']   || '',  // J
    'Nuevo'                                          // K
  ];
  sheet.appendRow(row);
}

// ============================================================
//  CONFIGURAR TRIGGERS (ejecutar UNA SOLA VEZ)
// ============================================================
function setupTrigger() {
  // Elimina todos los triggers anteriores
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });

  // Trigger 1: sincronización cada hora
  ScriptApp.newTrigger('checkNewSubmissions')
    .timeBased()
    .everyHours(1)
    .create();

  // Trigger 2: resumen cada lunes entre 8:00 y 9:00 AM
  ScriptApp.newTrigger('sendWeeklySummary')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  Logger.log('✓ Triggers configurados:');
  Logger.log('  - checkNewSubmissions: cada hora');
  Logger.log('  - sendWeeklySummary: cada lunes a las 8 AM');
}

// ============================================================
//  PRUEBA MANUAL DE CONEXIÓN
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
    Logger.log('  ID:     ' + last['_id']);
    Logger.log('  Nombre: ' + last['visitor_info/visitor_name']);
    Logger.log('  Fecha:  ' + last['visitor_info/visit_date']);
  }
}

// ============================================================
//  PRUEBA DEL CORREO SEMANAL (sin esperar el lunes)
// ============================================================
function testWeeklySummary() {
  Logger.log('Ejecutando prueba del correo semanal...');
  sendWeeklySummary();
}
