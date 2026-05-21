// ============================================================
//  CONFIGURACIÓN — se guarda en el Sheet, no en el código
// ============================================================
function getConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    KOBO_API_TOKEN:     props.getProperty('KOBO_API_TOKEN')     || '64cc018b88067397addd36b09288be8b6539cf39',
    NOTIFICATION_EMAIL: props.getProperty('NOTIFICATION_EMAIL') || 'adrian@creamosguatemala.org',
    KOBO_ASSET_UID:     'aJHSDMnJqjhZ6YPUsiyEze',
    BLOOMERANG_API_KEY: props.getProperty('BLOOMERANG_API_KEY') || '',
    BLOOMERANG_ENABLED: props.getProperty('BLOOMERANG_ENABLED') === 'true'
  };
}

var KOBO_BASE_URL = 'https://kf.kobotoolbox.org/api/v2/assets/';
var SHEET_NAME    = 'Ingresos Creamos';
var STATUS_COL    = 11;

// ============================================================
//  MENÚ — aparece automáticamente al abrir el Sheet
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Creamos')
    .addItem('Configurar credenciales', 'showSetupDialog')
    .addSeparator()
    .addItem('Activar triggers automaticos', 'setupTrigger')
    .addSeparator()
    .addItem('Probar conexion con Kobo', 'testConnection')
    .addItem('Probar correo semanal ahora', 'testWeeklySummary')
    .addItem('Probar contacto en Bloomerang', 'testBloomerang')
    .addToUi();
}

// ============================================================
//  DIALOGO DE CONFIGURACIÓN
// ============================================================
function showSetupDialog() {
  var config = getConfig();
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head>' +
    '<style>' +
    'body{font-family:Arial,sans-serif;padding:20px;color:#333;font-size:14px;}' +
    'h2{color:#4a86e8;margin:0 0 6px;font-size:18px;}' +
    'p.sub{color:#888;font-size:12px;margin:0 0 20px;}' +
    'label{display:block;font-weight:600;margin:14px 0 4px;font-size:13px;}' +
    'input[type=text],input[type=email]{width:100%;padding:8px 10px;border:1px solid #ddd;border-radius:6px;box-sizing:border-box;font-size:13px;}' +
    'input:focus{outline:none;border-color:#4a86e8;}' +
    '.toggle{display:flex;align-items:center;gap:8px;margin-top:14px;}' +
    '.toggle input{width:auto;}' +
    '.btn{display:block;width:100%;margin-top:20px;padding:10px;background:#4a86e8;color:#fff;border:none;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;}' +
    '.btn:hover{background:#3367d6;}' +
    '.ok{display:none;color:#2e7d32;font-weight:600;margin-top:12px;text-align:center;}' +
    '.section{background:#f5f7ff;border-radius:8px;padding:14px;margin-top:16px;}' +
    '.section-title{font-weight:700;font-size:12px;text-transform:uppercase;letter-spacing:.5px;color:#4a86e8;margin-bottom:2px;}' +
    '</style></head><body>' +

    '<h2>Configuracion de Creamos</h2>' +
    '<p class="sub">Tus datos se guardan de forma segura en este proyecto de Apps Script.</p>' +

    '<div class="section">' +
    '<div class="section-title">KoboToolbox</div>' +
    '<label>Token de API</label>' +
    '<input type="text" id="kobo" value="' + config.KOBO_API_TOKEN + '" placeholder="Pega tu token de KoboToolbox">' +
    '</div>' +

    '<div class="section">' +
    '<div class="section-title">Notificaciones</div>' +
    '<label>Correo que recibe el resumen semanal</label>' +
    '<input type="email" id="email" value="' + config.NOTIFICATION_EMAIL + '" placeholder="correo@creamos.org">' +
    '</div>' +

    '<div class="section">' +
    '<div class="section-title">Bloomerang (opcional)</div>' +
    '<label>API Key de Bloomerang</label>' +
    '<input type="text" id="bloom" value="' + config.BLOOMERANG_API_KEY + '" placeholder="Deja vacio si no lo usas aun">' +
    '<div class="toggle">' +
    '<input type="checkbox" id="bloomOn" ' + (config.BLOOMERANG_ENABLED ? 'checked' : '') + '>' +
    '<label style="margin:0;font-weight:400;">Activar sincronizacion con Bloomerang</label>' +
    '</div>' +
    '</div>' +

    '<button class="btn" onclick="save()">Guardar configuracion</button>' +
    '<div class="ok" id="ok">Guardado correctamente</div>' +

    '<script>' +
    'function save(){' +
    '  var d={' +
    '    kobo:document.getElementById("kobo").value,' +
    '    email:document.getElementById("email").value,' +
    '    bloom:document.getElementById("bloom").value,' +
    '    bloomOn:document.getElementById("bloomOn").checked' +
    '  };' +
    '  google.script.run.withSuccessHandler(function(){' +
    '    document.getElementById("ok").style.display="block";' +
    '    setTimeout(function(){google.script.host.close();},1200);' +
    '  }).saveConfig(d);' +
    '}' +
    '<\/script></body></html>'
  ).setWidth(460).setHeight(520);

  SpreadsheetApp.getUi().showModalDialog(html, 'Configuracion — Creamos');
}

function saveConfig(data) {
  var props = PropertiesService.getScriptProperties();
  if (data.kobo)  props.setProperty('KOBO_API_TOKEN',     data.kobo);
  if (data.email) props.setProperty('NOTIFICATION_EMAIL', data.email);
  props.setProperty('BLOOMERANG_API_KEY', data.bloom   || '');
  props.setProperty('BLOOMERANG_ENABLED', data.bloomOn ? 'true' : 'false');
}

// ============================================================
//  TRIGGER 1 (cada hora) — Sincroniza datos, NO envía correo
// ============================================================
function checkNewSubmissions() {
  var config = getConfig();
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    setupHeaders(sheet);
  }

  var processedIds = getProcessedIds(sheet);
  var submissions  = fetchKoboData(config);

  if (!submissions) {
    Logger.log('No se pudo obtener datos de KoboToolbox.');
    return;
  }

  var newCount = 0;

  submissions.forEach(function(submission) {
    var id = String(submission['_id']);
    if (!processedIds[id]) {
      appendToSheet(sheet, submission);
      if (config.BLOOMERANG_ENABLED) {
        createBloomerangContact(submission, config);
      }
      newCount++;
    }
  });

  Logger.log(newCount > 0
    ? newCount + ' nueva(s) respuesta(s) guardada(s).'
    : 'Sin respuestas nuevas.');
}

// ============================================================
//  TRIGGER 2 (cada lunes 8 AM) — Resumen solo si hay nuevos
// ============================================================
function sendWeeklySummary() {
  var config = getConfig();
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName(SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('No hay datos en el Sheet todavia.');
    return;
  }

  checkNewSubmissions();

  var lastRow = sheet.getLastRow();
  var allData = sheet.getRange(2, 1, lastRow - 1, STATUS_COL).getValues();
  var newRows = [];
  var rowNums = [];

  allData.forEach(function(row, i) {
    if (row[STATUS_COL - 1] === 'Nuevo') {
      newRows.push(row);
      rowNums.push(i + 2);
    }
  });

  if (newRows.length === 0) {
    Logger.log('Sin ingresos nuevos esta semana. No se envia correo.');
    return;
  }

  var today  = new Date();
  var opts   = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  var fecha  = today.toLocaleDateString('es-GT', opts);
  var subject = 'Resumen semanal de ingresos a Creamos — ' + fecha;

  GmailApp.sendEmail(
    config.NOTIFICATION_EMAIL,
    subject,
    buildWeeklySummaryPlain(newRows, fecha),
    { htmlBody: buildWeeklySummaryHtml(newRows, fecha) }
  );

  rowNums.forEach(function(rowNum) {
    sheet.getRange(rowNum, STATUS_COL).setValue('Enviado');
  });

  Logger.log('Correo enviado con ' + newRows.length + ' ingreso(s).');
}

// ============================================================
//  HELPERS DE FORMATO
// ============================================================
function fmtDate(val) {
  if (!val) return '—';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'America/Guatemala', 'dd/MM/yyyy');
  }
  var m = String(val).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[3] + '/' + m[2] + '/' + m[1];
  return String(val);
}

function fmtEmail(val) {
  if (!val || val === 'email not found' || String(val).trim() === '') {
    return '<span style="color:#bbb;font-style:italic;">Sin correo</span>';
  }
  return '<a href="mailto:' + val + '" style="color:#4a86e8;text-decoration:none;">' + val + '</a>';
}

function fmtNewsletter(val) {
  return val === 'yes'
    ? '<span style="color:#2e7d32;font-weight:bold;">&#10003; Si</span>'
    : '<span style="color:#999;">No</span>';
}

// ============================================================
//  CONSTRUIR HTML DEL CORREO SEMANAL
// ============================================================
function buildWeeklySummaryHtml(rows, fecha) {
  var rowsHtml = rows.map(function(r, i) {
    var bg = i % 2 === 0 ? '#f8f9fa' : '#ffffff';
    return '<tr style="background:' + bg + ';">' +
      '<td style="padding:10px 14px;border-bottom:1px solid #e8e8e8;color:#999;font-size:12px;">' + (i + 1) + '</td>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #e8e8e8;font-weight:600;color:#222;">' + (r[1] || '—') + '</td>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #e8e8e8;font-size:12px;">' + fmtEmail(r[2]) + '</td>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #e8e8e8;white-space:nowrap;">' + fmtDate(r[3]) + '</td>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #e8e8e8;">' + (r[6] || '—') + '</td>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #e8e8e8;white-space:nowrap;">' + (r[7] || '—') + '</td>' +
      '<td style="padding:10px 14px;border-bottom:1px solid #e8e8e8;text-align:center;">' + fmtNewsletter(r[4]) + '</td>' +
      '</tr>';
  }).join('');

  var siBoletin = rows.filter(function(r) { return r[4] === 'yes'; }).length;

  return '<div style="font-family:Arial,sans-serif;max-width:820px;margin:0 auto;background:#fff;">' +

    '<div style="background:linear-gradient(135deg,#4a86e8,#3367d6);padding:28px 32px;border-radius:10px 10px 0 0;">' +
    '<h2 style="color:#fff;margin:0;font-size:22px;">Resumen Semanal de Ingresos</h2>' +
    '<p style="color:#c8d8f8;margin:6px 0 0;font-size:14px;">Creamos Guatemala &nbsp;&middot;&nbsp; ' + fecha + '</p>' +
    '</div>' +

    '<div style="display:flex;border:1px solid #e0e0e0;border-top:none;">' +
    '<div style="flex:1;padding:18px 24px;border-right:1px solid #e0e0e0;text-align:center;">' +
    '<div style="font-size:32px;font-weight:700;color:#4a86e8;">' + rows.length + '</div>' +
    '<div style="font-size:12px;color:#888;margin-top:2px;">Ingresos nuevos</div>' +
    '</div>' +
    '<div style="flex:1;padding:18px 24px;text-align:center;">' +
    '<div style="font-size:32px;font-weight:700;color:#2e7d32;">' + siBoletin + '</div>' +
    '<div style="font-size:12px;color:#888;margin-top:2px;">Suscritos al boletin</div>' +
    '</div>' +
    '</div>' +

    '<div style="border:1px solid #e0e0e0;border-top:none;overflow-x:auto;">' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
    '<thead><tr style="background:#f5f7ff;">' +
    '<th style="padding:10px 14px;text-align:left;color:#666;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #e0e0e0;">#</th>' +
    '<th style="padding:10px 14px;text-align:left;color:#666;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #e0e0e0;">Nombre</th>' +
    '<th style="padding:10px 14px;text-align:left;color:#666;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #e0e0e0;">Correo</th>' +
    '<th style="padding:10px 14px;text-align:left;color:#666;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #e0e0e0;">Fecha Visita</th>' +
    '<th style="padding:10px 14px;text-align:left;color:#666;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #e0e0e0;">Contacto Emergencia</th>' +
    '<th style="padding:10px 14px;text-align:left;color:#666;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #e0e0e0;">Telefono</th>' +
    '<th style="padding:10px 14px;text-align:center;color:#666;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #e0e0e0;">Boletin</th>' +
    '</tr></thead>' +
    '<tbody>' + rowsHtml + '</tbody>' +
    '</table></div>' +

    '<div style="padding:14px 24px;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 10px 10px;background:#fafafa;">' +
    '<p style="font-size:11px;color:#aaa;margin:0;">Generado automaticamente cada lunes &nbsp;&middot;&nbsp; Ver detalle en Google Sheets</p>' +
    '</div></div>';
}

// ============================================================
//  CONSTRUIR TEXTO PLANO DEL CORREO
// ============================================================
function buildWeeklySummaryPlain(rows, fecha) {
  var lines = [
    'RESUMEN SEMANAL DE INGRESOS — CREAMOS GUATEMALA',
    fecha, '',
    rows.length + ' ingreso(s) nuevo(s) esta semana:',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'
  ];
  rows.forEach(function(r, i) {
    var email = (!r[2] || r[2] === 'email not found') ? 'Sin correo' : r[2];
    lines.push('');
    lines.push((i + 1) + '. ' + (r[1] || 'Sin nombre'));
    lines.push('   Correo:     ' + email);
    lines.push('   Visita:     ' + fmtDate(r[3]));
    lines.push('   Emergencia: ' + (r[6] || '—') + '  ' + (r[7] || ''));
    lines.push('   Boletin:    ' + (r[4] === 'yes' ? 'Si' : 'No'));
  });
  lines.push('');
  lines.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('Ver detalle completo en Google Sheets.');
  return lines.join('\n');
}

// ============================================================
//  OBTENER DATOS DE KOBOTOOLBOX
// ============================================================
function fetchKoboData(config) {
  config = config || getConfig();
  try {
    var options = {
      method: 'GET',
      headers: { 'Authorization': 'Token ' + config.KOBO_API_TOKEN },
      muteHttpExceptions: true
    };
    var allResults = [];
    var url = KOBO_BASE_URL + config.KOBO_ASSET_UID + '/data/?format=json&limit=5000';

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
    Logger.log('Error de conexion: ' + e.message);
    return null;
  }
}

// ============================================================
//  ENCABEZADOS DE LA HOJA
// ============================================================
function setupHeaders(sheet) {
  var headers = [
    'ID', 'Nombre del Visitante', 'Correo del Visitante',
    'Fecha de Visita', 'Boletin (newsletter)', 'Fecha de Envio',
    'Contacto de Emergencia', 'Telefono de Emergencia',
    'Permiso de Fotos', 'Exencion de Responsabilidad', 'Estado'
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
//  IDs YA PROCESADOS
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
  sheet.appendRow([
    s['_id']                                  || '',
    s['visitor_info/visitor_name']            || '',
    s['visitor_info/visitor_email']           || '',
    s['visitor_info/visit_date']              || '',
    s['visitor_info/newsletter']              || '',
    s['_submission_time']                     || '',
    s['emergency_contact/emergency_name']     || '',
    s['emergency_contact/emergency_phone']    || '',
    s['guidelines_section/photo_permission']  || '',
    s['liability_section/liability_waiver']   || '',
    'Nuevo'
  ]);
}

// ============================================================
//  CONFIGURAR TRIGGERS (ejecutar desde el menú)
// ============================================================
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkNewSubmissions')
    .timeBased().everyHours(1).create();
  ScriptApp.newTrigger('sendWeeklySummary')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();

  SpreadsheetApp.getUi().alert(
    'Listo',
    'Triggers activados:\n• Sincronizacion: cada hora\n• Resumen por correo: cada lunes a las 8 AM',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================
//  PRUEBA DE CONEXIÓN CON KOBO
// ============================================================
function testConnection() {
  var submissions = fetchKoboData();
  var ui = SpreadsheetApp.getUi();

  if (!submissions) {
    ui.alert('Error', 'No se pudo conectar con KoboToolbox.\nRevisa tu token en Configurar credenciales.', ui.ButtonSet.OK);
    return;
  }

  var last = submissions.length > 0 ? submissions[submissions.length - 1] : null;
  var msg  = 'Conexion exitosa.\nTotal de respuestas: ' + submissions.length;
  if (last) {
    msg += '\n\nUltima respuesta:\n  Nombre: ' + (last['visitor_info/visitor_name'] || '—') +
           '\n  Fecha:  ' + (last['visitor_info/visit_date'] || '—');
  }
  ui.alert('KoboToolbox OK', msg, ui.ButtonSet.OK);
}

// ============================================================
//  PRUEBA DEL CORREO SEMANAL
// ============================================================
function testWeeklySummary() {
  sendWeeklySummary();
  SpreadsheetApp.getUi().alert('Listo', 'Si habia ingresos nuevos, el correo fue enviado.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================
//  BLOOMERANG — Crear contacto
// ============================================================
function createBloomerangContact(submission, config) {
  config = config || getConfig();
  var fullName = (submission['visitor_info/visitor_name']  || '').trim();
  var email    = (submission['visitor_info/visitor_email'] || '').trim();

  if (!email) {
    Logger.log('Bloomerang: sin correo para ' + fullName);
    return;
  }

  var parts     = fullName.split(' ');
  var firstName = parts[0] || '';
  var lastName  = parts.slice(1).join(' ') || '';

  var options = {
    method:      'POST',
    contentType: 'application/json',
    headers:     { 'X-API-KEY': config.BLOOMERANG_API_KEY },
    payload:     JSON.stringify({
      Type:         'Individual',
      FirstName:    firstName,
      LastName:     lastName,
      PrimaryEmail: { Value: email, IsPrimary: true },
      Notes:        'Ingresado via formulario de visitas Creamos. Fecha: ' +
                    (submission['visitor_info/visit_date'] || '')
    }),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch('https://api.bloomerang.co/v2/constituent', options);
    var code     = response.getResponseCode();
    if (code === 200 || code === 201) {
      Logger.log('Bloomerang: contacto creado — ' + fullName);
    } else if (code === 409) {
      Logger.log('Bloomerang: ' + email + ' ya existe.');
    } else {
      Logger.log('Bloomerang ERROR (' + code + '): ' + response.getContentText());
    }
  } catch (e) {
    Logger.log('Bloomerang: error — ' + e.message);
  }
}

// ============================================================
//  PRUEBA DE BLOOMERANG
// ============================================================
function testBloomerang() {
  var config = getConfig();
  var ui     = SpreadsheetApp.getUi();

  if (!config.BLOOMERANG_ENABLED) {
    ui.alert('Bloomerang desactivado', 'Ve a Configurar credenciales y activa la opcion de Bloomerang.', ui.ButtonSet.OK);
    return;
  }

  createBloomerangContact({
    'visitor_info/visitor_name':  'Prueba Creamos',
    'visitor_info/visitor_email': config.NOTIFICATION_EMAIL,
    'visitor_info/visit_date':    new Date().toISOString().slice(0, 10)
  }, config);

  ui.alert('Prueba enviada', 'Revisa tu Bloomerang para confirmar que el contacto fue creado.', ui.ButtonSet.OK);
}
