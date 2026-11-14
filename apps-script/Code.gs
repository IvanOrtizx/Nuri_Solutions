/**
 * NURI SOLUTIONS — Website backend (Google Apps Script)
 * ------------------------------------------------------------------
 * Create this script FROM the Google Sheet (Extensions > Apps Script)
 * in the Google account dedicated to the website. Then:
 *   1. Fill in CONFIG below.
 *   2. Run setup() once (authorize when asked).
 *   3. Deploy > New deployment > Web app
 *        Execute as: Me   |   Who has access: Anyone
 *   4. Paste the /exec URL into NURI_CONFIG.endpoint in index.html.
 *
 * Actions received from the website (POST, JSON as text/plain):
 *   - lead       : partial contact info (name, brand, email, phone, consent)
 *   - discovery  : the full discovery form
 * Later phases (contract, payment, onboarding) will add more actions here.
 */

var CONFIG = {
  NOTIFY_EMAIL: 'nuriroman21@gmail.com',      // where Nuri gets the alerts
  BUSINESS_NAME: 'Nuri Solutions',
  SITE_URL: 'https://www.nurisolutions.si',
  ROOT_FOLDER_NAME: 'Nuri Solutions - Clients',
  SEND_CLIENT_CONFIRMATION: true
};

// Pipeline stages (order matters: a lead never moves backwards automatically)
var STAGES = [
  'Lead (form started)',
  'Discovery form',
  'Call',
  'Contract sent',
  'Contract signed',
  'Payment received',
  'Onboarding complete'
];

var SHEETS = {
  LEADS: {
    name: 'Tracker',
    headers: ['Lead ID', 'Created', 'Updated', 'Stage', 'Full name', 'Business / brand', 'Email', 'Phone',
      'Marketing consent', 'Language', 'Discovery submitted', 'Call date', 'Contract sent', 'Contract signed',
      'Payment received', 'Onboarding complete', 'Client folder', 'Notes', 'Source page']
  },
  DISCOVERY: {
    name: 'Discovery forms',
    headers: ['Submitted', 'Lead ID', 'Full name', 'Business / brand', 'Email', 'Phone', 'Website', 'Social media profiles',
      'Industry', 'Brand type', 'Services of interest', 'Other services', 'Primary goals', 'Biggest challenges',
      'Monthly budget', 'Start', 'Anything else', 'Marketing consent', 'Language', 'PDF']
  }
};

/* ============================ SETUP ============================ */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHEETS.LEADS, SHEETS.DISCOVERY].forEach(function (def) {
    var sh = ss.getSheetByName(def.name) || ss.insertSheet(def.name);
    sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers])
      .setFontWeight('bold').setBackground('#0E0E11').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.setFrozenColumns(1);
  });
  var tracker = ss.getSheetByName(SHEETS.LEADS.name);
  var stageCol = col_(SHEETS.LEADS, 'Stage');
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(STAGES, true).setAllowInvalid(false).build();
  tracker.getRange(2, stageCol, tracker.getMaxRows() - 1, 1).setDataValidation(rule);
  addStageColors_(tracker, stageCol);
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja 1');
  if (def && ss.getSheets().length > 2) ss.deleteSheet(def);
  rootFolder_();
  Logger.log('Setup complete. Now deploy as Web app.');
}

function addStageColors_(sh, stageCol) {
  var colors = { 'Lead (form started)': '#F3F1EC', 'Discovery form': '#FFE6F1', 'Call': '#FFD3E6',
    'Contract sent': '#FFF1D6', 'Contract signed': '#FFE3A8', 'Payment received': '#D7F2E3', 'Onboarding complete': '#B6E8CC' };
  var range = sh.getRange(2, stageCol, sh.getMaxRows() - 1, 1);
  var rules = Object.keys(colors).map(function (s) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s).setBackground(colors[s]).setRanges([range]).build();
  });
  sh.setConditionalFormatRules(rules);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Nuri Solutions')
    .addItem('Export leads with marketing consent (CSV)', 'exportConsentedLeads')
    .addItem('Export all leads (CSV)', 'exportAllLeads')
    .addToUi();
}

/* ============================ WEB APP ============================ */

function doGet() {
  return json_({ ok: true, service: CONFIG.BUSINESS_NAME + ' backend' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!data.leadId || !/^[A-Z0-9]{6,30}$/i.test(data.leadId)) return json_({ ok: false, error: 'bad id' });
    var f = data.fields || {};
    if (f.company_site) return json_({ ok: true }); // honeypot (bots)
    if (data.action === 'lead') return json_(handleLead_(data));
    if (data.action === 'discovery') return json_(handleDiscovery_(data));
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------- partial lead ---------- */
function handleLead_(data) {
  var f = data.fields || {};
  if (!f.email || !f.full_name) return { ok: false, error: 'missing contact' };
  upsertLead_(data.leadId, {
    'Full name': f.full_name, 'Business / brand': f.brand_name, 'Email': f.email, 'Phone': f.phone,
    'Marketing consent': f.marketing_consent ? 'Yes' : 'No', 'Language': (data.lang || 'en').toUpperCase(),
    'Source page': data.page || ''
  }, STAGES[0]);
  return { ok: true };
}

/* ---------- full discovery form ---------- */
function handleDiscovery_(data) {
  var f = data.fields || {};
  if (!f.full_name || !f.email || !f.brand_name || !f.phone) return { ok: false, error: 'missing required fields' };
  var now = new Date();
  var who = (f.brand_name || f.full_name);

  // Client folder + PDF copy in Drive
  var folder = clientFolder_(data.leadId, who, f.full_name);
  var html = discoveryHtml_(f, data.text, now);
  var pdf = Utilities.newBlob(html, 'text/html', 'discovery.html').getAs('application/pdf')
    .setName('Discovery form - ' + safe_(who) + ' - ' + Utilities.formatDate(now, tz_(), 'yyyy-MM-dd') + '.pdf');
  var pdfFile = folder.createFile(pdf);

  // Sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dsh = ss.getSheetByName(SHEETS.DISCOVERY.name);
  dsh.appendRow([now, data.leadId, f.full_name, f.brand_name, f.email, f.phone, f.website, f.socials, f.industry,
    f.brand_type, f.services, f.services_other, f.goals, f.challenges, f.budget, f.start, f.notes,
    f.marketing_consent ? 'Yes' : 'No', (data.lang || 'en').toUpperCase(), pdfFile.getUrl()].map(clean_));

  upsertLead_(data.leadId, {
    'Full name': f.full_name, 'Business / brand': f.brand_name, 'Email': f.email, 'Phone': f.phone,
    'Marketing consent': f.marketing_consent ? 'Yes' : 'No', 'Language': (data.lang || 'en').toUpperCase(),
    'Discovery submitted': now, 'Client folder': folder.getUrl(), 'Source page': data.page || ''
  }, STAGES[1]);

  // Email to Nuri
  var subject = 'New discovery form: ' + who + ' (' + f.full_name + ')';
  var body = 'A new discovery form was submitted on your website.\n\n' + (data.text || '') +
    '\n\nClient folder: ' + folder.getUrl() + '\nTracker: ' + ss.getUrl();
  MailApp.sendEmail({ to: CONFIG.NOTIFY_EMAIL, subject: subject, body: body, htmlBody: notifyHtml_(f, folder.getUrl(), ss.getUrl()),
    attachments: [pdfFile.getAs('application/pdf')], replyTo: f.email, name: CONFIG.BUSINESS_NAME + ' Website' });

  // Confirmation to the client (in their language)
  if (CONFIG.SEND_CLIENT_CONFIRMATION && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(f.email)) {
    var es = data.lang === 'es';
    var first = String(f.full_name).split(' ')[0];
    MailApp.sendEmail({
      to: f.email,
      subject: es ? 'Recibimos tus respuestas | Nuri Solutions' : 'We received your answers | Nuri Solutions',
      htmlBody: clientHtml_(first, es),
      body: es ? ('Hola ' + first + ', gracias por llenar el formulario. Nuri revisará tus respuestas y te contactará pronto para agendar tu llamada de descubrimiento.')
               : ('Hi ' + first + ', thank you for filling out the form. Nuri will review your answers and reach out soon to schedule your discovery call.'),
      name: CONFIG.BUSINESS_NAME, replyTo: CONFIG.NOTIFY_EMAIL,
      attachments: [pdfFile.getAs('application/pdf')]
    });
  }
  return { ok: true };
}

/* ============================ HELPERS ============================ */

function upsertLead_(leadId, values, stage) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.LEADS.name);
  var H = SHEETS.LEADS.headers;
  var now = new Date();
  var row = findRow_(sh, leadId);
  if (!row) {
    // same email already in the tracker? reuse that row so one person = one row
    row = findRowByEmail_(sh, values['Email']);
  }
  if (!row) {
    var arr = H.map(function () { return ''; });
    arr[0] = leadId; arr[H.indexOf('Created')] = now;
    sh.appendRow(arr);
    row = sh.getLastRow();
  }
  var current = sh.getRange(row, 1, 1, H.length).getValues()[0];
  values['Updated'] = now;
  var curStage = current[H.indexOf('Stage')];
  if (STAGES.indexOf(stage) > STAGES.indexOf(curStage)) values['Stage'] = stage;
  Object.keys(values).forEach(function (k) {
    var i = H.indexOf(k);
    if (i < 0) return;
    var v = values[k];
    if (v === '' || v == null) return; // never erase existing data with blanks
    current[i] = (v instanceof Date) ? v : clean_(v);
  });
  sh.getRange(row, 1, 1, H.length).setValues([current]);
}

function findRow_(sh, leadId) {
  var last = sh.getLastRow(); if (last < 2) return 0;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(leadId)) return i + 2;
  return 0;
}
function findRowByEmail_(sh, email) {
  if (!email) return 0;
  var last = sh.getLastRow(); if (last < 2) return 0;
  var c = col_(SHEETS.LEADS, 'Email');
  var vals = sh.getRange(2, c, last - 1, 1).getValues();
  var target = String(email).trim().toLowerCase();
  for (var i = vals.length - 1; i >= 0; i--) if (String(vals[i][0]).trim().toLowerCase() === target) return i + 2;
  return 0;
}

function rootFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('ROOT_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var it = DriveApp.getFoldersByName(CONFIG.ROOT_FOLDER_NAME);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME);
  props.setProperty('ROOT_FOLDER_ID', folder.getId());
  return folder;
}
function clientFolder_(leadId, brand, name) {
  var props = PropertiesService.getScriptProperties();
  var key = 'FOLDER_' + leadId;
  var id = props.getProperty(key);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var folder = rootFolder_().createFolder(safe_(brand) + ' - ' + safe_(name) + ' (' + leadId + ')');
  props.setProperty(key, folder.getId());
  return folder;
}

function col_(def, header) { return def.headers.indexOf(header) + 1; }
function tz_() { return Session.getScriptTimeZone() || 'America/Los_Angeles'; }
function safe_(s) { return String(s || '').replace(/[\\\/:*?"<>|#]/g, '-').slice(0, 80).trim(); }
function clean_(v) {
  if (v === true) return 'Yes'; if (v === false) return 'No';
  var s = String(v == null ? '' : v).slice(0, 5000);
  return /^[=+\-@]/.test(s) ? "'" + s : s; // block formula injection
}
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>'); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- HTML for PDF and emails ---------- */
var LABELS = [
  ['About you', [['Full name', 'full_name'], ['Business / brand', 'brand_name'], ['Email', 'email'], ['Phone', 'phone'], ['Marketing consent', 'marketing_consent']]],
  ['Your business', [['Website', 'website'], ['Industry', 'industry'], ['Social media profiles', 'socials'], ['Brand type', 'brand_type']]],
  ['Services', [['Services of interest', 'services'], ['Other services', 'services_other']]],
  ['Goals & challenges', [['Primary goals', 'goals'], ['Biggest challenges', 'challenges']]],
  ['Budget & timing', [['Monthly budget', 'budget'], ['Start', 'start'], ['Anything else', 'notes']]]
];

function discoveryHtml_(f, text, when) {
  var h = '<html><body style="font-family:Arial,Helvetica,sans-serif;color:#18181D;font-size:11pt;margin:0">' +
    '<div style="background:#0E0E11;color:#fff;padding:22px 28px;border-bottom:4px solid #FB4698">' +
    '<div style="color:#FB4698;font-size:8pt;font-weight:bold;letter-spacing:2px">NURI SOLUTIONS | DISCOVERY FORM</div>' +
    '<div style="font-size:20pt;font-weight:bold;margin-top:6px">' + esc_(f.brand_name || f.full_name) + '</div>' +
    '<div style="font-size:9pt;color:#ccc;margin-top:4px">Submitted ' + Utilities.formatDate(when, tz_(), 'MMMM d, yyyy h:mm a') + '</div></div>' +
    '<div style="padding:10px 28px">';
  LABELS.forEach(function (sec, i) {
    h += '<div style="background:#F7F5F1;padding:6px 10px;margin:16px 0 8px;font-weight:bold;font-size:12pt">' + (i + 1) + '. ' + esc_(sec[0]) + '</div>';
    sec[1].forEach(function (p) {
      var v = f[p[1]]; if (p[1] === 'marketing_consent') v = v ? 'Yes' : 'No';
      if (v === '' || v == null) return;
      h += '<div style="margin:0 0 9px"><div style="color:#DE2A7C;font-size:8pt;font-weight:bold;text-transform:uppercase">' + esc_(p[0]) + '</div><div>' + esc_(v) + '</div></div>';
    });
  });
  return h + '</div></body></html>';
}

function notifyHtml_(f, folderUrl, sheetUrl) {
  var h = '<div style="font-family:Arial,Helvetica,sans-serif;color:#18181D;max-width:620px">' +
    '<div style="background:#0E0E11;color:#fff;padding:20px 24px;border-radius:14px 14px 0 0;border-bottom:4px solid #FB4698">' +
    '<div style="color:#FB4698;font-size:11px;font-weight:bold;letter-spacing:2px">NEW DISCOVERY FORM</div>' +
    '<div style="font-size:22px;font-weight:bold;margin-top:6px">' + esc_(f.brand_name) + '</div>' +
    '<div style="color:#ccc;font-size:14px">' + esc_(f.full_name) + ' &middot; ' + esc_(f.email) + ' &middot; ' + esc_(f.phone) + '</div></div>' +
    '<div style="border:1px solid #E6E1D9;border-top:0;border-radius:0 0 14px 14px;padding:8px 24px 20px">';
  LABELS.slice(1).forEach(function (sec) {
    sec[1].forEach(function (p) {
      var v = f[p[1]]; if (!v) return;
      h += '<p style="margin:14px 0 0"><b style="color:#DE2A7C;font-size:11px;text-transform:uppercase">' + esc_(p[0]) + '</b><br>' + esc_(v) + '</p>';
    });
  });
  h += '<p style="margin:14px 0 0"><b style="color:#DE2A7C;font-size:11px;text-transform:uppercase">Marketing consent</b><br>' + (f.marketing_consent ? 'Yes' : 'No') + '</p>';
  h += '<p style="margin-top:22px"><a href="' + folderUrl + '" style="background:#FB4698;color:#fff;padding:10px 16px;border-radius:999px;text-decoration:none;font-weight:bold">Open client folder</a> &nbsp; ' +
    '<a href="' + sheetUrl + '" style="color:#0E0E11">Open tracker</a></p><p style="color:#888;font-size:12px">Reply to this email to answer the client directly.</p></div></div>';
  return h;
}

function clientHtml_(first, es) {
  var t = es ? {
    h: '¡Gracias, ' + esc_(first) + '!',
    p1: 'Recibimos tus respuestas. Nuri las revisará personalmente y te contactará pronto para agendar tu llamada de descubrimiento.',
    p2: 'Adjuntamos una copia en PDF de lo que compartiste.',
    s: 'Con cariño,<br>Nuri Roman<br>Nuri Solutions'
  } : {
    h: 'Thank you, ' + esc_(first) + '!',
    p1: 'We received your answers. Nuri will personally review them and reach out soon to schedule your discovery call.',
    p2: 'A PDF copy of what you shared is attached.',
    s: 'Warmly,<br>Nuri Roman<br>Nuri Solutions'
  };
  return '<div style="font-family:Arial,Helvetica,sans-serif;color:#18181D;max-width:560px">' +
    '<div style="background:#0E0E11;padding:22px 24px;border-radius:14px 14px 0 0;border-bottom:4px solid #FB4698;color:#fff;font-size:20px;font-weight:bold">' + t.h + '</div>' +
    '<div style="border:1px solid #E6E1D9;border-top:0;border-radius:0 0 14px 14px;padding:18px 24px;font-size:15px;line-height:1.55">' +
    '<p>' + t.p1 + '</p><p>' + t.p2 + '</p><p>' + t.s + '</p>' +
    '<p style="color:#888;font-size:12px"><a href="' + CONFIG.SITE_URL + '" style="color:#DE2A7C">' + CONFIG.SITE_URL.replace(/^https?:\/\//, '') + '</a></p></div></div>';
}

/* ============================ EXPORTS ============================ */

function exportConsentedLeads() { exportLeads_(true); }
function exportAllLeads() { exportLeads_(false); }
function exportLeads_(onlyConsent) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEADS.name);
  var H = SHEETS.LEADS.headers;
  var data = sh.getDataRange().getValues().slice(1);
  var ci = H.indexOf('Marketing consent');
  var rows = data.filter(function (r) { return r[0] && (!onlyConsent || r[ci] === 'Yes'); });
  // Format ready for Meta / Google custom audiences
  var out = [['email', 'phone', 'fn', 'ln', 'brand', 'stage', 'consent']];
  rows.forEach(function (r) {
    var parts = String(r[H.indexOf('Full name')]).trim().split(/\s+/);
    out.push([r[H.indexOf('Email')], r[H.indexOf('Phone')], parts[0] || '', parts.slice(1).join(' '),
      r[H.indexOf('Business / brand')], r[H.indexOf('Stage')], r[ci]]);
  });
  var csv = out.map(function (r) { return r.map(function (c) { return '"' + String(c).replace(/"/g, '""') + '"'; }).join(','); }).join('\n');
  var name = (onlyConsent ? 'Leads with consent' : 'All leads') + ' - ' + Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd') + '.csv';
  var file = rootFolder_().createFile(name, csv, MimeType.CSV);
  SpreadsheetApp.getUi().alert('Exported ' + rows.length + ' leads.\n\nFile saved in Drive:\n' + file.getUrl());
}
