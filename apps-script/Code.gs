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
 *   - lead         : partial contact info (name, brand, email, phone, consent)
 *   - discovery    : the full discovery form
 *   - onb_request  : client asks for a private onboarding link by email
 *   - onb_google   : client signs in to onboarding with Google
 *   - onb_session  : load client + saved draft for a private link
 *   - onb_save     : autosave onboarding draft
 *   - onb_upload   : upload one onboarding file (base64)
 *   - onb_delete   : remove an uploaded onboarding file
 *   - onb_submit   : submit the onboarding form
 *
 * IMPORTANT after replacing this code: run setup() again (it adds the new
 * tabs and asks for one new permission), then Deploy > Manage deployments >
 * edit > Version: New version > Deploy (the /exec URL stays the same).
 */

var CONFIG = {
  NOTIFY_EMAIL: 'nuriroman21@gmail.com',      // where Nuri gets the alerts
  BUSINESS_NAME: 'Nuri Solutions',
  SITE_URL: 'https://www.nurisolutions.si',
  ROOT_FOLDER_NAME: 'Nuri Solutions - Clients',
  SEND_CLIENT_CONFIRMATION: true,

  // ---- Onboarding ----
  // Clients can open onboarding once their stage is this one or later.
  ONBOARDING_MIN_STAGE: 'Payment received',
  // Where the private links point to. The first one is used by default.
  ONBOARDING_URLS: [
    'https://camott.com/Nuri_Solutions/onboarding/',
    'https://www.nurisolutions.si/onboarding/',
    'https://nurisolutions.si/onboarding/'
  ],
  LINK_DAYS: 30,                 // how long a private link works
  MAX_FILE_MB: 20,               // per file
  GOOGLE_CLIENT_ID: ''           // optional: OAuth Web client ID for "Continue with Google"
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
      'Monthly budget', 'Start', 'Anything else', 'Marketing consent', 'Language', 'PDF', 'Truth to awaken people to']
  },
  ONBOARDING: { name: 'Onboarding', headers: null }, // built from ONB_FIELDS below
  SESSIONS: {
    name: 'Access links',
    headers: ['Token', 'Email', 'Lead ID', 'Created', 'Expires', 'Last used', 'Method']
  }
};

// Onboarding fields: [key, label]. Same keys as onboarding/index.html
var ONB_FIELDS = [
  ['brand_name', 'Brand / business name'], ['one_liner', 'What you do (one sentence)'], ['mission', 'Mission & values'],
  ['colors', 'Brand colors'], ['fonts', 'Fonts'],
  ['tone', 'Tone of voice'], ['voice_desc', 'How you talk'], ['phrases', 'Words & phrases you use'],
  ['avoid', 'Topics or words to avoid'], ['inspiration', 'Brands or creators you admire'],
  ['audience', 'Target audience'], ['audience_age', 'Audience age range'], ['audience_location', 'Audience location'],
  ['audience_pain', 'Audience problems & desires'],
  ['instagram', 'Instagram'], ['tiktok', 'TikTok'], ['facebook', 'Facebook'], ['youtube', 'YouTube'],
  ['linkedin', 'LinkedIn'], ['x_twitter', 'X (Twitter)'], ['google_business', 'Google Business Profile'],
  ['website', 'Website'], ['other_social', 'Other accounts'], ['access_notes', 'Account access notes'],
  ['goals_90', 'Goals for the next 90 days'], ['kpis', 'Metrics that matter'], ['frequency', 'Preferred posting frequency'],
  ['products', 'Products / services list'], ['offer_focus', 'Main offer to promote'],
  ['media_links', 'Links to photos & videos'], ['notes', 'Anything else']
];
var ONB_FILE_CATS = { logo: 'Logo', font_files: 'Fonts', brand_guide: 'Brand guide', product_files: 'Products and services', media: 'Photos and videos' };
SHEETS.ONBOARDING.headers = ['Lead ID', 'Email', 'Business / brand', 'Status', 'Started', 'Updated', 'Submitted', 'Onboarding folder', 'PDF', 'Files uploaded']
  .concat(ONB_FIELDS.map(function (f) { return f[1]; })).concat(['Draft (do not edit)']);

/* ============================ SETUP ============================ */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHEETS.LEADS, SHEETS.DISCOVERY, SHEETS.ONBOARDING, SHEETS.SESSIONS].forEach(function (def) {
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
  ss.getSheetByName(SHEETS.SESSIONS.name).hideSheet();
  rootFolder_();
  // touch UrlFetchApp so the permission is requested now (used by Google sign-in)
  try { UrlFetchApp.getRequest('https://oauth2.googleapis.com/tokeninfo'); } catch (e) {}
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
    .addSeparator()
    .addItem('Send onboarding link to selected client', 'sendOnboardingLinkToSelected')
    .addToUi();
}

/* ============================ WEB APP ============================ */

function doGet() {
  return json_({ ok: true, service: CONFIG.BUSINESS_NAME + ' backend' });
}

function doPost(e) {
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var f = data.fields || {};
    if (f.company_site) return json_({ ok: true }); // honeypot (bots)
    var a = data.action;
    if (a === 'lead' || a === 'discovery') {
      if (!data.leadId || !/^[A-Z0-9]{6,30}$/i.test(data.leadId)) return json_({ ok: false, error: 'bad id' });
      return json_(withLock_(function () { return a === 'lead' ? handleLead_(data) : handleDiscovery_(data); }));
    }
    if (a === 'onb_request') return json_(withLock_(function () { return onbRequest_(data); }));
    if (a === 'onb_google') return json_(withLock_(function () { return onbGoogle_(data); }));
    if (a === 'onb_session') return json_(onbSession_(data));
    if (a === 'onb_save') return json_(withLock_(function () { return onbSave_(data); }));
    if (a === 'onb_upload') return json_(onbUpload_(data));
    if (a === 'onb_delete') return json_(onbDelete_(data));
    if (a === 'onb_submit') return json_(withLock_(function () { return onbSubmit_(data); }));
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { try { lock.releaseLock(); } catch (x) {} }
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
    f.marketing_consent ? 'Yes' : 'No', (data.lang || 'en').toUpperCase(), pdfFile.getUrl(), f.truth].map(clean_));

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
  ['Mission, goals & challenges', [['Truth to awaken people to', 'truth'], ['Primary goals', 'goals'], ['Biggest challenges', 'challenges']]],
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

/* ============================ ONBOARDING ============================ */

// Find the client in the Tracker by email (latest row). Returns null if not found.
function clientByEmail_(email) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEADS.name);
  var row = findRowByEmail_(sh, email);
  if (!row) return null;
  return clientFromRow_(sh, row);
}
function clientByLeadId_(leadId) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEADS.name);
  var row = findRow_(sh, leadId);
  return row ? clientFromRow_(sh, row) : null;
}
function clientFromRow_(sh, row) {
  var H = SHEETS.LEADS.headers;
  var r = sh.getRange(row, 1, 1, H.length).getValues()[0];
  var leadId = String(r[0] || '');
  if (!leadId) { // client added by hand in the Tracker: give it an ID
    leadId = 'L' + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 6).toUpperCase();
    sh.getRange(row, 1).setValue(leadId);
  }
  return {
    row: row, leadId: leadId, stage: String(r[H.indexOf('Stage')] || ''),
    name: String(r[H.indexOf('Full name')] || ''), brand: String(r[H.indexOf('Business / brand')] || ''),
    email: String(r[H.indexOf('Email')] || '').trim(), lang: String(r[H.indexOf('Language')] || 'EN')
  };
}
function eligible_(c) {
  return !!c && STAGES.indexOf(c.stage) >= STAGES.indexOf(CONFIG.ONBOARDING_MIN_STAGE);
}
function onbUrl_(requested) {
  var list = CONFIG.ONBOARDING_URLS;
  for (var i = 0; i < list.length; i++) if (requested && String(requested).indexOf(list[i]) === 0) return list[i];
  return list[0];
}

function newSession_(c, method) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SESSIONS.name);
  var token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  var now = new Date();
  var exp = new Date(now.getTime() + CONFIG.LINK_DAYS * 86400000);
  sh.appendRow([token, c.email, c.leadId, now, exp, '', method]);
  return token;
}
// Validates a token and returns the client, or null.
function sessionClient_(token) {
  if (!token || !/^[a-f0-9]{64}$/i.test(token)) return null;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SESSIONS.name);
  var last = sh.getLastRow(); if (last < 2) return null;
  var vals = sh.getRange(2, 1, last - 1, 7).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== token) continue;
    if (new Date(vals[i][4]).getTime() < Date.now()) return null;
    var c = clientByLeadId_(vals[i][2]);
    if (!eligible_(c)) return null;
    sh.getRange(i + 2, 6).setValue(new Date());
    return c;
  }
  return null;
}

function sendLinkEmail_(c, url, token) {
  var es = /es/i.test(c.lang);
  var first = (c.name || '').split(' ')[0] || (es ? 'hola' : 'there');
  var link = url + '?t=' + token;
  var t = es ? {
    s: 'Tu enlace privado de onboarding | Nuri Solutions',
    h: 'Hola ' + esc_(first) + ', bienvenido a Nuri Solutions',
    p: 'Este es tu enlace privado para completar el onboarding. Ahí compartes tu logo, colores, tono de voz, redes y archivos. Tus avances se guardan solos, así que puedes volver cuando quieras.',
    b: 'Abrir mi onboarding', n: 'El enlace funciona por ' + CONFIG.LINK_DAYS + ' días. No lo compartas: es solo para ti.'
  } : {
    s: 'Your private onboarding link | Nuri Solutions',
    h: 'Hi ' + esc_(first) + ', welcome to Nuri Solutions',
    p: 'Here is your private link to complete your onboarding. You will share your logo, colors, tone of voice, social accounts and files. Your progress saves automatically, so you can come back anytime.',
    b: 'Open my onboarding', n: 'This link works for ' + CONFIG.LINK_DAYS + ' days. Please do not share it: it is just for you.'
  };
  var html = '<div style="font-family:Arial,Helvetica,sans-serif;color:#18181D;max-width:560px">' +
    '<div style="background:#0E0E11;padding:22px 24px;border-radius:14px 14px 0 0;border-bottom:4px solid #FB4698;color:#fff;font-size:20px;font-weight:bold">' + t.h + '</div>' +
    '<div style="border:1px solid #E6E1D9;border-top:0;border-radius:0 0 14px 14px;padding:20px 24px;font-size:15px;line-height:1.55">' +
    '<p>' + t.p + '</p><p style="margin:24px 0"><a href="' + link + '" style="background:#FB4698;color:#fff;padding:13px 22px;border-radius:999px;text-decoration:none;font-weight:bold">' + t.b + '</a></p>' +
    '<p style="color:#888;font-size:12px">' + t.n + '<br>' + link + '</p></div></div>';
  MailApp.sendEmail({ to: c.email, subject: t.s, htmlBody: html, body: t.p + '\n\n' + link + '\n\n' + t.n, name: CONFIG.BUSINESS_NAME, replyTo: CONFIG.NOTIFY_EMAIL });
}

/* client types their email and asks for a link */
function onbRequest_(data) {
  var email = String(data.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'email' };
  var c = clientByEmail_(email);
  if (eligible_(c)) {
    if (data.lang) c.lang = data.lang;
    sendLinkEmail_(c, onbUrl_(data.page), newSession_(c, 'Email link'));
  }
  return { ok: true }; // same answer either way (does not reveal who is a client)
}

/* client signs in with Google */
function onbGoogle_(data) {
  if (!CONFIG.GOOGLE_CLIENT_ID) return { ok: false, error: 'google disabled' };
  var res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(data.credential || ''), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return { ok: false, error: 'invalid google token' };
  var info = JSON.parse(res.getContentText());
  if (info.aud !== CONFIG.GOOGLE_CLIENT_ID || String(info.email_verified) !== 'true') return { ok: false, error: 'invalid google token' };
  var c = clientByEmail_(info.email);
  if (!eligible_(c)) return { ok: false, error: 'not_client' };
  return { ok: true, token: newSession_(c, 'Google') };
}

function onbRow_(leadId) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ONBOARDING.name);
  return { sh: sh, row: findRow_(sh, leadId) };
}
function readDraft_(leadId) {
  var o = onbRow_(leadId);
  if (!o.row) return { draft: null, status: '' };
  var H = SHEETS.ONBOARDING.headers;
  var r = o.sh.getRange(o.row, 1, 1, H.length).getValues()[0];
  var d = null; try { d = JSON.parse(r[H.length - 1] || 'null'); } catch (e) {}
  return { draft: d, status: String(r[H.indexOf('Status')] || '') };
}

function onbSession_(data) {
  var c = sessionClient_(data.token);
  if (!c) return { ok: false, error: 'expired' };
  var d = readDraft_(c.leadId);
  return { ok: true, client: { name: c.name, brand: c.brand, email: c.email }, draft: d.draft, status: d.status,
    maxMb: CONFIG.MAX_FILE_MB };
}

function writeOnb_(c, draft, submitted, pdfUrl) {
  var o = onbRow_(c.leadId);
  var H = SHEETS.ONBOARDING.headers;
  var now = new Date();
  if (!o.row) {
    var blank = H.map(function () { return ''; });
    blank[0] = c.leadId; blank[H.indexOf('Started')] = now;
    o.sh.appendRow(blank); o.row = o.sh.getLastRow();
  }
  var r = o.sh.getRange(o.row, 1, 1, H.length).getValues()[0];
  var f = (draft && draft.fields) || {};
  var files = (draft && draft.files) || {};
  var nFiles = 0; Object.keys(files).forEach(function (k) { nFiles += (files[k] || []).length; });
  r[H.indexOf('Email')] = c.email;
  r[H.indexOf('Business / brand')] = clean_(f.brand_name || c.brand);
  r[H.indexOf('Updated')] = now;
  r[H.indexOf('Files uploaded')] = nFiles;
  r[H.indexOf('Onboarding folder')] = onbFolder_(c).getUrl();
  if (submitted) {
    r[H.indexOf('Status')] = r[H.indexOf('Submitted')] ? 'Updated after submit' : 'Submitted';
    r[H.indexOf('Submitted')] = now;
    if (pdfUrl) r[H.indexOf('PDF')] = pdfUrl;
  } else if (!r[H.indexOf('Status')]) {
    r[H.indexOf('Status')] = 'In progress';
  }
  ONB_FIELDS.forEach(function (fd) { r[H.indexOf(fd[1])] = clean_(f[fd[0]] || ''); });
  var json = JSON.stringify({ fields: f, files: files });
  r[H.length - 1] = json.length > 49000 ? json.slice(0, 49000) : json;
  o.sh.getRange(o.row, 1, 1, H.length).setValues([r]);
}

function onbSave_(data) {
  var c = sessionClient_(data.token);
  if (!c) return { ok: false, error: 'expired' };
  writeOnb_(c, data.draft || {}, false);
  return { ok: true };
}

function onbFolder_(c) {
  var parent = clientFolder_(c.leadId, c.brand || c.name, c.name);
  var it = parent.getFoldersByName('Onboarding');
  return it.hasNext() ? it.next() : parent.createFolder('Onboarding');
}

function onbUpload_(data) {
  var c = withLock_(function () { return sessionClient_(data.token); });
  if (!c) return { ok: false, error: 'expired' };
  var cat = ONB_FILE_CATS[data.category];
  if (!cat) return { ok: false, error: 'bad category' };
  var bytes = Utilities.base64Decode(String(data.data || ''));
  if (!bytes.length || bytes.length > CONFIG.MAX_FILE_MB * 1048576) return { ok: false, error: 'size' };
  var base = onbFolder_(c);
  var it = base.getFoldersByName(cat);
  var folder = it.hasNext() ? it.next() : base.createFolder(cat);
  var name = String(data.name || 'file').replace(/[\\\/:*?"<>|#]/g, '-').slice(0, 120);
  var file = folder.createFile(Utilities.newBlob(bytes, String(data.type || 'application/octet-stream'), name));
  return { ok: true, id: file.getId(), name: name, size: bytes.length };
}

function onbDelete_(data) {
  var c = withLock_(function () { return sessionClient_(data.token); });
  if (!c) return { ok: false, error: 'expired' };
  try {
    var file = DriveApp.getFileById(String(data.id || ''));
    var baseId = onbFolder_(c).getId();
    var parents = file.getParents();
    while (parents.hasNext()) {
      var p = parents.next(), gp = p.getParents();
      while (gp.hasNext()) if (gp.next().getId() === baseId) { file.setTrashed(true); return { ok: true }; }
    }
  } catch (e) {}
  return { ok: false, error: 'not found' };
}

function onbSubmit_(data) {
  var c = sessionClient_(data.token);
  if (!c) return { ok: false, error: 'expired' };
  var draft = data.draft || {};
  var f = draft.fields || {};
  var now = new Date();
  var folder = onbFolder_(c);
  var html = onboardingHtml_(c, draft, now);
  var pdf = folder.createFile(Utilities.newBlob(html, 'text/html', 'onboarding.html').getAs('application/pdf')
    .setName('Onboarding - ' + safe_(f.brand_name || c.brand || c.name) + ' - ' + Utilities.formatDate(now, tz_(), 'yyyy-MM-dd') + '.pdf'));
  writeOnb_(c, draft, true, pdf.getUrl());
  upsertLead_(c.leadId, { 'Onboarding complete': now, 'Client folder': clientFolder_(c.leadId, c.brand || c.name, c.name).getUrl() }, 'Onboarding complete');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  MailApp.sendEmail({
    to: CONFIG.NOTIFY_EMAIL, replyTo: c.email, name: CONFIG.BUSINESS_NAME + ' Website',
    subject: 'Onboarding completed: ' + (f.brand_name || c.brand || c.name),
    body: (c.name || '') + ' completed the onboarding form.\n\nFiles and PDF: ' + folder.getUrl() + '\nTracker: ' + ss.getUrl(),
    htmlBody: '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px"><div style="background:#0E0E11;color:#fff;padding:20px 24px;border-radius:14px 14px 0 0;border-bottom:4px solid #FB4698">' +
      '<div style="color:#FB4698;font-size:11px;font-weight:bold;letter-spacing:2px">ONBOARDING COMPLETED</div><div style="font-size:22px;font-weight:bold;margin-top:6px">' + esc_(f.brand_name || c.brand) + '</div>' +
      '<div style="color:#ccc;font-size:14px">' + esc_(c.name) + ' &middot; ' + esc_(c.email) + '</div></div>' +
      '<div style="border:1px solid #E6E1D9;border-top:0;border-radius:0 0 14px 14px;padding:18px 24px"><p>The client finished the onboarding. The PDF summary is attached and every file is in the client folder.</p>' +
      '<p><a href="' + folder.getUrl() + '" style="background:#FB4698;color:#fff;padding:10px 16px;border-radius:999px;text-decoration:none;font-weight:bold">Open files</a> &nbsp; <a href="' + ss.getUrl() + '" style="color:#0E0E11">Open tracker</a></p></div></div>',
    attachments: [pdf.getAs('application/pdf')]
  });
  var es = data.lang === 'es';
  var first = (c.name || '').split(' ')[0];
  MailApp.sendEmail({
    to: c.email, name: CONFIG.BUSINESS_NAME, replyTo: CONFIG.NOTIFY_EMAIL,
    subject: es ? 'Onboarding recibido | Nuri Solutions' : 'Onboarding received | Nuri Solutions',
    htmlBody: clientHtml_(first, es).replace(es ? 'Recibimos tus respuestas. Nuri las revisará personalmente y te contactará pronto para agendar tu llamada de descubrimiento.' : 'We received your answers. Nuri will personally review them and reach out soon to schedule your discovery call.',
      es ? 'Recibimos tu onboarding y tus archivos. Con esto empezamos a construir tu presencia digital. Nuri te contactará con los siguientes pasos.' : 'We received your onboarding and files. This is everything we need to start building your digital presence. Nuri will reach out with the next steps.'),
    attachments: [pdf.getAs('application/pdf')]
  });
  return { ok: true };
}

function onboardingHtml_(c, draft, when) {
  var f = draft.fields || {}, files = draft.files || {};
  var groups = [
    ['Your brand', ['brand_name', 'one_liner', 'mission']],
    ['Visual identity', ['colors', 'fonts']],
    ['Voice & personality', ['tone', 'voice_desc', 'phrases', 'avoid', 'inspiration']],
    ['Target audience', ['audience', 'audience_age', 'audience_location', 'audience_pain']],
    ['Social media accounts', ['instagram', 'tiktok', 'facebook', 'youtube', 'linkedin', 'x_twitter', 'google_business', 'website', 'other_social', 'access_notes']],
    ['Goals', ['goals_90', 'kpis', 'frequency']],
    ['Products & services', ['products', 'offer_focus']],
    ['Photos & videos', ['media_links']],
    ['Anything else', ['notes']]
  ];
  var label = {}; ONB_FIELDS.forEach(function (x) { label[x[0]] = x[1]; });
  var h = '<html><body style="font-family:Arial,Helvetica,sans-serif;color:#18181D;font-size:11pt;margin:0">' +
    '<div style="background:#0E0E11;color:#fff;padding:22px 28px;border-bottom:4px solid #FB4698">' +
    '<div style="color:#FB4698;font-size:8pt;font-weight:bold;letter-spacing:2px">NURI SOLUTIONS | CLIENT ONBOARDING</div>' +
    '<div style="font-size:20pt;font-weight:bold;margin-top:6px">' + esc_(f.brand_name || c.brand || c.name) + '</div>' +
    '<div style="font-size:9pt;color:#ccc;margin-top:4px">' + esc_(c.name) + ' &middot; ' + esc_(c.email) + ' &middot; ' + Utilities.formatDate(when, tz_(), 'MMMM d, yyyy') + '</div></div><div style="padding:10px 28px">';
  groups.forEach(function (g, i) {
    var items = g[1].filter(function (k) { return f[k]; });
    h += '<div style="background:#F7F5F1;padding:6px 10px;margin:16px 0 8px;font-weight:bold;font-size:12pt">' + (i + 1) + '. ' + esc_(g[0]) + '</div>';
    if (!items.length) h += '<div style="color:#999">-</div>';
    items.forEach(function (k) {
      h += '<div style="margin:0 0 9px"><div style="color:#DE2A7C;font-size:8pt;font-weight:bold;text-transform:uppercase">' + esc_(label[k]) + '</div><div>' + esc_(f[k]) + '</div></div>';
    });
  });
  h += '<div style="background:#F7F5F1;padding:6px 10px;margin:16px 0 8px;font-weight:bold;font-size:12pt">Files uploaded</div>';
  Object.keys(ONB_FILE_CATS).forEach(function (k) {
    var list = files[k] || []; if (!list.length) return;
    h += '<div style="margin:0 0 9px"><div style="color:#DE2A7C;font-size:8pt;font-weight:bold;text-transform:uppercase">' + esc_(ONB_FILE_CATS[k]) + '</div><div>' +
      list.map(function (x) { return esc_(x.name); }).join('<br>') + '</div></div>';
  });
  return h + '</div></body></html>';
}

/* Menu: send the private link to the client selected in the Tracker */
function sendOnboardingLinkToSelected() {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEETS.LEADS.name) { ui.alert('Open the "' + SHEETS.LEADS.name + '" tab and click on the client\'s row first.'); return; }
  var row = sh.getActiveRange().getRow();
  if (row < 2) { ui.alert('Click on the client\'s row first.'); return; }
  var c = clientFromRow_(sh, row);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(c.email)) { ui.alert('This row has no valid email.'); return; }
  if (!eligible_(c)) {
    var ans = ui.alert('Stage is "' + (c.stage || 'empty') + '"', 'Onboarding opens at "' + CONFIG.ONBOARDING_MIN_STAGE + '". Mark this client as "' + CONFIG.ONBOARDING_MIN_STAGE + '" and send the link?', ui.ButtonSet.YES_NO);
    if (ans !== ui.Button.YES) return;
    upsertLead_(c.leadId, { 'Payment received': new Date() }, CONFIG.ONBOARDING_MIN_STAGE);
    c = clientFromRow_(sh, row);
  }
  sendLinkEmail_(c, CONFIG.ONBOARDING_URLS[0], newSession_(c, 'Sent by Nuri'));
  ui.alert('Onboarding link sent to ' + c.email);
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
