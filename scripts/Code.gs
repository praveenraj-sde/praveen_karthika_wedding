// ═══════════════════════════════════════════════════════════════════
// Praveen & Karthika Wedding — RSVP Backend
// Google Apps Script Web App
// Sheet: https://docs.google.com/spreadsheets/d/1iJ4RXIxOxIA-bhHwOVN-gSjjKOCJaPlzzp2OGdYUx9U
// ═══════════════════════════════════════════════════════════════════

var SPREADSHEET_ID  = '1iJ4RXIxOxIA-bhHwOVN-gSjjKOCJaPlzzp2OGdYUx9U';
var RESPONSES_SHEET = 'Responses';
var SUMMARY_SHEET   = 'Summary';

// Column headers — order matches appendRow() in doPost
var HEADERS = [
  'Timestamp',
  'Name',
  'Phone',
  'Engagement (12 Sep)',
  'Wedding (13 Sep)',
  'Number of Guests',
  'Category',
  'Transport Required',
  'Room Arrangements',
  'Message',
  'User Agent'
];

// ── Sanitize: strip HTML, limit length ────────────────────────────
function sanitize(val, maxLen) {
  maxLen = maxLen || 500;
  if (val === null || val === undefined) return '';
  return String(val)
    .replace(/<[^>]*>/g, '')
    .replace(/[\x00-\x1F\x7F]/g, '')
    .trim()
    .substring(0, maxLen);
}

// ── POST handler ──────────────────────────────────────────────────
function doPost(e) {
  try {
    // Parse body
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut({ success: false, error: 'Empty request body' });
    }

    var data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (_) {
      return jsonOut({ success: false, error: 'Invalid payload' });
    }

    // Honeypot — silent discard
    if (data._honeypot && data._honeypot !== '') {
      return jsonOut({ success: true, message: 'Thank you!' });
    }

    // Validate required fields
    var name       = sanitize(data.name, 120);
    var phone      = sanitize(data.phone, 30);
    var engagement = sanitize(data.engagement, 20);
    var wedding    = sanitize(data.wedding, 20);
    var category   = sanitize(data.category, 60);

    var errors = [];
    if (name.length < 2)          errors.push('Name is required');
    if (phone.replace(/\D/g,'').length < 7) errors.push('Valid phone number is required');
    if (!engagement)              errors.push('Engagement attendance is required');
    if (!wedding)                 errors.push('Wedding attendance is required');
    if (!category)                errors.push('Please let us know your relation');

    if (errors.length > 0) {
      return jsonOut({ success: false, error: errors.join('. ') });
    }

    // Rate-limit: same phone can submit but flag rapid duplicates (>3 in 5 min)
    // (Optional — left simple for now; same person can submit multiple times as requested)

    // Write to sheet
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(RESPONSES_SHEET);

    if (!sheet) {
      return jsonOut({ success: false, error: 'Responses sheet not found. Please run setupSheets() first.' });
    }

    var guestsRaw = sanitize(data.guests, 30);
    if (guestsRaw === 'custom') {
      guestsRaw = sanitize(data.guests_custom, 10) || 'custom';
    }

    sheet.appendRow([
      new Date(),
      name,
      phone,
      engagement,
      wedding,
      guestsRaw,
      category,
      sanitize(data.transport, 10),
      sanitize(data.room, 10),
      sanitize(data.message, 1000),
      sanitize(data.userAgent, 300)
    ]);

    return jsonOut({ success: true, message: 'Response recorded!' });

  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return jsonOut({ success: false, error: 'Server error — please try again' });
  }
}

// ── GET handler (health check / CORS preflight fallback) ──────────
function doGet(e) {
  return jsonOut({ status: 'ok', service: 'Praveen & Karthika RSVP API' });
}

// ── JSON output helper ────────────────────────────────────────────
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════════
// ONE-TIME SETUP — run this manually once from the editor
// Creates headers on Responses sheet + sets up Summary sheet formulas
// ══════════════════════════════════════════════════════════════════
function setupSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // ── Responses sheet ────────────────────────────────────────────
  var responses = ss.getSheetByName(RESPONSES_SHEET);
  if (!responses) {
    responses = ss.insertSheet(RESPONSES_SHEET);
  }

  if (responses.getLastRow() === 0) {
    var hRange = responses.getRange(1, 1, 1, HEADERS.length);
    hRange.setValues([HEADERS]);
    hRange.setBackground('#4a2535');
    hRange.setFontColor('#f9f1e3');
    hRange.setFontWeight('bold');
    hRange.setFontSize(11);
    responses.setFrozenRows(1);

    // Column widths
    responses.setColumnWidth(1, 185);   // Timestamp
    responses.setColumnWidth(2, 160);   // Name
    responses.setColumnWidth(3, 130);   // Phone
    responses.setColumnWidth(4, 150);   // Engagement
    responses.setColumnWidth(5, 150);   // Wedding
    responses.setColumnWidth(6, 130);   // Guests
    responses.setColumnWidth(7, 170);   // Category
    responses.setColumnWidth(8, 150);   // Transport
    responses.setColumnWidth(9, 170);   // Room
    responses.setColumnWidth(10, 250);  // Message
    responses.setColumnWidth(11, 280);  // User Agent
  }

  // ── Summary sheet ──────────────────────────────────────────────
  var summary = ss.getSheetByName(SUMMARY_SHEET);
  if (!summary) {
    summary = ss.insertSheet(SUMMARY_SHEET);
  }
  summary.clearContents();
  summary.clearFormats();

  var rows = [
    ['Metric',                     'Value'],
    ['Total Responses',            '=COUNTA(Responses!B2:B)'],
    ['',                           ''],
    ['ENGAGEMENT (12 Sep)',         ''],
    ['  Attending',                '=COUNTIF(Responses!D2:D,"Yes")'],
    ["  Can't make it",            '=COUNTIF(Responses!D2:D,"No")'],
    ['',                           ''],
    ['WEDDING (13 Sep)',            ''],
    ['  Attending',                '=COUNTIF(Responses!E2:E,"Yes")'],
    ["  Can't make it",            '=COUNTIF(Responses!E2:E,"No")'],
    ['',                           ''],
    ['LOGISTICS',                  ''],
    ['  Transport Required',       '=COUNTIF(Responses!H2:H,"Yes")'],
    ['  Room Arrangements',        '=COUNTIF(Responses!I2:I,"Yes")'],
    ['  Total Guest Count',        '=SUMPRODUCT((ISNUMBER(VALUE(Responses!F2:F)))*IFERROR(VALUE(Responses!F2:F),0))'],
    ['',                           ''],
    ['CATEGORY',                   ''],
    ["  Bride's Friends/Family",   '=COUNTIF(Responses!G2:G,"Bride\'s Friend/Family")'],
    ["  Groom's Friends/Family",   '=COUNTIF(Responses!G2:G,"Groom\'s Friend/Family")'],
  ];

  var dataRange = summary.getRange(1, 1, rows.length, 2);
  dataRange.setValues(rows);

  // Header row styling
  var hdr = summary.getRange(1, 1, 1, 2);
  hdr.setBackground('#4a2535');
  hdr.setFontColor('#f9f1e3');
  hdr.setFontWeight('bold');
  hdr.setFontSize(12);
  summary.setFrozenRows(1);

  // Total Responses — bigger
  summary.getRange(2, 2).setFontSize(18).setFontWeight('bold').setFontColor('#4a2535');

  // Section headers — bold, coloured background
  [4, 8, 12, 17].forEach(function(r) {
    var row = summary.getRange(r, 1, 1, 2);
    row.setBackground('#f5ede0');
    row.setFontWeight('bold');
    row.setFontColor('#4a2535');
    row.setFontSize(10);
  });

  summary.setColumnWidth(1, 220);
  summary.setColumnWidth(2, 100);

  SpreadsheetApp.getUi().alert(
    '✅ Setup complete!\n\n' +
    '• "Responses" sheet — ready to receive submissions\n' +
    '• "Summary" sheet — auto-updates as responses come in\n\n' +
    'Deploy this script as a Web App next.'
  );
}
