// ═══════════════════════════════════════════════════════════════════
// Praveen & Karthika — RSVP + Visitor Analytics Backend
// Google Apps Script Web App
// Sheet: https://docs.google.com/spreadsheets/d/1iJ4RXIxOxIA-bhHwOVN-gSjjKOCJaPlzzp2OGdYUx9U
// ═══════════════════════════════════════════════════════════════════

// ── Sheet identifiers ──────────────────────────────────────────────
var SPREADSHEET_ID  = '1iJ4RXIxOxIA-bhHwOVN-gSjjKOCJaPlzzp2OGdYUx9U';

// RSVP
var RESPONSES_SHEET = 'Responses';
var SUMMARY_SHEET   = 'Summary';

// Analytics
var VISITORS_SHEET  = 'Visitors';
var VISITS_SHEET    = 'Visits';
var EVENTS_SHEET    = 'Events';
var DASHBOARD_SHEET = 'Dashboard';

// ── Column headers ─────────────────────────────────────────────────
var RSVP_HEADERS = [
  'Timestamp', 'Name', 'Phone',
  'Engagement (12 Sep)', 'Wedding (13 Sep)',
  'Number of Guests', 'Category',
  'Transport Required', 'Room Arrangements',
  'Message', 'User Agent'
];

var VISITOR_HEADERS = [
  'Visitor ID', 'First Visit', 'Last Visit', 'Visit Count',
  'Browser', 'OS', 'Device', 'Language', 'Timezone', 'First Referrer'
];

var VISIT_HEADERS = [
  'Timestamp', 'Visitor ID', 'Page URL', 'Page Title',
  'Browser', 'OS', 'Device', 'Language', 'Timezone',
  'Screen Resolution', 'Referrer'
];

var EVENT_HEADERS = ['Timestamp', 'Visitor ID', 'Event Name'];

// ══════════════════════════════════════════════════════════════════
// SHARED UTILITIES
// ══════════════════════════════════════════════════════════════════

function sanitize(val, maxLen) {
  maxLen = maxLen || 500;
  if (val === null || val === undefined) return '';
  return String(val)
    .replace(/<[^>]*>/g, '')
    .replace(/[\x00-\x1F\x7F]/g, '')
    .trim()
    .substring(0, maxLen);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════════
// ENTRY POINTS
// ══════════════════════════════════════════════════════════════════

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return jsonOut({ success: false, error: 'Empty request body' });
  }

  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (_) {
    return jsonOut({ success: false, error: 'Invalid JSON payload' });
  }

  // Route by type
  if (data.type === 'analytics') return handleAnalytics(data);
  if (data.type === 'event')     return handleEvent(data);
  return handleRsvp(data);
}

function doGet(e) {
  return jsonOut({ status: 'ok', service: 'Praveen & Karthika API' });
}

// ══════════════════════════════════════════════════════════════════
// RSVP HANDLER
// ══════════════════════════════════════════════════════════════════

function handleRsvp(data) {
  try {
    // Honeypot
    if (data._honeypot && data._honeypot !== '') {
      return jsonOut({ success: true, message: 'Thank you!' });
    }

    var name       = sanitize(data.name, 120);
    var phone      = sanitize(data.phone, 30);
    var engagement = sanitize(data.engagement, 20);
    var wedding    = sanitize(data.wedding, 20);
    var category   = sanitize(data.category, 60);

    var errors = [];
    if (name.length < 2)                        errors.push('Name is required');
    if (phone.replace(/\D/g,'').length < 7)     errors.push('Valid phone number is required');
    if (!engagement)                            errors.push('Engagement attendance is required');
    if (!wedding)                               errors.push('Wedding attendance is required');
    if (!category)                              errors.push('Please let us know your relation');

    if (errors.length > 0) {
      return jsonOut({ success: false, error: errors.join('. ') });
    }

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(RESPONSES_SHEET);
    if (!sheet) {
      return jsonOut({ success: false, error: 'Responses sheet not found. Run setupSheets() first.' });
    }

    var guests = sanitize(data.guests, 30);
    if (guests === 'custom') {
      guests = sanitize(data.guests_custom, 10) || 'custom';
    }

    sheet.appendRow([
      new Date(),
      name, phone, engagement, wedding, guests, category,
      sanitize(data.transport, 10),
      sanitize(data.room, 10),
      sanitize(data.message, 1000),
      sanitize(data.userAgent, 300)
    ]);

    return jsonOut({ success: true, message: 'Response recorded!' });

  } catch (err) {
    Logger.log('RSVP error: ' + err.toString());
    return jsonOut({ success: false, error: 'Server error — please try again' });
  }
}

// ══════════════════════════════════════════════════════════════════
// ANALYTICS HANDLER
// ══════════════════════════════════════════════════════════════════

function handleAnalytics(data) {
  // Acquire a script-level lock to prevent race conditions on the
  // Visitors sheet (concurrent visitors could create duplicate rows)
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var visitorId  = sanitize(data.visitorId, 36);
    var pageUrl    = sanitize(data.pageUrl, 300);
    var pageTitle  = sanitize(data.pageTitle, 200);
    var browser    = sanitize(data.browser, 50);
    var os         = sanitize(data.os, 50);
    var device     = sanitize(data.device, 20);
    var language   = sanitize(data.language, 20);
    var timezone   = sanitize(data.timezone, 60);
    var resolution = sanitize(data.resolution, 20);
    var referrer   = sanitize(data.referrer, 300);

    // Basic UUID validation (must be 32+ hex chars with hyphens)
    if (!visitorId || visitorId.length < 32) {
      return jsonOut({ success: false, error: 'Invalid visitor ID' });
    }

    var ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    var visitsSheet = ss.getSheetByName(VISITS_SHEET);
    var visitorsSheet = ss.getSheetByName(VISITORS_SHEET);

    if (!visitsSheet || !visitorsSheet) {
      return jsonOut({ success: false, error: 'Analytics sheets not found. Run setupAnalyticsSheets() first.' });
    }

    var now = new Date();

    // ── 1. Visits log — always append a new row ──────────────────
    visitsSheet.appendRow([
      now, visitorId, pageUrl, pageTitle,
      browser, os, device, language, timezone, resolution, referrer
    ]);

    // ── 2. Visitors — upsert ─────────────────────────────────────
    var lastRow = visitorsSheet.getLastRow();

    if (lastRow > 1) {
      // Read visitor ID + visit count columns at once (single API call)
      // Columns: 1=VisitorID, 4=VisitCount
      var existing = visitorsSheet.getRange(2, 1, lastRow - 1, 4).getValues();
      var found    = -1;

      for (var i = 0; i < existing.length; i++) {
        if (existing[i][0] === visitorId) {
          found = i + 2; // Convert to 1-based sheet row (i=0 → row 2)
          break;
        }
      }

      if (found > 0) {
        // Existing visitor: update Last Visit + increment Visit Count
        var newCount = (existing[found - 2][3] || 0) + 1;
        visitorsSheet.getRange(found, 3, 1, 2).setValues([[now, newCount]]);
      } else {
        // New visitor
        visitorsSheet.appendRow([
          visitorId, now, now, 1,
          browser, os, device, language, timezone, referrer
        ]);
      }
    } else {
      // Sheet is empty (only header) — first visitor ever
      visitorsSheet.appendRow([
        visitorId, now, now, 1,
        browser, os, device, language, timezone, referrer
      ]);
    }

    return jsonOut({ success: true });

  } catch (err) {
    Logger.log('Analytics error: ' + err.toString());
    return jsonOut({ success: false, error: 'Analytics error' });
  } finally {
    lock.releaseLock();
  }
}

// ══════════════════════════════════════════════════════════════════
// EVENT CLICK HANDLER
// ══════════════════════════════════════════════════════════════════

function handleEvent(data) {
  try {
    var visitorId = sanitize(data.visitorId, 36);
    var eventName = sanitize(data.event, 60);

    if (!visitorId || visitorId.length < 32) {
      return jsonOut({ success: false, error: 'Invalid visitor ID' });
    }
    if (!eventName) {
      return jsonOut({ success: false, error: 'Missing event name' });
    }

    var ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    var eventsSheet = ss.getSheetByName(EVENTS_SHEET);
    if (!eventsSheet) {
      return jsonOut({ success: false, error: 'Events sheet not found. Run setupAnalyticsSheets() first.' });
    }

    eventsSheet.appendRow([new Date(), visitorId, eventName]);
    return jsonOut({ success: true });

  } catch (err) {
    Logger.log('Event error: ' + err.toString());
    return jsonOut({ success: false, error: 'Event tracking error' });
  }
}

// ══════════════════════════════════════════════════════════════════
// SETUP — run manually once from the editor
// ══════════════════════════════════════════════════════════════════

// Run this first — sets up RSVP sheets
function setupSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var responses = ss.getSheetByName(RESPONSES_SHEET) || ss.insertSheet(RESPONSES_SHEET);
  if (responses.getLastRow() === 0) {
    var h = responses.getRange(1, 1, 1, RSVP_HEADERS.length);
    h.setValues([RSVP_HEADERS]);
    h.setBackground('#4a2535').setFontColor('#f9f1e3').setFontWeight('bold').setFontSize(11);
    responses.setFrozenRows(1);
    [185,160,130,150,150,130,170,150,170,250,280].forEach(function(w,i){ responses.setColumnWidth(i+1,w); });
  }

  _buildSummarySheet(ss);
  console.log('✅ RSVP sheets ready.');
}

// Run this after setupSheets() — sets up analytics sheets
function setupAnalyticsSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // ── Visitors sheet ─────────────────────────────────────────────
  var visitors = ss.getSheetByName(VISITORS_SHEET) || ss.insertSheet(VISITORS_SHEET);
  if (visitors.getLastRow() === 0) {
    var vh = visitors.getRange(1, 1, 1, VISITOR_HEADERS.length);
    vh.setValues([VISITOR_HEADERS]);
    vh.setBackground('#1a3a2a').setFontColor('#e8f5e9').setFontWeight('bold').setFontSize(11);
    visitors.setFrozenRows(1);
    [220,160,160,100,100,100,100,100,150,200].forEach(function(w,i){ visitors.setColumnWidth(i+1,w); });
  }

  // ── Visits sheet ───────────────────────────────────────────────
  var visits = ss.getSheetByName(VISITS_SHEET) || ss.insertSheet(VISITS_SHEET);
  if (visits.getLastRow() === 0) {
    var vth = visits.getRange(1, 1, 1, VISIT_HEADERS.length);
    vth.setValues([VISIT_HEADERS]);
    vth.setBackground('#1a3a2a').setFontColor('#e8f5e9').setFontWeight('bold').setFontSize(11);
    visits.setFrozenRows(1);
    [160,220,280,200,100,100,100,100,150,120,280].forEach(function(w,i){ visits.setColumnWidth(i+1,w); });
  }

  // ── Events sheet ───────────────────────────────────────────────
  var evSheet = ss.getSheetByName(EVENTS_SHEET) || ss.insertSheet(EVENTS_SHEET);
  if (evSheet.getLastRow() === 0) {
    var evh = evSheet.getRange(1, 1, 1, EVENT_HEADERS.length);
    evh.setValues([EVENT_HEADERS]);
    evh.setBackground('#1a3a2a').setFontColor('#e8f5e9').setFontWeight('bold').setFontSize(11);
    evSheet.setFrozenRows(1);
    [160, 220, 220].forEach(function(w,i){ evSheet.setColumnWidth(i+1,w); });
  }

  // ── Dashboard sheet ────────────────────────────────────────────
  _buildDashboard(ss);
  console.log('✅ Analytics sheets ready.');
}

// Convenience: run both at once
function setupAllSheets() {
  setupSheets();
  setupAnalyticsSheets();
  console.log('✅ All sheets ready. Deploy as Web App next.');
}

// ── Private: build RSVP summary ───────────────────────────────────
function _buildSummarySheet(ss) {
  var summary = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET);
  summary.clearContents().clearFormats ? summary.clearContents() : summary.clear();
  try { summary.clearFormats(); } catch(_) {}

  var rows = [
    ['Metric', 'Value'],
    ['Total Responses',         '=COUNTA(Responses!B2:B)'],
    ['', ''],
    ['ENGAGEMENT (12 Sep)', ''],
    ['  Attending',              '=COUNTIF(Responses!D2:D,"Yes")'],
    ["  Can't make it",          '=COUNTIF(Responses!D2:D,"No")'],
    ['', ''],
    ['WEDDING (13 Sep)', ''],
    ['  Attending',              '=COUNTIF(Responses!E2:E,"Yes")'],
    ["  Can't make it",          '=COUNTIF(Responses!E2:E,"No")'],
    ['', ''],
    ['LOGISTICS', ''],
    ['  Transport Required',     '=COUNTIF(Responses!H2:H,"Yes")'],
    ['  Room Arrangements',      '=COUNTIF(Responses!I2:I,"Yes")'],
    ['  Total Guest Count',      '=SUMPRODUCT((ISNUMBER(VALUE(Responses!F2:F)))*IFERROR(VALUE(Responses!F2:F),0))'],
    ['', ''],
    ['CATEGORY', ''],
    ["  Bride's Friends/Family", '=COUNTIF(Responses!G2:G,"Bride\'s Friend/Family")'],
    ["  Groom's Friends/Family", '=COUNTIF(Responses!G2:G,"Groom\'s Friend/Family")'],
  ];

  summary.getRange(1, 1, rows.length, 2).setValues(rows);
  var hdr = summary.getRange(1, 1, 1, 2);
  hdr.setBackground('#4a2535').setFontColor('#f9f1e3').setFontWeight('bold').setFontSize(12);
  summary.setFrozenRows(1);
  summary.getRange(2, 2).setFontSize(18).setFontWeight('bold').setFontColor('#4a2535');
  [4, 8, 12, 17].forEach(function(r) {
    summary.getRange(r, 1, 1, 2).setBackground('#f5ede0').setFontWeight('bold').setFontColor('#4a2535');
  });
  summary.setColumnWidth(1, 220).setColumnWidth ? summary.setColumnWidth(2, 100) : null;
  try { summary.setColumnWidth(1, 220); summary.setColumnWidth(2, 100); } catch(_) {}
}

// ── Private: build analytics Dashboard ────────────────────────────
function _buildDashboard(ss) {
  var db = ss.getSheetByName(DASHBOARD_SHEET) || ss.insertSheet(DASHBOARD_SHEET);
  db.clear();

  // ── Section A: Key Metrics ──────────────────────────────────────
  var metrics = [
    ['METRIC', 'VALUE'],
    ['Total Unique Visitors',    '=COUNTA(Visitors!A2:A)'],
    ['Total Page Visits',        '=COUNTA(Visits!A2:A)'],
    ['Avg Visits per Visitor',   '=IFERROR(ROUND(COUNTA(Visits!A2:A)/COUNTA(Visitors!A2:A),1),"—")'],
    ['Returning Visitors',       '=COUNTIF(Visitors!D2:D,">"&1)'],
    ['New Visitors',             '=COUNTA(Visitors!A2:A)-COUNTIF(Visitors!D2:D,">"&1)'],
    ['', ''],
    ['TODAY', ''],
    ["Today's Unique Visitors",  '=SUMPRODUCT((TEXT(Visits!A2:A,"YYYY-MM-DD")=TEXT(TODAY(),"YYYY-MM-DD"))*(1/COUNTIFS(Visits!B2:B,Visits!B2:B,TEXT(Visits!A2:A,"YYYY-MM-DD"),TEXT(TODAY(),"YYYY-MM-DD"))))'],
    ["Today's Total Visits",     '=COUNTIFS(Visits!A2:A,">="&TODAY(),Visits!A2:A,"<"&TODAY()+1)'],
    ['', ''],
    ['TOP VALUES', ''],
    ['Most Visited Page',        '=IFERROR(INDEX(Visits!C2:C,MATCH(MAX(COUNTIF(Visits!C2:C,Visits!C2:C)),COUNTIF(Visits!C2:C,Visits!C2:C),0)),"—")'],
    ['Top Browser',              '=IFERROR(INDEX(Visits!E2:E,MATCH(MAX(COUNTIF(Visits!E2:E,Visits!E2:E)),COUNTIF(Visits!E2:E,Visits!E2:E),0)),"—")'],
    ['Top Device',               '=IFERROR(INDEX(Visits!G2:G,MATCH(MAX(COUNTIF(Visits!G2:G,Visits!G2:G)),COUNTIF(Visits!G2:G,Visits!G2:G),0)),"—")'],
    ['Top Language',             '=IFERROR(INDEX(Visits!H2:H,MATCH(MAX(COUNTIF(Visits!H2:H,Visits!H2:H)),COUNTIF(Visits!H2:H,Visits!H2:H),0)),"—")'],
  ];

  db.getRange(1, 1, metrics.length, 2).setValues(metrics);

  // ── Section B: Daily Visits (last 7 days) ───────────────────────
  var startRow = metrics.length + 2;
  db.getRange(startRow, 1).setValue('DAILY VISITS (last 7 days)');
  db.getRange(startRow, 1, 1, 2).setBackground('#1a3a2a').setFontColor('#e8f5e9').setFontWeight('bold');
  db.getRange(startRow + 1, 1, 1, 2).setValues([['Date', 'Visits']]).setFontWeight('bold');
  for (var i = 0; i < 7; i++) {
    var offset = i; // 0 = today, 1 = yesterday, etc.
    var dateLabel = '=TEXT(TODAY()-' + offset + ',"DD MMM")';
    var countFormula = '=COUNTIFS(Visits!A$2:A$10000,">="&(TODAY()-' + offset + '),Visits!A$2:A$10000,"<"&(TODAY()-' + offset + '+1))';
    db.getRange(startRow + 2 + i, 1, 1, 2).setValues([[dateLabel, countFormula]]);
  }

  // ── Section C: Browser breakdown ───────────────────────────────
  var bRow = startRow + 11;
  db.getRange(bRow, 1).setValue('BROWSER BREAKDOWN');
  db.getRange(bRow, 1, 1, 2).setBackground('#1a3a2a').setFontColor('#e8f5e9').setFontWeight('bold');
  db.getRange(bRow + 1, 1, 1, 2).setValues([['Browser', 'Visits']]).setFontWeight('bold');
  var browsers = ['Chrome', 'Safari', 'Firefox', 'Edge', 'Opera', 'Other'];
  browsers.forEach(function(b, i) {
    db.getRange(bRow + 2 + i, 1, 1, 2).setValues([[b, '=COUNTIF(Visits!E$2:E$10000,"' + b + '")']]);
  });

  // ── Section D: Device breakdown ────────────────────────────────
  var dRow = bRow + 10;
  db.getRange(dRow, 1).setValue('DEVICE BREAKDOWN');
  db.getRange(dRow, 1, 1, 2).setBackground('#1a3a2a').setFontColor('#e8f5e9').setFontWeight('bold');
  db.getRange(dRow + 1, 1, 1, 2).setValues([['Device', 'Visits']]).setFontWeight('bold');
  ['Desktop', 'Mobile', 'Tablet'].forEach(function(d, i) {
    db.getRange(dRow + 2 + i, 1, 1, 2).setValues([[d, '=COUNTIF(Visits!G$2:G$10000,"' + d + '")']]);
  });

  // ── Section E: Click Events ─────────────────────────────────────
  var eRow = dRow + 7;
  db.getRange(eRow, 1).setValue('CLICK EVENTS');
  db.getRange(eRow, 1, 1, 2).setBackground('#1a3a2a').setFontColor('#e8f5e9').setFontWeight('bold');
  db.getRange(eRow + 1, 1, 1, 2).setValues([['Event', 'Count']]).setFontWeight('bold');
  var eventRows = [
    ['Get Directions (hero)',         'get_directions'],
    ['Venue Map — Engagement',        'venue_map_engagement'],
    ['Venue Map — Wedding',           'venue_map_wedding'],
    ['Google Calendar — Engagement',  'calendar_google_engagement'],
    ['Apple Calendar — Engagement',   'calendar_apple_engagement'],
    ['Google Calendar — Wedding',     'calendar_google_wedding'],
    ['Apple Calendar — Wedding',      'calendar_apple_wedding'],
    ['View Map (footer)',              'footer_view_map'],
  ];
  eventRows.forEach(function(ev, i) {
    db.getRange(eRow + 2 + i, 1, 1, 2).setValues([[ev[0], '=COUNTIF(Events!C$2:C$10000,"' + ev[1] + '")']]);
  });

  // ── Styling ─────────────────────────────────────────────────────
  var hdr = db.getRange(1, 1, 1, 2);
  hdr.setBackground('#1a3a2a').setFontColor('#e8f5e9').setFontWeight('bold').setFontSize(12);
  db.setFrozenRows(1);
  db.getRange(2, 2).setFontSize(18).setFontWeight('bold');
  [8, 12].forEach(function(r) {
    db.getRange(r, 1, 1, 2).setBackground('#e8f5e9').setFontWeight('bold');
  });
  try { db.setColumnWidth(1, 240); db.setColumnWidth(2, 120); } catch(_) {}
}
