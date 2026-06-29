// ════════════════════════════════════════════════════════════
//  IBI Settlement Tracker — Google Apps Script Backend
//  Paste into script.google.com → New Project → Save
//  Deploy → New Deployment → Web App
//    Execute as  : Me
//    Who can access : Anyone   ← IMPORTANT (not "Anyone with Google account")
//  Copy the Web App URL into the tool's GAS Endpoint field
//
//  UPDATING an existing deployment? Replace ONLY the doPost(e) function
//  below (that keeps your SHEET_ID intact), then:
//    Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy
// ════════════════════════════════════════════════════════════

const SHEET_ID  = 'YOUR_GOOGLE_SHEET_ID_HERE';  // ← your Settlements Sheet ID (leave your existing one if updating)
const SHEET_TAB = 'Settlements';

// ── Package Tracker sheet — read to fill proper product names by Order ID.
//    Runs as you, so only Order ID → Product name leaves the sheet (no buyer PII).
const TRACKER_SHEET_ID = '1VjK5oA6mCZVXZ2AhfZYtb0kVAkB549bSUB4iQ4eW7f0';
const TRACKER_TAB      = 'Package Tracker';

const HEADERS = [
  'Timestamp','Platform','Order ID','Order Date','Product',
  'SKU','Qty','Selling Price (₹)','Courier Charges (₹)','All Fees (₹)','Net Settlement (₹)',
  'Settlement Date','Status','Days Since Order','Remarks'
];

// ── CORS helper — must wrap every response ───────────────────
function corsResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Handle POST (data export from the web tool) ──────────────
function doPost(e) {
  try {
    // Works whether browser sends Content-Type: text/plain or application/json
    const raw     = e.postData ? e.postData.contents : '{}';
    const payload = JSON.parse(raw);

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let   sheet = ss.getSheetByName(SHEET_TAB);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_TAB);
      const hdrRange = sheet.getRange(1, 1, 1, HEADERS.length);
      hdrRange.setValues([HEADERS]);
      hdrRange.setBackground('#0D1B2A')
              .setFontColor('#00D4F0')
              .setFontWeight('bold')
              .setFontSize(10);
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(3, 160);   // Order ID
      sheet.setColumnWidth(5, 200);   // Product
    }

    const rows = payload.rows || [];
    const full = payload.full === true || payload.mode === 'replace';
    const ts   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss');

    const data = rows.map(r => [
      ts,
      r.platform              || '',
      r.orderId               || '',
      r.orderDate             || '',
      r.product               || '',
      r.sku                   || '',
      r.qty                   || '',
      Number(r.sellingPrice)  || 0,
      Number(r.courierCharge) || 0,
      Number(r.fees)          || 0,
      Number(r.netSettlement) || 0,
      r.settlementDate        || '',
      r.status                || '',
      r.daysSinceOrder        || '',
      r.remarks               || ''
    ]);

    // Full export → REPLACE: wipe existing data rows (keep the header) so the sheet
    // MIRRORS the current Tracker data instead of accumulating duplicate / stale rows.
    // A filtered (partial) export leaves full=false and still appends.
    if (full) {
      const last = sheet.getLastRow();
      if (last > 1) sheet.getRange(2, 1, last - 1, HEADERS.length).clearContent();
    }

    if (data.length > 0) {
      const lastRow = sheet.getLastRow();
      // Force the date columns to PLAIN TEXT before writing so Google Sheets can't
      // silently re-read a day-first "DD-MM-YYYY" string as a US MM-DD date.
      sheet.getRange(lastRow + 1, 4,  data.length, 1).setNumberFormat('@');   // Order Date
      sheet.getRange(lastRow + 1, 12, data.length, 1).setNumberFormat('@');   // Settlement Date
      sheet.getRange(lastRow + 1, 1, data.length, HEADERS.length).setValues(data);

      // Colour-code the Status column (col 13)
      data.forEach((row, i) => {
        const cell   = sheet.getRange(lastRow + 1 + i, 13);
        const status = row[12];
        if (status === 'Settled')   { cell.setBackground('#d4edda'); cell.setFontColor('#155724'); }
        if (status === 'Pending')   { cell.setBackground('#fff3cd'); cell.setFontColor('#856404'); }
        if (status === 'Returned')  { cell.setBackground('#f8d7da'); cell.setFontColor('#721c24'); }
        if (status === 'Cancelled') { cell.setBackground('#e2e3e5'); cell.setFontColor('#383d41'); }

        // Bold net settlement column
        sheet.getRange(lastRow + 1 + i, 10).setFontWeight('bold');
      });

      // Auto-resize amount columns
      [8,9,10].forEach(col => sheet.autoResizeColumn(col));
    }

    return corsResponse({ ok: true, rows: data.length, ts: ts });

  } catch(err) {
    return corsResponse({ ok: false, error: err.message });
  }
}

// ── Handle GET — health-check, or JSONP product-name lookup ──
function doGet(e) {
  const action   = e && e.parameter ? e.parameter.action : '';
  const callback = e && e.parameter ? e.parameter.callback : '';
  if (action === 'getTrackerNames') {
    const out = JSON.stringify(getTrackerNames_());
    // JSONP (a <script> tag read) — bypasses the GAS no-CORS limit for reads
    if (callback) return ContentService.createTextOutput(callback + '(' + out + ')')
                                       .setMimeType(ContentService.MimeType.JAVASCRIPT);
    return corsResponse(getTrackerNames_());
  }
  return corsResponse({ ok: true, msg: 'IBI Settlement Tracker GAS is live', ts: new Date().toISOString() });
}

// ── Order ID → Product name, read from the Package Tracker sheet ──────
function getTrackerNames_() {
  try {
    const ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
    const sh = ss.getSheetByName(TRACKER_TAB) || ss.getSheets()[0];
    if (!sh) return { ok: true, names: {} };
    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) return { ok: true, names: {} };
    const hdr = vals[0].map(function(h){ return String(h||'').toLowerCase().replace(/[\s_\-\/\\\.,\(\)\[\]:#&₹]/g,''); });
    function find(aliases){ for (var c=0;c<hdr.length;c++){ for (var a=0;a<aliases.length;a++){ if (hdr[c]===aliases[a]) return c; } } return -1; }
    const oi = find(['orderid','orderno','ordernumber']);
    const pi = find(['productssku','products','product','itemname','items']);
    if (oi < 0 || pi < 0) return { ok: false, error: 'Order ID / Products column not found in Package Tracker' };
    const names = {};
    for (var i=1;i<vals.length;i++){
      var id = String(vals[i][oi]||'').trim();
      var nm = String(vals[i][pi]||'').trim().replace(/\s*\|\s*SKU\s*:.*$/i,'').trim();
      if (id && nm) names[id] = nm;
    }
    return { ok: true, names: names, total: Object.keys(names).length };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}
