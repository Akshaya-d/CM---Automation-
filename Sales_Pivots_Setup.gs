/**
 * Sales Report pivots. Paste into the Sales Report sheet's OWN Apps Script
 * project, NOT CL Payments'. Separate on purpose — CL Payments just exports
 * rows into Bookings as part of its own Update Now; this one only touches
 * its own sheet, no cross-file spreadsheet ID juggling.
 *
 * Setup: Extensions > Apps Script on the Sales Report sheet, paste, save,
 * reload — adds a "Sales Pivots" menu. Click "Build/Rebuild Pivots" once
 * after CL Payments exports fresh data. Pivots are live native Sheets
 * pivots so they update on their own after that — only re-run to reset
 * layout (safe any time, just clears + rebuilds each tab).
 *
 * 11 tabs, each live native pivots: Revenue Trends, Accommodation
 * Performance, Guest & Occupancy Mix, Pet-Friendly Stays, Cancellations &
 * Refunds, Promo & Discount Effectiveness, Booking Lead Time, Segment &
 * Group Mix, Booking vs Travel Trends, Daily Booking Trend, OTA
 * Accommodation Mix.
 *
 * V2: currency/int formatting on Bookings numeric cols so pivots inherit
 * it, calculated KPIs per tab (Net Revenue After Refunds, Avg Rev/Night,
 * Refund Rate %, Discount Rate %, Avg Booking Value, Avg Lead Time, % of
 * Total Revenue), one chart-source mini pivot per tab feeding a native
 * chart, frozen header row/col + auto-sized cols, charts/pivots explicitly
 * cleared on rebuild (sh.clear() doesn't remove either on its own).
 *
 * Assumes Bookings is the 44-col layout exportToSales() writes — 39
 * original cols (A-AM) + Refund Amount (AN) + Refund Type (AO) + 3 hidden
 * pivot-helper cols (AP category, AQ lead time days, AR promo used). Update
 * BOOKINGS_COL below if that layout ever changes.
 */

// CUSTOM MENU

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sales Pivots')
    .addItem('📊 Build/Rebuild Pivots', 'buildSalesPivots')
    .addToUi();
}

// Expected header text, in column order — used to self-heal the header row
// before building pivots. A pivot table can't be built at all if its source
// header row contains a broken/error cell (e.g. "#REF!") anywhere — it
// poisons field detection for the WHOLE pivot, not just that one column,
// which is why a single bad header cell can make every pivot fail
// identically. Cheaper to just fix it than to explain it.
var SALES_BOOKINGS_HEADERS = [
  'SL', 'Booking Date', 'Booking Mth', 'Travel Start Date', 'Travel Mth',
  'Travel End Date', 'Travel End Mth', 'Booking ID', 'Campsite Name',
  'Original campsite Name', 'Accommodation', 'Camp Owner Name', 'Camper Name',
  'Camper phone', 'Camper email', 'TOTAL', 'Promocode Name', 'Promocode used',
  'Flat discount', 'Referral discount', 'Weekdays discount', 'Wallet amount used',
  'GST', 'GRAND TOTAL', 'Total units', 'Insta units', 'Total Adults', 'Total Kids',
  'Extra Adults', 'Extra Kids', 'Total guests', 'Total Pets', 'No of Nights',
  'Status', 'BLC vs CLC vs CMV', 'Group', 'Event Name', 'Events', 'Invoice Number',
  'Refund Amount', 'Refund Type',
  'Accommodation Category', 'Booking Lead Time (Days)', 'Promo Used'
];

// COLUMN LAYOUT — 1-based positions within Bookings (A=1, B=2, …)

// NOTE ON INDEXING — THIS WAS WRONG FOR A LONG TIME AND IS THE ROOT CAUSE OF
// EVERY "SUM of Camper email instead of SUM of TOTAL" / "column group shows
// Status instead of Segment" symptom ever seen in this file. Google's own
// docs are explicit: "sourceDataColumn — This index represents the absolute
// number of the column in the spreadsheet; 1 representing column A, 2
// representing column B, etc." — i.e. addRowGroup/addColumnGroup/
// addPivotValue/addFilter's sourceDataColumn is 1-BASED, the SAME as
// getRange()'s column argument, NOT zero-based as a previous version of
// this comment claimed. That wrong assumption meant every single field
// reference was reading the column ONE TO THE LEFT of the intended one —
// consistently, across every tab, regardless of how correct the column
// NAME resolution was, because the off-by-one was in this final numeric
// conversion, not in figuring out which column something lives in.
// Reference: https://developers.google.com/apps-script/reference/spreadsheet/pivot-table#addRowGroup(Integer)
// BOOKINGS_LAST_COL is a plain COUNT (feeds getRange()'s numColumns
// argument), unaffected by this — counts were never the problem.
var BOOKINGS_COL = {
  SL: 1, BOOKING_DATE: 2, BOOKING_MTH: 3, TRAVEL_START: 4, TRAVEL_MTH: 5,
  TRAVEL_END: 6, TRAVEL_END_MTH: 7, BOOKING_ID: 8, CAMPSITE: 9, ORIG_CAMPSITE: 10,
  ACCOMMODATION: 11, CAMP_OWNER: 12, CAMPER_NAME: 13, CAMPER_PHONE: 14, CAMPER_EMAIL: 15,
  TOTAL: 16, PROMO_NAME: 17, PROMO_DISCOUNT: 18, FLAT_DISC: 19, REFERRAL_DISC: 20,
  WEEKDAY_DISC: 21, WALLET_USED: 22, GST: 23, GRAND_TOTAL: 24, TOTAL_UNITS: 25,
  INSTA_UNITS: 26, TOTAL_ADULTS: 27, TOTAL_KIDS: 28, EXTRA_ADULTS: 29, EXTRA_KIDS: 30,
  TOTAL_GUESTS: 31, TOTAL_PETS: 32, NIGHTS: 33, STATUS: 34, SEGMENT: 35, GROUP: 36,
  EVENT_NAME: 37, EVENTS: 38, INVOICE: 39, REFUND_AMOUNT: 40, REFUND_TYPE: 41,
  ACCOM_CATEGORY: 42, LEAD_TIME: 43, PROMO_USED: 44
};
var BOOKINGS_LAST_COL = 44;

// Same key order as BOOKINGS_COL above and as SALES_BOOKINGS_HEADERS — used
// by resolveBookingsColumns() to know which key name goes with which
// expected header text, position by position in the canonical layout.
var BOOKINGS_COL_KEYS = [
  'SL', 'BOOKING_DATE', 'BOOKING_MTH', 'TRAVEL_START', 'TRAVEL_MTH',
  'TRAVEL_END', 'TRAVEL_END_MTH', 'BOOKING_ID', 'CAMPSITE', 'ORIG_CAMPSITE',
  'ACCOMMODATION', 'CAMP_OWNER', 'CAMPER_NAME', 'CAMPER_PHONE', 'CAMPER_EMAIL',
  'TOTAL', 'PROMO_NAME', 'PROMO_DISCOUNT', 'FLAT_DISC', 'REFERRAL_DISC',
  'WEEKDAY_DISC', 'WALLET_USED', 'GST', 'GRAND_TOTAL', 'TOTAL_UNITS',
  'INSTA_UNITS', 'TOTAL_ADULTS', 'TOTAL_KIDS', 'EXTRA_ADULTS', 'EXTRA_KIDS',
  'TOTAL_GUESTS', 'TOTAL_PETS', 'NIGHTS', 'STATUS', 'SEGMENT', 'GROUP',
  'EVENT_NAME', 'EVENTS', 'INVOICE', 'REFUND_AMOUNT', 'REFUND_TYPE',
  'ACCOM_CATEGORY', 'LEAD_TIME', 'PROMO_USED'
];

// Resolves each expected header (SALES_BOOKINGS_HEADERS) to WHEREVER its
// text actually appears in row 1, instead of trusting the fixed position in
// BOOKINGS_COL. This is the fix for a real bug: if a column ever gets
// inserted/deleted in Bookings upstream of this script (manually, or by an
// older/newer version of CL Payments' exportToSales()), every pivot value
// silently reads from the wrong column and shows a wrong-but-plausible
// label ("SUM of Camper email" instead of "SUM of TOTAL") — nothing errors,
// so it's easy to ship and hard to notice. A fixed-position map can't
// detect this at all; searching by name can.
// Returns 1-BASED column numbers throughout (matching BOOKINGS_COL's
// convention — see the note above it), so callers never need to remember
// to convert.
function resolveBookingsColumns(headerRowValues) {
  var textToIndex = {};
  for (var i = 0; i < headerRowValues.length; i++) {
    var v = headerRowValues[i];
    var text = (v === null || v === undefined) ? '' : v.toString().trim();
    if (text && !(text in textToIndex)) textToIndex[text] = i + 1; // +1: 1-based
  }
  var resolved = {};
  var missing = [];
  var drifted = [];
  for (var k = 0; k < BOOKINGS_COL_KEYS.length; k++) {
    var key = BOOKINGS_COL_KEYS[k];
    var expectedText = SALES_BOOKINGS_HEADERS[k];
    var defaultCol = k + 1; // 1-based
    if (expectedText in textToIndex) {
      resolved[key] = textToIndex[expectedText];
      if (resolved[key] !== defaultCol) {
        drifted.push(expectedText + ': expected col ' + defaultCol + ', actually at col ' + resolved[key]);
      }
    } else {
      resolved[key] = defaultCol; // nothing matched anywhere — fall back to the default position
      missing.push(expectedText + ' (expected col ' + (k + 1) + ')');
    }
  }
  return { map: resolved, missing: missing, drifted: drifted };
}

var SUM  = SpreadsheetApp.PivotTableSummarizeFunction.SUM;
var CNTA = SpreadsheetApp.PivotTableSummarizeFunction.COUNTA;
var AVG  = SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE;
var PCT_OF_GRAND_TOTAL = SpreadsheetApp.PivotValueDisplayType.PERCENT_OF_GRAND_TOTAL;

// Any 2D pivot's WIDTH is (row-label col) + (# column-group categories) ×
// (# pivot values) — with real data (31 months of history, half a dozen+
// campsites/channels, up to 4 values per pivot after the calculated-field
// additions) that alone can reach 30-100+ columns. A fixed column offset for
// a second pivot/mini-pivot placed "beside" the first (tried at col 30, then
// col 80) kept getting swallowed by that growth — that's what caused the
// "Array result was not expanded... AD3" / "#REF!" errors in several tabs.
//
// Row growth is far more bounded by comparison (a row-group-only pivot's
// row count is just its own category count — bookings-by-month tops out at
// ~31-40, promo codes maybe a few dozen). So instead of placing extra
// pivots to the RIGHT, everything extra now goes BELOW, at a fixed row deep
// enough to clear any plausible row-group's growth.
var SECOND_PIVOT_ROW = 60;   // second pivot within a tab (was "beside", now "below")
var MINI_PIVOT_ROW = 130;    // chart-source mini pivot (clears SECOND_PIVOT_ROW's own growth too)
var CHART_COL_OFFSET = 5;    // chart sits a few columns right of the mini pivot's own (narrow) output — floating/visual only, so this is purely cosmetic, never a collision risk

// This is a DIFFERENT concern from the row-stacking above: it's about
// whether the sheet's physical grid has enough columns for pt1 ITSELF to
// render without being clipped, regardless of where anything else sits.
// Stacking pt2/the mini pivot below (same column 1) means pt1's own width
// no longer risks colliding with them — but Accommodation Performance's pt1
// alone (31 travel months × 4 values) can still reach 120+ columns, and a
// live pivot table does NOT auto-grow the sheet's grid the way a plain cell
// write does. Generous on purpose.
var REQUIRED_COLS = 150;

function ensureColumns(sh, minCols) {
  var have = sh.getMaxColumns();
  if (have < minCols) sh.insertColumnsAfter(have, minCols - have);
}

// Applies currency/number formats to the Bookings source columns so every
// pivot's SUM/AVERAGE values inherit correct formatting (Sheets pivot output
// cells generally take on the number format of their source column) instead
// of showing bare unformatted decimals. Run once per build, before any tab
// is built. NOTE: this formats the Bookings sheet itself, not the pivot
// tabs — if a pivot value ever still renders unformatted, it's this mapping
// that needs a column added, not a per-tab fix.
function formatBookingsColumns(bookingsSheet, actualLastRow) {
  var numDataRows = actualLastRow - 1; // exclude header row
  if (numDataRows < 1) return;
  var currencyCols = [
    BOOKINGS_COL.TOTAL, BOOKINGS_COL.PROMO_DISCOUNT, BOOKINGS_COL.FLAT_DISC,
    BOOKINGS_COL.REFERRAL_DISC, BOOKINGS_COL.WEEKDAY_DISC, BOOKINGS_COL.WALLET_USED,
    BOOKINGS_COL.GST, BOOKINGS_COL.GRAND_TOTAL, BOOKINGS_COL.REFUND_AMOUNT
  ];
  var integerCols = [
    BOOKINGS_COL.TOTAL_UNITS, BOOKINGS_COL.INSTA_UNITS, BOOKINGS_COL.TOTAL_ADULTS,
    BOOKINGS_COL.TOTAL_KIDS, BOOKINGS_COL.EXTRA_ADULTS, BOOKINGS_COL.EXTRA_KIDS,
    BOOKINGS_COL.TOTAL_GUESTS, BOOKINGS_COL.TOTAL_PETS, BOOKINGS_COL.NIGHTS,
    BOOKINGS_COL.LEAD_TIME
  ];
  // Booking Mth / Travel Mth are genuinely one value per month (confirmed:
  // 5 distinct dates for 5 months of data) — but the underlying value is a
  // real date, so without this it renders as a full date like "26-Jan-2026"
  // everywhere it's grouped on, which reads as day-level data even though
  // it isn't. This is what made some tabs look "day-wise" and others
  // "monthly" — same granularity, inconsistent display. Daily Booking Trend
  // groups on the actual Booking Date column instead, so it's unaffected
  // and correctly keeps showing full dates.
  var monthLabelCols = [BOOKINGS_COL.BOOKING_MTH, BOOKINGS_COL.TRAVEL_MTH];
  // BOOKINGS_COL is already 1-based (matching getRange()'s column argument
  // directly) — no +1 conversion here.
  currencyCols.forEach(function(col) {
    bookingsSheet.getRange(2, col, numDataRows, 1).setNumberFormat('₹#,##0');
  });
  integerCols.forEach(function(col) {
    bookingsSheet.getRange(2, col, numDataRows, 1).setNumberFormat('#,##0');
  });
  monthLabelCols.forEach(function(col) {
    bookingsSheet.getRange(2, col, numDataRows, 1).setNumberFormat('mmm-yy');
  });
}

// GROUP / BOOKING ID COLOR CODING
//
// Single-cell tints only (Group cell + Booking ID cell) — deliberately NOT a
// full-row highlight, so the sheet stays readable instead of turning into a
// wall of color. Same "not too flashy" call already made on the OTA sheet
// and on CL Payments' column H.
//
// The 5 OTA-channel colors below are the EXACT hex values used for the same
// brand on the OTA Bookings sheet and on CL Payments' column H — so a
// GoMMT/ClearTrip/Airbnb/etc. booking looks the same wherever you see it
// across all three spreadsheets. The remaining categories (Direct Booking,
// Group Booking, BLC Group, CM Journeys, BLC/CLC Event Bookings, Dotpe) only
// exist in this Sales sheet, so they get their own new muted colors, chosen
// to stay visually distinct from both the OTA set and from CL Payments'
// existing status colors (green/amber/red/purple) so nothing gets misread.
var GROUP_ROW_COLORS = {
  // OTA channels — match OTA Bookings sheet / CL Payments column H exactly
  GoMMT: '#F0DDB8',
  ClearTrip: '#D0E6F5',
  Airbnb: '#F5C6CB',
  'Stories Collective': '#A8DDD2',
  'TSC Booking': '#A8DDD2',       // same brand as Stories Collective, just a different label in this sheet's data
  'ALIVE Booking': '#FCE8A6',
  // Sales-only categories — new colors, kept muted/pastel on purpose
  'Direct Booking': '#E0E0E0',     // neutral grey — the default bucket, deliberately unobtrusive since it's the bulk of rows
  'Group Booking': '#C9D6EA',      // soft steel blue
  'BLC Group': '#AAB8D6',          // deeper steel blue — same family as Group Booking since it's a subtype of it
  'CM Journeys': '#E0C9A8',        // caramel tan
  'BLC Event Bookings': '#D6C9EA', // soft violet
  'CLC Event Bookings': '#EAC9DC', // soft rose — paired with BLC Event Bookings' violet
  'Dotpe': '#C9B8A8'               // muted beige-brown
};

// Trims stray leading/trailing whitespace from Group cell text (e.g. "Dotpe "
// vs "Dotpe", "TSC Booking " vs "TSC Booking") — without this, pivots and
// this color-coding step both silently treat what should be ONE category as
// two, since Sheets/Excel text matching is exact, not trim-tolerant. Only
// rewrites cells that actually changed, and only ever trims — never alters
// otherwise-valid text — so this can't corrupt a legitimate value.
function normalizeGroupText(bookingsSheet, actualLastRow, groupCol) {
  var numDataRows = actualLastRow - 1;
  if (numDataRows < 1) return 0;
  var range = bookingsSheet.getRange(2, groupCol, numDataRows, 1);
  var values = range.getValues();
  var fixed = 0;
  for (var i = 0; i < values.length; i++) {
    var raw = values[i][0];
    if (typeof raw !== 'string') continue;
    var trimmed = raw.trim();
    if (trimmed !== raw) {
      values[i][0] = trimmed;
      fixed++;
    }
  }
  if (fixed > 0) range.setValues(values);
  return fixed;
}

// Resolves a Group cell's text to a color. Exact matches in GROUP_ROW_COLORS
// win first (covers every category actually seen in real data so far). For
// anything ending in "Event Bookings" that ISN'T an exact match — e.g.
// exportToSales() can now also produce "CMV Event Bookings" or "Agreggated
// Event Bookings" for segments other than BLC/CLC — fall back to a
// segment-keyed variant instead of leaving it uncolored, so a brand-new
// segment showing up later doesn't silently go back to "no color".
function classifyGroupColor(groupText) {
  var g = (groupText || '').toString().trim();
  if (GROUP_ROW_COLORS[g]) return GROUP_ROW_COLORS[g];
  if (/Event Bookings$/.test(g)) {
    if (g.indexOf('BLC') === 0) return '#D6C9EA';       // soft violet
    if (g.indexOf('CLC') === 0) return '#EAC9DC';       // soft rose
    if (g.indexOf('CMV') === 0) return '#C9E0DC';       // soft teal — new segment, kept in the same pastel family
    return '#DCD3EA';                                    // generic light violet fallback (e.g. "Agreggated Event Bookings")
  }
  return null;
}

// Applies the Group-cell and Booking-ID-cell tints described above. Run
// AFTER normalizeGroupText() so the lookup below matches on clean text —
// otherwise "Dotpe " would silently fall through to "no color" on the very
// same run that's supposed to fix it.
function colorCodeBookingsTab(bookingsSheet, actualLastRow, groupCol, bookingIdCol) {
  var numDataRows = actualLastRow - 1;
  if (numDataRows < 1) return;
  var groupRange = bookingsSheet.getRange(2, groupCol, numDataRows, 1);
  var groupValues = groupRange.getValues();
  var groupColors = [];
  var bookingIdColors = [];
  for (var i = 0; i < groupValues.length; i++) {
    var hex = classifyGroupColor(groupValues[i][0]);
    groupColors.push([hex]);
    bookingIdColors.push([hex]);
  }
  groupRange.setBackgrounds(groupColors);
  bookingsSheet.getRange(2, bookingIdCol, numDataRows, 1).setBackgrounds(bookingIdColors);
}

// Wipes whatever ad-hoc conditional-format rules already exist on Bookings.
// The rules found in the live sheet were built by hand in Excel/Sheets over
// a long time — fragmented into hundreds of tiny, mismatched row-ranges
// across the wrong columns (e.g. a "Booking Completed" rule scattered across
// columns Y through AH instead of just Status) — which is the actual cause
// of "gaps": most rows end up with no color, a few unrelated cells got
// tinted by accident, and any legacy rule left in place would still fire on
// top of (and hide) the clean single-cell tints applied above, since a
// matching conditional-format rule always wins over a plain background.
function clearLegacyConditionalFormatting(bookingsSheet) {
  bookingsSheet.setConditionalFormatRules([]);
}

// Cosmetic/readability pass applied at the end of every tab builder: freezes
// the title row and row-label column, and widens columns so labels aren't
// clipped. Wrapped in try/catch — a formatting hiccup shouldn't fail the
// whole tab when the pivot itself built fine.
function finishTab(sh) {
  try {
    sh.setFrozenRows(1);
    sh.setFrozenColumns(1);
    sh.autoResizeColumns(1, Math.min(sh.getMaxColumns(), 20));
  } catch (e) {
    console.log(sh.getName() + ': cosmetic formatting skipped — ' + e);
  }
}

// Adds one embedded chart reading off a fixed range. Wrapped in try/catch —
// charts are a bonus on top of the underlying pivot data, so a chart failure
// must never take down the tab it's decorating.
function addChart(sh, chartType, dataRange, title, anchorRow, anchorCol) {
  try {
    var chart = sh.newChart()
      .setChartType(chartType)
      .addRange(dataRange)
      .setPosition(anchorRow, anchorCol, 0, 0)
      .setOption('title', title)
      .setOption('legend', { position: 'bottom' })
      .setOption('width', 480)
      .setOption('height', 300)
      .build();
    sh.insertChart(chart);
  } catch (e) {
    console.log(sh.getName() + ': chart "' + title + '" failed — ' + e);
  }
}

// Always logs to Cloud Logs (View > Executions) AND tries a visible popup —
// so the outcome is checkable even when a script-editor run doesn't surface
// a dialog. Mirrors the (fixed) safeAlert pattern in CL_Payment_Setup.gs.
function safeAlert(msg) {
  console.log(msg);
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    console.log('(no UI context available to show a popup — see log above)');
  }
}

// MAIN ENTRY POINT

function buildSalesPivots() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bookingsSheet = ss.getSheetByName('Bookings');
  if (!bookingsSheet) {
    safeAlert(
      'ERROR: No "Bookings" tab found in this spreadsheet.\n\n' +
      'This script expects to be run from inside the Sales Report file, ' +
      'with a "Bookings" tab already populated by CL Payments\' export.'
    );
    return;
  }

  // Prove which exact file/tab this run is touching — if there's ever a
  // second "Bookings"-ish tab or a stale duplicate spreadsheet in the mix,
  // this line makes that visible instead of us guessing from a screenshot.
  var allSheetNames = ss.getSheets().map(function(s){ return s.getName(); });
  console.log(
    'Running against spreadsheet: ' + ss.getUrl() + '\n' +
    'Bookings tab gid: ' + bookingsSheet.getSheetId() + '\n' +
    'All tabs in this file: ' + allSheetNames.join(', ')
  );

  // Guard against building pivots off an empty/not-yet-exported sheet —
  // turns a cryptic pivot-API exception into a clear, actionable message.
  var actualLastCol = bookingsSheet.getLastColumn();
  var actualLastRow = bookingsSheet.getLastRow();
  if (actualLastCol < BOOKINGS_LAST_COL || actualLastRow < 2) {
    safeAlert(
      'Sales pivots skipped — Bookings doesn\'t look ready yet.\n\n' +
      'Bookings tab currently has ' + actualLastRow + ' row(s) and ' + actualLastCol +
      ' column(s) (expected at least 2 rows and ' + BOOKINGS_LAST_COL + ' columns).\n\n' +
      'Run "Export to Sales Report" from the CL Payments sheet first, then try this again.'
    );
    return;
  }

  // One-time cleanup of leftovers from earlier iterations of this script:
  // an older version wrote a "Booking Day of Month" helper column (col 45)
  // that a later version stopped using — nothing ever deletes a column just
  // because the code stopped writing to it, so it was still sitting there.
  // Same story for the "Day-of-Month Patterns" tab, renamed to "Daily
  // Booking Trend" — the old tab name isn't in the builders list below, so
  // getOrResetSalesTab() never touches it and it's just orphaned.
  var col45Header = bookingsSheet.getRange(1, 45).getValue();
  if ((col45Header || '').toString().trim() === 'Booking Day of Month') {
    bookingsSheet.getRange(1, 45, bookingsSheet.getMaxRows(), 1).clearContent();
    console.log('Cleaned up leftover "Booking Day of Month" helper column (col 45).');
  }
  var orphanedTab = ss.getSheetByName('Day-of-Month Patterns');
  if (orphanedTab) {
    ss.deleteSheet(orphanedTab);
    console.log('Deleted orphaned "Day-of-Month Patterns" tab (renamed to "Daily Booking Trend").');
  }

  // Read the header row WIDE (not just the expected 44) so a header that's
  // drifted rightward — e.g. someone inserted a column upstream — is still
  // visible to resolveBookingsColumns() below, instead of silently falling
  // outside the scan.
  //
  // ensureColumns() here is NOT optional: Bookings' actual grid width is
  // whatever it's grown to from real writes (currently 45 cols, confirmed
  // from a live export), which is narrower than HEADER_SCAN_COLS. Without
  // this, getRange(1,1,1,70) throws immediately — before ANY tab gets
  // built and before the final safeAlert ever runs, which is exactly what
  // "no alert, tabs just show old/empty content" looks like: the whole
  // function crashes here, on the very first thing it tries to do.
  var HEADER_SCAN_COLS = 70;
  ensureColumns(bookingsSheet, HEADER_SCAN_COLS);
  var headerRange = bookingsSheet.getRange(1, 1, 1, HEADER_SCAN_COLS);
  var headerRow = headerRange.getValues()[0];

  // Resolve every column BY NAME first, on the header text as it actually
  // is right now — before any repair touches it. This is what actually
  // protects against column drift: if "TOTAL" now lives at col 16 instead
  // of col 15 because something got inserted before it, every pivot in
  // every tab should read col 16, not blindly keep using col 15.
  var resolution = resolveBookingsColumns(headerRow);
  BOOKINGS_COL = resolution.map;

  // Self-heal is now narrowed to genuinely broken cells only — blank, or a
  // spreadsheet error string like "#REF!" (this is what originally caused
  // every pivot to fail identically: a poisoned header cell breaks field
  // detection for the WHOLE pivot, not just that column). It no longer
  // overwrites a cell just because it holds some OTHER valid header text at
  // an unexpected position — that's real drift, and resolveBookingsColumns
  // above already handles it correctly by finding the real position instead
  // of erasing the evidence.
  var ERROR_TEXTS = ['#REF!', '#N/A', '#VALUE!', '#NULL!', '#DIV/0!', '#NUM!', '#ERROR!', '#NAME?'];
  var fixedHeaders = [];
  var repaired = headerRow.slice();
  for (var hj = 0; hj < BOOKINGS_LAST_COL; hj++) {
    var expected = SALES_BOOKINGS_HEADERS[hj];
    var actual = headerRow[hj];
    var actualText = (actual === null || actual === undefined) ? '' : actual.toString().trim();
    var isBrokenCell = (actualText === '') || (ERROR_TEXTS.indexOf(actualText) !== -1);
    if (isBrokenCell && actualText !== expected) {
      repaired[hj] = expected;
      fixedHeaders.push('col ' + (hj + 1) + ' (' + expected + '): was "' + actualText + '"');
    }
  }
  if (fixedHeaders.length > 0) {
    headerRange.setValues([repaired]);
    // Force the write to commit to the actual spreadsheet NOW, before
    // anything below reads the sheet again. Apps Script can batch pending
    // writes rather than applying them instantly, and createPivotTable()
    // reading a stale (pre-repair) snapshot of row 1 would silently
    // reproduce the exact same "poisoned header" failure even though the
    // repair code above ran without error.
    SpreadsheetApp.flush();
    // Re-resolve after repair too, in case a just-fixed cell changes what
    // resolveBookingsColumns finds (cheap — one more pass over one row).
    resolution = resolveBookingsColumns(headerRange.getValues()[0]);
    BOOKINGS_COL = resolution.map;
    // Re-read directly from the sheet (not from our in-memory "repaired"
    // array) so the log/alert reflect what Sheets actually committed, not
    // what we intended to write.
    headerRow = headerRange.getValues()[0];
    console.log('Repaired ' + fixedHeaders.length + ' header cell(s) before building pivots:\n' + fixedHeaders.join('\n'));
    console.log('Post-repair, post-flush A1 now reads: "' + headerRow[0] + '"');
  }

  // Format Bookings' own numeric columns (currency vs. plain count) BEFORE
  // building any pivot, so every SUM/AVERAGE value across all 8 tabs renders
  // properly instead of as a bare unformatted decimal.
  formatBookingsColumns(bookingsSheet, actualLastRow);

  // Clean up the Group column's text (trims stray whitespace like "Dotpe "),
  // clear out the old fragmented hand-painted conditional formatting, then
  // apply clean single-cell color tints on Group + Booking ID only — see
  // the functions above for why each step exists.
  var groupTextFixed = normalizeGroupText(bookingsSheet, actualLastRow, BOOKINGS_COL.GROUP);
  clearLegacyConditionalFormatting(bookingsSheet);
  colorCodeBookingsTab(bookingsSheet, actualLastRow, BOOKINGS_COL.GROUP, BOOKINGS_COL.BOOKING_ID);
  if (groupTextFixed > 0) {
    console.log('Trimmed stray whitespace from ' + groupTextFixed + ' Group cell(s).');
  }

  // Source range sized off the ACTUAL row count + a working buffer, not a
  // flat 50,000 — pivoting over tens of thousands of empty rows is what
  // made a build take 10 minutes. This keeps room for growth without
  // dragging in a huge empty range. Width is HEADER_SCAN_COLS, not
  // BOOKINGS_LAST_COL — if a column drifted rightward, a resolved index
  // could be higher than 44, and the source range has to actually cover it.
  var sourceRows = actualLastRow + 2000;
  var sourceRange = bookingsSheet.getRange(1, 1, sourceRows, HEADER_SCAN_COLS);
  console.log(
    'Building pivots — source range: ' + sourceRange.getNumRows() + ' rows × ' +
    sourceRange.getNumColumns() + ' cols. Bookings sheet actual: ' +
    actualLastRow + ' rows × ' + actualLastCol + ' cols. ' +
    'Header row: ' + JSON.stringify(headerRow) + '.' +
    (fixedHeaders.length > 0 ? ' (' + fixedHeaders.length + ' header cell(s) auto-repaired this run.)' : ' Headers OK.') +
    (resolution.drifted.length > 0 ? '\nDRIFTED (found at a different column than default): ' + resolution.drifted.join('; ') : '') +
    (resolution.missing.length > 0 ? '\nMISSING (not found anywhere in row 1, using default position): ' + resolution.missing.join('; ') : '')
  );

  // Each tab wrapped individually so one bad pivot doesn't abort the rest,
  // and any failure names exactly which tab it was.
  var builders = [
    ['Revenue Trends', buildRevenueTrendsTab],
    ['Accommodation Performance', buildAccommodationPerformanceTab],
    ['Guest & Occupancy Mix', buildGuestOccupancyMixTab],
    ['Pet-Friendly Stays', buildPetFriendlyStaysTab],
    ['Cancellations & Refunds', buildCancellationsRefundsTab],
    ['Promo & Discount Effectiveness', buildPromoEffectivenessTab],
    ['Booking Lead Time', buildBookingLeadTimeTab],
    ['Segment & Group Mix', buildSegmentGroupMixTab],
    ['Booking vs Travel Trends', buildBookingVsTravelTrendsTab],
    ['Daily Booking Trend', buildDailyBookingTrendTab],
    ['OTA Accommodation Mix', buildOtaAccommodationMixTab]
  ];
  var built = [];
  var failed = [];
  for (var bi = 0; bi < builders.length; bi++) {
    try {
      builders[bi][1](ss, sourceRange);
      built.push(builders[bi][0]);
    } catch (e) {
      failed.push(builders[bi][0] + ': ' + e);
    }
  }

  var driftWarning = '';
  if (resolution.drifted.length > 0) {
    driftWarning += '\n⚠️ COLUMN DRIFT DETECTED — these fields were found at a different ' +
      'column than expected (likely a column was inserted/deleted in Bookings upstream). ' +
      'Every pivot was built using the ACTUAL position, so this run should be correct, but ' +
      'worth checking why the layout moved:\n' + resolution.drifted.join('\n') + '\n';
  }
  if (resolution.missing.length > 0) {
    driftWarning += '\n⚠️ NOT FOUND anywhere in row 1 (using the default fixed position, ' +
      'which may be wrong) — check these headers exist and are spelled exactly as expected:\n' +
      resolution.missing.join('\n') + '\n';
  }

  safeAlert(
    (failed.length === 0 && !driftWarning ? '✅ Sales pivots built!\n\n' : '⚠️ Sales pivots built with warnings\n\n') +
    'File: ' + ss.getUrl() + '\n' +
    'Bookings tab gid: ' + bookingsSheet.getSheetId() + '\n' +
    'A1 right now: "' + headerRow[0] + '"' +
    (fixedHeaders.length > 0 ? ' (auto-repaired ' + fixedHeaders.length + ' header cell(s) this run)' : ' (no repair needed)') + '\n\n' +
    'Built (' + built.length + '): ' + built.join(', ') + '\n\n' +
    (failed.length > 0 ? 'Failed (' + failed.length + '):\n' + failed.join('\n') + '\n\n' : '') +
    '• Replaced the old hand-painted conditional formatting with clean single-cell tints on Group + Booking ID (no full-row colors)\n' +
    (groupTextFixed > 0 ? '• Trimmed stray whitespace from ' + groupTextFixed + ' Group cell(s) (e.g. "Dotpe " → "Dotpe") so pivots count them correctly\n' : '') +
    '\n' +
    driftWarning +
    '\nThese are live native pivot tables — they recalculate automatically as ' +
    'Bookings data changes. Re-run this only if you want to reset their layout.'
  );
}

// Gets (or creates) a tab, removing any pivot tables and clearing it first
// so each rebuild starts clean — safe to run this repeatedly.
function getOrResetSalesTab(ss, name) {
  var sh = ss.getSheetByName(name);
  if (sh) {
    var pts = sh.getPivotTables();
    for (var i = 0; i < pts.length; i++) pts[i].remove();
    // Charts are sheet-level objects too, just like pivot tables — clear()
    // does NOT remove them (same reason pivot tables need an explicit
    // .remove() above). Without this, every rebuild would stack another
    // copy of the same chart on top of the last one.
    var charts = sh.getCharts();
    for (var j = 0; j < charts.length; j++) sh.removeChart(charts[j]);
    sh.clear();
  } else {
    sh = ss.insertSheet(name);
  }
  return sh;
}

// 1. Revenue Trends — revenue by month, campsite, accommodation category.
function buildRevenueTrendsTab(ss, sourceRange) {
  var sh = getOrResetSalesTab(ss, 'Revenue Trends');
  ensureColumns(sh, REQUIRED_COLS);
  sh.getRange(1, 1).setValue('Revenue by Booking Month × Campsite').setFontWeight('bold');
  console.log('Revenue Trends: about to call createPivotTable, sourceRange cols=' + sourceRange.getNumColumns());
  var pt1 = sh.getRange(3, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  // Ask the pivot table object itself what it thinks its source range is —
  // rather than trusting our own separately-built sourceRange variable. If
  // these disagree with what we requested, that's the actual bug.
  var pt1Src = pt1.getSourceDataRange();
  console.log(
    'Revenue Trends: createPivotTable OK. pt1.getSourceDataRange() = ' + pt1Src.getA1Notation() +
    ' on sheet "' + pt1Src.getSheet().getName() + '", ' + pt1Src.getNumRows() + ' rows × ' +
    pt1Src.getNumColumns() + ' cols. About to addRowGroup(' + BOOKINGS_COL.BOOKING_MTH + ')'
  );
  pt1.addRowGroup(BOOKINGS_COL.BOOKING_MTH);
  console.log('Revenue Trends: addRowGroup OK, about to addColumnGroup(' + BOOKINGS_COL.CAMPSITE + ')');
  pt1.addColumnGroup(BOOKINGS_COL.CAMPSITE);
  console.log('Revenue Trends: addColumnGroup OK, about to addPivotValue(' + BOOKINGS_COL.TOTAL + ')');
  pt1.addPivotValue(BOOKINGS_COL.TOTAL, SUM);
  console.log('Revenue Trends: first addPivotValue OK');
  pt1.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA);
  // Net of refunds — the headline "how much did we actually keep" number.
  pt1.addCalculatedPivotValue('Net Revenue After Refunds', "='GRAND TOTAL'-'Refund Amount'");
  // Share of company-wide revenue each campsite/month contributes.
  pt1.addPivotValue(BOOKINGS_COL.TOTAL, SUM)
    .setDisplayName('% of Total Revenue')
    .showAs(PCT_OF_GRAND_TOTAL);

  sh.getRange(SECOND_PIVOT_ROW - 2, 1).setValue('Revenue by Travel Month × Accommodation Category').setFontWeight('bold');
  var pt2 = sh.getRange(SECOND_PIVOT_ROW, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  pt2.addRowGroup(BOOKINGS_COL.TRAVEL_MTH);
  pt2.addColumnGroup(BOOKINGS_COL.ACCOM_CATEGORY);
  pt2.addPivotValue(BOOKINGS_COL.TOTAL, SUM);
  pt2.addPivotValue(BOOKINGS_COL.TOTAL_GUESTS, SUM);

  // Chart-source mini pivot: total revenue trend, one row per booking
  // month, no column split — the simplest possible shape to chart cleanly.
  sh.getRange(MINI_PIVOT_ROW, 1).setValue('Revenue Trend (chart source)').setFontWeight('bold');
  var pt3 = sh.getRange(MINI_PIVOT_ROW + 2, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush();
  pt3.addRowGroup(BOOKINGS_COL.BOOKING_MTH);
  pt3.addPivotValue(BOOKINGS_COL.TOTAL, SUM).setDisplayName('Total Revenue');
  pt3.addCalculatedPivotValue('Net Revenue After Refunds', "='GRAND TOTAL'-'Refund Amount'");
  addChart(sh, Charts.ChartType.LINE,
    sh.getRange(MINI_PIVOT_ROW + 2, 1,40, 3),
    'Revenue Trend by Booking Month', MINI_PIVOT_ROW, 1 + CHART_COL_OFFSET);

  finishTab(sh);
  sh.setTabColor('#4285f4');
}

// 2. Accommodation Performance — room nights by real accommodation type
// (Addon/Charge, Day Outing, and Event rows filtered out).
function buildAccommodationPerformanceTab(ss, sourceRange) {
  var sh = getOrResetSalesTab(ss, 'Accommodation Performance');
  ensureColumns(sh, REQUIRED_COLS);
  sh.getRange(1, 1).setValue('Room Nights by Accommodation Type × Travel Month (real accommodations only)').setFontWeight('bold');
  var pt = sh.getRange(3, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  var accomCategoryFilter = SpreadsheetApp.newFilterCriteria().whenTextEqualTo('Accommodation Type').build();
  pt.addRowGroup(BOOKINGS_COL.ACCOMMODATION);
  pt.addColumnGroup(BOOKINGS_COL.TRAVEL_MTH);
  pt.addPivotValue(BOOKINGS_COL.NIGHTS, SUM);
  pt.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA);
  pt.addPivotValue(BOOKINGS_COL.TOTAL, SUM);
  // Blended average daily rate — revenue earned per room-night, the number
  // that actually says whether pricing/mix is improving, not just volume.
  pt.addCalculatedPivotValue('Avg Revenue per Night (₹)', "=TOTAL/'No of Nights'");
  pt.addFilter(BOOKINGS_COL.ACCOM_CATEGORY, accomCategoryFilter);

  // Chart-source mini pivot: room nights ranked by accommodation type,
  // no column split, so it charts cleanly as a bar ranking.
  sh.getRange(MINI_PIVOT_ROW, 1).setValue('Room Nights by Type (chart source)').setFontWeight('bold');
  var pt2 = sh.getRange(MINI_PIVOT_ROW + 2, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush();
  pt2.addRowGroup(BOOKINGS_COL.ACCOMMODATION);
  pt2.addPivotValue(BOOKINGS_COL.NIGHTS, SUM).setDisplayName('Total Room Nights');
  pt2.addFilter(BOOKINGS_COL.ACCOM_CATEGORY, accomCategoryFilter);
  addChart(sh, Charts.ChartType.COLUMN,
    sh.getRange(MINI_PIVOT_ROW + 2, 1,30, 2),
    'Room Nights by Accommodation Type', MINI_PIVOT_ROW, 1 + CHART_COL_OFFSET);

  finishTab(sh);
  sh.setTabColor('#0f9d58');
}

// 3. Guest & Occupancy Mix — party composition trend over time.
function buildGuestOccupancyMixTab(ss, sourceRange) {
  var sh = getOrResetSalesTab(ss, 'Guest & Occupancy Mix');
  ensureColumns(sh, REQUIRED_COLS);
  sh.getRange(1, 1).setValue('Guest Mix by Travel Month').setFontWeight('bold');
  var pt = sh.getRange(3, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  pt.addRowGroup(BOOKINGS_COL.TRAVEL_MTH);
  pt.addPivotValue(BOOKINGS_COL.TOTAL_ADULTS, SUM);
  pt.addPivotValue(BOOKINGS_COL.TOTAL_KIDS, SUM);
  pt.addPivotValue(BOOKINGS_COL.EXTRA_ADULTS, SUM);
  pt.addPivotValue(BOOKINGS_COL.EXTRA_KIDS, SUM);
  pt.addPivotValue(BOOKINGS_COL.TOTAL_GUESTS, SUM);
  pt.addPivotValue(BOOKINGS_COL.TOTAL_GUESTS, AVG).setDisplayName('Avg Guests / Booking');
  pt.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA);

  // Chart-source mini pivot: total guests trended by travel month.
  sh.getRange(MINI_PIVOT_ROW, 1).setValue('Guest Volume Trend (chart source)').setFontWeight('bold');
  var pt2 = sh.getRange(MINI_PIVOT_ROW + 2, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush();
  pt2.addRowGroup(BOOKINGS_COL.TRAVEL_MTH);
  pt2.addPivotValue(BOOKINGS_COL.TOTAL_GUESTS, SUM).setDisplayName('Total Guests');
  pt2.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA).setDisplayName('Bookings');
  addChart(sh, Charts.ChartType.LINE,
    sh.getRange(MINI_PIVOT_ROW + 2, 1,40, 3),
    'Guest Volume by Travel Month', MINI_PIVOT_ROW, 1 + CHART_COL_OFFSET);

  finishTab(sh);
  sh.setTabColor('#f4b400');
}

// 4. Pet-Friendly Stays — pet volume trend, independent of accommodation
// naming (Total Pets is already clean numeric data).
function buildPetFriendlyStaysTab(ss, sourceRange) {
  var sh = getOrResetSalesTab(ss, 'Pet-Friendly Stays');
  ensureColumns(sh, REQUIRED_COLS);
  sh.getRange(1, 1).setValue('Pets by Campsite × Travel Month').setFontWeight('bold');
  var pt = sh.getRange(3, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  var petFilter = SpreadsheetApp.newFilterCriteria().whenNumberGreaterThan(0).build();
  pt.addRowGroup(BOOKINGS_COL.CAMPSITE);
  pt.addColumnGroup(BOOKINGS_COL.TRAVEL_MTH);
  pt.addPivotValue(BOOKINGS_COL.TOTAL_PETS, SUM);
  pt.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA).setDisplayName('Count of Bookings');
  pt.addPivotValue(BOOKINGS_COL.TOTAL_PETS, AVG).setDisplayName('Avg Pets / Booking (Pet Stays)');
  pt.addFilter(BOOKINGS_COL.TOTAL_PETS, petFilter);

  // Chart-source mini pivot: pet-friendly demand share by campsite.
  sh.getRange(MINI_PIVOT_ROW, 1).setValue('Pet Demand Share (chart source)').setFontWeight('bold');
  var pt2 = sh.getRange(MINI_PIVOT_ROW + 2, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush();
  pt2.addRowGroup(BOOKINGS_COL.CAMPSITE);
  pt2.addPivotValue(BOOKINGS_COL.TOTAL_PETS, SUM).setDisplayName('Total Pets');
  pt2.addFilter(BOOKINGS_COL.TOTAL_PETS, petFilter);
  addChart(sh, Charts.ChartType.PIE,
    sh.getRange(MINI_PIVOT_ROW + 2, 1,30, 2),
    'Pet-Friendly Demand Share by Campsite', MINI_PIVOT_ROW, 1 + CHART_COL_OFFSET);

  finishTab(sh);
  sh.setTabColor('#8e24aa');
}

// 5. Cancellations & Refunds — dedicated view, with the wallet-vs-source
// split from Refund Type.
function buildCancellationsRefundsTab(ss, sourceRange) {
  var sh = getOrResetSalesTab(ss, 'Cancellations & Refunds');
  ensureColumns(sh, REQUIRED_COLS);
  sh.getRange(1, 1).setValue('Cancellations & Refunds by Booking Month').setFontWeight('bold');
  var pt = sh.getRange(3, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  // whenTextEqualToAny() only works on DataSource-backed pivots ("non-data
  // source objects" in the error) — this is a plain range pivot, so a
  // multi-value equality check has to go through a formula filter instead.
  // $AH2 = Status column (BOOKINGS_COL.STATUS is col 34, 1-based, = AH),
  // row 2 = first data row under the header; Sheets adjusts the row per line.
  var cancelledOrRefundedFilter = SpreadsheetApp.newFilterCriteria()
    .whenFormulaSatisfied('=OR($AH2="Cancelled",$AH2="Refunded")')
    .build();
  pt.addRowGroup(BOOKINGS_COL.BOOKING_MTH);
  pt.addColumnGroup(BOOKINGS_COL.REFUND_TYPE);
  pt.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA).setDisplayName('Count');
  pt.addPivotValue(BOOKINGS_COL.REFUND_AMOUNT, SUM);
  // How much of that month/type's gross revenue came back out the door.
  pt.addCalculatedPivotValue('Refund Rate (%)', "=('Refund Amount'/'GRAND TOTAL')*100");
  pt.addFilter(BOOKINGS_COL.STATUS, cancelledOrRefundedFilter);

  // Chart-source mini pivot: refund ₹ trended by booking month.
  sh.getRange(MINI_PIVOT_ROW, 1).setValue('Refund Trend (chart source)').setFontWeight('bold');
  var pt2 = sh.getRange(MINI_PIVOT_ROW + 2, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush();
  pt2.addRowGroup(BOOKINGS_COL.BOOKING_MTH);
  pt2.addPivotValue(BOOKINGS_COL.REFUND_AMOUNT, SUM).setDisplayName('Total Refunded');
  pt2.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA).setDisplayName('Cancelled/Refunded Count');
  pt2.addFilter(BOOKINGS_COL.STATUS, cancelledOrRefundedFilter);
  addChart(sh, Charts.ChartType.COLUMN,
    sh.getRange(MINI_PIVOT_ROW + 2, 1,40, 3),
    'Refund Amount by Booking Month', MINI_PIVOT_ROW, 1 + CHART_COL_OFFSET);

  finishTab(sh);
  sh.setTabColor('#db4437');
}

// 6. Promo & Discount Effectiveness — booking volume/revenue with vs.
// without a promocode, plus discount value by promocode.
function buildPromoEffectivenessTab(ss, sourceRange) {
  var sh = getOrResetSalesTab(ss, 'Promo & Discount Effectiveness');
  ensureColumns(sh, REQUIRED_COLS);
  sh.getRange(1, 1).setValue('Bookings & Revenue by Promo Used × Booking Month').setFontWeight('bold');
  var pt1 = sh.getRange(3, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  pt1.addRowGroup(BOOKINGS_COL.BOOKING_MTH);
  pt1.addColumnGroup(BOOKINGS_COL.PROMO_USED);
  pt1.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA);
  pt1.addPivotValue(BOOKINGS_COL.TOTAL, SUM);
  // Does using a promo actually change average order value, up or down.
  pt1.addPivotValue(BOOKINGS_COL.TOTAL, AVG).setDisplayName('Avg Booking Value (₹)');
  // What share of gross revenue that month/segment gave up as discount.
  pt1.addCalculatedPivotValue('Discount Rate (%)', "=('Promocode used'/TOTAL)*100");

  var promoUsedFilter = SpreadsheetApp.newFilterCriteria().whenTextEqualTo('Yes').build();
  sh.getRange(SECOND_PIVOT_ROW - 2, 1).setValue('Discount Value by Promocode Name').setFontWeight('bold');
  var pt2 = sh.getRange(SECOND_PIVOT_ROW, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  pt2.addRowGroup(BOOKINGS_COL.PROMO_NAME);
  pt2.addPivotValue(BOOKINGS_COL.PROMO_DISCOUNT, SUM);
  pt2.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA);
  pt2.addFilter(BOOKINGS_COL.PROMO_USED, promoUsedFilter);

  // Chart-source mini pivot: which promocodes cost the most, ranked.
  sh.getRange(MINI_PIVOT_ROW, 1).setValue('Discount Spend by Promocode (chart source)').setFontWeight('bold');
  var pt3 = sh.getRange(MINI_PIVOT_ROW + 2, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush();
  pt3.addRowGroup(BOOKINGS_COL.PROMO_NAME);
  pt3.addPivotValue(BOOKINGS_COL.PROMO_DISCOUNT, SUM).setDisplayName('Total Discount (₹)');
  pt3.addFilter(BOOKINGS_COL.PROMO_USED, promoUsedFilter);
  addChart(sh, Charts.ChartType.COLUMN,
    sh.getRange(MINI_PIVOT_ROW + 2, 1,40, 2),
    'Discount Spend by Promocode', MINI_PIVOT_ROW, 1 + CHART_COL_OFFSET);

  finishTab(sh);
  sh.setTabColor('#ff6d00');
}

// 7. Booking Lead Time — how far in advance people book, bucketed into
// weekly bins via native histogram grouping, trended by booking month.
function buildBookingLeadTimeTab(ss, sourceRange) {
  var sh = getOrResetSalesTab(ss, 'Booking Lead Time');
  ensureColumns(sh, REQUIRED_COLS);
  sh.getRange(1, 1).setValue('Booking Lead Time (Days) Distribution × Booking Month').setFontWeight('bold');
  var pt = sh.getRange(3, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  var rowGrp = pt.addRowGroup(BOOKINGS_COL.LEAD_TIME);
  try { rowGrp.setHistogramGroupRule(7); } catch (e) { /* falls back to raw values if histogram unsupported */ }
  pt.addColumnGroup(BOOKINGS_COL.BOOKING_MTH);
  pt.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA).setDisplayName('Count of Bookings');

  // Chart-source mini pivot: are people booking further ahead or closer
  // to travel date over time — a single trended average, easy to read.
  sh.getRange(MINI_PIVOT_ROW, 1).setValue('Avg Lead Time Trend (chart source)').setFontWeight('bold');
  var pt2 = sh.getRange(MINI_PIVOT_ROW + 2, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush();
  pt2.addRowGroup(BOOKINGS_COL.BOOKING_MTH);
  pt2.addPivotValue(BOOKINGS_COL.LEAD_TIME, AVG).setDisplayName('Avg Lead Time (Days)');
  addChart(sh, Charts.ChartType.LINE,
    sh.getRange(MINI_PIVOT_ROW + 2, 1,40, 2),
    'Avg Booking Lead Time by Month', MINI_PIVOT_ROW, 1 + CHART_COL_OFFSET);

  finishTab(sh);
  sh.setTabColor('#00acc1');
}

// 8. Segment & Group Mix — BLC/CLC/CMV and Direct/Group booking mix.
function buildSegmentGroupMixTab(ss, sourceRange) {
  var sh = getOrResetSalesTab(ss, 'Segment & Group Mix');
  ensureColumns(sh, REQUIRED_COLS);
  sh.getRange(1, 1).setValue('Revenue by Booking Month × Segment (BLC/CLC/CMV)').setFontWeight('bold');
  var pt1 = sh.getRange(3, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  pt1.addRowGroup(BOOKINGS_COL.BOOKING_MTH);
  pt1.addColumnGroup(BOOKINGS_COL.SEGMENT);
  pt1.addPivotValue(BOOKINGS_COL.TOTAL, SUM);
  pt1.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA);
  pt1.addPivotValue(BOOKINGS_COL.TOTAL, AVG).setDisplayName('Avg Booking Value (₹)');

  sh.getRange(SECOND_PIVOT_ROW - 2, 1).setValue('Revenue by Booking Month × Group Type (Booking Channel)').setFontWeight('bold');
  var pt2 = sh.getRange(SECOND_PIVOT_ROW, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  pt2.addRowGroup(BOOKINGS_COL.BOOKING_MTH);
  pt2.addColumnGroup(BOOKINGS_COL.GROUP);
  pt2.addPivotValue(BOOKINGS_COL.TOTAL, SUM);
  pt2.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA);

  // Chart-source mini pivot: revenue share by booking channel (Airbnb,
  // Direct, GoMMT, Group Booking, etc. — the "Group" field is actually the
  // booking channel, not group-vs-individual). This is the channel-mix view
  // leadership currently only gets as a raw, uncharted table elsewhere.
  sh.getRange(MINI_PIVOT_ROW, 1).setValue('Revenue by Channel (chart source)').setFontWeight('bold');
  var pt3 = sh.getRange(MINI_PIVOT_ROW + 2, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush();
  pt3.addRowGroup(BOOKINGS_COL.GROUP);
  pt3.addPivotValue(BOOKINGS_COL.TOTAL, SUM).setDisplayName('Total Revenue');
  pt3.addPivotValue(BOOKINGS_COL.TOTAL, SUM)
    .setDisplayName('% of Total Revenue')
    .showAs(PCT_OF_GRAND_TOTAL);
  addChart(sh, Charts.ChartType.PIE,
    sh.getRange(MINI_PIVOT_ROW + 2, 1,20, 2),
    'Revenue Share by Booking Channel', MINI_PIVOT_ROW, 1 + CHART_COL_OFFSET);

  finishTab(sh);
  sh.setTabColor('#9e9d24');
}

// 9. Booking vs Travel Trends — the two questions "which month do we get
// a boost in bookings" (Booking Month = when the sale happens) and "which
// month do people actually want to travel" (Travel Month = seasonality of
// demand) are different questions and get mixed up when only one is shown.
// Side by side, each as a simple single-dimension pivot so it charts cleanly.
function buildBookingVsTravelTrendsTab(ss, sourceRange) {
  var sh = getOrResetSalesTab(ss, 'Booking vs Travel Trends');
  ensureColumns(sh, REQUIRED_COLS);

  sh.getRange(1, 1).setValue('Bookings & Revenue by BOOKING Month — when demand is generated').setFontWeight('bold');
  var pt1 = sh.getRange(3, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  pt1.addRowGroup(BOOKINGS_COL.BOOKING_MTH);
  pt1.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA).setDisplayName('Bookings');
  pt1.addPivotValue(BOOKINGS_COL.TOTAL, SUM).setDisplayName('Revenue');
  addChart(sh, Charts.ChartType.COLUMN,
    sh.getRange(3, 1, 40, 3),
    'Bookings Made, by Booking Month', 1, 25);

  sh.getRange(1, 12).setValue('Bookings & Revenue by TRAVEL Month — when guests actually stay').setFontWeight('bold');
  var pt2 = sh.getRange(3, 12).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  pt2.addRowGroup(BOOKINGS_COL.TRAVEL_MTH);
  pt2.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA).setDisplayName('Bookings');
  pt2.addPivotValue(BOOKINGS_COL.TOTAL, SUM).setDisplayName('Revenue');
  addChart(sh, Charts.ChartType.COLUMN,
    sh.getRange(3, 12, 40, 3),
    'Travel Demand, by Travel Month', 18, 25);

  finishTab(sh);
  sh.setTabColor('#3f51b5');
}

// 10. Daily Booking Trend — every actual calendar date (Booking Date
// itself, not a day-number bucket), so specific spurt/lull dates are visible
// directly rather than smeared across months. Over ~2-3 years of data this
// is easily 500-1000+ distinct dates, so it's charted as a LINE (a column
// chart with that many categories is unreadable) and given generous row
// headroom.
function buildDailyBookingTrendTab(ss, sourceRange) {
  var sh = getOrResetSalesTab(ss, 'Daily Booking Trend');
  ensureColumns(sh, REQUIRED_COLS);
  sh.getRange(1, 1).setValue('Bookings by Calendar Date — where the spurts and lulls actually fall').setFontWeight('bold');
  var pt = sh.getRange(3, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  pt.addRowGroup(BOOKINGS_COL.BOOKING_DATE);
  pt.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA).setDisplayName('Bookings');
  pt.addPivotValue(BOOKINGS_COL.TOTAL, SUM).setDisplayName('Revenue');
  addChart(sh, Charts.ChartType.LINE,
    sh.getRange(3, 1, 1100, 3),
    'Bookings by Date — peaks = spurts, dips = lulls', 1, 15);

  finishTab(sh);
  sh.setTabColor('#795548');
}

// 11. OTA Accommodation Mix — which accommodation types get booked most
// via OTA platforms (Airbnb/GoMMT/ClearTrip/ALIVE Booking/Dotpe), with a
// Direct Booking column alongside for comparison. The OTA channel list below
// is a stand-in until the dedicated OTA/commission-rate sheet is ready —
// once that exists, the commission-leakage and "website discount headroom"
// calculation (Excel Sales Summary sheet's ask) can be added as a
// calculated pivot value here, keyed off that sheet's per-channel rates.
function buildOtaAccommodationMixTab(ss, sourceRange) {
  var sh = getOrResetSalesTab(ss, 'OTA Accommodation Mix');
  ensureColumns(sh, REQUIRED_COLS);
  sh.getRange(1, 1).setValue('OTA Bookings by Accommodation Type').setFontWeight('bold');
  sh.getRange(2, 1).setValue('OTA channel list (Airbnb/GoMMT/ClearTrip/ALIVE Booking/Dotpe) is a placeholder — update once the dedicated OTA/commission sheet is ready.').setFontStyle('italic');

  // $AJ2 = Group/channel column (BOOKINGS_COL.GROUP is col 36, 1-based, =
  // AJ), row 2 = first data row; Sheets adjusts the row per line. TRIM()
  // guards against the trailing-space duplicates seen in the raw data
  // (e.g. "Dotpe" vs "Dotpe ", "TSC Booking " with no clean match).
  var otaFormula = '=OR(TRIM($AJ2)="Airbnb",TRIM($AJ2)="GoMMT",TRIM($AJ2)="ClearTrip",TRIM($AJ2)="ALIVE Booking",TRIM($AJ2)="Dotpe")';
  var otaFilter = SpreadsheetApp.newFilterCriteria().whenFormulaSatisfied(otaFormula).build();
  var directFilter = SpreadsheetApp.newFilterCriteria().whenFormulaSatisfied('=TRIM($AJ2)="Direct Booking"').build();

  var pt1 = sh.getRange(4, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  pt1.addRowGroup(BOOKINGS_COL.ACCOMMODATION);
  pt1.addColumnGroup(BOOKINGS_COL.GROUP);
  pt1.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA).setDisplayName('OTA Bookings');
  pt1.addPivotValue(BOOKINGS_COL.TOTAL, SUM).setDisplayName('OTA Revenue');
  pt1.addFilter(BOOKINGS_COL.GROUP, otaFilter);

  // Same shape for Direct Booking, stacked below (not beside — pt1's own
  // column-group width isn't bounded enough to safely place anything to its
  // right at a fixed column) for an immediate comparison.
  sh.getRange(SECOND_PIVOT_ROW - 2, 1).setValue('Direct Booking by Accommodation Type (comparison)').setFontWeight('bold');
  var pt2 = sh.getRange(SECOND_PIVOT_ROW, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush(); // needed before add*Group/addPivotValue or it can throw on a valid column index
  pt2.addRowGroup(BOOKINGS_COL.ACCOMMODATION);
  pt2.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA).setDisplayName('Direct Bookings');
  pt2.addPivotValue(BOOKINGS_COL.TOTAL, SUM).setDisplayName('Direct Revenue');
  pt2.addFilter(BOOKINGS_COL.GROUP, directFilter);

  // Chart-source mini pivot: OTA bookings ranked by accommodation type.
  sh.getRange(MINI_PIVOT_ROW, 1).setValue('OTA Bookings by Type (chart source)').setFontWeight('bold');
  var pt3 = sh.getRange(MINI_PIVOT_ROW + 2, 1).createPivotTable(sourceRange);
  SpreadsheetApp.flush();
  pt3.addRowGroup(BOOKINGS_COL.ACCOMMODATION);
  pt3.addPivotValue(BOOKINGS_COL.BOOKING_ID, CNTA).setDisplayName('OTA Bookings');
  pt3.addFilter(BOOKINGS_COL.GROUP, otaFilter);
  addChart(sh, Charts.ChartType.COLUMN,
    sh.getRange(MINI_PIVOT_ROW + 2, 1,20, 2),
    'Most-Booked Accommodation Type via OTA', MINI_PIVOT_ROW, 1 + CHART_COL_OFFSET);

  finishTab(sh);
  sh.setTabColor('#00695c');
}
