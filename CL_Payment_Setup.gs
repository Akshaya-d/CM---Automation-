/**
 * CL Payment Sheet script.
 *
 * Setup: Extensions > Apps Script, paste this in, save, run setupCLPayments()
 * once to grant perms, reload the sheet — adds a "CL Automation" menu. Day
 * to day just use "Update Now" — only rebuilds if something actually
 * changed (see smartUpdate()).
 *
 * Reads Raw Booking Data + Raw Experience Booking Data (merged, sorted by
 * Booking Date) into one 52-col CL Payments tab. Raw Experience Booking Data
 * auto-creates with headers if it doesn't exist. Only "confirmed" experience
 * rows make it into CL Payments — draft/failed/initiated stuff goes to
 * Experience Leads instead so it's not lost.
 *
 * Both raw tabs get deduped by Booking ID (latest paste wins) before being
 * read — the experience export has no date filter and gets re-pasted over
 * overlapping windows a lot.
 *
 * CMX bookings (day outing) get a light blue row tint, low priority so
 * status highlights still win.
 *
 * Event bookings (June 2026+, since inventory sync): most now also have a
 * real property booking for the same transaction (e.g. CM28192 in Raw
 * Booking Data = event CME6540). Razorpay only ever has the event ID, and
 * the property row's Amount is usually 0 (placeholder), so:
 *   - matched (same email/phone + Booking Date) → no second row, property
 *     row's Booking ID gets swapped to the event ID so Razorpay/downstream
 *     matching just works. O/V pulled straight from the event's own
 *     Amount/GST. Original property Booking ID goes into Comments 2,
 *     accommodation type into Comments 3, Accommodation col shows event
 *     title.
 *   - unmatched + confirmed → standalone row via EVENT_MAP, flat 18% GST
 *     reverse-calc since there's no accommodation name for the GST lookup.
 * Pre-June-2026 events (old e_xxx IDs) aren't handled here. CME rows get a
 * light green tint, same low priority as CMX.
 *
 * Wallet/promo: when TOTAL+GST != GRAND TOTAL, O and V get reverse
 * calculated off the Master GST Data divisor:
 *   O = AV * 100 / divisor      (divisor from Master GST Data)
 *   V = AV * (divisor-100) / divisor
 * W is always O+V. Wallet use noted in AS.
 *
 * Columns: A-AF booking data, AG status, AH Razorpay ID, AI-AV finance +
 * comments, AW final status.
 *
 * Safe to re-run — full rebuild each time. Manual notes in Comments 2/3
 * (AT/AU) get wiped on rebuild.
 */

// Safe alert — works whether the script is bound to a sheet or run from a trigger.
// Tries getUi().alert() first; falls back to console.log if UI isn't available.
// (Fixed: this used to call itself instead of calling the actual UI alert —
// an infinite-recursion bug that meant EVERY "alert" in this script was
// silently falling straight to the console.log fallback, never showing a
// real popup. That's very likely why several "nothing happened when I
// clicked X" reports never showed a visible dialog.)
function safeAlert(msg) {
  console.log(msg); // always log, regardless of whether the UI dialog below works
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    console.log('(no UI context available to show a popup — see log above)');
  }
}

// Dedupe a raw intake sheet by Booking ID
// Admin exports without a clean date-range filter mean the same booking gets
// re-downloaded (and re-pasted) across overlapping windows. This collapses
// repeat Booking IDs down to one row, keeping the LAST occurrence (i.e. the
// most recently pasted copy — the one most likely to carry an updated
// status, e.g. draft → confirmed). Shared by setupCLPayments() and
// exportExperienceLeads() so the raw tab is deduped no matter which runs first.
// Returns the number of duplicate rows removed.
function dedupeSheetByBookingId(sheet, bidCol) {
  if (!sheet) return 0;
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, last - 1, lastCol).getValues();

  var seenAt = {};   // bid -> index into `unique`
  var unique = [];
  for (var i = 0; i < data.length; i++) {
    var bid = (data[i][bidCol - 1] || '').toString().trim();
    if (!bid) { unique.push(data[i]); continue; } // no ID — leave as-is, can't dedupe
    if (seenAt.hasOwnProperty(bid)) {
      unique[seenAt[bid]] = data[i]; // later paste wins — overwrite with the newer row
    } else {
      seenAt[bid] = unique.length;
      unique.push(data[i]);
    }
  }

  var removed = data.length - unique.length;
  if (removed > 0) {
    sheet.getRange(2, 1, data.length, lastCol).clearContent();
    if (unique.length > 0) {
      sheet.getRange(2, 1, unique.length, lastCol).setValues(unique);
    }
  }
  return removed;
}

// Dedupe "Raw Razorpay Data" by Booking ID (col A) — like
// dedupeSheetByBookingId above, but completeness-aware instead of blindly
// keeping whichever copy was pasted last. The admin export gets
// re-downloaded across overlapping date windows (see file header comment),
// and a later re-paste of the SAME payment can genuinely be an earlier
// snapshot of it — taken before Razorpay finished settling, so Fee (C) /
// Paid to CM (D) are still blank/zero on that copy. Plain last-wins would
// silently overwrite an already-settled row with an unsettled one, zeroing
// out that Booking ID's O/V/W on the next rebuild even though the money
// already came through — this is the leading suspect for a real report
// (CMQR093: Razorpay showed ₹1,435 paid, CL Payments showed ₹0).
// When BOTH copies have Fee+Paid to CM populated, the later paste still
// wins as before — completeness only overrides recency when the later copy
// is the incomplete one, so a genuine status update (captured → refunded)
// is never blocked.
function dedupeRazorpaySheetByCompleteness(sheet) {
  if (!sheet) return { removed: 0, completenessOverrides: 0 };
  var last = sheet.getLastRow();
  if (last < 2) return { removed: 0, completenessOverrides: 0 };
  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, last - 1, lastCol).getValues();

  function isPopulated(v) {
    return v !== '' && v !== null && v !== undefined && !(typeof v === 'number' && v === 0);
  }
  function isComplete(row) {
    return isPopulated(row[2]) && isPopulated(row[3]); // C: Fee, D: Paid to CM
  }

  var seenAt = {};
  var unique = [];
  var completenessOverrides = 0;
  for (var i = 0; i < data.length; i++) {
    var bid = (data[i][0] || '').toString().trim();
    if (!bid) { unique.push(data[i]); continue; }
    if (seenAt.hasOwnProperty(bid)) {
      var idx = seenAt[bid];
      if (isComplete(data[i]) || !isComplete(unique[idx])) {
        unique[idx] = data[i]; // later paste is complete (or existing never was) — normal last-wins
      } else {
        completenessOverrides++; // later paste is the incomplete one — keep the existing, complete copy
      }
    } else {
      seenAt[bid] = unique.length;
      unique.push(data[i]);
    }
  }

  var removed = data.length - unique.length;
  if (removed > 0 || completenessOverrides > 0) {
    sheet.getRange(2, 1, data.length, lastCol).clearContent();
    if (unique.length > 0) {
      sheet.getRange(2, 1, unique.length, lastCol).setValues(unique);
    }
  }
  return { removed: removed, completenessOverrides: completenessOverrides };
}

// Force a raw sheet's amount-bearing columns to Plain Text so Sheets can't
// silently reinterpret a pasted number as a date on paste — same
// protection the "Created At" column below already had, extended to every
// column that ultimately feeds TOTAL/GST/GRAND TOTAL (O/V/W). This only
// stops FUTURE corruption on new pastes; a cell that was already
// auto-converted before this ran has already lost its original numeric
// value at the moment Sheets reinterpreted it, and can't be recovered from
// here — see countDateCorruptedCells() below, which reports how many such
// cells still need a manual re-paste from the original export (real report:
// CM29309's downloaded amount was silently converted to a date).
function forceColsPlainText(sheet, colLetters, numRows) {
  if (!sheet) return;
  colLetters.forEach(function(letter) {
    sheet.getRange(2, letterToIndexTop(letter), numRows, 1).setNumberFormat('@');
  });
}

// 1-based spreadsheet column letter → index, e.g. 'A'→1, 'Y'→25, 'AG'→33.
// Top-level copy of the identically-named helper nested inside
// setupCLPayments() — needed here because forceColsPlainText/
// countDateCorruptedCells run before that nested function is in scope.
function letterToIndexTop(letter) {
  var s = 0;
  for (var i = 0; i < letter.length; i++) {
    s = s * 26 + (letter.charCodeAt(i) - 64);
  }
  return s;
}

// Floors a parseDateForSort()-style ms value to that calendar day's
// midnight. Top-level (not nested in setupCLPayments) so both Step 0.B's
// event↔property match in setupCLPayments AND the identical match key in
// exportEventLeads() use one shared definition instead of two copies that
// could quietly drift apart. The email/phone+date join is meant to
// recognize "same real-world booking" — but if either source's Booking
// Date cell ever carries a time-of-day (a timestamp instead of a bare
// date), two rows for the SAME day get different raw ms values and
// silently fail to join, leaving a genuinely-matched event stuck as
// unmatched (or, in setupCLPayments, wrongly added as a standalone row
// instead of being folded into its property row). Flooring both sides to
// the day removes that noise without weakening the match — an exact match
// still matches, this only recovers cases the exact-ms comparison missed.
function toDayKey(ms) {
  if (!ms) return 0;
  var d = new Date(ms);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}

// Count cells in a raw sheet's amount-bearing columns that are CURRENTLY a
// Date object — i.e. already corrupted by Sheets' auto-date-conversion
// before forceColsPlainText() above could stop it. Informational only (see
// that function's comment for why these can't be auto-recovered) — surfaced
// in the "Setup complete" summary so the affected rows can be found and
// manually re-pasted from the original export.
function countDateCorruptedCells(sheet, colLetters) {
  if (!sheet) return 0;
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var count = 0;
  colLetters.forEach(function(letter) {
    var col = letterToIndexTop(letter);
    var values = sheet.getRange(2, col, last - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] instanceof Date) count++;
    }
  });
  return count;
}

function setupCLPayments(silent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Locate sheets
  // Flexible lookup: exact name first, then any tab whose name starts with
  // the prefix — handles "CL Payments- automated sheet" and similar variants.
  function findSheet(exactName, prefix) {
    var s = ss.getSheetByName(exactName);
    if (s) return s;
    var all = ss.getSheets();
    for (var i = 0; i < all.length; i++) {
      if (all[i].getName().indexOf(prefix) === 0) return all[i];
    }
    return null;
  }

  var clSheet  = findSheet('CL Payments',       'CL Pay');
  var rbdSheet = findSheet('Raw Booking Data',  'Raw Booking');
  var rrSheet  = findSheet('Raw Razorpay Data', 'Raw Razorpay');
  var gstSheet = findSheet('GST PAN CM COMM',   'GST PAN');

  // Collapse repeat Booking IDs (col G) before reading — same overlapping-
  // date admin-export problem as Raw Experience/Event Booking Data below:
  // the admin report gets re-downloaded across overlapping windows and the
  // same booking ends up pasted twice. Keeps the LAST occurrence (most
  // recently pasted copy, most likely to carry an updated status).
  var rbdDupesRemoved = dedupeSheetByBookingId(rbdSheet, 7);

  // Collapse repeat Booking IDs (col A) in the raw Razorpay export too — an
  // accidental re-paste of an overlapping payments export would otherwise
  // leave the same payment as two literal rows. This only touches rows that
  // already carry a real Razorpay-native ID at this point in the flow — the
  // QR/add-on ID-recovery steps below run afterward and fill in blanks, so
  // they're untouched by this pass. Completeness-aware (see
  // dedupeRazorpaySheetByCompleteness) — NOT the plain last-wins dedup used
  // for the other raw tabs, so a re-paste of an unsettled snapshot can't
  // silently zero out an already-settled payment.
  var rrDedupeResult = dedupeRazorpaySheetByCompleteness(rrSheet);
  var rrDupesRemoved = rrDedupeResult.removed;
  var rrCompletenessOverrides = rrDedupeResult.completenessOverrides;
  // Experience/day-outing raw tab. Auto-created (with the correct header
  // row) if missing, so it's ready to paste into right away.
  var redSheet = findSheet('Raw Experience Booking Data', 'Raw Experience');
  if (!redSheet) {
    redSheet = ss.insertSheet('Raw Experience Booking Data');
    var redHeaders = [
      'Booking Date', 'Booking Month', 'Travel Start Date', 'Travel Start Month',
      'Travel End Date', 'Travel End Month', 'Booking ID', 'Property',
      'Original Property Name', 'Property Owner Name', 'User', 'User phone',
      'User email', 'Amount', 'Promocode', 'Promocode discount', 'Flat discount',
      'GST', 'Total amount', 'Total adults', 'Total kids', 'Total Non Veg Food',
      'Total Veg Food', 'Total guests', 'Status', 'Invoice number'
    ];
    redSheet.getRange(1, 1, 1, redHeaders.length).setValues([redHeaders]);
    redSheet.getRange(1, 1, 1, redHeaders.length)
      .setBackground('#F7CAAC').setFontWeight('bold').setHorizontalAlignment('center');
    redSheet.setFrozenRows(1);
  }

  // Collapse repeat Booking IDs (col G) before reading — overlapping-date
  // admin downloads mean the same booking often gets pasted more than once.
  var redDupesRemoved = dedupeSheetByBookingId(redSheet, 7);

  // Event raw tab. Auto-created if missing, same pattern as Raw Experience
  // Booking Data. NOTE: events booked through inventory sync already have a
  // real property booking in Raw Booking Data — this tab is used to MATCH
  // and ANNOTATE those existing rows (see Step 0.B below), not to add
  // duplicate rows for them. Only genuinely standalone/unmatched event
  // bookings get a new CL Payments row of their own.
  var evtSheet = findSheet('Raw Event Booking Data', 'Raw Event');
  if (!evtSheet) {
    evtSheet = ss.insertSheet('Raw Event Booking Data');
    var evtHeaders = [
      'Booking Date', 'Booking Month', 'Travel Start Date', 'Travel Start Month',
      'Travel End Date', 'Travel End Month', 'Booking ID', 'Property',
      'Original Property Name', 'Event', 'Property Owner Name', 'User',
      'User phone', 'User email', 'Amount', 'Promocode', 'Promocode discount',
      'GST', 'Total amount', 'Total guests', 'Status', 'Invoice number'
    ];
    evtSheet.getRange(1, 1, 1, evtHeaders.length).setValues([evtHeaders]);
    evtSheet.getRange(1, 1, 1, evtHeaders.length)
      .setBackground('#F7CAAC').setFontWeight('bold').setHorizontalAlignment('center');
    evtSheet.setFrozenRows(1);
  }
  var evtDupesRemoved = dedupeSheetByBookingId(evtSheet, 7);

  // OTA raw tab — mirrored in by importQualifiedOtaBookings() (called from
  // smartUpdate() before this function runs). Auto-created empty here too,
  // so a manual "Force Full Rebuild" click doesn't error out if that import
  // step hasn't run yet in this session. Columns: A Booking ID, B Booking
  // Date, C Travel Start Date, D Travel End Date, E Campsite Name,
  // F Accommodation, G Camper Name, H Total Units, I No of Nights,
  // J Platform, K Host Payout, L Qualified Status, M Settlement Date.
  // M added so CL Payments can show the payout-received date as this row's
  // Booking Date instead of the OTA guest's original booking date — see
  // buildOtaRowFormulas below. B (this tab's own "Booking Date") still holds
  // the ORIGINAL OTA booking date untouched; that's what gets quoted into
  // Comments 2 on the CL Payments row.
  var otaSheet = findSheet('Raw OTA Data', 'Raw OTA');
  if (!otaSheet) {
    otaSheet = ss.insertSheet('Raw OTA Data');
    var otaRawHeaders = [
      'Booking ID', 'Booking Date', 'Travel Start Date', 'Travel End Date',
      'Campsite Name', 'Accommodation', 'Camper Name', 'Total Units',
      'No of Nights', 'Platform', 'Host Payout', 'Qualified Status', 'Settlement Date'
    ];
    otaSheet.getRange(1, 1, 1, otaRawHeaders.length).setValues([otaRawHeaders]);
    otaSheet.getRange(1, 1, 1, otaRawHeaders.length)
      .setBackground('#F7CAAC').setFontWeight('bold').setHorizontalAlignment('center');
    otaSheet.setFrozenRows(1);
  }

  if (!clSheet || !rbdSheet || !rrSheet) {
    safeAlert(
      'ERROR: Could not find one or more required sheets.\n' +
      'Expected tabs starting with: "CL Pay", "Raw Booking Data", "Raw Razorpay Data"\n\n' +
      'Found: ' + ss.getSheets().map(function(s){ return '"' + s.getName() + '"'; }).join(', ')
    );
    return;
  }

  // Protect every amount-bearing raw column from Sheets' auto-date-
  // conversion the same way col AE (Razorpay Created At) is protected
  // below — force Plain Text on paste so a value like an amount doesn't get
  // silently reinterpreted as a date (real report: CM29309's downloaded
  // amount came through as a date). Sized generously past each sheet's
  // current data so new pastes stay covered too.
  var AMOUNT_COLS_RBD = ['O', 'Q', 'R', 'S', 'T', 'U', 'V', 'W']; // amount, promo/flat/referral/weekdays discount, wallet used, GST, total amount
  var AMOUNT_COLS_RED = ['N', 'P', 'Q', 'R', 'S'];                // amount, promo/flat discount, GST, total amount
  var AMOUNT_COLS_EVT = ['O', 'Q', 'R', 'S'];                     // amount, promo discount, GST, total amount
  var AMOUNT_COLS_RR  = ['C', 'D'];                                // Razorpay Fee, Paid to CM
  function protectRows(sheet) { return Math.max(6000, (sheet ? sheet.getLastRow() : 0) + 200); }

  var amountFormatIssuesFound =
    countDateCorruptedCells(rbdSheet, AMOUNT_COLS_RBD) +
    countDateCorruptedCells(redSheet, AMOUNT_COLS_RED) +
    countDateCorruptedCells(evtSheet, AMOUNT_COLS_EVT) +
    countDateCorruptedCells(rrSheet, AMOUNT_COLS_RR);

  forceColsPlainText(rbdSheet, AMOUNT_COLS_RBD, protectRows(rbdSheet));
  forceColsPlainText(redSheet, AMOUNT_COLS_RED, protectRows(redSheet));
  forceColsPlainText(evtSheet, AMOUNT_COLS_EVT, protectRows(evtSheet));
  forceColsPlainText(rrSheet,  AMOUNT_COLS_RR,  protectRows(rrSheet));

  // STEP 0.0 — QRv2 payments: assign a permanent CMQR#### Booking ID, then
  // collect every CMQR-tagged row (this run's new ones + prior runs') as its
  // own standalone CL Payments row. Must run before "combined" is built below
  // so freshly-assigned IDs make it into this run's CL Payments rebuild.

  // Remove failed Razorpay rows entirely
  // Failed payments shouldn't linger in Raw Razorpay Data at all — delete any
  // row where col I says "failed" (case-insensitive), even if pasted in again
  // by accident later. Runs first, before ID assignment / date cleanup /
  // classification, so nothing downstream ever has to deal with them.
  var failedRowsRemoved = 0;
  var rrLastRowForFailedScan = rrSheet.getLastRow();
  if (rrLastRowForFailedScan > 1) {
    var iColValues = rrSheet.getRange(2, 9, rrLastRowForFailedScan - 1, 1).getValues(); // col I
    var failedRowNumbers = [];
    for (var fi = 0; fi < iColValues.length; fi++) {
      var iVal = (iColValues[fi][0] || '').toString().trim().toLowerCase();
      if (iVal === 'failed') failedRowNumbers.push(fi + 2); // actual sheet row number
    }
    if (failedRowNumbers.length > 0) {
      // Delete bottom-up so earlier row numbers stay valid as we go.
      for (var fr = failedRowNumbers.length - 1; fr >= 0; fr--) {
        rrSheet.deleteRow(failedRowNumbers[fr]);
      }
      failedRowsRemoved = failedRowNumbers.length;
    }
  }

  // Force col AE (Created At) to Plain Text
  // "DD/MM/YYYY" is ambiguous with "MM/DD/YYYY" whenever the day is ≤ 12 —
  // Google Sheets silently auto-converts such pasted text into an actual date
  // using its own (often US-style) guess, swapping day/month BEFORE any
  // formula ever runs. That's why "05/07/2026" (5 Jul) was landing as 7 May —
  // the cell itself already held the wrong date; no formula can recover it
  // after the fact. Forcing Plain Text on this column stops future pastes
  // from being silently corrupted. Applied to a generous row range so it
  // covers rows not pasted yet too.
  var rrLastRowForCleanup = rrSheet.getLastRow();
  rrSheet.getRange(2, 31, 5000, 1).setNumberFormat('@');

  // One-time cleanup: un-swap already-corrupted Created At values
  // A cell that's still a genuine Date object here got auto-converted by
  // Sheets BEFORE the Plain Text format above could stop it — that only ever
  // happens when the original day was ≤ 12 (ambiguous with a month), so any
  // Date-typed cell in this column is, by construction, one of these swapped
  // values: Sheets read "DD/MM/YYYY" as "MM/DD/YYYY", storing month=originalDay
  // and day=originalMonth. Swap them back and re-write as plain text (not a
  // Date object, so it stays fixed under the new Plain Text format and reads
  // identically to every never-corrupted row). Idempotent — once every row
  // has gone through this once, there are no more Date-typed cells left to
  // find, so this becomes a no-op on future runs.
  var swappedDatesFixed = 0;
  if (rrLastRowForCleanup > 1) {
    var aeRange = rrSheet.getRange(2, 31, rrLastRowForCleanup - 1, 1);
    var aeValues = aeRange.getValues();
    var aeChanged = false;
    var pad2 = function(n) { return (n < 10 ? '0' : '') + n; };
    for (var ai = 0; ai < aeValues.length; ai++) {
      var av = aeValues[ai][0];
      if (av instanceof Date) {
        var sDay = av.getDate();
        var sMonth = av.getMonth() + 1;
        if (sDay >= 1 && sDay <= 12) {
          // swap back: true day = sMonth, true month = sDay
          aeValues[ai][0] =
            pad2(sMonth) + '/' + pad2(sDay) + '/' + av.getFullYear() + ' ' +
            pad2(av.getHours()) + ':' + pad2(av.getMinutes()) + ':' + pad2(av.getSeconds());
          aeChanged = true;
          swappedDatesFixed++;
        }
      }
    }
    if (aeChanged) aeRange.setValues(aeValues);
  }

  // Assign IDs
  // These are direct QR-code payments with no existing admin/event/experience
  // booking record, so col A (Booking ID) is blank for them. We assign a
  // synthetic ID (CMQR001, CMQR002, ...) written as a literal VALUE (not a
  // formula) so it's stable across reruns — already-assigned rows keep their
  // ID, and only newly-appearing blank+QRv2 rows get the next number in
  // sequence, continuing from whatever the highest existing CMQR number is.
  var qrNewIdsAssigned = 0;
  var descIdsRecovered = 0;
  var rrLastRow = rrSheet.getLastRow();
  var rrLastCol = rrSheet.getLastColumn();

  // "status" can live anywhere in the pasted Razorpay export (same header
  // lookup the classifier formula below uses via MATCH("status",...)) — find
  // its actual column here so the script-side check matches the formula.
  var statusColIdx = -1; // 0-based
  if (rrLastCol > 0) {
    var rrHeader = rrSheet.getRange(1, 1, 1, rrLastCol).getValues()[0];
    for (var hi = 0; hi < rrHeader.length; hi++) {
      if ((rrHeader[hi] || '').toString().trim().toLowerCase() === 'status') { statusColIdx = hi; break; }
    }
  }

  var qrReadCols = Math.max(18, statusColIdx + 1); // at least col R (18)
  if (rrLastRow > 1 && rrLastCol >= 18) {
    var rrAR = rrSheet.getRange(2, 1, rrLastRow - 1, Math.min(qrReadCols, rrLastCol)).getValues();
    var maxCMQR = 0;
    for (var qi = 0; qi < rrAR.length; qi++) {
      var aVal = (rrAR[qi][0] || '').toString().trim();
      var mQr = aVal.match(/^CMQR(\d+)$/i);
      if (mQr) maxCMQR = Math.max(maxCMQR, parseInt(mQr[1], 10));
    }
    function padCMQR(n) {
      var s = String(n);
      while (s.length < 3) s = '0' + s;
      return s;
    }
    var colAUpdated = false;
    for (var qj = 0; qj < rrAR.length; qj++) {
      var aValJ = (rrAR[qj][0] || '').toString().trim();
      var rValRawJ = (rrAR[qj][17] || '').toString();       // col R, index 17, ORIGINAL case — needed below so extracted IDs keep their real casing
      var rValJ = rValRawJ.toLowerCase();
      var statusValJ = statusColIdx >= 0 ? (rrAR[qj][statusColIdx] || '').toString().trim().toLowerCase() : '';
      if (!aValJ && rValJ.indexOf('qrv2 payment') !== -1 && statusValJ === 'captured') {
        maxCMQR++;
        rrAR[qj][0] = 'CMQR' + padCMQR(maxCMQR);
        colAUpdated = true;
        qrNewIdsAssigned++;
      } else if (!aValJ) {
        // Recover a Booking ID that's sitting right in the payment
        // description text but never made it into col A — e.g. "Accomodation
        // booking - CM28470", "Event booking - CME6548", "Property_experience
        // booking - CMX6528". This is Razorpay's own recorded description at
        // payment time, not a guess: every one of these was a real captured
        // payment for a real booking, but without col A filled in it fell
        // through to "Manual payment link" and its Razorpay ID never reached
        // the matching row in CL Payments. Only fires on a recognizable
        // "<something> booking - <ID>" pattern — a description with no such
        // pattern (e.g. a bare Razorpay reference like "#TADQDm51oZUpUO")
        // is left alone and correctly stays classified as a manual payment.
        // Guard against "Addons booking - undefined" (a real, known pattern
        // for add-on purchases with genuinely no booking ID attached — the
        // literal word "undefined" is NOT a Booking ID and must not be
        // extracted as one).
        // Character class includes "-" so hyphenated IDs like "CM-ADMIN-00113"
        // (admin manual bookings) capture in full — without it, the capture
        // stopped at the first internal hyphen ("CM"), failed the trailing
        // \s*$ anchor against the leftover "-ADMIN-00113", and the whole
        // match silently failed, leaving the row classified as "Manual
        // payment link" instead of "Manual booking via admin".
        var descMatch = rValRawJ.match(/booking\s*-\s*([A-Za-z0-9_-]+)\s*$/i);
        if (descMatch && descMatch[1].toLowerCase() !== 'undefined') {
          rrAR[qj][0] = descMatch[1];
          colAUpdated = true;
          descIdsRecovered++;
        }
      }
    }
    if (colAUpdated) {
      var colAOnly = rrAR.map(function(row) { return [row[0]]; });
      rrSheet.getRange(2, 1, colAOnly.length, 1).setValues(colAOnly);
    }
  }

  // STEP 0.D — Add-on payments: recover a Booking ID via guest email/contact
  // match. Razorpay's own description for these is literally "Addons booking
  // - undefined" (no ID attached at payment time — see the guard above), and
  // there's no shared UUID between Raw Razorpay Data's notes column
  // (booking_uuid/payment_uuid) and Raw Booking Data, so email/contact is the
  // only available join key (confirmed 2026-08-xx — admin export has no
  // booking_uuid field to match against).
  //
  // A guest can have MULTIPLE bookings, so an email/contact match alone is
  // ambiguous. Resolved by picking whichever of the guest's bookings has a
  // Travel Start Date closest to this payment's Created At — but if the two
  // closest candidates are within ADDON_CLOSE_MATCH_DAYS of each other, it's
  // left UNRESOLVED (col A stays blank, same as today, still shows as ADDON
  // PAYMENT) rather than silently guessing and risking mis-attributed
  // revenue. Same "flag what can't be resolved, don't hide it" philosophy as
  // the OTA amount-mismatch handling elsewhere in this file.
  //
  // ID format: CMADDON-<original Booking ID>-<payment_uuid prefix>. The
  // payment_uuid (from THIS row's own notes JSON) is already unique per
  // payment, so it doubles as a collision-safe suffix when the same
  // guest/booking has more than one add-on payment — no separate counter
  // needed except on the rare row where notes has no payment_uuid ("{}").
  //
  // SCOPE (v1): only matches against Raw Booking Data (property bookings).
  // Experience/Event bookings aren't included yet — if add-ons commonly
  // attach to those too, extend the lookup build-out below the same way.
  var addonsMatched = 0, addonsAmbiguous = 0;
  var ADDON_CLOSE_MATCH_DAYS = 3; // candidates this close together = don't guess

  function normEmailForAddon(e) { return (e || '').toString().trim().toLowerCase(); }
  function normPhoneForAddon(p) {
    var digits = (p || '').toString().replace(/\D/g, '');
    return digits.length >= 10 ? digits.slice(-10) : digits; // last 10 digits — tolerates country-code prefixes differing between sources
  }

  var addonBookingsByEmail = {}, addonBookingsByPhone = {};
  if (rbdSheet) {
    var rbdLastRowForAddon = rbdSheet.getLastRow();
    if (rbdLastRowForAddon >= 2) {
      // A:N — need C (Travel Start Date), G (Booking ID), M (Phone), N (Email)
      var rbdForAddon = rbdSheet.getRange(2, 1, rbdLastRowForAddon - 1, 14).getValues();
      rbdForAddon.forEach(function(row) {
        var abId = (row[6] || '').toString().trim(); // G
        if (!abId) return;
        var entry = { bookingId: abId, travelStart: parseDateForSort(row[2]) }; // C
        var aem = normEmailForAddon(row[13]); // N
        var aph = normPhoneForAddon(row[12]); // M
        if (aem) (addonBookingsByEmail[aem] = addonBookingsByEmail[aem] || []).push(entry);
        if (aph) (addonBookingsByPhone[aph] = addonBookingsByPhone[aph] || []).push(entry);
      });
    }
  }

  var rrLastRowForAddon = rrSheet.getLastRow();
  var rrLastColForAddon = rrSheet.getLastColumn();
  if (rrLastRowForAddon > 1 && rrLastColForAddon >= 26) {
    // Fresh read (not reusing rrAR above, which may be narrower than col Z) —
    // A:Z: need A (Booking ID), R (description), X (email), Y (contact),
    // Z (notes), AE (Created At) is read separately below since col 31 > 26.
    var rrAddon = rrSheet.getRange(2, 1, rrLastRowForAddon - 1, Math.max(26, rrLastColForAddon >= 31 ? 31 : 26)).getValues();
    var addonColAUpdated = false;
    for (var ad = 0; ad < rrAddon.length; ad++) {
      var arow = rrAddon[ad];
      if ((arow[0] || '').toString().trim()) continue; // already has a Booking ID
      var descTextAddon = (arow[17] || '').toString().toLowerCase(); // R
      if (descTextAddon.indexOf('addons booking - undefined') === -1) continue;

      var guestEmailAddon = normEmailForAddon(arow[23]); // X
      var guestPhoneAddon = normPhoneForAddon(arow[24]); // Y
      var addonCandidates = (guestEmailAddon && addonBookingsByEmail[guestEmailAddon]) ||
                             (guestPhoneAddon && addonBookingsByPhone[guestPhoneAddon]) || [];
      if (addonCandidates.length === 0) continue; // no match at all — leave as-is

      var addonPaymentDate = parseDateForSort(arow[30]); // AE, Created At
      var addonScored = addonCandidates.map(function(c) {
        return { bookingId: c.bookingId, diff: Math.abs(c.travelStart - addonPaymentDate) };
      }).sort(function(x, y) { return x.diff - y.diff; });

      if (addonScored.length > 1) {
        var addonGapDays = (addonScored[1].diff - addonScored[0].diff) / 86400000;
        if (addonGapDays < ADDON_CLOSE_MATCH_DAYS) { addonsAmbiguous++; continue; } // too close to call
      }

      var addonNotesText = (arow[25] || '').toString(); // Z
      var addonUuidMatch = addonNotesText.match(/"payment_uuid"\s*:\s*"([^"]+)"/i);
      var addonSuffix = addonUuidMatch ? addonUuidMatch[1].split('-')[0] : ('x' + (ad + 1));
      rrAddon[ad][0] = 'CMADDON-' + addonScored[0].bookingId + '-' + addonSuffix;
      addonColAUpdated = true;
      addonsMatched++;
    }
    if (addonColAUpdated) {
      var addonColAOnly = rrAddon.map(function(row) { return [row[0]]; });
      rrSheet.getRange(2, 1, addonColAOnly.length, 1).setValues(addonColAOnly);
    }
  }

  // Collect standalone QR-v2 rows for CL Payments
  // Every CMQR-tagged row becomes its own CL Payments row (see
  // buildQRRowFormulas below) — B–G all show the same date (col AE, "Created
  // At"), H is the CMQR ID itself, J stays blank as a manual-fill marker.
  var qrNewRowItems = [];
  var rrLastRow2 = rrSheet.getLastRow(); // re-read: IDs may have just been written above
  var rrLastCol2 = rrSheet.getLastColumn();
  if (rrLastRow2 > 1 && rrLastCol2 >= 31) { // need at least col AE (31) for Created At
    var rrForCollect = rrSheet.getRange(2, 1, rrLastRow2 - 1, rrLastCol2).getValues();
    for (var qci = 0; qci < rrForCollect.length; qci++) {
      var qcBid = (rrForCollect[qci][0] || '').toString().trim();
      if (!/^CMQR\d+$/i.test(qcBid)) continue;
      qrNewRowItems.push({
        isQR: true,
        bid: qcBid,
        dateVal: parseDateForSort(rrForCollect[qci][30]), // col AE, index 30
        rrRow: qci + 2,
        email: ''
      });
    }
  }

  // Collect standalone CM-ADMIN rows for CL Payments
  // Same situation as the QR rows just above: a booking made manually via the
  // admin panel (Payment Type classifier calls these "Manual booking via
  // admin", col A prefix "CM-ADMIN-") has NO row in Raw Booking Data at all —
  // there's no admin booking report to match it against — so it was silently
  // dropped from CL Payments entirely until now. Reuses the same rrForCollect
  // read from the QR block above (no need to re-query the sheet). B–G all
  // show the same date (col AE, "Created At" — booking date, NOT a travel
  // date, since there's no real stay/travel record to pull one from), H is
  // the CM-ADMIN ID itself, J stays blank as a manual-fill marker, same as
  // QR-v2 rows.
  var adminNewRowItems = [];
  if (rrForCollect) {
    for (var aci = 0; aci < rrForCollect.length; aci++) {
      var acBid = (rrForCollect[aci][0] || '').toString().trim();
      if (acBid.toUpperCase().indexOf('CM-ADMIN') !== 0) continue; // same LEFT(A2,8)="CM-ADMIN" prefix check the classifier formula uses
      adminNewRowItems.push({
        isAdmin: true,
        bid: acBid,
        dateVal: parseDateForSort(rrForCollect[aci][30]), // col AE, index 30
        rrRow: aci + 2,
        email: ''
      });
    }
  }

  // Set of Booking IDs with an actual Razorpay payment
  // Built here (BEFORE Experience/Event inclusion is decided below), from
  // Raw Razorpay Data col A — already reflects the QR-ID assignment AND the
  // description-based ID recovery above, since both already ran by this
  // point. Used as a second, independent reason to include an Experience or
  // Event booking even when its status text isn't "confirmed": once a real
  // payment exists for a Booking ID, that booking must never disappear from
  // CL Payments/Sales again just because someone later changes its status to
  // Cancelled/Rejected (or moves the refund to wallet vs. source — same
  // logic either way). The "confirmed-only" filter is still exactly right
  // for its original purpose — excluding pre-payment leads (draft/initiated/
  // payment_pending/payment_failed/insta_pay_intiated) that never actually
  // paid — this only ADDS a case the status-only check couldn't see.
  var razorpayIdSet = {};
  if (rrSheet) {
    var rrLastForIdSet = rrSheet.getLastRow();
    if (rrLastForIdSet > 1) {
      var rrIdsForSet = rrSheet.getRange(2, 1, rrLastForIdSet - 1, 1).getValues();
      rrIdsForSet.forEach(function(row) {
        var id = (row[0] || '').toString().trim();
        if (id) razorpayIdSet[id] = true;
      });
    }
  }
  function hasRazorpayPayment(bid) {
    return !!razorpayIdSet[(bid || '').toString().trim()];
  }

  // STEP 0.A — Merge "Raw Booking Data" + "Raw Experience Booking Data",
  //            sorted by Booking Date (col A), oldest first.
  //            Tie-break on Booking ID for stable ordering on same-day bookings.
  function parseDateForSort(val) {
    // REVERTED (v53) — v51 added a blanket "day 1-12 means Sheets swapped
    // it" correction here, assuming any already-converted Date object was
    // mis-parsed. That assumption is only actually proven for ONE source
    // (Razorpay "Created At", handled separately by the swappedDatesFixed
    // cleanup below, which already normalizes that column's raw cells before
    // anything else reads them). Applying it to EVERY date source — property/
    // experience Booking Date, OTA Settlement Date — silently corrupted
    // already-correct dates: e.g. a real "01-Jun-26" got flipped to
    // "06-Jan-26" purely because its day (1) was <=12. Confirmed wrong
    // against real data (bookings only go back to June 2026, yet random
    // Jan/Feb/Mar/Apr/May dates appeared after this "fix"). Trusting an
    // already-real Date object as-is, with no reinterpretation, is the
    // correct default — there's no reliable way to tell from the value alone
    // whether Sheets' own parse was ever ambiguous in the first place.
    if (val instanceof Date) return val.getTime();
    if (typeof val === 'number') return val; // Sheets date serial
    var s = (val || '').toString().trim();
    if (!s) return 0;
    var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);          // YYYY-MM-DD
    if (iso) return new Date(iso[1], iso[2] - 1, iso[3]).getTime();
    var dmy = s.match(/^(\d{2})-(\d{2})-(\d{4})/);           // DD-MM-YYYY
    if (dmy) return new Date(dmy[3], dmy[2] - 1, dmy[1]).getTime();
    var dmySlash = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);    // DD/MM/YYYY (Razorpay "Created At")
    if (dmySlash) return new Date(dmySlash[3], dmySlash[2] - 1, dmySlash[1]).getTime();
    var d = new Date(s);
    return isNaN(d) ? 0 : d.getTime();
  }

  // Raw Booking Data and Raw Experience Booking Data have GENUINELY DIFFERENT
  // column layouts — the experience admin export doesn't include Accommodation,
  // Referral/Weekdays Discount, Wallet Amount Used, Total/Insta Units, Extra
  // Adults/Kids, or Total Pets, and its remaining columns sit at different
  // letters. Each source gets its own field→column map; null means "this
  // field doesn't exist in this source" (rendered as blank/0 in CL Payments).
  //
  // Field names below correspond 1:1 to the CL Payments columns that use them
  // (see buildRowFormulas). Date/month columns A–G are identical across both
  // sources, so they're not part of this map.
  var PROPERTY_MAP = {
    bookingId: 'G',
    campsiteName: 'H', accommodation: 'J', campOwner: 'K', camperName: 'L',
    camperPhone: 'M', camperEmail: 'N', amount: 'O', promoName: 'P', promoDiscount: 'Q',
    flatDiscount: 'R', referralDiscount: 'S', weekdaysDiscount: 'T', walletUsed: 'U', gst: 'V',
    totalAmount: 'W', totalUnits: 'X', instaUnits: 'Y', totalAdults: 'Z', totalKids: 'AA',
    extraAdults: 'AB', extraKids: 'AC', totalGuests: 'AD', totalPets: 'AE', noOfNights: 'AF',
    status: 'AG'
  };

  // Raw Experience Booking Data — 26 columns, per the actual admin export:
  // Booking Date, Booking Month, Travel Start Date, Travel Start Month,
  // Travel End Date, Travel End Month, Booking ID, Property, Original Property
  // Name, Property Owner Name, User, User phone, User email, Amount, Promocode,
  // Promocode discount, Flat discount, GST, Total amount, Total adults,
  // Total kids, Total Non Veg Food, Total Veg Food, Total guests, Status,
  // Invoice number.
  var EXPERIENCE_MAP = {
    bookingId: 'G',
    campsiteName: 'H', accommodation: null, campOwner: 'J', camperName: 'K',
    camperPhone: 'L', camperEmail: 'M', amount: 'N', promoName: 'O', promoDiscount: 'P',
    flatDiscount: 'Q', referralDiscount: null, weekdaysDiscount: null, walletUsed: null, gst: 'R',
    totalAmount: 'S', totalUnits: null, instaUnits: null, totalAdults: 'T', totalKids: 'U',
    extraAdults: null, extraKids: null, totalGuests: 'X', totalPets: null, noOfNights: null,
    status: 'Y',
    // No real Accommodation column for this source — synthesize a label
    // instead: "Day Outing at <Campsite Name>" (see clIdx===11 in buildRowFormulas).
    accommodationOverride: { type: 'label', prefix: 'Day Outing at ', field: 'campsiteName' }
  };

  // Raw Event Booking Data — 22 columns, per the actual admin export:
  // Booking Date, Booking Month, Travel Start Date, Travel Start Month,
  // Travel End Date, Travel End Month, Booking ID, Property, Original
  // Property Name, Event, Property Owner Name, User, User phone, User email,
  // Amount, Promocode, Promocode discount, GST, Total amount, Total guests,
  // Status, Invoice number.
  //
  // Only used to build a NEW CL Payments row for events that have NO
  // matching property booking (see Step 0.B) — standalone event tickets.
  // Events matched to an existing property booking never reach buildRowFormulas
  // via this map; they annotate the property row's Comments 2 instead.
  var EVENT_MAP = {
    bookingId: 'G',
    campsiteName: 'H', accommodation: null, campOwner: 'K', camperName: 'L',
    camperPhone: 'M', camperEmail: 'N', amount: 'O', promoName: 'P', promoDiscount: 'Q',
    flatDiscount: null, referralDiscount: null, weekdaysDiscount: null, walletUsed: null, gst: 'R',
    totalAmount: 'S', totalUnits: null, instaUnits: null, totalAdults: null, totalKids: null,
    extraAdults: null, extraKids: null, totalGuests: 'T', totalPets: null, noOfNights: null,
    status: 'U',
    eventName: 'J', // raw "Event" column — the descriptive title, e.g. "Bande Sunday at Camp Monk Bannerghatta (19 July 2026)"
    // No Accommodation column either, but unlike Experience, the raw data
    // already has a perfectly descriptive title — show it directly instead
    // of synthesizing one.
    accommodationOverride: { type: 'field', field: 'eventName' }
  };

  // 1-based spreadsheet column letter → index, e.g. 'A'→1, 'Y'→25, 'AG'→33.
  function letterToIndex(letter) {
    var s = 0;
    for (var i = 0; i < letter.length; i++) {
      s = s * 26 + (letter.charCodeAt(i) - 64);
    }
    return s;
  }

  // statusFilter (optional): function(rawStatusString, bid) → boolean. When
  // provided, only rows whose status field (or Booking ID, e.g. for a
  // has-a-Razorpay-payment override) passes the filter are collected — used
  // to restrict "Raw Experience Booking Data" to confirmed (or ever-paid)
  // bookings only.
  function collectSource(sheet, fieldMap, statusFilter) {
    var items = [];
    if (!sheet) return items;
    var last = sheet.getLastRow();
    if (last < 2) return items;
    var statusIdx = fieldMap.status ? letterToIndex(fieldMap.status) : 0;
    var emailIdx  = fieldMap.camperEmail ? letterToIndex(fieldMap.camperEmail) : 0;
    var width = Math.max(7, statusIdx, emailIdx); // need at least col A (date) & G (Booking ID)
    var data = sheet.getRange(2, 1, last - 1, width).getValues();
    var sheetRef = "'" + sheet.getName() + "'";
    for (var i = 0; i < data.length; i++) {
      var bid = data[i][6];
      if (!bid) continue;
      if (statusFilter) {
        var statusVal = statusIdx ? (data[i][statusIdx - 1] || '').toString() : '';
        if (!statusFilter(statusVal, bid.toString())) continue;
      }
      var emailVal = emailIdx ? (data[i][emailIdx - 1] || '').toString() : '';
      items.push({ sheetRef: sheetRef, row: i + 2, dateVal: parseDateForSort(data[i][0]), bid: bid.toString(), map: fieldMap, email: emailVal });
    }
    return items;
  }

  // Test-account bookings (staff/QA emails) should never inflate CL Payments
  // unless a real Razorpay payment actually exists for that booking.
  var TEST_EMAIL_DOMAINS = ['@think201.com', '@campmonk.com'];
  function isTestEmail(email) {
    var e = (email || '').toString().trim().toLowerCase();
    if (!e) return false;
    for (var ti = 0; ti < TEST_EMAIL_DOMAINS.length; ti++) {
      if (e.indexOf(TEST_EMAIL_DOMAINS[ti]) !== -1) return true;
    }
    return false;
  }

  // Property/event rows: unfiltered (Payment Pending / Rejected etc. are
  // still tracked in CL Payments for that source, as before).
  // Experience rows: confirmed only — everything else goes to Experience
  // Leads instead (see exportExperienceLeads()).
  // STEP 0.B — Match "Raw Event Booking Data" rows to their underlying
  // property booking in "Raw Booking Data" (same guest email/phone + same
  // Booking Date — inventory sync creates both records together).
  //
  //   MATCHED   → this event and its property booking are the SAME real
  //               transaction. Adding a second row would double-count
  //               revenue, but the property row's own Booking ID (CM...) is
  //               useless for Razorpay matching — Razorpay only ever has the
  //               event's ID (CME...) — and the property row's Amount is
  //               typically 0 (a placeholder). So the property row's H gets
  //               REPLACED with the event ID, and O/V pull straight from the
  //               event's own Amount/GST — see eventOverrideByBookingId,
  //               used inside buildRowFormulas via the financeOverride param.
  //
  //   UNMATCHED + confirmed → a standalone event ticket with no accommodation
  //               attached, i.e. genuinely not in Raw Booking Data at all.
  //               This gets its own new CL Payments row via EVENT_MAP, using
  //               its own Booking ID (e.g. "CME6551") directly for Razorpay
  //               matching — same pattern as standalone Experience bookings.
  //
  //   UNMATCHED + not confirmed → a dropped/failed event attempt with no
  //               accommodation tie. Not added to CL Payments (mirrors the
  //               confirmed-only rule for Experience) — a candidate for a
  //               future "Event Leads" tracker, not built yet.
  function readRawValues(sheet, width) {
    if (!sheet) return [];
    var last = sheet.getLastRow();
    if (last < 2) return [];
    return sheet.getRange(2, 1, last - 1, width).getValues();
  }

  // Raw Booking Data lookups, keyed by "email|dayKey" / "phone|dayKey".
  // toDayKey() is the top-level helper defined near letterToIndexTop above.
  // Only need cols A (date), G (Booking ID), M (phone), N (email) → width 14.
  var rbdValues = readRawValues(rbdSheet, 14);
  var propByEmailDate = {};
  var propByPhoneDate = {};
  for (var pi = 0; pi < rbdValues.length; pi++) {
    var prow = rbdValues[pi];
    var pbid = prow[6]; // G: Booking ID
    if (!pbid) continue;
    var pdate  = toDayKey(parseDateForSort(prow[0]));   // A: Booking Date, day-only
    var pemail = (prow[13] || '').toString().trim().toLowerCase(); // N: User email
    var pphone = (prow[12] || '').toString().trim();               // M: User phone
    if (pemail) {
      var ek = pemail + '|' + pdate;
      (propByEmailDate[ek] = propByEmailDate[ek] || []).push(pbid.toString());
    }
    if (pphone) {
      var pk = pphone + '|' + pdate;
      (propByPhoneDate[pk] = propByPhoneDate[pk] || []).push(pbid.toString());
    }
  }

  var evtValues = readRawValues(evtSheet, 22);
  // property Booking ID -> { sheetRef, row, map, bid, note } — Razorpay only
  // ever has the EVENT's ID (CME6540), never the property's (CM28192), so a
  // matched property row needs its Booking ID (H) REPLACED with the event ID
  // (that's what makes AH/AJ/AK/AV resolve against Raw Razorpay Data), its
  // TOTAL/GST (O/V) pulled from the event's own Amount/GST — because the
  // property side is typically a zero-value inventory placeholder, all the
  // real money sits on the event row — and the ORIGINAL property Booking ID
  // preserved in Comments 2 (AT), since H no longer shows it.
  var eventOverrideByBookingId = {};
  var eventNewRowItems = [];     // unmatched + confirmed -> new CL Payments rows
  var matchedEventCount = 0;
  var eventSkippedOld = 0;

  // Only June 2026 onward is relevant — older event data (including the
  // legacy "e_xxx" ID era) predates what this sheet needs to reconcile.
  var EVENT_CUTOFF = new Date(2026, 5, 1).getTime(); // June 1, 2026

  for (var ei = 0; ei < evtValues.length; ei++) {
    var erow = evtValues[ei];
    var ebid = erow[6]; // G: this event's own Booking ID
    if (!ebid) continue;
    var eDateVal = parseDateForSort(erow[0]);              // A: Booking Date
    if (eDateVal < EVENT_CUTOFF) { eventSkippedOld++; continue; }
    var eEmail   = (erow[13] || '').toString().trim().toLowerCase(); // N: User email
    var ePhone   = (erow[12] || '').toString().trim();               // M: User phone
    var eStatus  = (erow[20] || '').toString().toLowerCase().trim(); // U: Status
    var eDayKey  = toDayKey(eDateVal); // day-only, for the join — see toDayKey() above

    var matchedProps = (eEmail && propByEmailDate[eEmail + '|' + eDayKey]) ||
                        (ePhone && propByPhoneDate[ePhone + '|' + eDayKey]) ||
                        null;

    // Matching requires the EVENT to actually be confirmed OR already have a
    // real Razorpay payment — not just "eStatus === 'confirmed'" for
    // standalone events below, but here as well. Before this fix, a matched
    // property row's identity/status got overridden by the event's ID
    // regardless of the event's own payment state, so an event still sitting
    // at "insta_pay_intiated" (or "payment_failed", etc.) would silently
    // inherit the PROPERTY's "Booking Completed" status — showing a booking
    // as complete when the event add-on itself hadn't actually been paid
    // for. Caught from a real example: CME6549 (status "insta_pay_intiated")
    // was matched to Kedar Rao's property booking and showed "Booking
    // Completed" on CL Payments.
    //
    // The hasRazorpayPayment() half exists so a booking that WAS confirmed
    // and paid, then later manually changed to Cancelled/Rejected (refund to
    // wallet or source — same either way), never disappears from CL
    // Payments/Sales — a real payment is proof this was a genuine booking,
    // regardless of what the status text says now. A booking with neither
    // "confirmed" status NOR a payment (a pure pre-payment lead) is still
    // correctly left out — it'll pick up the override automatically on a
    // later run if it's ever actually confirmed or paid.
    if (matchedProps && matchedProps.length > 0 && (eStatus === 'confirmed' || hasRazorpayPayment(ebid.toString()))) {
      matchedProps.forEach(function(propBid) {
        // First match wins if a property row somehow matches more than one
        // event row (rare) — avoids overwriting an existing override.
        if (!eventOverrideByBookingId[propBid]) {
          eventOverrideByBookingId[propBid] = {
            sheetRef: "'" + evtSheet.getName() + "'",
            row: ei + 2,
            map: EVENT_MAP,
            bid: ebid.toString(),
            note: 'Property Booking ID: ' + propBid,
            email: eEmail
          };
        }
      });
      matchedEventCount++;
    } else if (!(matchedProps && matchedProps.length > 0) && (eStatus === 'confirmed' || hasRazorpayPayment(ebid.toString()))) {
      eventNewRowItems.push({
        sheetRef: "'" + evtSheet.getName() + "'",
        row: ei + 2,
        dateVal: eDateVal,
        bid: ebid.toString(),
        map: EVENT_MAP,
        email: eEmail
      });
    }
  }

  // Collect OTA rows (already-qualified, from Raw OTA Data)
  // One row per Booking ID (= OTA Confirmation Code) — Host Payout (net,
  // post-commission) is the "matched total" AV pulls via VLOOKUP, same role
  // Razorpay's AJ+AK plays for QR rows. See buildOtaRowFormulas below.
  var otaNewRowItems = [];
  var otaLastRow = otaSheet.getLastRow();
  if (otaLastRow > 1) {
    // 13 cols (A:M), not 12 — need column M (Settlement Date) for the sort
    // key below. Reading only A:L (12 cols) was the bug: it silently made
    // otaValues[oti][12] undefined, so dateVal fell back to sorting by the
    // GUEST's original booking date while the actual displayed "Booking
    // Date" formula in buildOtaRowFormulas() shows Settlement Date (col M) —
    // two different dates driving sort order vs. display, which is exactly
    // what scrambled the CL Payments row sequence out of chronological order.
    var otaValues = otaSheet.getRange(2, 1, otaLastRow - 1, 13).getValues();
    for (var oti = 0; oti < otaValues.length; oti++) {
      var otaBid = (otaValues[oti][0] || '').toString().trim();
      if (!otaBid) continue;
      otaNewRowItems.push({
        isOta: true,
        bid: otaBid,
        // Sort key MUST match what column B actually displays for this row
        // (Settlement Date, col M) — not the OTA guest's original booking
        // date (col B) — otherwise sort order and displayed dates diverge.
        dateVal: parseDateForSort(otaValues[oti][12]), // M: Settlement Date
        otaRow: oti + 2,
        platform: (otaValues[oti][9] || '').toString().trim(), // J: Platform
        email: ''
      });
    }
  }

  var combined = collectSource(rbdSheet, PROPERTY_MAP)
    .concat(collectSource(redSheet, EXPERIENCE_MAP, function(status, bid) {
      return status.toLowerCase().trim() === 'confirmed' || hasRazorpayPayment(bid);
    }))
    .concat(eventNewRowItems)
    .concat(qrNewRowItems)
    .concat(adminNewRowItems)
    .concat(otaNewRowItems);

  // Attach the matched-event override onto property-sourced items only —
  // this is what makes buildRowFormulas swap in the CME Booking ID + event
  // financials for these specific rows (see Step 0.B above).
  combined.forEach(function(item) {
    if (item.map === PROPERTY_MAP && eventOverrideByBookingId[item.bid]) {
      item.financeOverride = eventOverrideByBookingId[item.bid];
    }
  });

  // Exclude test-account bookings — UNLESS a real Razorpay payment exists
  // Staff/QA bookings (@think201.com, @campmonk.com) shouldn't inflate CL
  // Payments, but if money genuinely moved through Razorpay for one (e.g. a
  // real test purchase), it stays — that's real revenue regardless of whose
  // email is on it. Checked against the EFFECTIVE Booking ID (the swapped-in
  // event ID for matched rows, since that's what H will actually show).
  // Reuses razorpayIdSet/hasRazorpayPayment() built earlier (before
  // Experience/Event inclusion was decided) — same lookup, no need to
  // rebuild it a second time.
  var testBookingsExcluded = 0;
  combined = combined.filter(function(item) {
    // For matched rows, the event side's email is the one that actually drove
    // the swapped Booking ID + financials — check that, not the (possibly
    // different/blank) email on the original property row.
    var effectiveEmail = item.financeOverride ? item.financeOverride.email : item.email;
    if (!isTestEmail(effectiveEmail)) return true;
    var effectiveBid = item.financeOverride ? item.financeOverride.bid : item.bid;
    if (razorpayIdSet[effectiveBid]) return true; // real payment exists — keep it
    testBookingsExcluded++;
    return false;
  });

  combined.sort(function(a, b) {
    if (a.dateVal !== b.dateVal) return a.dateVal - b.dateVal;
    return a.bid < b.bid ? -1 : (a.bid > b.bid ? 1 : 0);
  });

  var dataRowCount = combined.length;
  var NUM_ROWS = Math.max(1000, dataRowCount); // formula/format rows (rows 2–NUM_ROWS+1)

  // STEP 0 — Create "Master GST Data" tab if it doesn't exist
  var gstRateSheet = ss.getSheetByName('Master GST Data');
  if (!gstRateSheet) {
    gstRateSheet = ss.insertSheet('Master GST Data');

    gstRateSheet.getRange(1, 1, 1, 3).setValues([['Accommodation Name', 'GST Divisor (105 or 118)', 'Notes']]);
    gstRateSheet.getRange(1, 1, 1, 3)
      .setBackground('#E2EFDA')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    // Divisor = 100 + GST rate: 105 = 5% GST, 118 = 18% GST
    gstRateSheet.getRange(2, 1, 4, 3).setValues([
      ['[Paste accommodation name here]', 105, '5% GST — tariff < ₹7500/night'],
      ['[Paste accommodation name here]', 118, '18% GST — tariff ≥ ₹7500/night'],
      ['[Add more rows as needed — one per accommodation]', '', ''],
      ['', '', '']
    ]);

    gstRateSheet.setColumnWidth(1, 280);
    gstRateSheet.setColumnWidth(2, 120);
    gstRateSheet.setColumnWidth(3, 240);
    gstRateSheet.setFrozenRows(1);
  }

  // STEP 0.C — Snapshot manually-overwritten cells before wiping CL Payments,
  // so hand-typed corrections survive this rebuild. A cell counts as "manual"
  // when it currently holds a non-blank literal value with NO formula behind
  // it — every auto-generated cell in this sheet is always a formula (even
  // the "blank" ones push a formula that resolves to ""), so this only ever
  // catches genuine hand-typed edits. Keyed by Booking ID (col H) rather than
  // row number, since row position shifts whenever the date-sorted order
  // changes between runs.
  var manualOverridesByBookingId = {};
  var oldCLLastRow = clSheet.getLastRow();
  if (oldCLLastRow > 1) {
    // Bounded by the sheet's ACTUAL current column count, not the target 52 —
    // the very first run after a column-count change (e.g. adding AZ here)
    // finds the sheet still at its old width, and requesting more columns
    // than physically exist throws. Anything beyond the old width simply
    // didn't exist yet, so there's nothing to snapshot there anyway.
    var oldMaxCol = Math.min(52, clSheet.getMaxColumns());
    var oldCLFormulas = clSheet.getRange(2, 1, oldCLLastRow - 1, oldMaxCol).getFormulas();
    var oldCLValues   = clSheet.getRange(2, 1, oldCLLastRow - 1, oldMaxCol).getValues();
    for (var oi = 0; oi < oldCLValues.length; oi++) {
      var oldBid = (oldCLValues[oi][7] || '').toString().trim(); // col H = index 7
      if (!oldBid) continue;
      var rowOverrides = null;
      for (var oc = 0; oc < oldMaxCol; oc++) {
        var oldFormulaStr = oldCLFormulas[oi][oc];
        var hasFormula = oldFormulaStr && oldFormulaStr.toString().charAt(0) === '=';
        var oldVal = oldCLValues[oi][oc];
        var isBlank = (oldVal === '' || oldVal === null || typeof oldVal === 'undefined');
        if (!hasFormula && !isBlank) {
          if (!rowOverrides) rowOverrides = {};
          rowOverrides[oc] = oldVal;
        }
      }
      if (rowOverrides) manualOverridesByBookingId[oldBid] = rowOverrides;
    }
  }

  // STEP 1 — Clear CL Payments and resize to 52 columns
  clSheet.clearContents();
  clSheet.clearFormats();
  clSheet.clearConditionalFormatRules();

  var maxCol = clSheet.getMaxColumns();
  if (maxCol > 52) {
    clSheet.deleteColumns(53, maxCol - 52);
  }
  if (clSheet.getMaxColumns() < 52) {
    clSheet.insertColumnsAfter(clSheet.getMaxColumns(), 52 - clSheet.getMaxColumns());
  }

  // STEP 2 — Write headers (row 1)
  //
  // 52-column layout:
  //   A(1)  : Sr No
  //   B(2)  : Booking Date
  //   C(3)  : Booking Mth
  //   D(4)  : Travel Start Date
  //   E(5)  : Travel Mth
  //   F(6)  : Travel End Date
  //   G(7)  : Travel End Mth
  //   H(8)  : Booking ID
  //   I(9)  : Campsite Name
  //   J(10) : Accommodation          ← GST lookup key (Master GST Data)
  //   K(11) : Camp Owner Name
  //   L(12) : Camper Name
  //   M(13) : Camper Phone
  //   N(14) : Camper Email           ← used in insta_pay conversion check
  //   O(15) : TOTAL                  ← reverse-calced from AV when promo discount detected
  //   P(16) : Promocode Name
  //   Q(17) : Promocode Discount
  //   R(18) : Flat Discount
  //   S(19) : Referral Discount
  //   T(20) : Weekdays Discount
  //   U(21) : Wallet Amount Used
  //   V(22) : GST                    ← reverse-calced from AV when promo discount detected
  //   W(23) : GRAND TOTAL            = O + V
  //   X(24) : Total Units
  //   Y(25) : Insta Units
  //   Z(26) : Total Adults
  //   AA(27): Total Kids
  //   AB(28): Extra Adults
  //   AC(29): Extra Kids
  //   AD(30): Total Guests
  //   AE(31): Total Pets
  //   AF(32): No of Nights
  //   AG(33): Booking Status         ← auto-classified
  //   AH(34): Razorpay ID            ← actual Booking ID from Razorpay, or N/A
  //   AI(35): PAN No
  //   AJ(36): Razorpay Fee
  //   AK(37): Paid to CM
  //   AL(38): CL Has GST
  //   AM(39): CM Comm %
  //   AN(40): CM Comm (excl GST)     = O × AM
  //   AO(41): CM GST
  //   AP(42): CL Transfer before TDS = W − AN − AO
  //   AQ(43): TDS Deducted
  //   AR(44): CL Transfer after TDS  = AP − AQ
  //   AS(45): Comments 1             ← auto: booking type + wallet note
  //   AT(46): Comments 2             ← manual
  //   AU(47): Comments 3             ← manual
  //   AV(48): Match Total            = AJ + AK (total Razorpay received)
  //   AW(49): Final Status
  //   AX(50): Refund Amount          ← manual, entered per cancelled/refunded
  //                                     booking (partial or full) — the ₹ figure
  //   AY(51): Refund Type            ← manual dropdown: "Wallet Credit" (stays
  //                                     as reusable customer credit, doesn't
  //                                     leave the business) or "Source Refund"
  //                                     (sent back to original payment method,
  //                                     actually leaves the business). Wallet
  //                                     credit that later gets spent shows up
  //                                     as Wallet Amount Used (U) on a
  //                                     DIFFERENT row — the two are only
  //                                     linkable by netting the totals, not
  //                                     row by row
  //   AZ(52): Influencer Booking     ← manual dropdown: "Yes"/"No" (blank =
  //                                     No). "Yes" forces O/V/W (TOTAL/GST/
  //                                     GRAND TOTAL) to 0 on that row, adds
  //                                     "Influencer Booking" to Comments 1
  //                                     (AS) in place of any wallet note, and
  //                                     the row gets a dedicated color that
  //                                     takes priority over the Wallet/LP
  //                                     full-row highlight where both would
  //                                     otherwise apply.

  var headers = [
    'Sr No',                   // A  (1)
    'Booking Date',            // B  (2)
    'Booking Mth',             // C  (3)
    'Travel Start Date',       // D  (4)
    'Travel Mth',              // E  (5)
    'Travel End Date',         // F  (6)
    'Travel End Mth',          // G  (7)
    'Booking ID',              // H  (8)
    'Campsite Name',           // I  (9)
    'Accomodation',            // J  (10)
    'Camp Owner Name',         // K  (11)
    'Camper Name',             // L  (12)
    'Camper Phone',            // M  (13)
    'Camper Email',            // N  (14)
    'TOTAL',                   // O  (15) ← effective TOTAL (reverse-calced if promo)
    'Promocode Name',          // P  (16)
    'Promocode Discount',      // Q  (17)
    'Flat Discount',           // R  (18)
    'Referral Discount',       // S  (19)
    'Weekdays Discount',       // T  (20)
    'Wallet Amount Used',      // U  (21)
    'GST',                     // V  (22) ← effective GST (reverse-calced if promo)
    'GRAND TOTAL',             // W  (23) = O + V
    'Total Units',             // X  (24)
    'Insta Units',             // Y  (25)
    'Total Adults',            // Z  (26)
    'Total Kids',              // AA (27)
    'Extra Adults',            // AB (28)
    'Extra Kids',              // AC (29)
    'Total Guests',            // AD (30)
    'Total Pets',              // AE (31)
    'No of Nights',            // AF (32)
    'Booking Status',          // AG (33) ← auto-classified
    'Razorpay ID',             // AH (34) ← Booking ID from Razorpay, or N/A
    'PAN No',                  // AI (35)
    'Razorpay Fee',            // AJ (36)
    'Paid to CM',              // AK (37)
    'CL Has GST',              // AL (38)
    'CM Comm %',               // AM (39)
    'CM Comm (excl GST)',      // AN (40)
    'CM GST',                  // AO (41)
    'CL Transfer before TDS',  // AP (42)
    'TDS Deducted',            // AQ (43)
    'CL Transfer after TDS',   // AR (44)
    'Comments 1',              // AS (45) ← auto: booking type + wallet note
    'Comments 2',              // AT (46) ← manual
    'Comments 3',              // AU (47) ← manual
    'Match Total',             // AV (48) = AJ + AK
    'Final Status',            // AW (49)
    'Refund Amount',           // AX (50) ← manual, ₹ amount for Cancelled/Refunded rows
    'Refund Type',             // AY (51) ← manual dropdown: Wallet Credit / Source Refund
    'Influencer Booking'       // AZ (52) ← manual dropdown: Yes/No — Yes zeroes O/V/W
  ];

  clSheet.getRange(1, 1, 1, 52).setValues([headers]);

  // STEP 3 — Build one formula-row per merged/sorted booking
  //
  // Unlike a plain "write row 2, copy down" template, every row here can point
  // at a DIFFERENT source sheet (Raw Booking Data or Raw Experience Booking
  // Data) and a DIFFERENT source row — whichever booking sits at that position
  // in the date-sorted "combined" list built in Step 0.A. buildRowFormulas()
  // takes the destination CL Payments row (r), the source sheet reference
  // (RAW, already quoted), and the source row (sr), and returns the same 52
  // formulas as before, just parameterized instead of hardcoded to row 2 /
  // 'Raw Booking Data'.

  // Helper: 1-based column index → letter(s)
  function col(n) {
    var s = '';
    while (n > 0) {
      n--;
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26);
    }
    return s;
  }

  // Raw Booking Data / Raw Experience Booking Data share the same column map (1-indexed):
  //   A=Booking Date, B=Booking Mth, C=Travel Start Date, D=Travel Start Mth,
  //   E=Travel End Date, F=Travel End Mth, G=Booking ID, H=Campsite Name,
  //   I=Original Campsite (skipped in CL Payments), J=Accommodation,
  //   K=Camp Owner, L=Camper Name, M=Phone, N=Email,
  //   O=TOTAL, P=Promocode Name, Q=Promo Disc, R=Flat Disc,
  //   S=Referral Disc, T=Weekdays Disc, U=Wallet Used, V=GST,
  //   W=GRAND TOTAL, X=Total Units, Y=Insta Units,
  //   Z=Total Adults, AA=Total Kids, AB=Extra Adults, AC=Extra Kids,
  //   AD=Total Guests, AE=Total Pets, AF=No of Nights, AG=Status

  function buildRowFormulas(r, RAW, sr, map, bid, financeOverride) {
    var f = [];

    // Reference to a mapped raw field: RAW!<col><sr>, or a literal fallback
    // ("" / 0) when this source doesn't have that field at all.
    function rawRef(fieldName, fallback) {
      var letter = map[fieldName];
      return letter ? (RAW + "!" + letter + sr) : fallback;
    }

    // Same, but into the matched EVENT source instead of the primary RAW
    // source — used only when financeOverride is set (see Step 0.B).
    function financeRef(fieldName, fallback) {
      if (!financeOverride) return rawRef(fieldName, fallback);
      var letter = financeOverride.map[fieldName];
      return letter ? (financeOverride.sheetRef + "!" + letter + financeOverride.row) : fallback;
    }

    // A(1): Sr No
    f.push("=IF(" + RAW + "!G" + sr + "=\"\",\"\",ROW()-1)");

    // Date & Month columns B–G
    // Identical positions (A/C/E) in both Raw Booking Data and Raw Experience
    // Booking Data, so no map lookup needed here.
    // REVERTED (v53) — v52 added a day<=12 "un-swap" whenever this cell was
    // already a real date, on the theory that Sheets had silently mis-parsed
    // it (same assumption as the QR fix). That assumption isn't actually
    // proven for this source and it corrupted already-correct dates (e.g.
    // real June 2026 bookings showing as January/February/March/April/May)
    // — confirmed wrong against live data. Trusting an already-real date
    // as-is is correct here. What DOES stay from v52: for text values, try
    // the unambiguous manual position-based parse BEFORE DATEVALUE() (which
    // has its own locale-guessing risk) rather than after — that ordering
    // change was never the problem, only the ISNUMBER swap was.
    function parseDateExpr(rawCol) {
      var ref = RAW + "!" + rawCol + sr;
      return "IF(ISNUMBER(" + ref + ")," + ref + "," +
               "IFERROR(DATE(VALUE(RIGHT(" + ref + ",4)),VALUE(MID(" + ref + ",4,2)),VALUE(LEFT(" + ref + ",2)))," +
                 "IFERROR(DATEVALUE(" + ref + "),\"\")))";
    }
    function dateFormula(rawCol) {
      return "=IF(" + RAW + "!G" + sr + "=\"\",\"\"," +
             "TEXT(" + parseDateExpr(rawCol) + ",\"DD- MMM- YY\"))";
    }
    function monthFormula(rawCol) {
      return "=IF(" + RAW + "!G" + sr + "=\"\",\"\"," +
             "TEXT(" + parseDateExpr(rawCol) + ",\"MMM-YY\"))";
    }

    f.push(dateFormula("A"));   // B(2): Booking Date
    f.push(monthFormula("A"));  // C(3): Booking Mth
    f.push(dateFormula("C"));   // D(4): Travel Start Date
    f.push(monthFormula("C"));  // E(5): Travel Mth
    f.push(dateFormula("E"));   // F(6): Travel End Date
    f.push(monthFormula("E"));  // G(7): Travel End Mth

    // H(8) through AF(32): booking data cols, via the source's field map
    // Each entry: [CL field name, is-text-when-missing (vs numeric 0)]
    var plan = [
      // 'bookingId' (H) handled specially below — literal event ID override
      // for matched rows, real value otherwise
      ['campsiteName', ''],
      // 'accommodation' (J) handled specially below — real value for
      // property/event, synthesized "Day Outing at <Campsite>" for experience
      ['campOwner', ''], ['camperName', ''],
      ['camperPhone', ''], ['camperEmail', ''],
      // 'amount' (O) handled specially below — TOTAL reverse-calc
      ['promoName', ''], ['promoDiscount', 0], ['flatDiscount', 0],
      ['referralDiscount', 0], ['weekdaysDiscount', 0], ['walletUsed', 0],
      // 'gst' (V) and 'totalAmount' (W) handled specially below
      ['totalUnits', 0], ['instaUnits', 0], ['totalAdults', 0], ['totalKids', 0],
      ['extraAdults', 0], ['extraKids', 0], ['totalGuests', 0], ['totalPets', 0], ['noOfNights', 0]
    ];
    var planIdx = 0;

    for (var clIdx = 8; clIdx <= 33; clIdx++) {
      if (clIdx === 10) continue; // skip Original Campsite Name — not in CL Payments

      if (clIdx === 8) {
        // H(8): Booking ID — for a matched property row, Razorpay only ever
        // has the EVENT's ID, never the property's, so this gets REPLACED
        // with the event's Booking ID (financeOverride.bid) instead of the
        // property's own. Everything downstream (AH Razorpay lookup, AS
        // booking-type classification, the Event color code) keys off H, so
        // this one swap makes all of it resolve correctly automatically.
        if (financeOverride) {
          f.push("=IF(" + RAW + "!G" + sr + "=\"\",\"\",\"" + financeOverride.bid.replace(/"/g, '""') + "\")");
        } else {
          f.push("=IF(" + RAW + "!G" + sr + "=\"\",\"\"," + rawRef('bookingId', '""') + ")");
        }
      } else if (clIdx === 11) {
        // J(10): Accommodation — property sources have a real accommodation
        // column. Sources without one (experience, standalone events) fall
        // back to map.accommodationOverride: either a synthesized label
        // ("Day Outing at <Campsite Name>") or a direct field passthrough
        // (the raw Event title, already descriptive on its own). Matched
        // rows (financeOverride set) ALWAYS show the event title here too —
        // that's the more useful piece of information for a booking that's
        // fundamentally an event booking, same treatment as standalone events.
        if (financeOverride) {
          var evtTitleRef = financeRef('eventName', '""');
          f.push("=IF(" + RAW + "!G" + sr + "=\"\",\"\"," + evtTitleRef + ")");
        } else if (map.accommodation) {
          f.push("=IF(" + RAW + "!G" + sr + "=\"\",\"\"," + RAW + "!" + map.accommodation + sr + ")");
        } else if (map.accommodationOverride && map.accommodationOverride.type === 'label') {
          var labelRef = rawRef(map.accommodationOverride.field, '""');
          f.push("=IF(" + RAW + "!G" + sr + "=\"\",\"\",\"" + map.accommodationOverride.prefix + "\"&" + labelRef + ")");
        } else if (map.accommodationOverride && map.accommodationOverride.type === 'field') {
          var fieldRef = rawRef(map.accommodationOverride.field, '""');
          f.push("=IF(" + RAW + "!G" + sr + "=\"\",\"\"," + fieldRef + ")");
        } else {
          f.push('""');
        }
      } else if (clIdx === 16) {
        // O(15): TOTAL — every branch below is wrapped in an outer
        // IF($AZ<row>="Yes",0,...) so an Influencer Booking always reads 0
        // here regardless of source. AZ (col 52) is this row's OWN cell, not
        // a raw-source field, hence the plain "$AZ"+r self-reference.
        if (financeOverride) {
          // Matched event row: the property side is typically a zero-value
          // inventory placeholder — the real base amount lives on the event
          // row, so pull it directly rather than reverse-calcing from Razorpay.
          f.push("=IF(" + RAW + "!G" + sr + "=\"\",\"\",IF($AZ" + r + "=\"Yes\",0,ROUND(" + financeRef('amount', '0') + ",2)))");
        } else if (map === EVENT_MAP) {
          // Standalone event booking: J holds the event title, not an
          // accommodation name, so there's nothing to look up in Master GST
          // Data. Event bookings currently carry a flat 18% GST rate, so
          // reverse-calc (when needed) uses that directly instead of a VLOOKUP.
          var amtRefE = rawRef('amount', '0');
          var totRefE = rawRef('totalAmount', '0');
          f.push(
            "=IF(" + RAW + "!G" + sr + "=\"\",\"\",IF($AZ" + r + "=\"Yes\",0," +
            "IF(AND(ABS(" + amtRefE + "-" + totRefE + ")>2,ISNUMBER(AV" + r + "),AV" + r + ">0)," +
              "ROUND(AV" + r + "*100/118,2)," +
              "ROUND(" + amtRefE + ",2))))"
          );
        } else {
          // effective, reverse-calced when promo discount + Razorpay present
          var amtRef = rawRef('amount', '0');
          var walletRef = rawRef('walletUsed', '0');
          var totRef = rawRef('totalAmount', '0');
          f.push(
            "=IF(" + RAW + "!G" + sr + "=\"\",\"\",IF($AZ" + r + "=\"Yes\",0," +
            "IF(AND(ABS(" + amtRef + "+" + walletRef + "-" + totRef + ")>2,ISNUMBER(AV" + r + "),AV" + r + ">0)," +
              "ROUND(IFERROR(AV" + r + "*100/VLOOKUP(J" + r + ",'Master GST Data'!A:B,2,0),AV" + r + "*100/118),2)," +
              "ROUND(" + amtRef + ",2))))"
          );
        }
      } else if (clIdx === 23) {
        // V(22): GST — same $AZ<row>="Yes" → 0 wrapper as O(15) above.
        if (financeOverride) {
          f.push("=IF(" + RAW + "!G" + sr + "=\"\",\"\",IF($AZ" + r + "=\"Yes\",0,ROUND(" + financeRef('gst', '0') + ",2)))");
        } else if (map === EVENT_MAP) {
          // Same flat-18% reasoning as O(15) above.
          var amtRefE2 = rawRef('amount', '0');
          var totRefE2 = rawRef('totalAmount', '0');
          var gstRefE = rawRef('gst', '0');
          f.push(
            "=IF(" + RAW + "!G" + sr + "=\"\",\"\",IF($AZ" + r + "=\"Yes\",0," +
            "IF(AND(ABS(" + amtRefE2 + "-" + totRefE2 + ")>2,ISNUMBER(AV" + r + "),AV" + r + ">0)," +
              "ROUND(AV" + r + "*18/118,2)," +
              "ROUND(" + gstRefE + ",2))))"
          );
        } else {
          // effective, reverse-calced when promo discount + Razorpay present
          var amtRef2 = rawRef('amount', '0');
          var walletRef2 = rawRef('walletUsed', '0');
          var totRef2 = rawRef('totalAmount', '0');
          var gstRef = rawRef('gst', '0');
          f.push(
            "=IF(" + RAW + "!G" + sr + "=\"\",\"\",IF($AZ" + r + "=\"Yes\",0," +
            "IF(AND(ABS(" + amtRef2 + "+" + walletRef2 + "-" + totRef2 + ")>2,ISNUMBER(AV" + r + "),AV" + r + ">0)," +
              "ROUND(IFERROR(AV" + r + "*(VLOOKUP(J" + r + ",'Master GST Data'!A:B,2,0)-100)/VLOOKUP(J" + r + ",'Master GST Data'!A:B,2,0),AV" + r + "*18/118),2)," +
              "ROUND(" + gstRef + ",2))))"
          );
        }
      } else if (clIdx === 24) {
        // W(23): GRAND TOTAL = O + V
        f.push("=IF(" + RAW + "!G" + sr + "=\"\",\"\",O" + r + "+V" + r + ")");
      } else {
        var entry = plan[planIdx++];
        var fieldName = entry[0];
        var fallback = entry[1] === '' ? '""' : String(entry[1]);
        var ref = rawRef(fieldName, fallback);
        if (map[fieldName]) {
          f.push("=IF(" + RAW + "!G" + sr + "=\"\",\"\"," + ref + ")");
        } else {
          f.push(ref); // constant literal — this field doesn't exist in this source
        }
      }
    }

    // AG(33): Booking Status — normalise raw admin statuses, but a real
    // Razorpay payment always wins over a stale "pending" admin status.
    //
    // The admin export's own status is checked FIRST for confirmed/rejected
    // — those are explicit admin decisions, always trusted as-is. Next comes
    // insta_pay (has to stay ahead of the Razorpay check below, since AW's
    // Final Status logic keys off AG containing "insta_pay" to do its own,
    // separate Converted/Not-Initiated split for those rows). Only THEN does
    // a Razorpay match get checked — independently of AH, via a direct
    // COUNTIF against 'Raw Razorpay Data' rather than referencing AH itself
    // (AH's own formula reads AG, so AG reading AH back would be a circular
    // reference). If the Booking ID shows up there at all, the booking is
    // treated as completed even if the admin export you last pasted still
    // says pending — you don't have to re-download/re-paste Raw Booking Data
    // just to catch up once Razorpay confirms the money actually arrived.
    // Only after all of that does "pending" fall back to "Payment Pending".
    //
    // v60 also widened the pending match itself to a SEARCH/contains check
    // (not exact-equals), since the admin export's actual raw value is
    // "payment_pending" (underscore) — the old exact-match on "pending" /
    // "payment pending" (space) never matched it, so it fell through
    // unclassified to the raw-text fallback below.
    var statusRef = function() { return rawRef('status', '""'); };
    f.push(
      "=IF(H" + r + "=\"\",\"\"," +
      "IF(LOWER(" + statusRef() + ")=\"confirmed\",\"Booking Completed\"," +
      "IF(LOWER(" + statusRef() + ")=\"rejected\",\"Rejected by Administrator\"," +
      "IF(ISNUMBER(SEARCH(\"insta_pay\",LOWER(" + statusRef() + "))),\"insta_payiated\"," +
      "IF(COUNTIF('Raw Razorpay Data'!A:A,H" + r + ")>0,\"Booking Completed\"," +
      "IF(ISNUMBER(SEARCH(\"pending\",LOWER(" + statusRef() + "))),\"Payment Pending\"," +
      statusRef() + "))))))"
    );

    // AH(34): Razorpay ID
    f.push(
      "=IF(H" + r + "=\"\",\"\"," +
      "IF(ISNUMBER(SEARCH(\"insta_pay\",LOWER(AG" + r + "))),\"N/A\"," +
      "IFERROR(VLOOKUP(H" + r + ",'Raw Razorpay Data'!A:A,1,0),\"N/A\")))"
    );

    // AI(35): PAN No
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(I" + r + ",'GST PAN CM COMM'!A:B,2,0),\"\"))");

    // AJ(36): Razorpay Fee
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(H" + r + ",'Raw Razorpay Data'!A:C,3,0),\"\"))");

    // AK(37): Paid to CM
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(H" + r + ",'Raw Razorpay Data'!A:D,4,0),\"\"))");

    // AL(38): CL Has GST
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(I" + r + ",'GST PAN CM COMM'!A:C,3,0),\"No\"))");

    // AM(39): CM Comm %
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(I" + r + ",'GST PAN CM COMM'!S:T,2,0),0))");

    // AN(40): CM Comm (excl GST)
    f.push("=IF(H" + r + "=\"\",\"\",O" + r + "*AM" + r + ")");

    // AO(41): CM GST on commission
    f.push("=IF(H" + r + "=\"\",\"\",IF(AL" + r + "=\"Yes\",(AN" + r + "*18%),0)+IF(AL" + r + "=\"No\",V" + r + ",0))");

    // AP(42): CL Transfer before TDS
    f.push("=IF(H" + r + "=\"\",\"\",W" + r + "-AN" + r + "-AO" + r + ")");

    // AQ(43): TDS Deducted
    f.push("=IF(H" + r + "=\"\",\"\",IF(AL" + r + "=\"No\",AP" + r + "*2%,0))");

    // AR(44): CL Transfer after TDS
    f.push("=IF(H" + r + "=\"\",\"\",AP" + r + "-AQ" + r + ")");

    // AS(45): Comments 1 — auto: booking type from ID pattern + wallet note if applicable
    var asBookingType =
      "IF(LEFT(H" + r + ",8)=\"CM-ADMIN\",\"Manual booking via admin\"," +
      "IF(LEFT(H" + r + ",3)=\"CME\",\"Event booking\"," +
      "IF(LEFT(H" + r + ",3)=\"CMX\",\"Experience booking\"," +
      "IF(LEFT(H" + r + ",8)=\"BLC Misc\",\"QR payment - Bannerghatta\"," +
      "IF(LEFT(H" + r + ",8)=\"CMV Misc\",\"QR payment - Vasind\"," +
      "IF(LEFT(H" + r + ",4)=\"CMMB\",\"Manual booking\",\"\"))))))";

    // Influencer Booking (AZ="Yes") takes priority over the wallet note —
    // an influencer row's O/V/W are already forced to 0 above, and it never
    // shows "Wallet used" even if the raw data happens to carry a wallet
    // figure, so the Wallet/LP row-color trigger (which keys off this AS
    // text containing "Wallet") naturally never fires for these rows either.
    var asWalletNote =
      "IF($AZ" + r + "=\"Yes\",\"Influencer Booking\"," +
      "IF(AND(ISNUMBER(U" + r + "),U" + r + ">0)," +
      "IF(ABS(" + rawRef('amount', '0') + "+" + rawRef('walletUsed', '0') + "-" + rawRef('totalAmount', '0') + ")>2," +
      "\"Wallet used - Recalc done\",\"Wallet used\"),\"\"))";

    f.push(
      "=IF(H" + r + "=\"\",\"\"," +
      "LET(bt," + asBookingType + "," +
      "wn," + asWalletNote + "," +
      "IF(AND(bt<>\"\",wn<>\"\"),bt&\" | \"&wn,bt&wn)))"
    );

    // AT(46): Comments 2 — manual, EXCEPT auto-filled with the ORIGINAL
    // property Booking ID ("Property Booking ID: CM28192") when this row's
    // own Booking ID (H) has been swapped to the matched event's ID (see
    // financeOverride, built in Step 0.B) — H no longer shows the property
    // ID, so this is where it's preserved for traceability. Blank/manual for
    // every other row, same as before.
    // setFormulas() treats every entry as a formula, not literal text — a
    // bare string like "Property Booking ID: CM28192" gets an implicit "="
    // prepended and fails to parse (colons/spaces aren't valid formula
    // syntax), hence #ERROR!. Wrap it as a quoted string formula instead,
    // with internal quotes doubled per Sheets' escaping rule.
    f.push(financeOverride ? '="' + financeOverride.note.replace(/"/g, '""') + '"' : "");

    // AU(47): manual, EXCEPT auto-filled with the real accommodation/room
    // type for matched rows — J now shows the event title instead (see
    // clIdx===11 above), so the actual room type would otherwise disappear
    // from this row entirely. Blank/manual for every other row, as before.
    if (financeOverride && map.accommodation) {
      var realAccomRef = rawRef('accommodation', '""');
      f.push("=IF(" + RAW + "!G" + sr + "=\"\",\"\"," + realAccomRef + ")");
    } else {
      f.push("");
    }

    // AV(48): Match Total
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(AJ" + r + "+AK" + r + ",\"\"))");

    // AW(49): Final Status
    var convCheck =
      "SUMPRODUCT(" +
        "(N$2:N$" + (NUM_ROWS + 1) + "=N" + r + ")*" +
        "(LOWER(AG$2:AG$" + (NUM_ROWS + 1) + ")=\"booking completed\")*" +
        "(H$2:H$" + (NUM_ROWS + 1) + "<>H" + r + ")*" +
        "((IFERROR(" +
            "(VALUE(MID(H$2:H$" + (NUM_ROWS + 1) + ",3,10))>=VALUE(MID(H" + r + ",3,10))+1)*" +
            "(VALUE(MID(H$2:H$" + (NUM_ROWS + 1) + ",3,10))<=VALUE(MID(H" + r + ",3,10))+5)" +
          ",0))" +
         "+(IFERROR(" +
            "(DATEVALUE(B$2:B$" + (NUM_ROWS + 1) + ")>=DATEVALUE(B" + r + "))*" +
            "(DATEVALUE(B$2:B$" + (NUM_ROWS + 1) + ")<=DATEVALUE(B" + r + ")+3)" +
          ",0))" +
        ")" +
      ")>0";

    f.push(
      "=IF(H" + r + "=\"\",\"\"," +
      "IF(ISNUMBER(SEARCH(\"insta_pay\",LOWER(AG" + r + ")))," +
        "IF(" + convCheck + ",\"Converted Booking\",\"Payment Not Initiated\")," +
      "IF(AG" + r + "=\"Payment Pending\"," +
        "IF(AH" + r + "<>\"N/A\",\"Payment Received ✓\",\"Cancelled\")," +
      "IF(AH" + r + "=\"N/A\",\"N/A\"," +
      "IF(AG" + r + "=\"Booking Completed\",\"Booking Completed ✓\",\"MATCH\")))))"
    );

    // AX(50): Refund Amount — manual, blank by default
    f.push("");

    // AY(51): Refund Type — manual dropdown, blank by default
    f.push("");

    // AZ(52): Influencer Booking — manual dropdown, blank by default (blank
    // is treated as "not an influencer booking" everywhere it's checked).
    f.push("");

    return f;
  }

  // Dedicated formula-builder for standalone QR-v2 payment rows — these have
  // no property/experience/event booking behind them at all, so none of the
  // per-source field maps apply. B–G all show the same date (Raw Razorpay
  // Data col AE, "Created At"), H is the CMQR ID itself (so AH/AJ/AK resolve
  // automatically — H matches its own row in Raw Razorpay Data), J
  // (Accommodation) stays blank as a manual-fill marker, Q–U and X–AF are 0,
  // AG is a flat "Booking Completed", and O/V reverse-calc off the
  // Razorpay-matched total (AV = AJ+AK) at a flat 5% GST.
  function buildQRRowFormulas(r, rrRow) {
    var f = [];
    var idRef = "'Raw Razorpay Data'!A" + rrRow;
    var dateRef = "'Raw Razorpay Data'!AE" + rrRow;

    // Raw Razorpay Data col AE ("Created At") is a fixed "DD/MM/YYYY
    // HH:MM:SS" text string — pull the date straight from those positions
    // rather than relying on DATEVALUE (which fails to parse this format and
    // was silently falling through to a broken end-of-string guess, giving a
    // correct day/month but a garbage year read off the trailing time).
    //
    // REVERTED (v53) — an earlier version of this formula un-swapped
    // day/month whenever the cell was already a real date (day<=12), on the
    // theory that Sheets always mis-parses this column via its MM/DD locale.
    // That guess isn't reliable enough to apply blindly — proven wrong
    // elsewhere in this same fix pass (it corrupted already-correct property/
    // OTA dates), so it's removed here too rather than trusted on unverified
    // evidence. The Raw Razorpay Data AE cleanup (swappedDatesFixed, above in
    // this file) already normalizes this specific column's raw values before
    // this formula ever runs — that's the intended fix for AE specifically.
    // What stays: never trust dateRef*1 or DATEVALUE()'s own locale-guessing
    // on TEXT — go straight to the unambiguous manual position parse.
    function parseDateExpr() {
      return "IF(ISNUMBER(" + dateRef + ")," + dateRef + "," +
               "IFERROR(DATE(VALUE(MID(" + dateRef + ",7,4)),VALUE(MID(" + dateRef + ",4,2)),VALUE(LEFT(" + dateRef + ",2))),\"\"))";
    }
    function dateFormula() {
      return "=IF(" + idRef + "=\"\",\"\",TEXT(" + parseDateExpr() + ",\"DD- MMM- YY\"))";
    }
    function monthFormula() {
      return "=IF(" + idRef + "=\"\",\"\",TEXT(" + parseDateExpr() + ",\"MMM-YY\"))";
    }

    f.push("=IF(" + idRef + "=\"\",\"\",ROW()-1)");                          // A: Sr No
    f.push(dateFormula());                                                    // B: Booking Date
    f.push(monthFormula());                                                   // C: Booking Mth
    f.push(dateFormula());                                                    // D: Travel Start Date (same date)
    f.push(monthFormula());                                                   // E: Travel Mth
    f.push(dateFormula());                                                    // F: Travel End Date (same date)
    f.push(monthFormula());                                                   // G: Travel End Mth
    f.push("=IF(" + idRef + "=\"\",\"\"," + idRef + ")");                     // H: Booking ID
    f.push("");                                                               // I: Campsite Name
    f.push("");                                                               // J: Accommodation — blank marker
    f.push("");                                                               // K: Camp Owner Name
    f.push("");                                                               // L: Camper Name
    f.push("");                                                               // M: Camper Phone
    f.push("");                                                               // N: Camper Email
    f.push("=IF(" + idRef + "=\"\",\"\",IF($AZ" + r + "=\"Yes\",0,ROUND(AV" + r + "*100/105,2)))"); // O: TOTAL (flat 5% reverse-calc; 0 if Influencer Booking)
    f.push("");                                                               // P: Promocode Name
    f.push(0);                                                                // Q: Promocode Discount
    f.push(0);                                                                // R: Flat Discount
    f.push(0);                                                                // S: Referral Discount
    f.push(0);                                                                // T: Weekdays Discount
    f.push(0);                                                                // U: Wallet Amount Used
    f.push("=IF(" + idRef + "=\"\",\"\",IF($AZ" + r + "=\"Yes\",0,ROUND(AV" + r + "*5/105,2)))"); // V: GST (flat 5%; 0 if Influencer Booking)
    f.push("=IF(" + idRef + "=\"\",\"\",O" + r + "+V" + r + ")");             // W: GRAND TOTAL
    f.push(0); f.push(0); f.push(0); f.push(0);                               // X, Y, Z, AA
    f.push(0); f.push(0); f.push(0); f.push(0); f.push(0);                    // AB, AC, AD, AE, AF
    f.push("=IF(" + idRef + "=\"\",\"\",\"Booking Completed\")");             // AG: Booking Status

    f.push("=IF(H" + r + "=\"\",\"\",IF(ISNUMBER(SEARCH(\"insta_pay\",LOWER(AG" + r + "))),\"N/A\",IFERROR(VLOOKUP(H" + r + ",'Raw Razorpay Data'!A:A,1,0),\"N/A\")))"); // AH: Razorpay ID
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(I" + r + ",'GST PAN CM COMM'!A:B,2,0),\"\"))");   // AI: PAN No
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(H" + r + ",'Raw Razorpay Data'!A:C,3,0),\"\"))"); // AJ: Razorpay Fee
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(H" + r + ",'Raw Razorpay Data'!A:D,4,0),\"\"))"); // AK: Paid to CM
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(I" + r + ",'GST PAN CM COMM'!A:C,3,0),\"No\"))"); // AL: CL Has GST
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(I" + r + ",'GST PAN CM COMM'!S:T,2,0),0))");      // AM: CM Comm %
    f.push("=IF(H" + r + "=\"\",\"\",O" + r + "*AM" + r + ")");               // AN: CM Comm (excl GST)
    f.push("=IF(H" + r + "=\"\",\"\",IF(AL" + r + "=\"Yes\",(AN" + r + "*18%),0)+IF(AL" + r + "=\"No\",V" + r + ",0))"); // AO: CM GST
    f.push("=IF(H" + r + "=\"\",\"\",W" + r + "-AN" + r + "-AO" + r + ")");   // AP: CL Transfer before TDS
    f.push("=IF(H" + r + "=\"\",\"\",IF(AL" + r + "=\"No\",AP" + r + "*2%,0))"); // AQ: TDS Deducted
    f.push("=IF(H" + r + "=\"\",\"\",AP" + r + "-AQ" + r + ")");              // AR: CL Transfer after TDS
    f.push("=IF(H" + r + "=\"\",\"\",\"QR payment - v2\")");                  // AS: Comments 1 — flags it as a QR payment
    f.push("");                                                               // AT: Comments 2 — manual
    f.push("");                                                               // AU: Comments 3 — manual
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(AJ" + r + "+AK" + r + ",\"\"))"); // AV: Match Total
    f.push(
      "=IF(H" + r + "=\"\",\"\"," +
      "IF(AG" + r + "=\"Booking Completed\",\"Booking Completed ✓\",\"MATCH\"))"
    ); // AW: Final Status
    f.push("");                                                               // AX: Refund Amount — manual
    f.push("");                                                               // AY: Refund Type — manual
    f.push("");                                                               // AZ: Influencer Booking — manual

    return f;
  }

  // Dedicated formula-builder for standalone CM-ADMIN (manual admin booking)
  // rows — same rationale as buildQRRowFormulas above: these have no
  // property/experience/event booking record behind them (no admin booking
  // report exists for a manually-created booking), so B–G all show the same
  // date (Raw Razorpay Data col AE, "Created At" — this is a BOOKING date,
  // not a travel date, since there's no separate travel record to pull one
  // from), H is the CM-ADMIN ID itself, J (Accommodation) stays blank as a
  // manual-fill marker.
  //
  // Unlike QR rows (which stay flat 5% — those are misc BLC/CMV Misc payments
  // tied to a fixed low-GST bucket), an admin-panel booking can be for ANY
  // accommodation, and GST is 5% or 18% depending on which one (Master GST
  // Data is the single source of truth for that, same as every other row
  // type). So O/V use the exact same VLOOKUP(J, 'Master GST Data'!A:B, 2, 0)
  // pattern property/OTA rows already use: once someone types the
  // accommodation name into J, the formula picks up the right divisor
  // automatically and recalculates — no manual GST entry needed. While J is
  // still blank (nothing typed in yet), VLOOKUP fails and IFERROR falls back
  // to 18% as a placeholder, same fallback property/OTA rows use for an
  // unrecognized accommodation name.
  function buildAdminRowFormulas(r, rrRow) {
    var f = [];
    var idRef = "'Raw Razorpay Data'!A" + rrRow;
    var dateRef = "'Raw Razorpay Data'!AE" + rrRow;

    function parseDateExpr() {
      return "IF(ISNUMBER(" + dateRef + ")," + dateRef + "," +
               "IFERROR(DATE(VALUE(MID(" + dateRef + ",7,4)),VALUE(MID(" + dateRef + ",4,2)),VALUE(LEFT(" + dateRef + ",2))),\"\"))";
    }
    function dateFormula() {
      return "=IF(" + idRef + "=\"\",\"\",TEXT(" + parseDateExpr() + ",\"DD- MMM- YY\"))";
    }
    function monthFormula() {
      return "=IF(" + idRef + "=\"\",\"\",TEXT(" + parseDateExpr() + ",\"MMM-YY\"))";
    }

    f.push("=IF(" + idRef + "=\"\",\"\",ROW()-1)");                          // A: Sr No
    f.push(dateFormula());                                                    // B: Booking Date
    f.push(monthFormula());                                                   // C: Booking Mth
    f.push(dateFormula());                                                    // D: Travel Start Date (same date)
    f.push(monthFormula());                                                   // E: Travel Mth
    f.push(dateFormula());                                                    // F: Travel End Date (same date)
    f.push(monthFormula());                                                   // G: Travel End Mth
    f.push("=IF(" + idRef + "=\"\",\"\"," + idRef + ")");                     // H: Booking ID
    f.push("");                                                               // I: Campsite Name
    f.push("");                                                               // J: Accommodation — blank marker, fill in to unlock the correct GST% below
    f.push("");                                                               // K: Camp Owner Name
    f.push("");                                                               // L: Camper Name
    f.push("");                                                               // M: Camper Phone
    f.push("");                                                               // N: Camper Email
    f.push("=IF(" + idRef + "=\"\",\"\",IF($AZ" + r + "=\"Yes\",0,ROUND(IFERROR(AV" + r + "*100/VLOOKUP(J" + r + ",'Master GST Data'!A:B,2,0),AV" + r + "*100/118),2)))"); // O: TOTAL — VLOOKUP GST% from Master GST Data once J is filled in; 0 if Influencer Booking
    f.push("");                                                               // P: Promocode Name
    f.push(0);                                                                // Q: Promocode Discount
    f.push(0);                                                                // R: Flat Discount
    f.push(0);                                                                // S: Referral Discount
    f.push(0);                                                                // T: Weekdays Discount
    f.push(0);                                                                // U: Wallet Amount Used
    f.push("=IF(" + idRef + "=\"\",\"\",IF($AZ" + r + "=\"Yes\",0,ROUND(IFERROR(AV" + r + "*(VLOOKUP(J" + r + ",'Master GST Data'!A:B,2,0)-100)/VLOOKUP(J" + r + ",'Master GST Data'!A:B,2,0),AV" + r + "*18/118),2)))"); // V: GST — same VLOOKUP; 0 if Influencer Booking
    f.push("=IF(" + idRef + "=\"\",\"\",O" + r + "+V" + r + ")");             // W: GRAND TOTAL
    f.push(0); f.push(0); f.push(0); f.push(0);                               // X, Y, Z, AA
    f.push(0); f.push(0); f.push(0); f.push(0); f.push(0);                    // AB, AC, AD, AE, AF
    f.push("=IF(" + idRef + "=\"\",\"\",\"Booking Completed\")");             // AG: Booking Status

    f.push("=IF(H" + r + "=\"\",\"\",IF(ISNUMBER(SEARCH(\"insta_pay\",LOWER(AG" + r + "))),\"N/A\",IFERROR(VLOOKUP(H" + r + ",'Raw Razorpay Data'!A:A,1,0),\"N/A\")))"); // AH: Razorpay ID
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(I" + r + ",'GST PAN CM COMM'!A:B,2,0),\"\"))");   // AI: PAN No
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(H" + r + ",'Raw Razorpay Data'!A:C,3,0),\"\"))"); // AJ: Razorpay Fee
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(H" + r + ",'Raw Razorpay Data'!A:D,4,0),\"\"))"); // AK: Paid to CM
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(I" + r + ",'GST PAN CM COMM'!A:C,3,0),\"No\"))"); // AL: CL Has GST
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(I" + r + ",'GST PAN CM COMM'!S:T,2,0),0))");      // AM: CM Comm %
    f.push("=IF(H" + r + "=\"\",\"\",O" + r + "*AM" + r + ")");               // AN: CM Comm (excl GST)
    f.push("=IF(H" + r + "=\"\",\"\",IF(AL" + r + "=\"Yes\",(AN" + r + "*18%),0)+IF(AL" + r + "=\"No\",V" + r + ",0))"); // AO: CM GST
    f.push("=IF(H" + r + "=\"\",\"\",W" + r + "-AN" + r + "-AO" + r + ")");   // AP: CL Transfer before TDS
    f.push("=IF(H" + r + "=\"\",\"\",IF(AL" + r + "=\"No\",AP" + r + "*2%,0))"); // AQ: TDS Deducted
    f.push("=IF(H" + r + "=\"\",\"\",AP" + r + "-AQ" + r + ")");              // AR: CL Transfer after TDS
    f.push("=IF(H" + r + "=\"\",\"\",\"Manual booking via admin\")");         // AS: Comments 1 — flags it as an admin-panel manual booking
    f.push("");                                                               // AT: Comments 2 — manual
    f.push("");                                                               // AU: Comments 3 — manual
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(AJ" + r + "+AK" + r + ",\"\"))"); // AV: Match Total
    f.push(
      "=IF(H" + r + "=\"\",\"\"," +
      "IF(AG" + r + "=\"Booking Completed\",\"Booking Completed ✓\",\"MATCH\"))"
    ); // AW: Final Status
    f.push("");                                                               // AX: Refund Amount — manual
    f.push("");                                                               // AY: Refund Type — manual
    f.push("");                                                               // AZ: Influencer Booking — manual

    return f;
  }

  // Dedicated formula-builder for qualified OTA bookings, pulled in from Raw
  // OTA Data (mirrored there by importQualifiedOtaBookings() from the
  // separate OTA sheet — see OTA_Payment_Setup.gs). Unlike QR rows, these DO
  // have a real Accommodation name, so O/V reverse-calc uses the same Master
  // GST Data VLOOKUP every property-sourced row uses — not a flat rate —
  // per instruction: GST must always come from Master GST Data, regardless
  // of whatever GST/tax figures the OTA's own export shows. AV (Match Total)
  // is a VLOOKUP into Raw OTA Data's Host Payout column (K) — the net,
  // post-commission figure — playing the same role Razorpay's AJ+AK plays
  // for other sources. Commission/gross/settlement detail intentionally
  // stays on the OTA sheet only; it never crosses over here.
  // Same brand color scheme as the OTA sheet itself (OTA_Payment_Setup.gs) —
  // GH/NH prefixes on the Booking ID distinguish Goibibo vs MakeMyTrip within
  // the same GoMMT account; everything else just uses its own Platform value
  // straight from Raw OTA Data. Applied to column H only (Booking ID cell),
  // not the whole row, same "not too flashy" call made on the OTA sheet.
  // GoMMT and Manual/Stories Collective were nudged away from #FCE4D6 (the
  // Missing-Razorpay flag peach) and #C6EFCE (the Booking Completed green) —
  // both already carry a specific meaning elsewhere in this sheet, so a
  // near-identical hue on column H would be easy to misread at a glance.
  var OTA_BRAND_ROW_COLORS = {
    Goibibo: '#F9CB9C',
    MakeMyTrip: '#A2C4C9',
    GoMMT: '#F0DDB8',
    ClearTrip: '#D0E6F5',
    Airbnb: '#F5C6CB',
    'Stories Collective': '#A8DDD2',
    'ALIVE Booking': '#FCE8A6',
    Manual: '#D6D6E8'
  };
  function classifyOtaBrand(bookingId, platform) {
    var id = (bookingId || '').toString().trim();
    if (id.indexOf('GH') === 0) return 'Goibibo';
    if (id.indexOf('NH') === 0) return 'MakeMyTrip';
    return platform || 'Manual';
  }

  function buildOtaRowFormulas(r, otaRow) {
    var f = [];
    var idRef = "'Raw OTA Data'!A" + otaRow;

    function dateFormula(col) {
      return "=IF(" + idRef + "=\"\",\"\",TEXT('Raw OTA Data'!" + col + otaRow + ",\"DD- MMM- YY\"))";
    }
    function monthFormula(col) {
      return "=IF(" + idRef + "=\"\",\"\",TEXT('Raw OTA Data'!" + col + otaRow + ",\"MMM-YY\"))";
    }
    function otaRef(col, fallback) {
      return "=IF(" + idRef + "=\"\",\"\",IF('Raw OTA Data'!" + col + otaRow + "=\"\"," + fallback + ",'Raw OTA Data'!" + col + otaRow + "))";
    }

    f.push("=IF(" + idRef + "=\"\",\"\",ROW()-1)");                       // A: Sr No
    // B/C: Booking Date/Mth — sourced from Settlement Date (col M: the date
    // the OTA payout was actually RECEIVED), not the guest's original OTA
    // booking date (col B). Sales data is meant to reflect when payment
    // actually landed, and for OTA bookings that can be weeks/months after
    // the guest booked — using the guest's booking date here would misalign
    // this row's date against every other (payment-dated) source. The
    // ORIGINAL OTA booking date is preserved in Comments 2 (AT) below so
    // it isn't lost, just relocated. Travel dates (D–G) are UNCHANGED —
    // still the guest's real travel dates, still from cols C/D.
    f.push(dateFormula("M"));                                             // B: Booking Date ← Settlement Date
    f.push(monthFormula("M"));                                            // C: Booking Mth ← Settlement Date
    f.push(dateFormula("C"));                                             // D: Travel Start Date
    f.push(monthFormula("C"));                                            // E: Travel Mth
    f.push(dateFormula("D"));                                             // F: Travel End Date
    f.push(monthFormula("D"));                                            // G: Travel End Mth
    f.push("=IF(" + idRef + "=\"\",\"\"," + idRef + ")");                 // H: Booking ID
    f.push(otaRef("E", '""'));                                            // I: Campsite Name
    f.push(otaRef("F", '""'));                                            // J: Accommodation ← GST lookup key
    f.push("=IF(" + idRef + "=\"\",\"\",\"Camp Monk\")");                 // K: Camp Owner Name — same convention every other row uses
    f.push(otaRef("G", '""'));                                            // L: Camper Name
    f.push("");                                                           // M: Camper Phone — not in OTA exports
    f.push("");                                                           // N: Camper Email — not in OTA exports
    // O: TOTAL — reverse-calc off AV via Master GST Data, same formula every
    // property-sourced row uses, NOT a flat rate — this is the whole point.
    // 0 if Influencer Booking (AZ="Yes").
    f.push("=IF(" + idRef + "=\"\",\"\",IF($AZ" + r + "=\"Yes\",0,ROUND(IFERROR(AV" + r + "*100/VLOOKUP(J" + r + ",'Master GST Data'!A:B,2,0),AV" + r + "*100/118),2)))");
    f.push("");                                                           // P: Promocode Name — N/A
    f.push(0); f.push(0); f.push(0); f.push(0); f.push(0);                // Q–U: discounts/wallet — N/A, 0
    // V: GST — same Master GST Data reverse-calc pattern as O. 0 if Influencer Booking.
    f.push("=IF(" + idRef + "=\"\",\"\",IF($AZ" + r + "=\"Yes\",0,ROUND(IFERROR(AV" + r + "*(VLOOKUP(J" + r + ",'Master GST Data'!A:B,2,0)-100)/VLOOKUP(J" + r + ",'Master GST Data'!A:B,2,0),AV" + r + "*18/118),2)))");
    f.push("=IF(" + idRef + "=\"\",\"\",O" + r + "+V" + r + ")");         // W: GRAND TOTAL
    f.push(otaRef("H", 0));                                               // X: Total units
    f.push(0);                                                            // Y: Insta units — N/A
    f.push(0); f.push(0); f.push(0); f.push(0); f.push(0); f.push(0);     // Z–AE: guest/kid/pet counts — not in OTA exports
    f.push(otaRef("I", 0));                                               // AF: No of Nights
    f.push("=IF(" + idRef + "=\"\",\"\",\"Booking Completed\")");         // AG: Booking Status
    f.push("=IF(" + idRef + "=\"\",\"\",\"N/A\")");                       // AH: Razorpay ID — not applicable
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(I" + r + ",'GST PAN CM COMM'!A:B,2,0),\"\"))");   // AI: PAN No
    f.push("");                                                           // AJ: Razorpay Fee — N/A
    f.push("");                                                           // AK: Paid to CM — N/A
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(I" + r + ",'GST PAN CM COMM'!A:C,3,0),\"No\"))"); // AL: CL Has GST
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(I" + r + ",'GST PAN CM COMM'!S:T,2,0),0))");      // AM: CM Comm %
    f.push("=IF(H" + r + "=\"\",\"\",O" + r + "*AM" + r + ")");           // AN: CM Comm (excl GST)
    f.push("=IF(H" + r + "=\"\",\"\",IF(AL" + r + "=\"Yes\",(AN" + r + "*18%),0)+IF(AL" + r + "=\"No\",V" + r + ",0))"); // AO: CM GST
    f.push("=IF(H" + r + "=\"\",\"\",W" + r + "-AN" + r + "-AO" + r + ")"); // AP: CL Transfer before TDS
    f.push("=IF(H" + r + "=\"\",\"\",IF(AL" + r + "=\"No\",AP" + r + "*2%,0))"); // AQ: TDS Deducted
    f.push("=IF(H" + r + "=\"\",\"\",AP" + r + "-AQ" + r + ")");          // AR: CL Transfer after TDS
    // AS: Comments 1 — auto: platform + a mismatch flag pulled straight from
    // the OTA sheet's own Qualified Status text, so a flagged reconciliation
    // discrepancy is visible here too, not just on the OTA sheet.
    f.push(
      "=IF(" + idRef + "=\"\",\"\",\"OTA booking - \"&'Raw OTA Data'!J" + otaRow +
      "&IF(ISNUMBER(SEARCH(\"Mismatch\",'Raw OTA Data'!L" + otaRow + ")),\" | Amount mismatch — review on OTA sheet\",\"\"))"
    );
    // AT: Comments 2 — the ORIGINAL OTA booking date (guest's actual booking
    // date, before the Settlement-Date swap above), so it's still visible
    // even though it no longer drives this row's Booking Date/Mth.
    f.push(
      "=IF(" + idRef + "=\"\",\"\",IFERROR(\"Original OTA booking date: \"&TEXT('Raw OTA Data'!B" + otaRow + ",\"DD-MMM-YY\"),\"\"))"
    );
    f.push("");                                                           // AU: Comments 3 — manual
    // AV: Match Total — VLOOKUP into Raw OTA Data's Host Payout (col K, the
    // 11th column), same role Razorpay's AJ+AK plays for other sources.
    f.push("=IF(H" + r + "=\"\",\"\",IFERROR(VLOOKUP(H" + r + ",'Raw OTA Data'!A:K,11,0),\"\"))");
    f.push("=IF(H" + r + "=\"\",\"\",\"Booking Completed ✓\")");          // AW: Final Status
    f.push("");                                                           // AX: Refund Amount — manual
    f.push("");                                                           // AY: Refund Type — manual
    f.push("");                                                           // AZ: Influencer Booking — manual

    return f;
  }

  // Build and write one formula row per combined/sorted booking
  var manualOverridesApplied = 0;
  if (dataRowCount > 0) {
    var allFormulas = [];
    var pendingOverrideWrites = []; // { row, col (1-based), value }
    var bookingIdColors = []; // parallel to allFormulas — column H brand tint, OTA rows only
    for (var idx = 0; idx < dataRowCount; idx++) {
      var item = combined[idx];
      allFormulas.push(
        item.isQR
          ? buildQRRowFormulas(idx + 2, item.rrRow)
          : item.isAdmin
          ? buildAdminRowFormulas(idx + 2, item.rrRow)
          : item.isOta
          ? buildOtaRowFormulas(idx + 2, item.otaRow)
          : buildRowFormulas(idx + 2, item.sheetRef, item.row, item.map, item.bid, item.financeOverride)
      );
      bookingIdColors.push(
        item.isOta
          ? [OTA_BRAND_ROW_COLORS[classifyOtaBrand(item.bid, item.platform)] || null]
          : [null]
      );
      var effBidForOverride = item.financeOverride ? item.financeOverride.bid : item.bid;
      var ov = manualOverridesByBookingId[effBidForOverride];
      if (ov) {
        for (var ovCol in ov) {
          pendingOverrideWrites.push({ row: idx + 2, col: parseInt(ovCol, 10) + 1, value: ov[ovCol] });
        }
      }
    }
    clSheet.getRange(2, 1, dataRowCount, 52).setFormulas(allFormulas);
    clSheet.getRange(2, 8, dataRowCount, 1).setBackgrounds(bookingIdColors); // col H: Booking ID

    // Re-apply hand-typed overrides on top of the fresh formulas — must
    // happen AFTER setFormulas() above, since that call rewrites every cell
    // in the range including the ones we're about to override again.
    manualOverridesApplied = pendingOverrideWrites.length;
    pendingOverrideWrites.forEach(function(w) {
      clSheet.getRange(w.row, w.col).setValue(w.value);
    });
  }

  // STEP 4 — Header formatting

  // Peach for data cols A–AG (1–33)
  clSheet.getRange(1, 1, 1, 33)
    .setBackground('#F7CAAC')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Light blue for calc cols AH–AZ (34–52) — 19 columns
  clSheet.getRange(1, 34, 1, 19)
    .setBackground('#BDD7EE')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  clSheet.setFrozenRows(1);
  clSheet.setFrozenColumns(1);

  // STEP 5 — Conditional Formatting

  var rules   = [];
  var agRange = clSheet.getRange(2, 33, NUM_ROWS, 1); // AG: Booking Status
  var ahRange = clSheet.getRange(2, 34, NUM_ROWS, 1); // AH: Razorpay ID
  var awRange = clSheet.getRange(2, 49, NUM_ROWS, 1); // AW: Final Status
  var axRange = clSheet.getRange(2, 50, NUM_ROWS, 1); // AX: Refund Amount
  var ayRange = clSheet.getRange(2, 51, NUM_ROWS, 1); // AY: Refund Type
  var azRange = clSheet.getRange(2, 52, NUM_ROWS, 1); // AZ: Influencer Booking
  var uRange  = clSheet.getRange(2, 21, NUM_ROWS, 1); // U:  Wallet Amount Used
  // Full row range A:AZ for row-level highlights
  var fullRowRange = clSheet.getRange(2, 1, NUM_ROWS, 52);

  // AG: Booking Status
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Booking Completed')
    .setBackground('#C6EFCE').setFontColor('#276221')
    .setRanges([agRange]).build());

  // Amber for EOD review — Payment Pending only
  // (insta_payiated rows are tracked in the separate InstaPay Tracker sheet)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Payment Pending')
    .setBackground('#FFE699').setFontColor('#7F6000')
    .setRanges([agRange]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('cancelled')
    .setBackground('#FFCCCC').setFontColor('#CC0000')
    .setRanges([agRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Cancelled')
    .setBackground('#FFCCCC').setFontColor('#CC0000')
    .setRanges([agRange]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Rejected')
    .setBackground('#FF7474').setFontColor('#7B0000')
    .setRanges([agRange]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('refunded')
    .setBackground('#FFE0E0').setFontColor('#993333')
    .setRanges([agRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Refunded')
    .setBackground('#FFE0E0').setFontColor('#993333')
    .setRanges([agRange]).build());

  // AH: Razorpay ID
  // Green = matched; orange = Booking Completed but NO Razorpay (needs action);
  // grey = N/A (insta_pay or Payment Pending — expected, not urgent)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(AH2<>"",AH2<>"N/A")')
    .setBackground('#C6EFCE').setFontColor('#276221')
    .setRanges([ahRange]).build());

  // Orange flag on AH — completed booking with no Razorpay, excluding wallet-paid
  // AND OTA-sourced rows (both are legitimately N/A here — OTA bookings never
  // have a Razorpay match at all, that's expected, not something to review).
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(AH2="N/A",AG2="Booking Completed",NOT(ISNUMBER(SEARCH("Wallet",AS2))),NOT(ISNUMBER(SEARCH("OTA booking",AS2))))')
    .setBackground('#F4B942').setFontColor('#7F3F00')
    .setRanges([ahRange]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('N/A')
    .setBackground('#D9D9D9').setFontColor('#595959')
    .setRanges([ahRange]).build());

  // Full-row highlights for EOD attention
  // Payment Pending — full row amber so it's impossible to miss at EOD
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$AG2="Payment Pending"')
    .setBackground('#FFF2CC').setFontColor('#7F6000')
    .setRanges([fullRowRange]).build());

  // Missing Razorpay on a completed booking — full row orange flag
  // Excludes wallet-paid rows AND OTA-sourced rows (AS col contains "Wallet"
  // or "OTA booking") — both are legitimately N/A, not something to review.
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($AH2="N/A",$AG2="Booking Completed",NOT(ISNUMBER(SEARCH("Wallet",$AS2))),NOT(ISNUMBER(SEARCH("OTA booking",$AS2))))')
    .setBackground('#FCE4D6').setFontColor('#7F3F00')
    .setRanges([fullRowRange]).build());

  // Influencer Booking — cell + full-row highlight
  // Trigger: AZ (col 52) = "Yes". Placed ABOVE the Wallet/LP layers below —
  // in Google Sheets, when two conditional format rules both match the same
  // cell for the same property (here: row background), the rule EARLIER in
  // the rules array wins. This row's AS text will already read "Influencer
  // Booking" rather than "Wallet used" (see buildRowFormulas' asWalletNote),
  // so in practice the wallet layers won't even fire on these rows — this
  // ordering is a deliberate belt-and-suspenders guarantee of the same
  // "influencer supersedes wallet" rule, not the only thing enforcing it.
  var influencerFormula = '=$AZ2="Yes"';

  // AZ cell itself — bold teal, distinct from every other color on the sheet
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(influencerFormula)
    .setBackground('#4DB6AC').setFontColor('#FFFFFF').setBold(true)
    .setRanges([azRange]).build());

  // Full row — light teal stripe
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(influencerFormula)
    .setBackground('#D0ECE7')
    .setRanges([fullRowRange]).build());

  // Wallet / LP usage — three-layer highlight
  // Trigger: AS (col 45, Comments 1) contains "Wallet" — this text is always
  // computed by the formula so it's the most reliable trigger.
  //
  // Wallet highlight — only when Booking Completed (AG col 33 = $AG2)
  var asRange = clSheet.getRange(2, 45, NUM_ROWS, 1); // AS: Comments 1
  var walletFormula = '=AND(ISNUMBER(SEARCH("Wallet",$AS2)),$AG2="Booking Completed")';

  // Layer 1: AS cell — bold purple
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(walletFormula)
    .setBackground('#C9A8F0').setFontColor('#3B006F').setBold(true)
    .setRanges([asRange]).build());

  // Layer 2: U cell (Wallet Amount Used) — purple
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(walletFormula)
    .setBackground('#C9A8F0').setFontColor('#3B006F')
    .setRanges([uRange]).build());

  // Layer 3: full row stripe — lavender
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(walletFormula)
    .setBackground('#EAD9F7')
    .setRanges([fullRowRange]).build());

  // AW: Final Status
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Booking Completed ✓')
    .setBackground('#C6EFCE').setFontColor('#276221')
    .setRanges([awRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('MATCH')
    .setBackground('#C6EFCE').setFontColor('#276221')
    .setRanges([awRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('N/A')
    .setBackground('#D9D9D9').setFontColor('#595959')
    .setRanges([awRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Converted Booking')
    .setBackground('#D9B3FF').setFontColor('#4B0082')
    .setRanges([awRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Payment Not Initiated')
    .setBackground('#D9D9D9').setFontColor('#595959')
    .setRanges([awRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Payment Received ✓')
    .setBackground('#C6EFCE').setFontColor('#276221')
    .setRanges([awRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Cancelled')
    .setBackground('#F4CCCC').setFontColor('#990000')
    .setRanges([awRange]).build());

  // AX: Refund Amount — flag whenever a refund figure has been entered
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($AX2),$AX2>0)')
    .setBackground('#F4CCCC').setFontColor('#990000').setBold(true)
    .setRanges([axRange]).build());

  // AY: Refund Type — Wallet Credit (purple, stays with the business as
  // reusable customer credit) vs Source Refund (red, actual cash leaves)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Wallet Credit')
    .setBackground('#E5D4F7').setFontColor('#4B0082')
    .setRanges([ayRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Source Refund')
    .setBackground('#F4CCCC').setFontColor('#990000')
    .setRanges([ayRange]).build());

  // AY: Refund Type — dropdown so the two values stay exact-match clean
  // for pivots (typos would otherwise silently fall out of any Refund Type
  // pivot/sum). Not strictly enforced (allowInvalid true) so a booking that
  // genuinely doesn't fit either bucket can still take a manual note.
  var refundTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Wallet Credit', 'Source Refund'], true)
    .setAllowInvalid(true)
    .setHelpText('Wallet Credit = stays as reusable customer credit. Source Refund = sent back to original payment method.')
    .build();
  ayRange.setDataValidation(refundTypeRule);

  // AZ: Influencer Booking — dropdown, same not-strictly-enforced pattern
  // as AY. "Yes" zeroes O/V/W on that row and swaps the AS note to
  // "Influencer Booking" (see buildRowFormulas). Blank/"No" = normal booking.
  var influencerRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .setAllowInvalid(true)
    .setHelpText('Yes = comp/influencer booking — Total, GST and Grand Total are forced to 0 on this row.')
    .build();
  azRange.setDataValidation(influencerRule);

  // Day Outing color code — CMX (experience) bookings
  // Light-blue full-row tint so experience/day-outing bookings stand out
  // from property/event rows in the merged list. Added last = lowest
  // priority, so Payment Pending / Wallet / missing-Razorpay row highlights
  // above still win where they overlap.
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=LEFT($H2,3)="CMX"')
    .setBackground('#D6EAF8')
    .setRanges([fullRowRange]).build());

  // Event color code
  // Light-green full-row tint for any row whose Booking ID is CME-prefixed —
  // covers both standalone event bookings AND matched property rows, since
  // Step 0.B replaces H with the event's ID for matched rows too. Same
  // low-priority placement as Day Outing — specific status highlights above
  // still win where they overlap.
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=LEFT($H2,3)="CME"')
    .setBackground('#D5F5E3')
    .setRanges([fullRowRange]).build());

  clSheet.setConditionalFormatRules(rules);

  // STEP 6 — Update Raw Razorpay Data col B → Payment Type classifier

  rrSheet.getRange(1, 2).setValue('Payment Type [auto]');

  var ptFormula =
    "=IF(LOWER(I2)=\"failed\",\"failed payment\"," +
      "IF(ISNUMBER(SEARCH(\"addons booking - undefined\",R2)),\"ADDON PAYMENT\"," +
      "IF(A2=\"\"," +
        "IF(ISNUMBER(SEARCH(\"qrv2 payment\",LOWER(R2))),\"\",\"Manual payment link\")," +
        "IF(LEFT(A2,4)=\"CMQR\",\"QR payment - v2\"," +
        "IF(LEFT(A2,8)=\"CM-ADMIN\",\"Manual booking via admin\"," +
        "IF(LEFT(A2,3)=\"CME\",\"Event booking\"," +
        "IF(LEFT(A2,3)=\"CMX\",\"Experience booking\"," +
        "IF(LEFT(A2,8)=\"BLC Misc\",\"QR payment - Bannerghatta\"," +
        "IF(LEFT(A2,8)=\"CMV Misc\",\"QR payment - Vasind\"," +
        "IF(LEFT(A2,4)=\"CMMB\",\"Manual booking\",\"\"))))))))))";

  rrSheet.getRange(2, 2).setFormula(ptFormula);
  rrSheet.getRange(2, 2).copyTo(rrSheet.getRange(3, 2, NUM_ROWS - 1, 1));

  // STEP 7 — Color Legend tab (static reference, rebuilt every full run)
  buildColorLegendTab(ss, OTA_BRAND_ROW_COLORS);

  // DONE
  // amountFormatIssuesFound/rrCompletenessOverrides are data-integrity
  // warnings, not routine status — they matter just as much on a silent
  // smartUpdate() run as on a manual Force Full Rebuild, so they're handed
  // back to the caller instead of only living in the detailed alert below
  // (which smartUpdate() suppresses). See the combined alert in
  // smartUpdate() for where these surface on a normal "Update Now" click.
  if (silent) return { amountFormatIssuesFound: amountFormatIssuesFound, rrCompletenessOverrides: rrCompletenessOverrides };

  safeAlert(
    '✅ Setup complete!\n\n' +
    '• CL Payments rebuilt with 52 columns (A–AZ), ' + dataRowCount + ' bookings\n' +
    '• AX = Refund Amount (manual, ₹ figure); AY = Refund Type (dropdown: Wallet Credit / Source Refund); AZ = Influencer Booking (dropdown: Yes/No — Yes zeroes Total/GST/Grand Total, notes "Influencer Booking" in Comments 1, and teal row color takes priority over the Wallet/LP highlight)\n' +
    '• Merged Raw Booking Data + Raw Experience Booking Data (confirmed only) + standalone Raw Event Booking Data, sorted by Booking Date\n' +
    (rbdDupesRemoved > 0 ? '• Removed ' + rbdDupesRemoved + ' duplicate Booking ID row(s) from Raw Booking Data\n' : '') +
    (rrDupesRemoved > 0 ? '• Removed ' + rrDupesRemoved + ' duplicate Booking ID row(s) from Raw Razorpay Data\n' : '') +
    (redDupesRemoved > 0 ? '• Removed ' + redDupesRemoved + ' duplicate Booking ID row(s) from Raw Experience Booking Data\n' : '') +
    (evtDupesRemoved > 0 ? '• Removed ' + evtDupesRemoved + ' duplicate Booking ID row(s) from Raw Event Booking Data\n' : '') +
    '• Events: ' + matchedEventCount + ' matched to an existing property booking (Booking ID swapped to the Event ID for accurate Razorpay matching, original property ID noted in Comments 2), ' + eventNewRowItems.length + ' standalone events added as new rows' + (eventSkippedOld > 0 ? ', ' + eventSkippedOld + ' older than June 2026 skipped' : '') + '\n' +
    (testBookingsExcluded > 0 ? '• Excluded ' + testBookingsExcluded + ' test-account booking(s) (@think201.com / @campmonk.com) with no Razorpay payment — see "Test Bookings" tab in the InstaPay Tracker\n' : '') +
    '• Day Outing bookings (CMX-prefixed) get a light-blue row tint; Event bookings (CME-prefixed) get a light-green tint\n' +
    '• AH (Razorpay ID): actual Booking ID if matched, else N/A\n' +
    '• AS (Comments 1): booking type + wallet note auto-combined\n' +
    '• AW (Final Status): Booking Completed ✓ / MATCH / N/A / Cancelled / Payment Received ✓\n' +
    '• O/V/W auto reverse-calc from Razorpay amount when promo discount detected\n' +
    '• Raw Razorpay Data col B updated to Payment Type classifier (incl. Admin Add-on Payment + QR payment - v2)\n' +
    (qrNewIdsAssigned > 0 ? '• Assigned ' + qrNewIdsAssigned + ' new CMQR#### Booking ID(s) to QRv2 payment rows in Raw Razorpay Data col A\n' : '') +
    (descIdsRecovered > 0 ? '• Recovered ' + descIdsRecovered + ' Booking ID(s) into Raw Razorpay Data col A from the payment description text (e.g. "Accomodation booking - CM28470") — these were "Manual payment link" before\n' : '') +
    (qrNewRowItems.length > 0 ? '• ' + qrNewRowItems.length + ' standalone QR-v2 payment row(s) added to CL Payments (J left blank for manual fill-in)\n' : '') +
    (adminNewRowItems.length > 0 ? '• ' + adminNewRowItems.length + ' standalone CM-ADMIN (manual booking via admin) row(s) added to CL Payments — Booking Date = Razorpay Created At, not a travel date (J left blank for manual fill-in)\n' : '') +
    (swappedDatesFixed > 0 ? '• Fixed ' + swappedDatesFixed + ' Created At date(s) in Raw Razorpay Data col AE that Sheets had silently swapped day/month on\n' : '') +
    (failedRowsRemoved > 0 ? '• Removed ' + failedRowsRemoved + ' failed-payment row(s) (col I = "failed") from Raw Razorpay Data\n' : '') +
    (rrCompletenessOverrides > 0 ? '⚠️  Kept the fully-settled copy over a more recent but incomplete re-paste for ' + rrCompletenessOverrides + ' Razorpay Booking ID(s) (Fee/Paid to CM blank on the later paste) — review those rows if the amount still looks wrong\n' : '') +
    (amountFormatIssuesFound > 0 ? '⚠️  ' + amountFormatIssuesFound + ' amount-column cell(s) across the raw tabs are currently stored as a DATE, not a number (Sheets auto-converted them on an earlier paste, before this run\'s Plain Text protection). These can\'t be recovered automatically — re-paste the affected row(s) from the original export. Future pastes are now protected.\n' : '') +
    (manualOverridesApplied > 0 ? '• Restored ' + manualOverridesApplied + ' hand-typed cell override(s) on top of this rebuild\n' : '') +
    '\n' +
    '⚠️  Fill "Master GST Data" tab with accommodation names + GST divisors\n' +
    '    (105 = 5% GST, 118 = 18% GST) before relying on O/V/W recalc.\n\n' +
    'Next: paste your admin report into "Raw Booking Data" (row 1 = headers),\n' +
    'your experience/day-outing export into "Raw Experience Booking Data",\n' +
    'your events export into "Raw Event Booking Data",\n' +
    'and your Razorpay Payments export into "Raw Razorpay Data" starting at col F.\n\n' +
    'Non-confirmed experience bookings (draft/failed/initiated/rejected) don\'t\n' +
    'appear here — run exportExperienceLeads() to pull them into a follow-up\n' +
    'tab inside the InstaPay Tracker sheet.'
  );
}

// COLOR LEGEND — plain-language reference for every color code used on the
// CL Payments tab. Rebuilt on every full setupCLPayments() run so it never
// drifts out of sync with the actual conditional-format rules above. Static
// content only (no formulas) — safe to fully clear and rewrite every time.
function buildColorLegendTab(ss, otaBrandColors) {
  var sh = ss.getSheetByName('Color Legend');
  if (!sh) {
    sh = ss.insertSheet('Color Legend');
  } else {
    sh.clear();
    sh.clearFormats();
  }

  var rows = []; // [Column, Color swatch text, Meaning]

  // Track hex per row index (1-based, matching final sheet row) so we can
  // paint swatches after writing all the text.
  var swatches = []; // { row, hex }

  var r = 1; // header row
  rows.push(['Column / Trigger', 'Color', 'Meaning']);

  function push(col, hex, meaning, bold) {
    r++;
    rows.push([col, hex ? '' : '', meaning]);
    if (hex) swatches.push({ row: r, hex: hex, bold: !!bold });
  }
  function pushSection(title) {
    r++;
    rows.push([title, '', '']);
    swatches.push({ row: r, hex: '#4A4A4A', font: '#FFFFFF', bold: true, isSection: true });
  }

  pushSection('AG — Booking Status');
  push('AG', '#C6EFCE', '"Booking Completed" — booking is confirmed, nothing to action.');
  push('AG', '#FFE699', '"Payment Pending" — awaiting payment, also tinted amber on the whole row.');
  push('AG', '#FFCCCC', 'Contains "cancelled" (any case) — booking called off.');
  push('AG', '#FF7474', 'Contains "Rejected" — deeper red, distinguishes a rejected payment from a plain cancellation.');
  push('AG', '#FFE0E0', 'Contains "refunded" (any case) — money already sent back to guest.');

  pushSection('AH — Razorpay ID');
  push('AH', '#C6EFCE', 'A real Razorpay ID is matched (not blank, not "N/A").');
  push('AH', '#F4B942', 'Orange flag — "Booking Completed" but Razorpay ID is "N/A" and it is not a Wallet payment or an OTA booking. Genuinely needs review.');
  push('AH', '#D9D9D9', 'Plain "N/A" (Payment Pending, wallet, or OTA row) — expected, not urgent.');

  pushSection('Full-row highlights');
  push('Row', '#FFF2CC', 'AG = "Payment Pending" — full row amber so it is impossible to miss at EOD.');
  push('Row', '#FCE4D6', 'Missing Razorpay ID on a completed booking (same rule as the AH orange flag above), excluding Wallet and OTA rows — genuinely needs review.');
  push('Row', '#D0ECE7', 'Influencer Booking (AZ = "Yes") — full row teal stripe. Takes priority over the Wallet/LP stripe below where both would apply.');
  push('Row', '#EAD9F7', 'Wallet/LP redemption row — full row lavender stripe (see Wallet layers below).');
  push('Row', '#D6EAF8', 'Day Outing / Experience booking — Booking ID starts with "CMX".');
  push('Row', '#D5F5E3', 'Event booking — Booking ID starts with "CME".');

  pushSection('AZ — Influencer Booking');
  push('AZ cell', '#4DB6AC', '"Yes" — bold white-on-teal. Zeroes Total/GST/Grand Total on this row and swaps Comments 1 (AS) to "Influencer Booking" instead of any wallet note.');

  pushSection('Wallet / LP usage (3-layer highlight, triggers together)');
  push('AS cell', '#C9A8F0', 'Comments 1 mentions "Wallet" and booking is completed — bold purple on the comment itself.');
  push('U cell', '#C9A8F0', 'Same trigger — purple on the Wallet Amount Used figure.');
  push('Row', '#EAD9F7', 'Same trigger — the whole row gets a lighter lavender stripe (lowest priority, other row rules above override it — including Influencer Booking, which also suppresses the "Wallet" text this trigger depends on).');

  pushSection('AW — Final Status');
  push('AW', '#C6EFCE', '"Booking Completed ✓" / "MATCH" / "Payment Received ✓" — all-clear.');
  push('AW', '#D9D9D9', '"N/A" / "Payment Not Initiated" — nothing to reconcile yet.');
  push('AW', '#D9B3FF', '"Converted Booking" — started as a lead/InstaPay, converted to a real booking.');
  push('AW', '#F4CCCC', '"Cancelled".');

  pushSection('AX / AY — Refunds');
  push('AX', '#F4CCCC', 'Refund Amount > 0 — bold red, flags that money was refunded on this booking.');
  push('AY', '#E5D4F7', '"Wallet Credit" — refund stays with the business as reusable customer credit.');
  push('AY', '#F4CCCC', '"Source Refund" — actual cash sent back to the original payment method.');

  pushSection('Column H — OTA Booking ID brand tint (OTA-sourced rows only)');
  push('H', otaBrandColors.Goibibo, 'Goibibo booking (Booking ID starts with "GH").');
  push('H', otaBrandColors.MakeMyTrip, 'MakeMyTrip booking (Booking ID starts with "NH").');
  push('H', otaBrandColors.GoMMT, 'GoMMT booking — same account, but not specifically Goibibo/MakeMyTrip-prefixed.');
  push('H', otaBrandColors.ClearTrip, 'ClearTrip booking.');
  push('H', otaBrandColors.Airbnb, 'Airbnb booking (manual bank-transfer entry on the OTA sheet).');
  push('H', otaBrandColors['Stories Collective'], 'Stories Collective booking (manual bank-transfer entry).');
  push('H', otaBrandColors['ALIVE Booking'], 'ALIVE Booking (manual bank-transfer entry).');
  push('H', otaBrandColors.Manual, 'Any other manually-typed OTA/channel not listed above.');

  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.setColumnWidths(1, 1, 260);
  sh.setColumnWidths(2, 1, 90);
  sh.setColumnWidths(3, 1, 560);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#F7CAAC');
  sh.setFrozenRows(1);

  for (var i = 0; i < swatches.length; i++) {
    var s = swatches[i];
    var range = s.isSection ? sh.getRange(s.row, 1, 1, 3) : sh.getRange(s.row, 2);
    range.setBackground(s.hex);
    if (s.isSection) {
      range.setFontColor(s.font || '#000000').setFontWeight('bold');
    }
    if (s.bold) range.setFontWeight('bold');
  }

  sh.getRange(1, 1, rows.length, 1).setWrap(false);
  sh.getRange(1, 3, rows.length, 1).setWrap(true);
}

// INSTAPAY EXPORT
//
// HOW TO USE:
//   Run this function (exportInstaPay) separately after setupCLPayments.
//   First run: creates a new standalone Google Sheet and shows its URL — share
//              that URL with your team as the live InstaPay tracker.
//   Subsequent runs: finds the same sheet (ID saved in Script Properties) and
//              REFRESHES the auto-filled columns while PRESERVING everything
//              your agents have typed in: AGENT, UPDATE, Converted Booking ID,
//              CALL-UPDATES.
//
// WHAT IT EXPORTS (from Raw Booking Data, insta_payiated rows only):
//   Col A : AGENT              ← manual, preserved across re-runs
//   Col B : BOOKING ID         ← auto
//   Col C : CAMPSITE NAME      ← auto
//   Col D : ACCOMMODATION      ← auto
//   Col E : TRAVEL START DATE  ← auto
//   Col F : GUEST NAME         ← auto
//   Col G : PHONE NUMBER       ← auto
//   Col H : E-MAIL             ← auto
//   Col I : UPDATE             ← manual (Already Booked / Lead to call / Duplicate…)
//   Col J : Converted Booking ID ← manual
//   Col K : INSTA PAY ATTEMPTS ← auto (count of attempts per person)
//   Col L : CALL-UPDATES       ← dropdown (Change of plan / Date N/A / OTA / Location / Others) + free text

function exportInstaPay(silent) {
  // Always open the main CL sheet directly by ID — works from editor, triggers, anywhere
  var MAIN_SS_ID = PropertiesService.getScriptProperties().getProperty('MAIN_SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(MAIN_SS_ID);

  // Flexible tab lookup — handles name variants
  var rbdSheet = null;
  var allSheets = ss.getSheets();
  for (var si = 0; si < allSheets.length; si++) {
    if (allSheets[si].getName().indexOf('Raw Booking') === 0) {
      rbdSheet = allSheets[si];
      break;
    }
  }

  if (!rbdSheet) {
    safeAlert('ERROR: Raw Booking Data tab not found.\nTabs: ' +
      allSheets.map(function(s){ return s.getName(); }).join(', '));
    return;
  }

  // AGENTS CONFIG
  // Edit this list to add/remove agents. New bookings are distributed evenly
  // in round-robin order. Already-assigned bookings are never touched.
  var AGENTS = ['AK', 'MS'];

  // 1. Get or create the standalone InstaPay Google Sheet
  var props      = PropertiesService.getScriptProperties();
  var ipSheetId  = props.getProperty('INSTAPAY_SHEET_ID');
  var ipSS, ipSheet;

  if (ipSheetId) {
    try {
      ipSS    = SpreadsheetApp.openById(ipSheetId);
      ipSheet = ipSS.getSheets()[0];
    } catch (e) {
      ipSheetId = null; // stale ID — will create fresh below
    }
  }

  if (!ipSheetId) {
    ipSS    = SpreadsheetApp.create('InstaPay Tracker - Camp Monk');
    ipSheet = ipSS.getSheets()[0];
    ipSheet.setName('InstaPay');
    props.setProperty('INSTAPAY_SHEET_ID', ipSS.getId());

    // Move into the same Drive folder as the main CL sheet
    try {
      var mainFile   = DriveApp.getFileById(ss.getId());
      var parentIter = mainFile.getParents();
      if (parentIter.hasNext()) {
        var folder  = parentIter.next();
        var newFile = DriveApp.getFileById(ipSS.getId());
        folder.addFile(newFile);
        DriveApp.getRootFolder().removeFile(newFile); // remove from root
      }
    } catch (e) {
      console.log('Could not move InstaPay sheet to folder: ' + e);
    }
  }

  // 2. Snapshot existing manual data keyed by Booking ID
  //    Columns (1-indexed in ipSheet):
  //      A=1 AGENT, B=2 Booking ID, I=9 UPDATE, J=10 Converted ID, K=11 CALL-UPDATES
  var manual = {};  // { bookingId: { agent, update, converted, notes } }
  var ipLastRow = ipSheet.getLastRow();
  if (ipLastRow > 1) {
    var existing = ipSheet.getRange(2, 1, ipLastRow - 1, 12).getValues();
    existing.forEach(function(r) {
      var bid = r[1]; // col B
      if (bid) {
        manual[bid] = {
          agent:     r[0],   // A
          update:    r[8],   // I
          converted: r[9],   // J
          // col K (index 10) = INSTA PAY ATTEMPTS — auto, not preserved
          notes:     r[11]   // L: CALL-UPDATES
        };
      }
    });
  }

  // 3. Read Raw Booking Data and filter insta_pay rows
  // RBD column indices (0-based):
  //   A=0 BookingDate, C=2 TravelStart, G=6 BookingID, H=7 Campsite,
  //   J=9 Accommodation, L=11 CamperName, M=12 Phone, N=13 Email, AG=32 Status
  var rbdAll  = rbdSheet.getDataRange().getValues();
  var newRows = [];

  // Build lookups of confirmed bookings keyed by email and phone
  // Each entry stores { numId, bid } so we can return the actual booking ID
  var confirmedByEmail = {};
  var confirmedByPhone = {};

  function numericId(bid) {
    var m = (bid || '').toString().match(/(\d+)$/);
    return m ? parseInt(m[1], 10) : null;
  }

  for (var j = 1; j < rbdAll.length; j++) {
    var rj  = rbdAll[j];
    var st  = (rj[32] || '').toString().toLowerCase();
    if (st !== 'confirmed') continue;
    var nid = numericId(rj[6]);
    if (nid === null) continue;
    var entry = { numId: nid, bid: rj[6] };
    var email = (rj[13] || '').toString().trim().toLowerCase();
    var phone = (rj[12] || '').toString().trim();
    if (email) { confirmedByEmail[email] = confirmedByEmail[email] || []; confirmedByEmail[email].push(entry); }
    if (phone) { confirmedByPhone[phone] = confirmedByPhone[phone] || []; confirmedByPhone[phone].push(entry); }
  }

  // Returns the converted Booking ID if a nearby confirmed booking exists,
  // or null if no conversion found.
  // (same email or phone, confirmed ID within next 5 booking IDs)
  function getConvertedId(bid, email, phone) {
    var nid = numericId(bid);
    if (nid === null) return null;
    var WINDOW = 5;
    function findNearby(entries) {
      if (!entries) return null;
      for (var k = 0; k < entries.length; k++) {
        if (entries[k].numId > nid && entries[k].numId <= nid + WINDOW) return entries[k].bid;
      }
      return null;
    }
    var e = (email || '').toString().trim().toLowerCase();
    var p = (phone || '').toString().trim();
    return findNearby(confirmedByEmail[e]) || findNearby(confirmedByPhone[p]);
  }

  for (var i = 1; i < rbdAll.length; i++) {
    var r      = rbdAll[i];
    var bid    = r[6];
    var status = (r[32] || '').toString().toLowerCase();
    // Match "insta_pay_intiated" (typo in raw data) and "insta_payiated" (correct)
    if (!bid || status.indexOf('insta_pay') === -1) continue;

    var convertedId = getConvertedId(bid, r[13], r[12]);
    var converted   = !!convertedId;

    // Format travel start date as DD-MMM-YY
    var rawDate  = r[2];
    var dateStr  = '';
    if (rawDate) {
      var d = (rawDate instanceof Date) ? rawDate
              : new Date(rawDate);
      if (!isNaN(d)) {
        var months = ['Jan','Feb','Mar','Apr','May','Jun',
                      'Jul','Aug','Sep','Oct','Nov','Dec'];
        dateStr = ('0' + d.getDate()).slice(-2) + '-' +
                  months[d.getMonth()] + '-' +
                  String(d.getFullYear()).slice(-2);
      } else {
        dateStr = rawDate; // keep as-is if unparseable
      }
    }

    var m = manual[bid] || {};

    // Converted rows: auto-mark "Already Booked", no agent
    // Non-converted rows: leave UPDATE blank for agent, mark agent as null (to be assigned below)
    // In both cases, if manual data already exists, preserve it
    var autoUpdate = converted ? 'Already Booked' : 'Lead to call';
    var updateVal  = m.update !== undefined && m.update !== '' ? m.update : autoUpdate;
    // Agent: converted rows get no agent; non-converted get assigned below (null = unassigned)
    var agentVal   = converted ? (m.agent || '') : m.agent; // converted → keep blank; else assign

    newRows.push([
      agentVal,             // A: AGENT
      bid,                  // B: BOOKING ID
      r[7],                 // C: CAMPSITE NAME
      r[9],                 // D: ACCOMMODATION
      dateStr,              // E: TRAVEL START DATE
      r[11],                // F: GUEST NAME
      r[12],                // G: PHONE NUMBER
      r[13],                // H: E-MAIL
      updateVal,            // I: UPDATE
      m.converted || convertedId || '',  // J: Converted Booking ID
      0,                    // K: INSTA PAY ATTEMPTS (filled in dedup step below)
      m.notes     || ''    // L: CALL-UPDATES (manual)
    ]);
    // Tag converted rows so the agent-assignment step skips them
    newRows[newRows.length - 1]._converted = converted;
  }

  // Deduplicate by person (email or phone)
  // One person may have multiple insta_pay attempts. Keep only the FIRST row
  // per unique email/phone, count all their attempts, store in col K.
  // This prevents the same person being assigned to multiple agents.
  var personKey = function(row) {
    var email = (row[7] || '').toString().trim().toLowerCase(); // col H
    var phone = (row[6] || '').toString().trim();               // col G
    return email || phone;
  };

  // Count attempts per person
  var attemptCount = {};
  newRows.forEach(function(row) {
    var key = personKey(row);
    if (key) attemptCount[key] = (attemptCount[key] || 0) + 1;
  });

  // Keep first occurrence per person, inject attempt count into col K
  var seenPersons = {};
  newRows = newRows.filter(function(row) {
    var key = personKey(row);
    if (!key) return true;
    if (seenPersons[key]) return false;
    seenPersons[key] = true;
    row[10] = attemptCount[key]; // col K: INSTA PAY ATTEMPTS (replaces CALL-UPDATES slot)
    return true;
  });

  // Auto-assign agents to new (unassigned) rows only
  // Converted rows are never assigned. Only genuine leads (no nearby confirmed
  // booking) with no existing agent get distributed between AGENTS.
  var unassigned = [];
  newRows.forEach(function(row, idx) {
    if (!row[0] && !row._converted) unassigned.push(idx);
  });

  if (AGENTS.length > 0 && unassigned.length > 0) {
    // Count how many each agent already owns in the existing data
    var tally = {};
    AGENTS.forEach(function(a) { tally[a] = 0; });
    newRows.forEach(function(row) {
      var a = row[0];
      if (a && tally.hasOwnProperty(a)) tally[a]++;
    });

    // Assign each unassigned row to whoever has the lowest count (greedy balance)
    unassigned.forEach(function(idx) {
      var pick = AGENTS.reduce(function(best, a) {
        return tally[a] < tally[best] ? a : best;
      }, AGENTS[0]);
      newRows[idx][0] = pick;
      tally[pick]++;
    });
  }

  // 4. Write to InstaPay sheet
  ipSheet.clearContents();
  ipSheet.clearFormats();
  ipSheet.getDataRange().clearDataValidations();

  // Resize to exactly 12 columns
  var curCols = ipSheet.getMaxColumns();
  if (curCols > 12) ipSheet.deleteColumns(13, curCols - 12);
  if (ipSheet.getMaxColumns() < 12) ipSheet.insertColumnsAfter(ipSheet.getMaxColumns(), 12 - ipSheet.getMaxColumns());

  var headers = [
    'AGENT', 'BOOKING ID', 'CAMPSITE NAME', 'ACCOMMODATION',
    'TRAVEL START DATE', 'GUEST NAME', 'PHONE NUMBER', 'E-MAIL',
    'UPDATE', 'Converted Booking ID', 'INSTA PAY ATTEMPTS', 'CALL-UPDATES'
  ];
  ipSheet.getRange(1, 1, 1, 12).setValues([headers]);

  if (newRows.length > 0) {
    ipSheet.getRange(2, 1, newRows.length, 12).setValues(newRows);
  }

  // 5. Formatting
  ipSheet.getRange(1, 1, 1, 12)
    .setFontWeight('bold').setHorizontalAlignment('center');
  ipSheet.getRange(1, 1, 1, 1)
    .setBackground('#BDD7EE');  // A: AGENT — manual, blue
  ipSheet.getRange(1, 2, 1, 7)
    .setBackground('#F7CAAC');  // B–H: auto cols — peach
  ipSheet.getRange(1, 9, 1, 4)
    .setBackground('#BDD7EE');  // I–L: manual/info — blue

  // Borders
  if (newRows.length > 0) {
    ipSheet.getRange(1, 1, newRows.length + 1, 12)
      .setBorder(true, true, true, true, true, true,
                 '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  }

  // Column widths
  ipSheet.setColumnWidth(1, 70);   // AGENT
  ipSheet.setColumnWidth(2, 100);  // BOOKING ID
  ipSheet.setColumnWidth(3, 180);  // CAMPSITE NAME
  ipSheet.setColumnWidth(4, 200);  // ACCOMMODATION
  ipSheet.setColumnWidth(5, 120);  // TRAVEL START DATE
  ipSheet.setColumnWidth(6, 160);  // GUEST NAME
  ipSheet.setColumnWidth(7, 120);  // PHONE NUMBER
  ipSheet.setColumnWidth(8, 220);  // E-MAIL
  ipSheet.setColumnWidth(9, 130);  // UPDATE
  ipSheet.setColumnWidth(10, 160); // Converted Booking ID
  ipSheet.setColumnWidth(11, 80);  // INSTA PAY ATTEMPTS
  ipSheet.setColumnWidth(12, 300); // CALL-UPDATES

  // Dropdown for CALL-UPDATES (col L = 12)
  // Allows free text (setAllowInvalid true) so agents can add notes when
  // selecting "Others". Applied to a large range so it persists on new rows.
  var callUpdateOptions = [
    'Change of plan',
    'Desired date not available',
    'Booked via OTA',
    'Location not desirable',
    'Others'
  ];
  var callUpdateValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(callUpdateOptions, true)
    .setAllowInvalid(true)
    .setHelpText('Select a reason. Choose "Others" and add a note if needed.')
    .build();
  ipSheet.getRange(2, 12, 500, 1).setDataValidation(callUpdateValidation);

  ipSheet.setFrozenRows(1);

  // Dropdown for UPDATE (col I = 9)
  // "Converted after follow-up" is the key status for measuring team impact.
  // "Already Booked" is auto-set by the script but agents can also set it manually.
  var updateOptions = [
    'Lead to call',
    'Converted after follow-up',
    'Not reachable',
    'Duplicate'
  ];
  var updateValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(updateOptions, true)
    .setAllowInvalid(true)
    .setHelpText('"Converted after follow-up" = team win. "Already Booked" rows have no dropdown — auto-set by script.')
    .build();
  // Apply dropdown only to non-"Already Booked" rows
  for (var ri = 0; ri < newRows.length; ri++) {
    var cellRow = ri + 2;
    var updateCell = ipSheet.getRange(cellRow, 9);
    if (updateCell.getValue() === 'Already Booked') {
      updateCell.clearDataValidations();  // no dropdown for auto-converted rows
    } else {
      updateCell.setDataValidation(updateValidation);
    }
  }

  // Agent row colour coding
  var agentColors = {
    'AK': '#FCE4EC',  // light pink
    'MS': '#E8F5E9'   // light green
  };
  var agentFullRange = ipSheet.getRange(2, 1, Math.max(newRows.length, 1), 12);
  var agentRules = [];
  Object.keys(agentColors).forEach(function(agent) {
    agentRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="' + agent + '"')
      .setBackground(agentColors[agent])
      .setRanges([agentFullRange]).build());
  });

  // Conditional formatting on UPDATE col (I = col 9)
  var updateRange = ipSheet.getRange(2, 9, Math.max(newRows.length, 1), 1);
  // Agent row rules first (lowest priority — UPDATE cell rules override on col I)
  var ipRules = agentRules.slice();
  ipRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Converted after follow-up')
    .setBackground('#B7E1CD').setFontColor('#0B5345').setBold(true)  // deep teal — team win
    .setRanges([updateRange]).build());
  ipRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Already Booked')
    .setBackground('#C6EFCE').setFontColor('#276221')  // lighter green — auto-detected
    .setRanges([updateRange]).build());
  ipRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Duplicate')
    .setBackground('#F4CCCC').setFontColor('#990000')
    .setRanges([updateRange]).build());
  ipRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Lead to call')
    .setBackground('#FFE699').setFontColor('#7F6000')
    .setRanges([updateRange]).build());
  ipRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Not reachable')
    .setBackground('#EFEFEF').setFontColor('#666666')
    .setRanges([updateRange]).build());
  // Grey out CALL-UPDATES (col L) for Already Booked rows — not relevant for pre-conversions
  ipRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$I2="Already Booked"')
    .setBackground('#EEEEEE').setFontColor('#AAAAAA')
    .setRanges([ipSheet.getRange(2, 12, 500, 1)]).build());
  ipSheet.setConditionalFormatRules(ipRules);

  // 6. Summary tab in InstaPay Tracker
  // Creates (or resets) a "Summary" tab with the pivot table.
  // Formulas reference the main "InstaPay Leads" tab by name.
  var LEADS_TAB = ipSheet.getName();
  var sumTab = ipSS.getSheetByName('Summary');
  if (!sumTab) {
    sumTab = ipSS.insertSheet('Summary');
  } else {
    sumTab.clearContents();
    sumTab.clearFormats();
  }

  // Build summary values — formulas cross-reference the leads tab
  var ref = function(col, val) {
    return '=COUNTIF(\'' + LEADS_TAB + '\'!I2:I,"' + val + '")';
  };
  var totalRef  = '=COUNTA(\'' + LEADS_TAB + '\'!B2:B)';
  var convRef   = ref('I', 'Converted after follow-up');
  var rateRef   = '=IFERROR(' + convRef + '/' + totalRef + ',"—")';

  var sumData = [
    ['METRIC',                      'COUNT'],
    ['Total Leads',                  totalRef],
    ['Lead to call',                 ref('I', 'Lead to call')],
    ['Not reachable',                ref('I', 'Not reachable')],
    ['Auto-converted (pre-call)',    ref('I', 'Already Booked')],
    ['Converted after follow-up',    convRef],
    ['Follow-up conversion rate',    rateRef]
  ];

  // Write header separately then formulas
  sumTab.getRange(1, 1, 1, 2).setValues([sumData[0]]);
  sumTab.getRange(2, 1, 6, 1).setValues(sumData.slice(1).map(function(r){ return [r[0]]; }));
  sumTab.getRange(2, 2, 6, 1).setFormulas(sumData.slice(1).map(function(r){ return [r[1]]; }));

  // Format rate as percentage
  sumTab.getRange(7, 2).setNumberFormat('0.0%');

  // Column widths
  sumTab.setColumnWidth(1, 230);
  sumTab.setColumnWidth(2, 100);

  // Header styling
  sumTab.getRange(1, 1, 1, 2)
    .setBackground('#4A4A4A').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');

  // Row styling
  var rowStyles = [
    '#F8F9FA', '#FFFFFF', '#F8F9FA', '#FFFFFF',  // rows 2-5
  ];
  rowStyles.forEach(function(bg, i) {
    sumTab.getRange(i + 2, 1, 1, 2).setBackground(bg);
  });
  // Converted after follow-up — teal highlight
  sumTab.getRange(6, 1, 1, 2)
    .setBackground('#B7E1CD').setFontColor('#0B5345').setFontWeight('bold');
  // Rate row
  sumTab.getRange(7, 1, 1, 2)
    .setBackground('#E8F5E9').setFontWeight('bold');

  // Right-align COUNT column
  sumTab.getRange(1, 2, 7, 1).setHorizontalAlignment('right');

  // Border
  sumTab.getRange(1, 1, 7, 2).setBorder(
    true, true, true, true, true, false,
    '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // Freeze header
  sumTab.setFrozenRows(1);

  // Move Summary tab to position 2 (right after the leads tab)
  ipSS.setActiveSheet(sumTab);
  ipSS.moveActiveSheet(2);

  // 8. Done
  if (silent) return; // called from smartUpdate() — it shows one combined alert instead

  safeAlert(
    '✅ InstaPay Tracker updated!\n\n' +
    newRows.length + ' leads exported.\n' +
    '"Summary" tab refreshed.\n\n' +
    '🔗 Sheet URL:\n' + ipSS.getUrl() + '\n\n' +
    'Run exportInstaPay() any time to refresh — agent notes preserved.'
  );
}

// EXPERIENCE LEADS EXPORT
//
// HOW TO USE:
//   Run this function (exportExperienceLeads) separately, same pattern as
//   exportInstaPay. Writes to a NEW TAB called "Experience Leads" inside the
//   same standalone InstaPay Tracker sheet (creates the tab on first run).
//
// WHAT IT EXPORTS (from "Raw Experience Booking Data", non-confirmed rows only):
//   Any CMX booking whose Status isn't "confirmed" — draft, payment_failed,
//   initiated, rejected, etc. — is an attempt that never converted, and
//   therefore never makes it into CL Payments (setupCLPayments only pulls
//   confirmed experience bookings in). This tab is the follow-up repository
//   for those attempts, so nothing falls through the cracks.
//
//   Col A : AGENT              ← manual, preserved across re-runs
//   Col B : BOOKING ID         ← auto
//   Col C : CAMPSITE NAME      ← auto
//   Col D : TRAVEL START DATE  ← auto
//   Col E : GUEST NAME         ← auto
//   Col F : PHONE NUMBER       ← auto
//   Col G : E-MAIL             ← auto
//   Col H : STATUS             ← auto (Draft = attempt started, no conversion /
//                                 Payment Failed / Insta Pay Initiated = reached
//                                 Razorpay but didn't finish / Rejected…)
//   Col I : UPDATE             ← manual (Already Booked / Lead to call / Duplicate…)
//   Col J : Converted Booking ID ← manual/auto
//   Col K : ATTEMPTS           ← auto (count of attempts per person)
//   Col L : CALL-UPDATES       ← dropdown (Change of plan / Date N/A / OTA / Location / Others) + free text
//
//   "Converted" detection mirrors exportInstaPay: if the same email/phone has
//   a CONFIRMED experience booking within the next 5 Booking IDs, that row is
//   auto-marked "Already Booked" instead of left as an open lead.

function exportExperienceLeads(silent) {
  // Always open the main CL sheet directly by ID — works from editor, triggers, anywhere
  var MAIN_SS_ID = PropertiesService.getScriptProperties().getProperty('MAIN_SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(MAIN_SS_ID);

  // Flexible tab lookup — handles name variants
  var redSheet = null;
  var allSheets = ss.getSheets();
  for (var si = 0; si < allSheets.length; si++) {
    if (allSheets[si].getName().indexOf('Raw Experience') === 0) {
      redSheet = allSheets[si];
      break;
    }
  }

  if (!redSheet) {
    safeAlert('ERROR: Raw Experience Booking Data tab not found.\nTabs: ' +
      allSheets.map(function(s){ return s.getName(); }).join(', '));
    return;
  }

  // Collapse repeat Booking IDs before reading — same fix as setupCLPayments(),
  // in case this runs first / on its own without setupCLPayments having deduped yet.
  var redDupesRemoved = dedupeSheetByBookingId(redSheet, 7);

  // AGENTS CONFIG — same roster as InstaPay
  var AGENTS = ['AK', 'MS'];

  // 1. Get or create the standalone InstaPay Google Sheet (shared)
  var props     = PropertiesService.getScriptProperties();
  var ipSheetId = props.getProperty('INSTAPAY_SHEET_ID');
  var ipSS;

  if (ipSheetId) {
    try {
      ipSS = SpreadsheetApp.openById(ipSheetId);
    } catch (e) {
      ipSheetId = null; // stale ID — will create fresh below
    }
  }

  if (!ipSheetId) {
    ipSS = SpreadsheetApp.create('InstaPay Tracker - Camp Monk');
    ipSS.getSheets()[0].setName('InstaPay');
    props.setProperty('INSTAPAY_SHEET_ID', ipSS.getId());

    try {
      var mainFile   = DriveApp.getFileById(ss.getId());
      var parentIter = mainFile.getParents();
      if (parentIter.hasNext()) {
        var folder  = parentIter.next();
        var newFile = DriveApp.getFileById(ipSS.getId());
        folder.addFile(newFile);
        DriveApp.getRootFolder().removeFile(newFile);
      }
    } catch (e) {
      console.log('Could not move InstaPay sheet to folder: ' + e);
    }
  }

  var elSheet = ipSS.getSheetByName('Experience Leads');
  if (!elSheet) elSheet = ipSS.insertSheet('Experience Leads');

  // 2. Snapshot existing manual data keyed by Booking ID
  var manual = {};
  var elLastRow = elSheet.getLastRow();
  if (elLastRow > 1) {
    var existing = elSheet.getRange(2, 1, elLastRow - 1, 12).getValues();
    existing.forEach(function(r) {
      var bid = r[1]; // col B
      if (bid) {
        manual[bid] = { agent: r[0], update: r[8], converted: r[9], notes: r[11] };
      }
    });
  }

  // 3. Read Raw Experience Booking Data and pull non-confirmed rows
  // Columns (0-based): A=0 BookingDate, C=2 TravelStart, G=6 BookingID,
  // H=7 Campsite/Property, K=10 GuestName(User), L=11 Phone, M=12 Email,
  // Y=24 Status
  var redAll = redSheet.getDataRange().getValues();

  var confirmedByEmail = {};
  var confirmedByPhone = {};

  function numericId(bid) {
    var m = (bid || '').toString().match(/(\d+)$/);
    return m ? parseInt(m[1], 10) : null;
  }

  for (var j = 1; j < redAll.length; j++) {
    var rj = redAll[j];
    var st = (rj[24] || '').toString().toLowerCase().trim();
    if (st !== 'confirmed') continue;
    var nid = numericId(rj[6]);
    if (nid === null) continue;
    var entry = { numId: nid, bid: rj[6] };
    var email = (rj[12] || '').toString().trim().toLowerCase();
    var phone = (rj[11] || '').toString().trim();
    if (email) { confirmedByEmail[email] = confirmedByEmail[email] || []; confirmedByEmail[email].push(entry); }
    if (phone) { confirmedByPhone[phone] = confirmedByPhone[phone] || []; confirmedByPhone[phone].push(entry); }
  }

  // Returns the converted Booking ID if a nearby confirmed booking exists,
  // or null if no conversion found (same email or phone, confirmed ID within
  // next 5 booking IDs) — same window logic as getConvertedId in exportInstaPay.
  function getConvertedId(bid, email, phone) {
    var nid = numericId(bid);
    if (nid === null) return null;
    var WINDOW = 5;
    function findNearby(entries) {
      if (!entries) return null;
      for (var k = 0; k < entries.length; k++) {
        if (entries[k].numId > nid && entries[k].numId <= nid + WINDOW) return entries[k].bid;
      }
      return null;
    }
    var e = (email || '').toString().trim().toLowerCase();
    var p = (phone || '').toString().trim();
    return findNearby(confirmedByEmail[e]) || findNearby(confirmedByPhone[p]);
  }

  // 'draft'     = admin-side booking attempt started, guest never completed/paid
  // 'initiated' = Insta Pay initiated — guest reached the Razorpay payment
  //               step but didn't finish (mirrors the property/event
  //               insta_pay_initiated flow, hence the matching purple color
  //               and the same next-5-ID conversion check below)
  var statusLabels = {
    'draft': 'Draft',
    'payment_failed': 'Payment Failed',
    'initiated': 'Insta Pay Initiated',
    'rejected': 'Rejected'
  };

  var newRows = [];
  for (var i = 1; i < redAll.length; i++) {
    var r      = redAll[i];
    var bid    = r[6];
    var status = (r[24] || '').toString().toLowerCase().trim();
    if (!bid || status === '' || status === 'confirmed') continue;

    var convertedId = getConvertedId(bid, r[12], r[11]);
    var converted   = !!convertedId;

    // Format travel start date as DD-MMM-YY
    var rawDate = r[2];
    var dateStr = '';
    if (rawDate) {
      var d = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      if (!isNaN(d)) {
        var months = ['Jan','Feb','Mar','Apr','May','Jun',
                      'Jul','Aug','Sep','Oct','Nov','Dec'];
        dateStr = ('0' + d.getDate()).slice(-2) + '-' +
                  months[d.getMonth()] + '-' +
                  String(d.getFullYear()).slice(-2);
      } else {
        dateStr = rawDate;
      }
    }

    var m = manual[bid] || {};
    var autoUpdate = converted ? 'Already Booked' : 'Lead to call';
    var updateVal  = m.update !== undefined && m.update !== '' ? m.update : autoUpdate;
    var agentVal   = converted ? (m.agent || '') : m.agent;

    newRows.push([
      agentVal,                          // A: AGENT
      bid,                                // B: BOOKING ID
      r[7],                               // C: CAMPSITE NAME
      dateStr,                            // D: TRAVEL START DATE
      r[10],                              // E: GUEST NAME
      r[11],                              // F: PHONE NUMBER
      r[12],                              // G: E-MAIL
      statusLabels[status] || status,     // H: STATUS
      updateVal,                          // I: UPDATE
      m.converted || convertedId || '',   // J: Converted Booking ID
      0,                                  // K: ATTEMPTS (filled in dedup step)
      m.notes || ''                       // L: CALL-UPDATES (manual)
    ]);
    newRows[newRows.length - 1]._converted = converted;
  }

  // Deduplicate by person (email or phone)
  var personKey = function(row) {
    var email = (row[6] || '').toString().trim().toLowerCase(); // col G
    var phone = (row[5] || '').toString().trim();                // col F
    return email || phone;
  };

  var attemptCount = {};
  newRows.forEach(function(row) {
    var key = personKey(row);
    if (key) attemptCount[key] = (attemptCount[key] || 0) + 1;
  });

  var seenPersons = {};
  newRows = newRows.filter(function(row) {
    var key = personKey(row);
    if (!key) return true;
    if (seenPersons[key]) return false;
    seenPersons[key] = true;
    row[10] = attemptCount[key]; // col K: ATTEMPTS
    return true;
  });

  // Auto-assign agents to new (unassigned) rows only
  var unassigned = [];
  newRows.forEach(function(row, idx) {
    if (!row[0] && !row._converted) unassigned.push(idx);
  });

  if (AGENTS.length > 0 && unassigned.length > 0) {
    var tally = {};
    AGENTS.forEach(function(a) { tally[a] = 0; });
    newRows.forEach(function(row) {
      var a = row[0];
      if (a && tally.hasOwnProperty(a)) tally[a]++;
    });
    unassigned.forEach(function(idx) {
      var pick = AGENTS.reduce(function(best, a) {
        return tally[a] < tally[best] ? a : best;
      }, AGENTS[0]);
      newRows[idx][0] = pick;
      tally[pick]++;
    });
  }

  // 4. Write to Experience Leads sheet
  elSheet.clearContents();
  elSheet.clearFormats();
  elSheet.getDataRange().clearDataValidations();

  var curCols = elSheet.getMaxColumns();
  if (curCols > 12) elSheet.deleteColumns(13, curCols - 12);
  if (elSheet.getMaxColumns() < 12) elSheet.insertColumnsAfter(elSheet.getMaxColumns(), 12 - elSheet.getMaxColumns());

  var elHeaders = [
    'AGENT', 'BOOKING ID', 'CAMPSITE NAME', 'TRAVEL START DATE',
    'GUEST NAME', 'PHONE NUMBER', 'E-MAIL', 'STATUS',
    'UPDATE', 'Converted Booking ID', 'ATTEMPTS', 'CALL-UPDATES'
  ];
  elSheet.getRange(1, 1, 1, 12).setValues([elHeaders]);

  if (newRows.length > 0) {
    elSheet.getRange(2, 1, newRows.length, 12).setValues(newRows);
  }

  // 5. Formatting (mirrors InstaPay tab)
  elSheet.getRange(1, 1, 1, 12)
    .setFontWeight('bold').setHorizontalAlignment('center');
  elSheet.getRange(1, 1, 1, 1)
    .setBackground('#BDD7EE');  // A: AGENT — manual, blue
  elSheet.getRange(1, 2, 1, 7)
    .setBackground('#F7CAAC');  // B–H: auto cols — peach
  elSheet.getRange(1, 9, 1, 4)
    .setBackground('#BDD7EE');  // I–L: manual/info — blue

  if (newRows.length > 0) {
    elSheet.getRange(1, 1, newRows.length + 1, 12)
      .setBorder(true, true, true, true, true, true,
                 '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  }

  elSheet.setColumnWidth(1, 70);   // AGENT
  elSheet.setColumnWidth(2, 100);  // BOOKING ID
  elSheet.setColumnWidth(3, 180);  // CAMPSITE NAME
  elSheet.setColumnWidth(4, 130);  // TRAVEL START DATE
  elSheet.setColumnWidth(5, 160);  // GUEST NAME
  elSheet.setColumnWidth(6, 120);  // PHONE NUMBER
  elSheet.setColumnWidth(7, 220);  // E-MAIL
  elSheet.setColumnWidth(8, 120);  // STATUS
  elSheet.setColumnWidth(9, 130);  // UPDATE
  elSheet.setColumnWidth(10, 160); // Converted Booking ID
  elSheet.setColumnWidth(11, 80);  // ATTEMPTS
  elSheet.setColumnWidth(12, 300); // CALL-UPDATES

  // Dropdown for CALL-UPDATES (col L = 12)
  var callUpdateOptions = [
    'Change of plan',
    'Desired date not available',
    'Booked via OTA',
    'Location not desirable',
    'Others'
  ];
  var callUpdateValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(callUpdateOptions, true)
    .setAllowInvalid(true)
    .setHelpText('Select a reason. Choose "Others" and add a note if needed.')
    .build();
  elSheet.getRange(2, 12, 500, 1).setDataValidation(callUpdateValidation);

  elSheet.setFrozenRows(1);

  // Dropdown for UPDATE (col I = 9)
  var updateOptions = [
    'Lead to call',
    'Converted after follow-up',
    'Not reachable',
    'Duplicate'
  ];
  var updateValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(updateOptions, true)
    .setAllowInvalid(true)
    .setHelpText('"Converted after follow-up" = team win. "Already Booked" rows have no dropdown — auto-set by script.')
    .build();
  for (var ri = 0; ri < newRows.length; ri++) {
    var updateCell = elSheet.getRange(ri + 2, 9);
    if (updateCell.getValue() === 'Already Booked') {
      updateCell.clearDataValidations();
    } else {
      updateCell.setDataValidation(updateValidation);
    }
  }

  // Agent row colour coding
  var agentColors = {
    'AK': '#FCE4EC',  // light pink
    'MS': '#E8F5E9'   // light green
  };
  var agentFullRange = elSheet.getRange(2, 1, Math.max(newRows.length, 1), 12);
  var agentRules = [];
  Object.keys(agentColors).forEach(function(agent) {
    agentRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="' + agent + '"')
      .setBackground(agentColors[agent])
      .setRanges([agentFullRange]).build());
  });

  // Conditional formatting on STATUS col (H = col 8)
  var statusRange = elSheet.getRange(2, 8, Math.max(newRows.length, 1), 1);
  var elRules = agentRules.slice();
  elRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Payment Failed')
    .setBackground('#F4CCCC').setFontColor('#990000')
    .setRanges([statusRange]).build());
  elRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Rejected')
    .setBackground('#FF7474').setFontColor('#7B0000')
    .setRanges([statusRange]).build());
  elRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Draft')
    .setBackground('#FFE699').setFontColor('#7F6000')
    .setRanges([statusRange]).build());
  // Purple matches the "Converted Booking" insta_pay color family used for
  // property/event bookings in CL Payments — same concept, same tab family.
  elRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Insta Pay Initiated')
    .setBackground('#E0CFFC').setFontColor('#4B0082')
    .setRanges([statusRange]).build());

  // Conditional formatting on UPDATE col (I = col 9)
  var updateRange = elSheet.getRange(2, 9, Math.max(newRows.length, 1), 1);
  elRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Converted after follow-up')
    .setBackground('#B7E1CD').setFontColor('#0B5345').setBold(true)
    .setRanges([updateRange]).build());
  elRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Already Booked')
    .setBackground('#C6EFCE').setFontColor('#276221')
    .setRanges([updateRange]).build());
  elRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Duplicate')
    .setBackground('#F4CCCC').setFontColor('#990000')
    .setRanges([updateRange]).build());
  elRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Lead to call')
    .setBackground('#FFE699').setFontColor('#7F6000')
    .setRanges([updateRange]).build());
  elRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Not reachable')
    .setBackground('#EFEFEF').setFontColor('#666666')
    .setRanges([updateRange]).build());
  // Grey out CALL-UPDATES (col L) for Already Booked rows
  elRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$I2="Already Booked"')
    .setBackground('#EEEEEE').setFontColor('#AAAAAA')
    .setRanges([elSheet.getRange(2, 12, 500, 1)]).build());
  elSheet.setConditionalFormatRules(elRules);

  // Move Experience Leads tab to position 3 (after InstaPay + Summary)
  ipSS.setActiveSheet(elSheet);
  ipSS.moveActiveSheet(3);

  // 6. Done
  if (silent) return; // called from smartUpdate() — it shows one combined alert instead

  safeAlert(
    '✅ Experience Leads updated!\n\n' +
    newRows.length + ' non-converted experience attempts exported.\n' +
    (redDupesRemoved > 0 ? redDupesRemoved + ' duplicate Booking ID row(s) removed from Raw Experience Booking Data.\n' : '') +
    '\n🔗 Sheet URL:\n' + ipSS.getUrl() + '\n\n' +
    'Run exportExperienceLeads() any time to refresh — agent notes preserved.'
  );
}

// EVENT LEADS EXPORT
//
// HOW TO USE:
//   Run this function (exportEventLeads) separately, same pattern as
//   exportExperienceLeads. Writes to a NEW TAB called "Event Leads" inside
//   the same standalone InstaPay Tracker sheet (creates the tab on first run).
//
// WHAT COUNTS AS A LEAD HERE (this is narrower than Experience Leads):
//   setupCLPayments() already accounts for every event row in one of two
//   ways: MATCHED events (same guest + same Booking Date as an existing Raw
//   Booking Data row) get annotated onto that property row's Comments 2, and
//   UNMATCHED + confirmed events get their own new CL Payments row. Neither
//   of those needs follow-up — they're already represented.
//
//   The gap is UNMATCHED + NOT confirmed: a guest who attempted an event
//   booking, has no accommodation tied to it, and never completed payment
//   (payment_failed / insta_pay initiated / cancelled / pending / etc).
//   Those rows appear nowhere else, so this tab is their repository.
//
//   Same June 2026+ cutoff as setupCLPayments() — older data is out of scope.
//
//   Col A : AGENT              ← manual, preserved across re-runs
//   Col B : BOOKING ID         ← auto
//   Col C : EVENT              ← auto (the raw Event title)
//   Col D : TRAVEL START DATE  ← auto
//   Col E : GUEST NAME         ← auto
//   Col F : PHONE NUMBER       ← auto
//   Col G : E-MAIL             ← auto
//   Col H : STATUS             ← auto (Payment Failed / Insta Pay Initiated / Cancelled / Pending / Cancellation Requested…)
//   Col I : UPDATE             ← manual (Already Booked / Lead to call / Duplicate…)
//   Col J : Converted Booking ID ← manual/auto
//   Col K : ATTEMPTS           ← auto (count of attempts per person)
//   Col L : CALL-UPDATES       ← dropdown + free text
//
//   "Converted" detection: if the same email/phone has a CONFIRMED event row
//   (matched or not — a real conversion either way) within the next 5
//   Booking IDs, that row is auto-marked "Already Booked" instead of left
//   as an open lead.

function exportEventLeads(silent) {
  var MAIN_SS_ID = PropertiesService.getScriptProperties().getProperty('MAIN_SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(MAIN_SS_ID);

  var rbdSheet = null;
  var evtSheet = null;
  var allSheets = ss.getSheets();
  for (var si = 0; si < allSheets.length; si++) {
    var nm = allSheets[si].getName();
    if (!rbdSheet && nm.indexOf('Raw Booking') === 0) rbdSheet = allSheets[si];
    if (!evtSheet && nm.indexOf('Raw Event') === 0) evtSheet = allSheets[si];
  }

  if (!evtSheet) {
    safeAlert('ERROR: Raw Event Booking Data tab not found.\nTabs: ' +
      allSheets.map(function(s){ return s.getName(); }).join(', '));
    return;
  }

  // Collapse repeat Booking IDs before reading — same fix as elsewhere.
  var evtDupesRemoved = dedupeSheetByBookingId(evtSheet, 7);

  var AGENTS = ['AK', 'MS'];
  var EVENT_CUTOFF = new Date(2026, 5, 1).getTime(); // June 1, 2026 — same as setupCLPayments()

  function parseDateForSortLocal(val) {
    if (val instanceof Date) return val.getTime();
    if (typeof val === 'number') return val;
    var s = (val || '').toString().trim();
    if (!s) return 0;
    var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return new Date(iso[1], iso[2] - 1, iso[3]).getTime();
    var dmy = s.match(/^(\d{2})-(\d{2})-(\d{4})/);
    if (dmy) return new Date(dmy[3], dmy[2] - 1, dmy[1]).getTime();
    var d = new Date(s);
    return isNaN(d) ? 0 : d.getTime();
  }

  // 1. Get or create the standalone InstaPay Google Sheet (shared)
  var props     = PropertiesService.getScriptProperties();
  var ipSheetId = props.getProperty('INSTAPAY_SHEET_ID');
  var ipSS;

  if (ipSheetId) {
    try {
      ipSS = SpreadsheetApp.openById(ipSheetId);
    } catch (e) {
      ipSheetId = null;
    }
  }

  if (!ipSheetId) {
    ipSS = SpreadsheetApp.create('InstaPay Tracker - Camp Monk');
    ipSS.getSheets()[0].setName('InstaPay');
    props.setProperty('INSTAPAY_SHEET_ID', ipSS.getId());
    try {
      var mainFile   = DriveApp.getFileById(ss.getId());
      var parentIter = mainFile.getParents();
      if (parentIter.hasNext()) {
        var folder  = parentIter.next();
        var newFile = DriveApp.getFileById(ipSS.getId());
        folder.addFile(newFile);
        DriveApp.getRootFolder().removeFile(newFile);
      }
    } catch (e) {
      console.log('Could not move InstaPay sheet to folder: ' + e);
    }
  }

  var evlSheet = ipSS.getSheetByName('Event Leads');
  if (!evlSheet) evlSheet = ipSS.insertSheet('Event Leads');

  // 2. Snapshot existing manual data keyed by Booking ID
  var manual = {};
  var evlLastRow = evlSheet.getLastRow();
  if (evlLastRow > 1) {
    var existing = evlSheet.getRange(2, 1, evlLastRow - 1, 12).getValues();
    existing.forEach(function(r) {
      var bid = r[1];
      if (bid) manual[bid] = { agent: r[0], update: r[8], converted: r[9], notes: r[11] };
    });
  }

  // 3. Raw Booking Data lookup — same match key as Step 0.B in setupCLPayments
  var rbdAll = rbdSheet ? rbdSheet.getDataRange().getValues() : [];
  var propByEmailDate = {};
  var propByPhoneDate = {};
  for (var pi = 1; pi < rbdAll.length; pi++) {
    var prow = rbdAll[pi];
    if (!prow[6]) continue; // G: Booking ID
    var pdate  = toDayKey(parseDateForSortLocal(prow[0])); // day-only — see toDayKey() near letterToIndexTop
    var pemail = (prow[13] || '').toString().trim().toLowerCase();
    var pphone = (prow[12] || '').toString().trim();
    if (pemail) { var ek = pemail + '|' + pdate; propByEmailDate[ek] = true; }
    if (pphone) { var pk = pphone + '|' + pdate; propByPhoneDate[pk] = true; }
  }

  // 4. Raw Event Booking Data — confirmed lookups (for conversion check)
  var evtAll = evtSheet.getDataRange().getValues();

  function numericId(bid) {
    var m = (bid || '').toString().match(/(\d+)$/);
    return m ? parseInt(m[1], 10) : null;
  }

  var confirmedByEmail = {};
  var confirmedByPhone = {};
  for (var ci = 1; ci < evtAll.length; ci++) {
    var crow = evtAll[ci];
    var cstatus = (crow[20] || '').toString().toLowerCase().trim(); // U: Status
    if (cstatus !== 'confirmed') continue;
    var cnid = numericId(crow[6]);
    if (cnid === null) continue;
    var centry = { numId: cnid, bid: crow[6] };
    var cemail = (crow[13] || '').toString().trim().toLowerCase();
    var cphone = (crow[12] || '').toString().trim();
    if (cemail) { confirmedByEmail[cemail] = confirmedByEmail[cemail] || []; confirmedByEmail[cemail].push(centry); }
    if (cphone) { confirmedByPhone[cphone] = confirmedByPhone[cphone] || []; confirmedByPhone[cphone].push(centry); }
  }

  function getConvertedId(bid, email, phone) {
    var nid = numericId(bid);
    if (nid === null) return null;
    var WINDOW = 5;
    function findNearby(entries) {
      if (!entries) return null;
      for (var k = 0; k < entries.length; k++) {
        if (entries[k].numId > nid && entries[k].numId <= nid + WINDOW) return entries[k].bid;
      }
      return null;
    }
    var e = (email || '').toString().trim().toLowerCase();
    var p = (phone || '').toString().trim();
    return findNearby(confirmedByEmail[e]) || findNearby(confirmedByPhone[p]);
  }

  var statusLabels = {
    'payment_failed': 'Payment Failed',
    'cancelled': 'Cancelled',
    'failed': 'Failed',
    'pending': 'Pending',
    'request_cancellation': 'Cancellation Requested'
  };

  // 5. Build leads: UNMATCHED + NOT confirmed + June 2026 onward only
  var newRows = [];
  for (var i = 1; i < evtAll.length; i++) {
    var r      = evtAll[i];
    var bid    = r[6]; // G
    if (!bid) continue;
    var dateVal = parseDateForSortLocal(r[0]);
    if (dateVal < EVENT_CUTOFF) continue;

    var status = (r[20] || '').toString().toLowerCase().trim(); // U
    if (status === 'confirmed') continue; // already its own CL Payments row

    var email = (r[13] || '').toString().trim().toLowerCase();
    var phone = (r[12] || '').toString().trim();
    var dayKey = toDayKey(dateVal); // must match the day-only key propByEmailDate/propByPhoneDate were built with above
    var isMatched = (email && propByEmailDate[email + '|' + dayKey]) ||
                    (phone && propByPhoneDate[phone + '|' + dayKey]);
    if (isMatched) continue; // already annotated onto the property row

    var convertedId = getConvertedId(bid, email, phone);
    var converted    = !!convertedId;

    var rawDate = r[2]; // C: Travel Start Date
    var dateStr = '';
    if (rawDate) {
      var d = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      if (!isNaN(d)) {
        var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        dateStr = ('0' + d.getDate()).slice(-2) + '-' + months[d.getMonth()] + '-' + String(d.getFullYear()).slice(-2);
      } else {
        dateStr = rawDate;
      }
    }

    var isInstaPay = status.indexOf('insta_pay') !== -1;
    var statusLabel = isInstaPay ? 'Insta Pay Initiated' : (statusLabels[status] || status);

    var m = manual[bid] || {};
    var autoUpdate = converted ? 'Already Booked' : 'Lead to call';
    var updateVal  = m.update !== undefined && m.update !== '' ? m.update : autoUpdate;
    var agentVal   = converted ? (m.agent || '') : m.agent;

    newRows.push([
      agentVal,                          // A: AGENT
      bid,                                // B: BOOKING ID
      (r[9] || '').toString(),            // C: EVENT (raw title)
      dateStr,                            // D: TRAVEL START DATE
      r[11],                              // E: GUEST NAME (L: User)
      phone,                              // F: PHONE NUMBER
      email,                              // G: E-MAIL
      statusLabel,                        // H: STATUS
      updateVal,                          // I: UPDATE
      m.converted || convertedId || '',   // J: Converted Booking ID
      0,                                  // K: ATTEMPTS (filled in dedup step)
      m.notes || ''                       // L: CALL-UPDATES (manual)
    ]);
    newRows[newRows.length - 1]._converted = converted;
  }

  // Deduplicate by person
  var personKey = function(row) {
    var email = (row[6] || '').toString().trim().toLowerCase(); // col G
    var phone = (row[5] || '').toString().trim();                // col F
    return email || phone;
  };
  var attemptCount = {};
  newRows.forEach(function(row) {
    var key = personKey(row);
    if (key) attemptCount[key] = (attemptCount[key] || 0) + 1;
  });
  var seenPersons = {};
  newRows = newRows.filter(function(row) {
    var key = personKey(row);
    if (!key) return true;
    if (seenPersons[key]) return false;
    seenPersons[key] = true;
    row[10] = attemptCount[key];
    return true;
  });

  // Auto-assign agents
  var unassigned = [];
  newRows.forEach(function(row, idx) {
    if (!row[0] && !row._converted) unassigned.push(idx);
  });
  if (AGENTS.length > 0 && unassigned.length > 0) {
    var tally = {};
    AGENTS.forEach(function(a) { tally[a] = 0; });
    newRows.forEach(function(row) {
      var a = row[0];
      if (a && tally.hasOwnProperty(a)) tally[a]++;
    });
    unassigned.forEach(function(idx) {
      var pick = AGENTS.reduce(function(best, a) { return tally[a] < tally[best] ? a : best; }, AGENTS[0]);
      newRows[idx][0] = pick;
      tally[pick]++;
    });
  }

  // 6. Write to Event Leads sheet
  evlSheet.clearContents();
  evlSheet.clearFormats();
  evlSheet.getDataRange().clearDataValidations();

  var curCols = evlSheet.getMaxColumns();
  if (curCols > 12) evlSheet.deleteColumns(13, curCols - 12);
  if (evlSheet.getMaxColumns() < 12) evlSheet.insertColumnsAfter(evlSheet.getMaxColumns(), 12 - evlSheet.getMaxColumns());

  var evlHeaders = [
    'AGENT', 'BOOKING ID', 'EVENT', 'TRAVEL START DATE',
    'GUEST NAME', 'PHONE NUMBER', 'E-MAIL', 'STATUS',
    'UPDATE', 'Converted Booking ID', 'ATTEMPTS', 'CALL-UPDATES'
  ];
  evlSheet.getRange(1, 1, 1, 12).setValues([evlHeaders]);
  if (newRows.length > 0) {
    evlSheet.getRange(2, 1, newRows.length, 12).setValues(newRows);
  }

  // 7. Formatting (mirrors Experience Leads / InstaPay tabs)
  evlSheet.getRange(1, 1, 1, 12).setFontWeight('bold').setHorizontalAlignment('center');
  evlSheet.getRange(1, 1, 1, 1).setBackground('#BDD7EE');
  evlSheet.getRange(1, 2, 1, 7).setBackground('#F7CAAC');
  evlSheet.getRange(1, 9, 1, 4).setBackground('#BDD7EE');

  if (newRows.length > 0) {
    evlSheet.getRange(1, 1, newRows.length + 1, 12)
      .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  }

  evlSheet.setColumnWidth(1, 70);
  evlSheet.setColumnWidth(2, 100);
  evlSheet.setColumnWidth(3, 260);
  evlSheet.setColumnWidth(4, 130);
  evlSheet.setColumnWidth(5, 160);
  evlSheet.setColumnWidth(6, 120);
  evlSheet.setColumnWidth(7, 220);
  evlSheet.setColumnWidth(8, 150);
  evlSheet.setColumnWidth(9, 130);
  evlSheet.setColumnWidth(10, 160);
  evlSheet.setColumnWidth(11, 80);
  evlSheet.setColumnWidth(12, 300);

  var callUpdateOptions = ['Change of plan','Desired date not available','Booked via OTA','Location not desirable','Others'];
  var callUpdateValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(callUpdateOptions, true).setAllowInvalid(true)
    .setHelpText('Select a reason. Choose "Others" and add a note if needed.').build();
  evlSheet.getRange(2, 12, 500, 1).setDataValidation(callUpdateValidation);

  evlSheet.setFrozenRows(1);

  var updateOptions = ['Lead to call', 'Converted after follow-up', 'Not reachable', 'Duplicate'];
  var updateValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(updateOptions, true).setAllowInvalid(true)
    .setHelpText('"Converted after follow-up" = team win. "Already Booked" rows have no dropdown — auto-set by script.')
    .build();
  for (var ri = 0; ri < newRows.length; ri++) {
    var updateCell = evlSheet.getRange(ri + 2, 9);
    if (updateCell.getValue() === 'Already Booked') {
      updateCell.clearDataValidations();
    } else {
      updateCell.setDataValidation(updateValidation);
    }
  }

  var agentColors = { 'AK': '#FCE4EC', 'MS': '#E8F5E9' };
  var agentFullRange = evlSheet.getRange(2, 1, Math.max(newRows.length, 1), 12);
  var agentRules = [];
  Object.keys(agentColors).forEach(function(agent) {
    agentRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="' + agent + '"')
      .setBackground(agentColors[agent]).setRanges([agentFullRange]).build());
  });

  var statusRange = evlSheet.getRange(2, 8, Math.max(newRows.length, 1), 1);
  var evlRules = agentRules.slice();
  evlRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Payment Failed').setBackground('#F4CCCC').setFontColor('#990000')
    .setRanges([statusRange]).build());
  evlRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Failed').setBackground('#FF7474').setFontColor('#7B0000')
    .setRanges([statusRange]).build());
  evlRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Cancelled').setBackground('#F4CCCC').setFontColor('#990000')
    .setRanges([statusRange]).build());
  evlRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Cancellation Requested').setBackground('#FFE0E0').setFontColor('#993333')
    .setRanges([statusRange]).build());
  evlRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pending').setBackground('#FFE699').setFontColor('#7F6000')
    .setRanges([statusRange]).build());
  evlRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Insta Pay Initiated').setBackground('#E0CFFC').setFontColor('#4B0082')
    .setRanges([statusRange]).build());

  var updateRange = evlSheet.getRange(2, 9, Math.max(newRows.length, 1), 1);
  evlRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Converted after follow-up')
    .setBackground('#B7E1CD').setFontColor('#0B5345').setBold(true)
    .setRanges([updateRange]).build());
  evlRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Already Booked')
    .setBackground('#C6EFCE').setFontColor('#276221')
    .setRanges([updateRange]).build());
  evlRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Duplicate')
    .setBackground('#F4CCCC').setFontColor('#990000')
    .setRanges([updateRange]).build());
  evlRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Lead to call')
    .setBackground('#FFE699').setFontColor('#7F6000')
    .setRanges([updateRange]).build());
  evlRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Not reachable')
    .setBackground('#EFEFEF').setFontColor('#666666')
    .setRanges([updateRange]).build());
  evlRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$I2="Already Booked"')
    .setBackground('#EEEEEE').setFontColor('#AAAAAA')
    .setRanges([evlSheet.getRange(2, 12, 500, 1)]).build());
  evlSheet.setConditionalFormatRules(evlRules);

  // Move Event Leads tab to position 4 (after InstaPay + Summary + Experience Leads)
  ipSS.setActiveSheet(evlSheet);
  ipSS.moveActiveSheet(4);

  // 8. Done
  if (silent) return;

  safeAlert(
    '✅ Event Leads updated!\n\n' +
    newRows.length + ' unmatched, non-converted event attempts exported ' +
    '(June 2026 onward only).\n' +
    (evtDupesRemoved > 0 ? evtDupesRemoved + ' duplicate Booking ID row(s) removed from Raw Event Booking Data.\n' : '') +
    '\n🔗 Sheet URL:\n' + ipSS.getUrl() + '\n\n' +
    'Run exportEventLeads() any time to refresh — agent notes preserved.'
  );
}

// TEST BOOKINGS EXPORT
//
// HOW TO USE:
//   Run this function (exportTestBookings) separately, same pattern as the
//   other exports. Writes to a NEW TAB called "Test Bookings" inside the
//   same standalone InstaPay Tracker sheet (creates the tab on first run).
//
// WHAT IT EXPORTS:
//   Every booking across Raw Booking Data, Raw Experience Booking Data, and
//   Raw Event Booking Data whose guest email ends in @think201.com or
//   @campmonk.com (staff/QA test accounts) — regardless of whether it also
//   made it into CL Payments. This is a full audit trail, not just the
//   excluded ones: the IN CL PAYMENTS column shows Yes/No so you can see at
//   a glance which test bookings had a real Razorpay payment attached
//   (kept in CL Payments — real money moved) vs. which didn't (excluded from
//   CL Payments — see setupCLPayments()'s test-email filter).
//
//   No manual columns here — this is a pure read-only log, fully rebuilt
//   every run, unlike the Leads tabs which preserve agent notes.
//
//   Col A : SOURCE          (Property / Experience / Event)
//   Col B : BOOKING ID
//   Col C : CAMPSITE / EVENT
//   Col D : GUEST NAME
//   Col E : PHONE
//   Col F : EMAIL
//   Col G : AMOUNT
//   Col H : STATUS
//   Col I : BOOKING DATE
//   Col J : IN CL PAYMENTS  (Yes/No)

function exportTestBookings(silent) {
  var MAIN_SS_ID = PropertiesService.getScriptProperties().getProperty('MAIN_SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(MAIN_SS_ID);

  var rbdSheet = null, redSheet = null, evtSheet = null, rrSheet = null;
  var allSheets = ss.getSheets();
  for (var si = 0; si < allSheets.length; si++) {
    var nm = allSheets[si].getName();
    if (!rbdSheet && nm.indexOf('Raw Booking') === 0) rbdSheet = allSheets[si];
    if (!redSheet && nm.indexOf('Raw Experience') === 0) redSheet = allSheets[si];
    if (!evtSheet && nm.indexOf('Raw Event') === 0) evtSheet = allSheets[si];
    if (!rrSheet && nm.indexOf('Raw Razorpay') === 0) rrSheet = allSheets[si];
  }

  var TEST_EMAIL_DOMAINS = ['@think201.com', '@campmonk.com'];
  function isTestEmail(email) {
    var e = (email || '').toString().trim().toLowerCase();
    if (!e) return false;
    for (var i = 0; i < TEST_EMAIL_DOMAINS.length; i++) {
      if (e.indexOf(TEST_EMAIL_DOMAINS[i]) !== -1) return true;
    }
    return false;
  }

  function parseDateForSortLocal(val) {
    if (val instanceof Date) return val.getTime();
    if (typeof val === 'number') return val;
    var s = (val || '').toString().trim();
    if (!s) return 0;
    var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return new Date(iso[1], iso[2] - 1, iso[3]).getTime();
    var dmy = s.match(/^(\d{2})-(\d{2})-(\d{4})/);
    if (dmy) return new Date(dmy[3], dmy[2] - 1, dmy[1]).getTime();
    var d = new Date(s);
    return isNaN(d) ? 0 : d.getTime();
  }

  function formatDateStr(rawDate) {
    if (!rawDate) return '';
    var d = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
    if (isNaN(d)) return rawDate.toString();
    var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return ('0' + d.getDate()).slice(-2) + '-' + months[d.getMonth()] + '-' + String(d.getFullYear()).slice(-2);
  }

  // Razorpay ID lookup — same booking-ID key CL Payments' AH column matches on.
  var razorpaySet = {};
  if (rrSheet) {
    var rrLast = rrSheet.getLastRow();
    if (rrLast > 1) {
      var rrIds = rrSheet.getRange(2, 1, rrLast - 1, 1).getValues();
      rrIds.forEach(function(r) {
        var id = (r[0] || '').toString().trim();
        if (id) razorpaySet[id] = true;
      });
    }
  }

  var rows = [];

  // cols are 0-based indices into each source's own column layout.
  function scan(sheet, sourceLabel, cols) {
    if (!sheet) return;
    var last = sheet.getLastRow();
    if (last < 2) return;
    var data = sheet.getRange(2, 1, last - 1, 26).getValues();
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      var bid = r[cols.bookingId];
      if (!bid) continue;
      var email = (r[cols.email] || '').toString();
      if (!isTestEmail(email)) continue;
      if (cols.cutoff && parseDateForSortLocal(r[0]) < cols.cutoff) continue;

      var bidStr = bid.toString().trim();
      rows.push([
        sourceLabel,
        bidStr,
        (r[cols.title] || '').toString(),
        (r[cols.guestName] || '').toString(),
        (r[cols.phone] || '').toString(),
        email,
        r[cols.amount],
        (r[cols.status] || '').toString(),
        formatDateStr(r[0]),
        razorpaySet[bidStr] ? 'Yes' : 'No'
      ]);
    }
  }

  // Column positions match PROPERTY_MAP / EXPERIENCE_MAP / EVENT_MAP in setupCLPayments().
  scan(rbdSheet, 'Property',   { bookingId: 6, title: 7,  guestName: 11, phone: 12, email: 13, amount: 14, status: 32 });
  scan(redSheet, 'Experience', { bookingId: 6, title: 7,  guestName: 10, phone: 11, email: 12, amount: 13, status: 24 });
  scan(evtSheet, 'Event',      { bookingId: 6, title: 9,  guestName: 11, phone: 12, email: 13, amount: 14, status: 20,
                                  cutoff: new Date(2026, 5, 1).getTime() }); // June 1, 2026 — same scope as setupCLPayments()

  // Clean up the old "Test Bookings" tab from a prior version of this script,
  // which lived in the InstaPay Tracker sheet — not needed there since this
  // data isn't a follow-up worklist like the Leads tabs.
  try {
    var oldIpId = PropertiesService.getScriptProperties().getProperty('INSTAPAY_SHEET_ID');
    if (oldIpId) {
      var oldIpSS = SpreadsheetApp.openById(oldIpId);
      var staleTb = oldIpSS.getSheetByName('Test Bookings');
      if (staleTb) oldIpSS.deleteSheet(staleTb);
    }
  } catch (e) {
    console.log('Could not clean up stale Test Bookings tab in InstaPay Tracker: ' + e);
  }

  // Test Bookings lives in the CL Payments workbook itself (not the
  // InstaPay Tracker) — no follow-up workflow needed on this data, it's a
  // pure audit log, so it stays next to CL Payments where it's easy to check.
  var tbSheet = ss.getSheetByName('Test Bookings');
  if (!tbSheet) tbSheet = ss.insertSheet('Test Bookings');

  // Write
  tbSheet.clearContents();
  tbSheet.clearFormats();

  var curCols = tbSheet.getMaxColumns();
  if (curCols > 10) tbSheet.deleteColumns(11, curCols - 10);
  if (tbSheet.getMaxColumns() < 10) tbSheet.insertColumnsAfter(tbSheet.getMaxColumns(), 10 - tbSheet.getMaxColumns());

  var headers = [
    'SOURCE', 'BOOKING ID', 'CAMPSITE / EVENT', 'GUEST NAME', 'PHONE',
    'EMAIL', 'AMOUNT', 'STATUS', 'BOOKING DATE', 'IN CL PAYMENTS'
  ];
  tbSheet.getRange(1, 1, 1, 10).setValues([headers]);
  if (rows.length > 0) {
    tbSheet.getRange(2, 1, rows.length, 10).setValues(rows);
  }

  // Formatting
  tbSheet.getRange(1, 1, 1, 10).setBackground('#F7CAAC').setFontWeight('bold').setHorizontalAlignment('center');
  if (rows.length > 0) {
    tbSheet.getRange(1, 1, rows.length + 1, 10)
      .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  }
  tbSheet.setColumnWidth(1, 90);
  tbSheet.setColumnWidth(2, 100);
  tbSheet.setColumnWidth(3, 260);
  tbSheet.setColumnWidth(4, 160);
  tbSheet.setColumnWidth(5, 120);
  tbSheet.setColumnWidth(6, 220);
  tbSheet.setColumnWidth(7, 90);
  tbSheet.setColumnWidth(8, 130);
  tbSheet.setColumnWidth(9, 110);
  tbSheet.setColumnWidth(10, 120);
  tbSheet.setFrozenRows(1);

  var inClRange = tbSheet.getRange(2, 10, Math.max(rows.length, 1), 1);
  var tbRules = [];
  tbRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Yes').setBackground('#C6EFCE').setFontColor('#276221')
    .setRanges([inClRange]).build());
  tbRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('No').setBackground('#EFEFEF').setFontColor('#666666')
    .setRanges([inClRange]).build());
  tbSheet.setConditionalFormatRules(tbRules);

  // Move Test Bookings tab to the end
  ss.setActiveSheet(tbSheet);
  ss.moveActiveSheet(ss.getSheets().length);

  // Done
  if (silent) return;

  var inCLCount = rows.filter(function(r) { return r[9] === 'Yes'; }).length;
  safeAlert(
    '✅ Test Bookings updated!\n\n' +
    rows.length + ' test-account bookings found (@think201.com / @campmonk.com).\n' +
    inCLCount + ' of those have a Razorpay payment and are also in CL Payments.\n\n' +
    'See the "Test Bookings" tab in this workbook.\n\n' +
    'Run exportTestBookings() any time to refresh.'
  );
}

// DAILY TRIGGER SETUP
//
// Run createInstapayTrigger() ONCE to schedule runDailyExports() every day at 9 AM
// — this runs InstaPay + Experience Leads + Event Leads + Test Bookings exports in sequence.
// Safe to re-run — removes any existing trigger first to avoid duplicates.
// To remove the trigger entirely, run deleteInstapayTrigger().

// Runs all daily exports in one call — used by the trigger, but also handy
// to run manually if you want to refresh every tracker tab at once.
function runDailyExports() {
  exportInstaPay();
  exportExperienceLeads();
  exportEventLeads();
  exportTestBookings();
}

// DEBUG — run this to diagnose why InstaPay sheet isn't populating
function debugInstaPay() {
  var MAIN_SS_ID = PropertiesService.getScriptProperties().getProperty('MAIN_SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(MAIN_SS_ID);

  var tabs = ss.getSheets().map(function(s){ return s.getName(); });
  console.log('Tabs found: ' + tabs.join(', '));

  var rbd = null;
  ss.getSheets().forEach(function(s){
    if (s.getName().indexOf('Raw Booking') === 0) rbd = s;
  });

  if (!rbd) { console.log('ERROR: Raw Booking Data tab not found'); return; }
  console.log('Raw Booking Data tab: ' + rbd.getName());

  var data = rbd.getDataRange().getValues();
  console.log('Total rows in Raw Booking Data: ' + (data.length - 1));

  // Print header row to confirm column positions
  console.log('Headers: ' + data[0].join(' | '));

  // Scan ALL columns for any cell containing "insta"
  var instaCount = 0;
  for (var i = 1; i < data.length; i++) {
    for (var c = 0; c < data[i].length; c++) {
      if ((data[i][c] || '').toString().toLowerCase().indexOf('insta') !== -1) {
        instaCount++;
        console.log('Found "insta" at row ' + (i+1) + ' col ' + (c+1) + ' (' + data[0][c] + '): ' + data[i][c]);
        break;
      }
    }
  }
  console.log('Total rows containing "insta": ' + instaCount);

  var props = PropertiesService.getScriptProperties();
  console.log('Saved InstaPay Sheet ID: ' + props.getProperty('INSTAPAY_SHEET_ID'));
}

function createInstapayTrigger() {
  // Remove any existing triggers for either the old exportInstaPay-only
  // trigger or a previous runDailyExports trigger, to avoid duplicates.
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'exportInstaPay' || fn === 'runDailyExports') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create daily trigger at 9:00–10:00 AM (Sheets picks exact minute in the window)
  ScriptApp.newTrigger('runDailyExports')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  safeAlert('✅ Daily trigger set!\n\nrunDailyExports() — InstaPay + Experience Leads + Event Leads + Test Bookings — will run automatically every morning between 9–10 AM.\nRun deleteInstapayTrigger() if you ever want to stop it.');
}

function deleteInstapayTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'exportInstaPay' || fn === 'runDailyExports') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  safeAlert(removed > 0 ? '✅ Trigger removed.' : 'No daily export trigger found.');
}

// SMART UPDATE — one menu button, safe for anyone on the team to click
//
// WHY THIS EXISTS:
//   Teammates other than the person who built this shouldn't need to know
//   that setupCLPayments() handles bookings, exportInstaPay() handles insta
//   pay leads, and exportExperienceLeads() handles day-outing attempts. This
//   button figures out which raw tabs changed and runs whatever's needed.
//
// HOW IT DECIDES SOMETHING CHANGED:
//   Hashes the full content of Raw Booking Data, Raw Experience Booking Data,
//   and Raw Razorpay Data (not just row count — catches edits to existing
//   rows too, not only new pastes at the bottom) and compares against the
//   hash saved after the last update. If all three match, nothing has
//   changed since the last click and it exits immediately with no rebuild.
//
// IF SOMETHING CHANGED:
//   Runs setupCLPayments(), exportInstaPay(), and exportExperienceLeads() in
//   sequence, all in "silent" mode (their individual success popups are
//   suppressed), then shows ONE combined confirmation at the end instead of
//   three separate ones to click through.
//
// This is the item that should be at the top of the CL Automation menu.
// The individual functions are still available below it (see onOpen) for
// cases where someone wants to force a rebuild even with no new data.

function getSheetSignature(sheet) {
  if (!sheet) return '';
  var last = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (last < 1 || lastCol < 1) return '';
  var values = sheet.getRange(1, 1, last, lastCol).getValues();
  var str = values.map(function(row) { return row.join('|'); }).join('\n');
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str);
  return Utilities.base64Encode(digest);
}

// Bump this any time the SCRIPT LOGIC changes (not the data) — it forces
// "Update Now" to rebuild on its next click even if none of the raw tabs
// changed, since a logic change can produce different output from the same
// input. Without this, pasting an updated script and clicking Update Now
// would silently no-op if nothing in the raw data itself had changed since
// the last run — you'd have to remember to use "Force Full Rebuild" instead.
var SCRIPT_VERSION = 'v62'; // v62: FIX — three gaps found while investigating Mohini's 13-Sep-2026 review (CM29309 amount showing as a date, CMQR093 showing ₹0 despite a captured ₹1,435 Razorpay payment, and CME6561/CME6562 not syncing). (1) Amount-bearing raw columns (Raw Booking/Experience/Event Data's amount/discount/GST/total fields, Raw Razorpay Data's Fee/Paid to CM) are now forced to Plain Text on paste — same protection col AE already had — so Sheets can't silently reinterpret a pasted number as a date; a new countDateCorruptedCells() reports any cell already corrupted before this ran (can't be auto-recovered, needs a manual re-paste from the original export — surfaced in both the Force Full Rebuild alert and the Update Now combined alert). (2) Raw Razorpay Data's dedupe (an overlapping-window re-paste can leave the same Booking ID twice) is no longer blind last-wins — dedupeRazorpaySheetByCompleteness() keeps whichever copy has Fee+Paid to CM populated, so a re-download taken before a payment settled can't silently zero out an already-settled row. (3) The event↔property match key (Step 0.B, and the identical key in exportEventLeads()) now floors both sides to the calendar day via a new toDayKey() before joining, instead of comparing raw ms timestamps — a stray time-of-day on either source's Booking Date cell could previously make a same-day match silently fail.
// v61: CHANGE — AG (Booking Status) now treats a matched Razorpay payment as "Booking Completed" even when the admin export you last pasted still says pending — checked via a direct COUNTIF against 'Raw Razorpay Data' (not via AH, which would be circular since AH's own formula reads AG). Priority order: confirmed/rejected (explicit admin decisions) first, then insta_pay (has to stay ahead of the Razorpay check so AW's Converted/Not-Initiated split for insta_pay rows still works), then the new Razorpay-match check, then plain pending as the final fallback. Means a stale "payment_pending" row self-corrects the moment Raw Razorpay Data shows the payment, with no need to re-download/re-paste Raw Booking Data. Requested after CM29094 stayed stuck on Payment Pending despite AH already showing a matched Razorpay ID.
// v60: FIX — AG (Booking Status) pending check was an exact-match on "pending"/"payment pending" only, so a raw admin status of "payment_pending" (underscore, what the export actually sends) fell through unclassified to the raw-text fallback. That in turn broke AW's "Payment Pending + Razorpay match → Payment Received ✓" upgrade, since AW only fires on an exact "Payment Pending" in AG. Now a SEARCH/contains check on "pending", so any variant normalizes correctly and the Razorpay-match upgrade works again.
// v59: NEW — CM-ADMIN (manual booking via admin) Razorpay rows now get added to CL Payments as their own standalone rows, same as CMQR/QR-v2 rows already were — previously these had no row in Raw Booking Data (no admin booking report exists for a manually-created booking) so they were silently dropped from CL Payments entirely despite having a real captured payment. New buildAdminRowFormulas() mirrors buildQRRowFormulas() for dates/ID/structure: B–G all show Raw Razorpay Data col AE ("Created At") as the Booking Date (NOT a travel date — there's no travel record to pull one from), H is the CM-ADMIN-##### ID itself, J (Accommodation) stays blank as a manual-fill marker, AS (Comments 1) reads "Manual booking via admin". UNLIKE QR rows, O/V do NOT use a flat 5% GST — an admin-panel booking can be for any accommodation (5% or 18% depending which), so O/V use the same VLOOKUP(J,'Master GST Data'!A:B,2,0) pattern property/OTA rows already use: typing the accommodation name into J after the fact auto-picks up the right GST% and recalculates, falling back to 18% while J is still blank. Collected via a new adminNewRowItems array (reuses the same rrForCollect read the QR block already does), concatenated into `combined` alongside qrNewRowItems, and dispatched via a new item.isAdmin branch in the row-formula-building loop. Setup-complete summary now reports how many CM-ADMIN rows were added.
// v57: NEW — Added an "Influencer Booking" flag (CL Payments col AZ, 52nd column, manual dropdown Yes/No) for comp/influencer bookings with no real payment behind them. "Yes" forces O (TOTAL) and V (GST) to 0 on that row (W/GRAND TOTAL follows automatically since it's O+V) across all three row builders (property/experience/event, standalone QR-v2, OTA), and swaps Comments 1 (AS) to read "Influencer Booking" in place of any wallet note — which also means the Wallet/LP row-color trigger (keyed off AS containing "Wallet") never fires on these rows. AZ also gets its own dedicated teal cell + full-row highlight, added ahead of the Wallet/LP full-row layer in the conditional-format rule list so it wins on the rare case both would otherwise apply to the same row. Column count everywhere on this tab goes 51→52 (A–AY to A–AZ): header write, resize, formula write, full-row conditional-format range, and the manual-override snapshot (widened defensively to whatever the sheet's actual current width is, so the very first run after this deploy — while the sheet is still 51 cols — doesn't error trying to read a 52nd column that doesn't exist yet).
// v56: NEW — Raw Booking Data and Raw Razorpay Data now get the same dedupe pass Raw Experience/Event Booking Data already had: dedupeSheetByBookingId() runs on both right after the sheets are located (Raw Booking Data col G, Raw Razorpay Data col A), collapsing accidental repeat-paste rows from overlapping admin/Razorpay export downloads down to the last (most recent) occurrence. Raw Razorpay Data's pass only touches rows that already carry a real Razorpay-native ID at that point — it runs before the QR/add-on ID-recovery steps that fill in blank IDs, so those are unaffected. Counts surface in the "Setup complete" summary alongside the existing Experience/Event dedupe counts.
// v55: NEW — Add-on payments ("Addons booking - undefined" rows, no Booking ID attached by Razorpay at payment time) now get matched to their originating booking via guest email/contact against Raw Booking Data, disambiguated by closest Travel Start Date when a guest has multiple bookings (left unresolved rather than guessed if two candidates are within 3 days of each other). Assigns 'CMADDON-<original Booking ID>-<payment_uuid prefix>' into Raw Razorpay Data col A when resolved. v1 scope: Raw Booking Data (property) only, not Experience/Event. Does NOT fold the add-on payment's amount into the original booking's CL Payments TOTAL/GRAND TOTAL — this only makes the add-on traceable to its booking, it doesn't change any revenue figures. Payment Type (col B) still shows "ADDON PAYMENT" either way (matched or not) — check col A to see whether it resolved.
// v54: Fixed the description-based Booking ID recovery regex in Raw Razorpay Data col A — it only matched [A-Za-z0-9_], so hyphenated admin IDs like "CM-ADMIN-00113" (from a description like "Property booking - CM-ADMIN-00113") silently failed to extract, leaving those rows misclassified as "Manual payment link" instead of "Manual booking via admin". Character class now includes "-".
// v53: REVERTED the day<=12 "un-swap" logic added in v51/v52 across parseDateForSort(), buildRowFormulas()'s parseDateExpr(), and buildQRRowFormulas()'s parseDateExpr(). That logic assumed any already-real Date cell with day 1-12 must have been mis-parsed by Sheets (MM/DD vs DD/MM) and swapped it back — but this was NOT a proven assumption for property/experience Booking Date or OTA Settlement Date, only for Razorpay Created At (which has its own separate, pre-existing, narrowly-scoped fix: the swappedDatesFixed cleanup). Applying it everywhere silently corrupted already-correct dates — confirmed against live data: real bookings only go back to June 2026, yet random Jan/Feb/Mar/Apr/May 2026 dates appeared after v51/v52 shipped (e.g. a real 01-Jun-26 became 06-Jan-26). v53 trusts already-real date cells as-is with no reinterpretation everywhere. What's KEPT from v51/v52: OTA rows' sort key now correctly reads Settlement Date (col M, 13-col range) to match what's displayed; and all three date formulas prefer the unambiguous manual DD/MM/YYYY position-parse over dateRef*1 / DATEVALUE() (which both risk Sheets' own locale-guessing) when the source is text.
// v52 (SUPERSEDED by v53): Found a third occurrence of the same day/month-swap bug fixed in v51 — this time in buildRowFormulas()'s parseDateExpr(), the date formula used for EVERY property/experience/event row's Booking/Travel Start/Travel End dates (B/D/F).
// v51: Fixed two date bugs that were scrambling CL Payments' row order out of chronological sequence. (1) OTA rows' sort key was computed from the guest's original OTA booking date (Raw OTA Data col B), but the DISPLAYED "Booking Date" for OTA rows shows Settlement Date (col M) per v-earlier's payout-date change — sort key now reads col M too, and the getRange width was widened from 12 to 13 cols to actually include it. (2) QR rows' "Created At" date (Raw Razorpay Data col AE) could show an absurd date like 07-Dec-26 for what should've been 12-Jul-26: the display formula's IFERROR(dateRef*1, ...) fast path trusted Sheets' auto-converted Date value blindly whenever the cell had already been (mis-)coerced from DD/MM/YYYY text using Sheets' MM/DD locale; now explicitly un-swaps day/month (day 1-12 range) instead of trusting it as-is. parseDateForSort() got the same correction so the sort key and the displayed date always agree.
// v50: Experience and Event bookings now ALSO stay in CL Payments/Sales if they have a real Razorpay payment (hasRazorpayPayment()), even when status later changes to Cancelled/Rejected — previously the "confirmed-only" filter would silently drop them on the next rebuild. Pre-payment leads (draft/initiated/payment_pending/payment_failed/insta_pay_intiated with no payment) are still correctly excluded.

// SALES REPORT EXPORT — mirrors CL Payments (+ a few raw-data-only fields)
// into the "Bookings" tab of a separate "Sales Report (Automated)" Google
// Sheet, so Sales stays live off CL Payments instead of manual copy/paste.
// Runs as part of smartUpdate() (Update Now), same one-click flow as
// everything else — no separate export step to remember.

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
  'Accommodation Category', 'Booking Lead Time (Days)', 'Promo Used' // hidden, cols 42-44
];

// Resolves the Sales report spreadsheet ID. First run: CREATES a brand new
// Google Sheet (with a "Bookings" tab + header row already in place) in the
// same Drive folder as the main CL Payments sheet, and remembers its ID in
// Script Properties from then on — same self-configuring, auto-create
// pattern already used for the InstaPay Tracker sheet (see exportInstaPay).
// No manual link-pasting needed.
function getSalesSheetId() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SALES_SHEET_ID');

  if (id) {
    try {
      SpreadsheetApp.openById(id); // confirm it still exists/is reachable
      return id;
    } catch (e) {
      id = null; // stale ID — fall through and create fresh below
    }
  }

  var MAIN_SS_ID = PropertiesService.getScriptProperties().getProperty('MAIN_SPREADSHEET_ID');
  var salesSS = SpreadsheetApp.create('Sales Report (Automated) - Camp Monk');
  var bookingsSheet = salesSS.getSheets()[0];
  bookingsSheet.setName('Bookings');
  bookingsSheet.getRange(1, 1, 1, SALES_BOOKINGS_HEADERS.length).setValues([SALES_BOOKINGS_HEADERS]);
  bookingsSheet.getRange(1, 1, 1, SALES_BOOKINGS_HEADERS.length)
    .setBackground('#F7CAAC').setFontWeight('bold').setHorizontalAlignment('center');
  bookingsSheet.setFrozenRows(1);
  try { bookingsSheet.hideColumns(42, 3); } catch (e) { /* ignore if it can't hide yet */ }

  // Move into the same Drive folder as the main CL sheet, same as InstaPay Tracker
  try {
    var mainFile   = DriveApp.getFileById(MAIN_SS_ID);
    var parentIter = mainFile.getParents();
    if (parentIter.hasNext()) {
      var folder  = parentIter.next();
      var newFile = DriveApp.getFileById(salesSS.getId());
      folder.addFile(newFile);
      DriveApp.getRootFolder().removeFile(newFile); // remove from root
    }
  } catch (e) {
    console.log('Could not move Sales report to folder: ' + e);
  }

  props.setProperty('SALES_SHEET_ID', salesSS.getId());
  return salesSS.getId();
}

// CL Payments dates are TEXT, formatted "DD- MMM- YY" (e.g. "18- Jul- 26") —
// parses that back into a real Date so lead time can be computed. Returns
// null if the text doesn't match (blank cell, unexpected format, etc.).
var CL_MONTH_MAP = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };
function parseCLDateText(text) {
  if (!text) return null;
  var m = text.toString().match(/(\d{1,2})-\s*([A-Za-z]{3})-\s*(\d{2})/);
  if (!m) return null;
  var mon = CL_MONTH_MAP[m[2].toLowerCase()];
  if (mon === undefined) return null;
  return new Date(2000 + parseInt(m[3], 10), mon, parseInt(m[1], 10));
}

// Buckets an Accommodation value into one of four categories, so pivots can
// filter down to just real room/tent types without a hand-maintained list —
// new one-off event names or charge line items get bucketed automatically
// as they show up, instead of silently falling through as "accommodation."
var ACCOMMODATION_CHARGE_KEYWORDS = [
  'charge', 'charges', 'food', 'misc', 'miscellaneous', 'gift card', 'fuel',
  'damage', 'upgrade', 'candlelight', 'participation ticket', 'add on', 'add ons', 'addon'
];
function classifyAccommodation(accomVal, eventName) {
  var v = (accomVal || '').toString().toLowerCase();
  if (eventName) return 'Event'; // already linked to a real Event Name
  if (v.indexOf('day outing') !== -1) return 'Day Outing';
  for (var k = 0; k < ACCOMMODATION_CHARGE_KEYWORDS.length; k++) {
    if (v.indexOf(ACCOMMODATION_CHARGE_KEYWORDS[k]) !== -1) return 'Addon/Charge';
  }
  return 'Accommodation Type';
}

// Returns a status object so callers (including smartUpdate, which runs
// this silently) always know what actually happened instead of assuming
// success — a silent skip/failure here used to get reported to the user as
// "refreshed" in the Update Now summary, which was misleading. Shape:
// { ok: true, count: N }  or  { ok: false, reason: '...' }
function exportToSales(silent) {
  var MAIN_SS_ID = PropertiesService.getScriptProperties().getProperty('MAIN_SPREADSHEET_ID');
  var mainSS = SpreadsheetApp.openById(MAIN_SS_ID);

  // Flexible lookup: exact name first, then any tab whose name starts with
  // the prefix — same pattern setupCLPayments() already relies on, since
  // the real tab names here are variants like "CL Payments- automated sheet"
  // rather than an exact match.
  function findMainSheet(exactName, prefix) {
    var s = mainSS.getSheetByName(exactName);
    if (s) return s;
    var all = mainSS.getSheets();
    for (var i = 0; i < all.length; i++) {
      if (all[i].getName().indexOf(prefix) === 0) return all[i];
    }
    return null;
  }

  var clSheet = findMainSheet('CL Payments', 'CL Pay');
  if (!clSheet) {
    var r1 = 'Sales export skipped — "CL Payments" tab not found.';
    if (!silent) safeAlert(r1);
    return { ok: false, reason: r1 };
  }

  var rbdSheet = findMainSheet('Raw Booking Data', 'Raw Booking');

  var salesId = getSalesSheetId();
  if (!salesId) {
    var r2 = 'Sales export skipped — could not get/create the Sales report sheet. Try again, or run "⚙️ Create New Sales Report" from the menu.';
    if (!silent) safeAlert(r2);
    return { ok: false, reason: r2 };
  }

  var salesSS, bookingsSheet;
  try {
    salesSS = SpreadsheetApp.openById(salesId);
    bookingsSheet = salesSS.getSheetByName('Bookings');
  } catch (e) {
    var r3 = 'ERROR opening Sales report — the saved link may be stale.\n\n' + e;
    if (!silent) safeAlert(r3);
    return { ok: false, reason: r3 };
  }
  if (!bookingsSheet) {
    var r4 = 'ERROR: "Bookings" tab not found in Sales report.';
    if (!silent) safeAlert(r4);
    return { ok: false, reason: r4 };
  }

  // Read CL Payments (already rebuilt earlier in this run by setupCLPayments)
  var clLastRow = clSheet.getLastRow();
  if (clLastRow < 2) {
    var r5 = 'Sales export skipped — no CL Payments data yet.';
    if (!silent) safeAlert(r5);
    return { ok: false, reason: r5 };
  }
  var clValues = clSheet.getRange(2, 1, clLastRow - 1, 51).getValues(); // A:AY

  // Booking ID → raw-only fields lookup, from Raw Booking Data
  // Raw Booking Data columns: G(7)=Booking ID, I(9)=Original property Name,
  // AH(34)=Invoice number, AJ(36)=Event name, AK(37)=Event booking ID.
  // Rows never in Raw Booking Data (QR payments, manual admin bookings,
  // future OTA imports) simply won't be in this map — that's what drives
  // the Invoice Number "NA" fallback below.
  var rawLookup = {};
  if (rbdSheet) {
    var rbdLastRow = rbdSheet.getLastRow();
    if (rbdLastRow >= 2) {
      var rbdValues = rbdSheet.getRange(2, 1, rbdLastRow - 1, 37).getValues(); // A:AK
      for (var ri = 0; ri < rbdValues.length; ri++) {
        var bid = (rbdValues[ri][6] || '').toString().trim(); // G, index 6
        if (!bid) continue;
        rawLookup[bid] = {
          originalCampsite: rbdValues[ri][8],   // I, index 8
          invoiceNumber:    rbdValues[ri][33],  // AH, index 33
          eventName:        rbdValues[ri][35],  // AJ, index 35
          eventBookingId:   rbdValues[ri][36]   // AK, index 36
        };
      }
    }
  }

  // Booking ID → Event title lookup, from Raw Event Booking Data
  // Keyed by the EVENT's own Booking ID (Raw Event Booking Data col G) — this
  // is what CL Payments' col H actually holds for an event row, for BOTH a
  // matched property row (H gets SWAPPED to the event's ID, see Step 0.B in
  // setupCLPayments) and a standalone event row (H = the event's own ID from
  // the start). rawLookup above can't resolve event rows at all because it's
  // keyed by the PROPERTY's original Booking ID, which no longer matches a
  // swapped row's H — this second lookup is what actually lets Event
  // Name/Events resolve below.
  var eventLookup = {};
  var evtSheetForSales = findMainSheet('Raw Event Booking Data', 'Raw Event');
  if (evtSheetForSales) {
    var evtLastRowForSales = evtSheetForSales.getLastRow();
    if (evtLastRowForSales >= 2) {
      var evtValuesForSales = evtSheetForSales.getRange(2, 1, evtLastRowForSales - 1, 10).getValues(); // A:J
      for (var evi = 0; evi < evtValuesForSales.length; evi++) {
        var evBid = (evtValuesForSales[evi][6] || '').toString().trim(); // G: Booking ID
        if (!evBid) continue;
        eventLookup[evBid] = (evtValuesForSales[evi][9] || '').toString(); // J: Event (title)
      }
    }
  }

  // Coerces a blank/empty raw cell to 0 for the numeric discount-family
  // fields (Promo/Flat/Referral/Weekdays discount, Wallet used). CL
  // Payments leaves these genuinely blank (not 0) whenever the source admin
  // report's own cell was blank — correct for CL Payments itself, but on
  // the Sales side a currency-formatted column showing nothing instead of
  // ₹0 reads as "data missing" when it actually just means "none applied".
  function numOrZero(v) {
    return (v === '' || v === null || v === undefined) ? 0 : v;
  }

  // Booking ID → Platform lookup, from Raw OTA Data
  // OTA-sourced rows need Group set to the actual platform (GoMMT/ClearTrip/
  // etc.), not the grandTotal/totalAdults business-rule buckets below —
  // those rules were designed for direct/website bookings and don't apply.
  var otaLookup = {};
  var otaDataSheetForSales = findMainSheet('Raw OTA Data', 'Raw OTA');
  if (otaDataSheetForSales) {
    var otaLastRowForSales = otaDataSheetForSales.getLastRow();
    if (otaLastRowForSales >= 2) {
      var otaValuesForSales = otaDataSheetForSales.getRange(2, 1, otaLastRowForSales - 1, 10).getValues(); // A:J
      for (var oi2 = 0; oi2 < otaValuesForSales.length; oi2++) {
        var otaBidForSales = (otaValuesForSales[oi2][0] || '').toString().trim(); // A: Booking ID
        if (!otaBidForSales) continue;
        otaLookup[otaBidForSales] = otaValuesForSales[oi2][9]; // J: Platform
      }
    }
  }

  // Build Sales Bookings rows (41 cols, A–AO)
  var out = [];
  for (var i = 0; i < clValues.length; i++) {
    var row = clValues[i];
    var bookingId = (row[7] || '').toString().trim(); // H, index 7
    if (!bookingId) continue;

    var campsiteName  = row[8];  // I
    var accommodation = row[9];  // J
    var grandTotal     = row[22]; // W
    var totalAdults    = row[25]; // Z

    // For a MATCHED event, bookingId here is the swapped-in EVENT id, not
    // the property's own — so rawLookup[bookingId] (keyed by property IDs)
    // never resolves. The original property Booking ID is still recorded as
    // plain text in Comments 2 ("Property Booking ID: CM18272" — see Step
    // 0.B's financeOverride.note), so pull it back out and fall back to
    // looking THAT up instead. Fixes Original Campsite Name / Invoice
    // Number going blank for every matched event row.
    var comments2 = (row[45] || '').toString(); // AT: Comments 2
    var propIdMatch = comments2.match(/Property Booking ID:\s*(\S+)/);
    var matchedPropertyBid = propIdMatch ? propIdMatch[1] : null;
    var lk = rawLookup[bookingId] || (matchedPropertyBid ? rawLookup[matchedPropertyBid] : undefined);
    var eventTitle = eventLookup[bookingId] || (lk ? lk.eventName : '');

    // BLC vs CLC vs CMV — same rule as the sheet's existing formula
    var csLower = (campsiteName || '').toString().toLowerCase();
    var segment =
      csLower === 'camp monk bannerghatta' ? 'BLC' :
      csLower === 'coffee lake camp sakleshpur' ? 'CLC' :
      csLower === 'camp monk vasind' ? 'CMV' : 'Agreggated';

    // Group — OTA-sourced rows show the actual platform (GoMMT/ClearTrip/
    // etc.); event bookings show "{segment} Event Bookings" (e.g. "BLC Event
    // Bookings"), matching the naming this sheet has always used for events;
    // everything else falls back to the existing business-rule buckets.
    //
    // Event detection was MISSING entirely until this fix — every CME-
    // prefixed event booking (Bande Sunday, Parent-Child Wilderness,
    // Campfire Singalong, etc.) silently fell through to "Direct Booking"
    // instead, because none of the business rules above ever checked for
    // it. Caught from a real example: CME6538/CME6540/etc. at Camp Monk
    // Bannerghatta were showing "Direct Booking" instead of "BLC Event
    // Bookings".
    var isEventBooking = bookingId.indexOf('CME') === 0 || !!eventTitle;
    var group = otaLookup[bookingId] ||
      (isEventBooking ? (segment + ' Event Bookings') :
      (typeof grandTotal === 'number' && grandTotal >= 25000) ? 'Group Booking' :
      (typeof totalAdults === 'number' && totalAdults > 10) ? 'BLC Group' : 'Direct Booking');

    // Invoice Number — NA whenever this Booking ID has no raw admin record
    // (manual/QR bookings, or OTA platforms — GoMMT/Airbnb/ClearTrip — none
    // of which flow through Raw Booking Data) or the raw record's invoice
    // field is itself blank.
    var invoiceNumber = (lk && lk.invoiceNumber) ? lk.invoiceNumber : 'NA';

    // Extra hidden columns (AP–AR), used only as pivot grouping fields
    var accomCategory = classifyAccommodation(accommodation, eventTitle);

    var bookingDateObj = parseCLDateText(row[1]);  // B, Booking Date
    var travelDateObj  = parseCLDateText(row[3]);  // D, Travel Start Date
    var leadTimeDays = (bookingDateObj && travelDateObj)
      ? Math.round((travelDateObj - bookingDateObj) / 86400000)
      : '';

    var promoName = (row[15] || '').toString().trim(); // P, Promocode Name
    var promoUsed = (promoName && promoName.toLowerCase() !== 'n/a') ? 'Yes' : 'No';

    // Event bookings' guest breakdown (Adults/Kids/Extra Adults/Extra Kids)
    // can legitimately show as 0/0/0/0 even when Total Guests is correct —
    // confirmed the raw Event report itself has NO age breakdown at all
    // (only one lump "Total guests" field), and even a matched property
    // record can have 0 in these fields while its own Total Guests is
    // right. Showing "0 adults, 2 guests" reads as missing data, so for
    // event bookings specifically: if the breakdown is all-zero but Total
    // Guests isn't, treat the unbroken-down headcount as all adults — same
    // assumption used elsewhere when age split isn't tracked. Scoped to
    // event bookings only; non-event 0/0/0/0 rows are left alone since
    // that gap hasn't been diagnosed the same way.
    var totalKidsVal    = row[26]; // AA
    var extraAdultsVal  = row[27]; // AB
    var extraKidsVal    = row[28]; // AC
    var totalGuestsVal  = row[29]; // AD
    var guestBreakdownIsZero =
      (typeof totalAdults === 'number' ? totalAdults === 0 : true) &&
      (typeof totalKidsVal === 'number' ? totalKidsVal === 0 : true) &&
      (typeof extraAdultsVal === 'number' ? extraAdultsVal === 0 : true) &&
      (typeof extraKidsVal === 'number' ? extraKidsVal === 0 : true);
    var displayTotalAdults =
      (isEventBooking && guestBreakdownIsZero && typeof totalGuestsVal === 'number' && totalGuestsVal > 0)
        ? totalGuestsVal
        : totalAdults;

    out.push([
      row[0],                                          // A  SL
      row[1], row[2], row[3], row[4], row[5], row[6],  // B–G  dates/months
      bookingId,                                        // H  Booking ID
      campsiteName,                                     // I  Campsite Name
      lk ? lk.originalCampsite : '',                    // J  Original campsite Name
      accommodation,                                     // K  Accommodation
      row[10], row[11], row[12], row[13],                // L–O  Camp Owner/Camper Name/Phone/Email
      row[14],                                           // P  TOTAL
      row[15],                                            // Q  Promocode Name
      numOrZero(row[16]), numOrZero(row[17]), numOrZero(row[18]), numOrZero(row[19]), numOrZero(row[20]), // R–V  Promo/Flat/Referral/Weekdays discount, Wallet used — blank raw cell now reads as ₹0, not empty
      row[21],                                            // W  GST
      grandTotal,                                        // X  GRAND TOTAL
      row[23], row[24],                                  // Y, Z  Total/Insta Units
      displayTotalAdults,                                  // AA  Total Adults
      row[26], row[27], row[28], row[29], row[30], row[31], // AB–AG  Total Kids..No of Nights
      row[32],                                            // AH  Status
      segment,                                            // AI  BLC vs CLC vs CMV
      group,                                              // AJ  Group
      eventTitle,                                          // AK  Event Name
      isEventBooking ? 'Yes' : '',                        // AL  Events — "Yes" flag used for pivots, blank = not an event
      invoiceNumber,                                      // AM  Invoice Number
      row[49],                                            // AN  Refund Amount ← CL Payments AX
      row[50],                                            // AO  Refund Type ← CL Payments AY
      accomCategory,                                       // AP  Accommodation Category — hidden, pivot use only
      leadTimeDays,                                        // AQ  Booking Lead Time (Days) — hidden, pivot use only
      promoUsed                                            // AR  Promo Used (Yes/No) — hidden, pivot use only
    ]);
  }

  // Write: full rebuild, same pattern as CL Payments
  var lastCol = 44; // A..AR (39 original + Refund Amount/Type + 3 hidden pivot-helper cols)
  if (bookingsSheet.getMaxColumns() < lastCol) {
    bookingsSheet.insertColumnsAfter(bookingsSheet.getMaxColumns(), lastCol - bookingsSheet.getMaxColumns());
  }

  // Ensure the 3 extra headers exist (harmless to rewrite each run) and stay hidden
  bookingsSheet.getRange(1, 42, 1, 3).setValues([[
    'Accommodation Category', 'Booking Lead Time (Days)', 'Promo Used'
  ]]);
  try { bookingsSheet.hideColumns(42, 3); } catch (e) { /* already hidden, ignore */ }

  var oldLastRow = bookingsSheet.getLastRow();
  if (oldLastRow > 1) {
    bookingsSheet.getRange(2, 1, oldLastRow - 1, Math.max(lastCol, bookingsSheet.getLastColumn())).clearContent();
  }
  if (out.length > 0) {
    bookingsSheet.getRange(2, 1, out.length, lastCol).setValues(out);
  }

  if (!silent) {
    safeAlert('✅ Sales export complete!\n\n' + out.length + ' bookings written to the Bookings tab.');
  }
  return { ok: true, count: out.length };
}

// Pivot-building has moved OUT of this script — it now lives as its own,
// separate Apps Script project pasted directly into the Sales Report
// spreadsheet (see Sales_Pivots_Setup.gs). This file's only cross-file job
// is exporting rows into Bookings; building the 8 analysis tabs from that
// data is a local, same-file operation over there, run from a "Sales
// Pivots" menu inside the Sales Report sheet itself.

// OTA IMPORT — mirrors "Qualified"/"Qualified - Mismatch" bookings from the
// separate OTA sheet (OTA_Payment_Setup.gs) into a local "Raw OTA Data" tab,
// so setupCLPayments() can treat OTA as a genuine 5th source (see
// buildOtaRowFormulas). Runs as part of smartUpdate(), same one-click flow
// as everything else. Always a full mirror/overwrite — Raw OTA Data is 100%
// derived, never hand-edited, so there's no upsert/diff logic to get wrong;
// a hand-typed correction on an OTA-derived CL Payments ROW is still
// preserved separately by setupCLPayments()'s existing
// manualOverridesByBookingId mechanism.

// Resolves the OTA sheet's ID. Unlike the Sales report, the OTA sheet
// already exists as its own separate spreadsheet — it can't be auto-created
// here, so the first run prompts once for its URL and remembers the ID from
// then on.
function getOtaSheetId() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('OTA_SHEET_ID');

  if (id) {
    try {
      SpreadsheetApp.openById(id); // confirm it's still reachable
      return id;
    } catch (e) {
      id = null; // stale ID — fall through and re-prompt
    }
  }

  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'Link the OTA Sheet',
    'Paste the OTA Sheet\'s URL or Spreadsheet ID (one-time setup):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return null;
  var input = resp.getResponseText().trim();
  var match = input.match(/[-\w]{25,}/); // extract ID out of a pasted URL, or use as-is
  var extractedId = match ? match[0] : input;
  if (!extractedId) return null;

  try {
    SpreadsheetApp.openById(extractedId); // validate before saving
  } catch (e) {
    safeAlert('Could not open that sheet — check the URL/ID and try again.\n\n' + e);
    return null;
  }

  props.setProperty('OTA_SHEET_ID', extractedId);
  return extractedId;
}

// Menu-only — forces re-linking if the OTA sheet ever moves or the saved ID
// goes stale, same pattern as reconnectSalesSheet().
function relinkOtaSheet() {
  PropertiesService.getScriptProperties().deleteProperty('OTA_SHEET_ID');
  var id = getOtaSheetId();
  safeAlert(id ? '✅ OTA Sheet linked.' : 'OTA Sheet not linked — try again from the menu.');
}

function importQualifiedOtaBookings(silent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function findLocalSheet(exactName, prefix) {
    var s = ss.getSheetByName(exactName);
    if (s) return s;
    var all = ss.getSheets();
    for (var i = 0; i < all.length; i++) {
      if (all[i].getName().indexOf(prefix) === 0) return all[i];
    }
    return null;
  }

  var otaId = getOtaSheetId();
  if (!otaId) {
    var r1 = 'OTA import skipped — OTA Sheet not linked yet.';
    if (!silent) safeAlert(r1);
    return { ok: false, reason: r1, count: 0 };
  }

  var otaSS, otaBookingsSheet;
  try {
    otaSS = SpreadsheetApp.openById(otaId);
    otaBookingsSheet = otaSS.getSheetByName('OTA Bookings');
  } catch (e) {
    var r2 = 'OTA import skipped — could not open the linked OTA Sheet. Try "Re-link OTA Sheet" from the menu.\n\n' + e;
    if (!silent) safeAlert(r2);
    return { ok: false, reason: r2, count: 0 };
  }
  if (!otaBookingsSheet) {
    var r3 = 'OTA import skipped — "OTA Bookings" tab not found in the linked OTA Sheet.';
    if (!silent) safeAlert(r3);
    return { ok: false, reason: r3, count: 0 };
  }

  var otaRawHeaders = [
    'Booking ID', 'Booking Date', 'Travel Start Date', 'Travel End Date',
    'Campsite Name', 'Accommodation', 'Camper Name', 'Total Units',
    'No of Nights', 'Platform', 'Host Payout', 'Qualified Status', 'Settlement Date'
  ];
  var rawOtaSheet = findLocalSheet('Raw OTA Data', 'Raw OTA');
  if (!rawOtaSheet) {
    // setupCLPayments() normally creates this — guard in case this import
    // ever runs before that has happened even once in this session.
    rawOtaSheet = ss.insertSheet('Raw OTA Data');
    rawOtaSheet.getRange(1, 1, 1, otaRawHeaders.length).setValues([otaRawHeaders]);
    rawOtaSheet.getRange(1, 1, 1, otaRawHeaders.length)
      .setBackground('#F7CAAC').setFontWeight('bold').setHorizontalAlignment('center');
    rawOtaSheet.setFrozenRows(1);
  }

  var lastRow = otaBookingsSheet.getLastRow();
  var lastCol = otaBookingsSheet.getLastColumn();
  var outRows = [];

  if (lastRow >= 2) {
    var data = otaBookingsSheet.getRange(1, 1, lastRow, lastCol).getValues();
    var header = data[0];
    var idx = {};
    for (var h = 0; h < header.length; h++) {
      var text = (header[h] || '').toString().trim();
      if (text && !(text in idx)) idx[text] = h;
    }
    function val(row, name) { var i = idx[name]; return (i === undefined) ? '' : row[i]; }

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var status = (val(row, 'Qualified Status') || '').toString();
      if (status.indexOf('Qualified') !== 0) continue; // only Qualified / Qualified - Mismatch — never Pending Payment

      outRows.push([
        val(row, 'Confirmation Code'),
        val(row, 'Booking Date'),
        val(row, 'Travel Start Date'),
        val(row, 'Travel End Date'),
        val(row, 'Campsite Name'),
        val(row, 'Accommodation'),
        val(row, 'Camper Name'),
        val(row, 'Total units'),
        val(row, 'No of Nights'),
        val(row, 'Platform'),
        val(row, 'Host Payout (Expected)'),
        status,
        val(row, 'Settlement Date') // payout-received date — used as this row's Booking Date in CL Payments
      ]);
    }
  }

  // Full mirror — clear and rewrite every time, since this tab is entirely
  // derived from the OTA sheet's current state, never hand-edited.
  var oldLastRow = rawOtaSheet.getLastRow();
  if (oldLastRow > 1) {
    rawOtaSheet.getRange(2, 1, oldLastRow - 1, rawOtaSheet.getMaxColumns()).clearContent();
  }
  if (outRows.length > 0) {
    rawOtaSheet.getRange(2, 1, outRows.length, otaRawHeaders.length).setValues(outRows);
  }

  if (!silent) {
    safeAlert('✅ OTA import complete!\n\n' + outRows.length + ' qualified OTA bookings mirrored into Raw OTA Data.');
  }
  return { ok: true, count: outRows.length };
}

// Menu-only wrapper for a manual, non-silent run.
function importQualifiedOtaBookingsManual() {
  importQualifiedOtaBookings(false);
}

function smartUpdate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function findSheet(exactName, prefix) {
    var s = ss.getSheetByName(exactName);
    if (s) return s;
    var all = ss.getSheets();
    for (var i = 0; i < all.length; i++) {
      if (all[i].getName().indexOf(prefix) === 0) return all[i];
    }
    return null;
  }

  // Mirror OTA data BEFORE the change-detection signature check below — this
  // must run every single time (cheap: one cross-spreadsheet read + a local
  // tab rewrite), never gated behind "did anything else change", otherwise a
  // newly-qualified OTA booking could silently get missed on a run where
  // none of the other 4 sources happened to change. Its own signature is
  // folded into curSig/prevSig right below, so a change here alone still
  // triggers the full rebuild.
  importQualifiedOtaBookings(true);

  var rbdSheet = findSheet('Raw Booking Data', 'Raw Booking');
  var redSheet = findSheet('Raw Experience Booking Data', 'Raw Experience');
  var evtSheet = findSheet('Raw Event Booking Data', 'Raw Event');
  var rrSheet  = findSheet('Raw Razorpay Data', 'Raw Razorpay');
  var otaDataSheet = findSheet('Raw OTA Data', 'Raw OTA');

  var props = PropertiesService.getScriptProperties();
  var curSig = {
    rbd: getSheetSignature(rbdSheet),
    red: getSheetSignature(redSheet),
    evt: getSheetSignature(evtSheet),
    rr:  getSheetSignature(rrSheet),
    ota: getSheetSignature(otaDataSheet),
    ver: SCRIPT_VERSION
  };
  var prevSig = {
    rbd: props.getProperty('SIG_RBD') || '',
    red: props.getProperty('SIG_RED') || '',
    evt: props.getProperty('SIG_EVT') || '',
    rr:  props.getProperty('SIG_RR')  || '',
    ota: props.getProperty('SIG_OTA') || '',
    ver: props.getProperty('SIG_VER') || ''
  };

  // Row-COUNT fallback, alongside the content-hash comparison above. Belt
  // and suspenders: the hash should already catch every real change, but a
  // row count is a much simpler, harder-to-get-wrong signal — if a raw
  // tab's row count differs at ALL from last run, that alone forces a
  // rebuild, no matter what the hash says. Added after a real report of
  // "Update Now" saying no new data despite fresh rows being visibly
  // present in Raw Booking Data.
  function rowCountOf(sheet) { return sheet ? sheet.getLastRow() : 0; }
  var curRows = {
    rbd: rowCountOf(rbdSheet), red: rowCountOf(redSheet), evt: rowCountOf(evtSheet),
    rr: rowCountOf(rrSheet), ota: rowCountOf(otaDataSheet)
  };
  var prevRows = {
    rbd: parseInt(props.getProperty('ROWS_RBD'), 10) || 0,
    red: parseInt(props.getProperty('ROWS_RED'), 10) || 0,
    evt: parseInt(props.getProperty('ROWS_EVT'), 10) || 0,
    rr:  parseInt(props.getProperty('ROWS_RR'),  10) || 0,
    ota: parseInt(props.getProperty('ROWS_OTA'), 10) || 0
  };
  var rowCountChanged = curRows.rbd !== prevRows.rbd || curRows.red !== prevRows.red ||
                         curRows.evt !== prevRows.evt || curRows.rr !== prevRows.rr ||
                         curRows.ota !== prevRows.ota;

  var changed = curSig.rbd !== prevSig.rbd || curSig.red !== prevSig.red ||
                curSig.evt !== prevSig.evt || curSig.rr !== prevSig.rr ||
                curSig.ota !== prevSig.ota || curSig.ver !== prevSig.ver ||
                rowCountChanged;

  if (!changed) {
    safeAlert(
      '✅ Already up to date!\n\n' +
      'No new data detected in Raw Booking Data, Raw Experience Booking Data, ' +
      'Raw Event Booking Data, Raw Razorpay Data, or the OTA sheet since the last update — nothing to rebuild.\n\n' +
      'Current row counts (incl. header): Raw Booking Data ' + curRows.rbd + ', Raw Experience Booking Data ' +
      curRows.red + ', Raw Event Booking Data ' + curRows.evt + ', Raw Razorpay Data ' + curRows.rr +
      ', Raw OTA Data ' + curRows.ota + '.\n\n' +
      'If you just pasted new data and this still looks wrong, use "Force Full Rebuild" instead — that always rebuilds regardless of this check.'
    );
    return;
  }

  var clResult = setupCLPayments(true);
  exportInstaPay(true);
  exportExperienceLeads(true);
  exportEventLeads(true);
  exportTestBookings(true);

  // Sales export runs last, off the CL Payments rows just rebuilt above.
  // Wrapped so a Sales-side hiccup (sheet not yet connected, link gone
  // stale, etc.) never blocks the CL Payments update itself. exportToSales
  // now always returns a status object — a graceful "skipped" (e.g. not
  // connected yet) is NOT a thrown exception, so it must be checked
  // explicitly here too, not just caught via try/catch, otherwise a silent
  // skip would get reported below as a success.
  var salesResult;
  try {
    salesResult = exportToSales(true);
  } catch (e) {
    salesResult = { ok: false, reason: String(e) };
  }

  // Re-hash AFTER running — dedup (Raw Experience/Event Booking Data) and the
  // Payment Type formula write (Raw Razorpay Data col B) both change tab
  // content, so the baseline for the NEXT comparison must be captured now,
  // not from the pre-run signature above.
  // Re-find Raw OTA Data — importQualifiedOtaBookings() may have just created
  // it if this is the very first run.
  otaDataSheet = findSheet('Raw OTA Data', 'Raw OTA');

  props.setProperty('SIG_RBD', getSheetSignature(rbdSheet));
  props.setProperty('SIG_RED', getSheetSignature(redSheet));
  props.setProperty('SIG_EVT', getSheetSignature(evtSheet));
  props.setProperty('SIG_RR',  getSheetSignature(rrSheet));
  props.setProperty('SIG_OTA', getSheetSignature(otaDataSheet));
  props.setProperty('SIG_VER', SCRIPT_VERSION);
  props.setProperty('ROWS_RBD', String(rowCountOf(rbdSheet)));
  props.setProperty('ROWS_RED', String(rowCountOf(redSheet)));
  props.setProperty('ROWS_EVT', String(rowCountOf(evtSheet)));
  props.setProperty('ROWS_RR',  String(rowCountOf(rrSheet)));
  props.setProperty('ROWS_OTA', String(rowCountOf(otaDataSheet)));

  safeAlert(
    '✅ Update complete!\n\n' +
    'New data was found, so CL Payments, the InstaPay Tracker, Experience ' +
    'Leads, Event Leads, and Test Bookings have all been refreshed.\n\n' +
    (salesResult && salesResult.ok
      ? '• Sales report Bookings tab refreshed too (' + salesResult.count + ' bookings)\n\n'
      : '⚠️ Sales report NOT updated — ' + (salesResult ? salesResult.reason : 'unknown error') + '\n\n') +
    (clResult && clResult.rrCompletenessOverrides > 0
      ? '⚠️  Kept a fully-settled Razorpay row over a more recent but incomplete re-paste for ' + clResult.rrCompletenessOverrides + ' Booking ID(s) — review those in Raw Razorpay Data if an amount still looks wrong.\n\n'
      : '') +
    (clResult && clResult.amountFormatIssuesFound > 0
      ? '⚠️  ' + clResult.amountFormatIssuesFound + ' amount-column cell(s) in the raw tabs are stored as a DATE, not a number (auto-converted on an earlier paste). Re-paste the affected row(s) from the original export — this can\'t be fixed automatically.\n\n'
      : '') +
    'Check each tab for details.'
  );
}

// CUSTOM MENU — adds a "CL Automation" menu to the Sheets UI on open, so
// nobody needs to go into Extensions > Apps Script to run anything.

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CL Automation')
    .addItem('🔄 Update Now (checks for new data)', 'smartUpdate')
    .addSeparator()
    .addItem('Force Full Rebuild', 'forceFullRebuildAndExportSales')
    .addItem('Export InstaPay Leads', 'exportInstaPay')
    .addItem('Export Experience Leads', 'exportExperienceLeads')
    .addItem('Export Event Leads', 'exportEventLeads')
    .addItem('Export Test Bookings', 'exportTestBookings')
    .addItem('Export to Sales Report', 'exportToSalesManual')
    .addItem('Import Qualified OTA Bookings', 'importQualifiedOtaBookingsManual')
    .addSeparator()
    .addItem('⚙️ Create New Sales Report (replaces current)', 'reconnectSalesSheet')
    .addItem('⚙️ Re-link OTA Sheet', 'relinkOtaSheet')
    .addSeparator()
    .addItem('Enable Daily Auto-Export (9AM)', 'createInstapayTrigger')
    .addItem('Disable Daily Auto-Export', 'deleteInstapayTrigger')
    .addToUi();
}

// Menu-only wrapper — exportToSales(silent) is also called internally by
// smartUpdate(), so this thin wrapper is what the menu item points at to
// get the non-silent (alert-showing) behavior on a manual click.
function exportToSalesManual() {
  exportToSales(false);
}

// Menu-only wrapper — "Force Full Rebuild" now also re-exports to the Sales
// Report immediately afterward, so a forced CL Payments rebuild never
// leaves the Sales Report stale relative to it (previously Force Full
// Rebuild only touched CL Payments — Sales Report required a separate
// "Export to Sales Report" click, which was easy to forget). Two separate
// completion alerts (CL Payments' own, then Sales export's own) rather than
// one combined message — setupCLPayments() already alerts internally, and
// refactoring it to return a summary instead of alerting was more change
// than this needed. exportToSales(false) is wrapped in try/catch, same
// pattern smartUpdate() already uses, so a Sales-side hiccup (not linked
// yet, stale ID, etc.) can't leave the user thinking the whole run failed —
// CL Payments' own rebuild already succeeded by that point regardless.
function forceFullRebuildAndExportSales() {
  setupCLPayments(false);
  try {
    exportToSales(false);
  } catch (e) {
    safeAlert('⚠️ CL Payments rebuilt successfully, but the Sales Report export failed:\n\n' + String(e));
  }
}

// Menu-only wrapper to force-create a brand new Sales Report sheet — use
// this only if the current one gets deleted/lost and needs replacing.
// Forgets the old ID first so getSalesSheetId() creates fresh rather than
// reusing a link that may no longer resolve.
function reconnectSalesSheet() {
  PropertiesService.getScriptProperties().deleteProperty('SALES_SHEET_ID');
  var id = getSalesSheetId();
  if (id) {
    var url = 'https://docs.google.com/spreadsheets/d/' + id + '/edit';
    safeAlert('✅ New Sales report created:\n\n' + url);
  }
}
