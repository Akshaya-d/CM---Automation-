/**
 * OTA_Payment_Setup.gs — separate sheet from CL Payments on purpose, so CL
 * Payments doesn't get bloated with OTA-specific raw data.
 *
 * OTAs (MMT/Goibibo etc) don't settle instantly like Razorpay. They send two
 * reports at different times — a booking report (commercial terms, not cash
 * yet) and a settlement report that shows up later, sometimes months later,
 * confirming money actually landed. A booking only "qualifies" once matched
 * to a settlement row.
 *
 * Tabs:
 * - Raw MMT Bookings - <Property>: one per property, paste GoMMT's New
 *   Bookings Report CSV as-is (header row first). Per-property because
 *   MMT's own export doesn't include a Property Name column.
 * - Raw MMT Net Earnings: one shared tab, append-only — paste new periods
 *   below what's already there, dedupes by Parent Booking ID.
 * - OTA Bookings: output, rebuilt fresh from the two raw tabs each run. One
 *   row per Parent Booking ID (multi-room bookings summed, since they
 *   settle as one reservation). Qualified Status = Pending Payment until
 *   matched, then Qualified (or Qualified - Mismatch if the settled amount
 *   doesn't match expected payout — still populated, just flagged).
 *
 * Not wired into CL Payments yet. Eventually: pull Qualified / Qualified-
 * Mismatch rows over with Group = platform, revenue = net Host Payout.
 * Commission/gross-net breakdown stays on this sheet only.
 *
 * New platform → add to PLATFORM_CONFIG below, rest is platform-agnostic.
 *
 * OTA_COL is generated from OTA_BOOKINGS_HEADERS, not hand-typed — avoids
 * the 0/1-based index bug that bit BOOKINGS_COL in Sales_Pivots_Setup.gs.
 * Don't hand-maintain two parallel column-index structures.
 */

// CONFIG

var PLATFORM_CONFIG = {
  GoMMT: {
    label: 'GoMMT',
    vendorAliases: ['MakeMyTrip', 'Goibibo', 'MMT'], // raw "Booking Vendor"/"Brand" values that map to this platform
    netEarningsTabName: 'Raw MMT Net Earnings', // periodic (e.g. monthly) consolidated settlement pull, shared across properties — the single MMT payment-report source.
    bookingsTabs: [
      // One entry per property MMT sends a separate "New Bookings Report" export for.
      // Add more rows here as more properties come online with MMT.
      { tabName: 'Raw MMT Bookings - Bannerghatta', property: 'Camp Monk Bannerghatta' },
      { tabName: 'Raw MMT Bookings - Vasind',       property: 'Camp Monk Vasind' }
    ]
  },
  ClearTrip: {
    label: 'ClearTrip',
    // Single shared tabs — ClearTrip's own export already includes hotel_name
    // per row, so no per-property split is needed (unlike MMT).
    bookingsTabName: 'Raw ClearTrip Bookings',
    settlementsTabName: 'Raw ClearTrip Settlements'
  },
  Manual: {
    // Airbnb, Stories Collective, ALIVE Booking, and any other channel with
    // no downloadable report — these are bank transfers only, hand-entered.
    // One shared tab, one row per booking, updated in place once payout
    // clears (no separate booking-report/payment-report pair to reconcile —
    // there's only ever the one manual row).
    bookingsTabName: 'Raw manual OTAs'
  }
};

// Exact header row GoMMT's "New Bookings Report" CSV export uses. Used only to
// bootstrap a blank tab so pasting is a straight paste with no re-typing.
// Columns are resolved BY NAME at parse time, not by position, so this is
// safe even if MMT reorders/adds columns later.
var MMT_BOOKING_RAW_HEADERS = [
  'BookingID', 'Vendor Booking ID', 'Parent Booking ID (only for MultiRoom booking)',
  'Booking Vendor', 'Booking Status', 'Customer Name', 'Check-in', 'Check-out', 'Booked On',
  'PAH Booking', 'Payment Status', 'Room name (Rate Plan)', 'No. of Rooms', 'Length of Stay',
  'Total Room Nights', 'Base Price', 'Room Charges (A)', 'Extra Adult/Child Charges (B)',
  'Service Charge (T)', 'Hotel Taxes (C)', 'Hotel Gross Charges (A+B+C+T)', 'Commission',
  'GST on Commission', 'Total Commission (D)', 'Gross Payable (A+B+C+T-D)', 'TCS Amount',
  'TDS Amount', 'Amount Payable to Hotel', 'Amount to be collected from customer (only for PAH booking)',
  'Amount Paid', 'Payment Ref.', 'Payment Date', 'Amount adjusted', 'Adjustment Ref.',
  'GSTN Assured', 'Customer GSTN', 'Customer Company Name', 'Hotel GSTN', 'Sales Channel Name',
  'Pre-buy ID'
];

// Exact header row GoMMT's "net earnings" settlement export uses (a second,
// alternative payment-report style — periodic/consolidated rather than
// per-transaction — used by some MMT/Goibibo seller accounts instead of, or
// alongside, bank_ref_detail_report).
var MMT_NET_EARNINGS_RAW_HEADERS = [
  'PNR', 'Brand', 'Booking ID', 'Parent Booking ID', 'Booking Status', 'Customer Name',
  'Checkin Date', 'Checkout Date', 'Booking Date', 'Pay At Property Booking', 'Payment Status*',
  'Room name', 'No. of Rooms', 'Length of Stay', 'Total Room Nights', 'Base Price', 'Service Tax',
  'Extra Adult/Child Charges', 'Property Taxes', 'Property Gross Charges', 'Commission',
  'GST on Commission', 'Total Commission', 'Gross Payable', 'TCS Amount', 'TDS Amount',
  'Amount Payable to Property', 'Amount to be collected from customer', 'Payment Ref.',
  'Payment Date', 'Adjustment Date', 'Adjustment Type', 'Amount adjusted', 'Adjusted Against PNR',
  'Net Earnings', 'Customer GSTN', 'Customer Company Name', 'Property GSTN', 'Sales Channel Name',
  'Pre Buy ID', 'Currency', 'Payment Type'
];

// Exact header row ClearTrip's booking export uses.
var CLEARTRIP_BOOKING_RAW_HEADERS = [
  'trip_id', 'ct_hotel_id', 'event_name', 'invoice_external_ref_id', 'hotel_name', 'city_name',
  'city_id', 'booking_date', 'fusion_status', 'cancellation_time', 'guest_name', 'room_name',
  'rate_plan_name', 'room_nights', 'pax', 'check_in', 'check_out', 'supplier_id',
  'supplier_voucher_number', 'booking_channel', 'payment_term', 'amount_to_pay', 'currency',
  'payment_method', 'settlement_ref_id', 'payment_due_date', 'rate_type', 'rate_category',
  'invoice_ref_id', 'discount_amount', 'esp_amount', 'extraAdultChildPrice_amount',
  'barPrice_amount', 'fees_amount', 'grossCharge_amount', 'commission_amount',
  'commissionGst_amount', 'gst_amount', 'tds_amount', 'tcs_amount', 'hcpFees_amount',
  'hcpFeesGst_amount', 'netPayable_amount', 'settlement_invoice_ref_id', 'utr_no',
  'payment_date', 'payment_status', 'transfer_amount'
];

// Exact header row ClearTrip's settlement/payment export uses. CORRECTED
// against a real export (2026-07-29) — the previous 7-column guess
// ('settlement_external_ref_id', 'amount', etc.) doesn't exist in ClearTrip's
// actual file, so every ClearTrip booking was silently stuck on "Pending
// Payment" forever, even after being paid. The real export is a near-mirror
// of the bookings report (CLEARTRIP_BOOKING_RAW_HEADERS above), just with
// payment fields moved to the front, invoice-only columns dropped (pax,
// rate_plan_name, invoice_ref_id, invoice_external_ref_id, hcpFees_amount,
// hcpFeesGst_amount), and — critically — it includes event_name so REFUND
// rows show up here too, keyed by the same trip_id as the booking. See the
// comment above parseClearTripSettlements() for how REFUND/BOOKING rows for
// the same trip_id are reconciled.
var CLEARTRIP_SETTLEMENT_RAW_HEADERS = [
  'settlement_invoice_ref_id', 'utr_no', 'payment_date', 'payment_status', 'transfer_amount',
  'trip_id', 'event_name', 'ct_hotel_id', 'hotel_name', 'city_id', 'city_name', 'booking_date',
  'fusion_status', 'cancellation_time', 'guest_name', 'room_name', 'check_in', 'check_out',
  'supplier_id', 'supplier_voucher_number', 'booking_channel', 'payment_term', 'amount_to_pay',
  'currency', 'payment_method', 'settlement_ref_id', 'payment_due_date', 'rate_type',
  'rate_category', 'discount_amount', 'esp_amount', 'extraAdultChildPrice_amount',
  'barPrice_amount', 'fees_amount', 'grossCharge_amount', 'commission_amount',
  'commissionGst_amount', 'gst_amount', 'tds_amount', 'tcs_amount', 'netPayable_amount'
];

// Manual bank-transfer bookings (Airbnb, Stories Collective, ALIVE Booking,
// etc.) — hand-filled, one row per booking. Booking Ref is whatever unique
// label you choose (e.g. guest name + date). Leave Payout Date/Payout blank
// until the bank transfer actually clears — that's what flips the row from
// "Pending Payment" to "Qualified" on the next run.
var MANUAL_BOOKING_HEADERS = [
  'Booking Ref', 'Platform', 'Campsite Name', 'Accommodation', 'Camper Name',
  'Booking Date', 'Travel Start Date', 'Travel End Date',
  'Customer Paid', 'Payout Date', 'Payout'
];

var OTA_BOOKINGS_TAB = 'OTA Bookings';

// ACCOMMODATION / CAMPSITE NAME NORMALIZATION
//
// CL Payments' GST reverse-calc is a VLOOKUP keyed on the exact Accommodation
// name against the "Master GST Data" tab, and Campsite Name is keyed against
// "GST PAN CM COMM". OTA raw exports use their own room-name text (e.g.
// "Sunny Container Home At Camp Monk Bannerghatta-CP"), so it has to be
// normalized to CL Payments' canonical names here — otherwise the future
// downstream VLOOKUPs just silently fail to match and GST/PAN come back
// blank. Canonical names below were pulled directly from the live CL
// Payments "Bookings" and "GST PAN CM COMM" tabs, not guessed.
//
// Campsite Name is simpler — it's supplied directly (bookingsTabs[].property
// for MMT, hotel_name for ClearTrip) and already matches CL Payments' exact
// strings ("Camp Monk Bannerghatta" / "Camp Monk Vasind"), so no alias table
// is needed for that one.

// Each rule: anyOf = matches if the raw text contains ANY of these (OR).
// withAll (optional) = ADDITIONALLY must contain ALL of these (AND) — used
// for qualifiers like "(Pet Friendly)" that narrow an otherwise-shared name.
// Rules are checked in order per property; first match wins.
var ACCOMMODATION_ALIASES = {
  'Camp Monk Bannerghatta': [
    { anyOf: ['sunny container home'], canonical: 'Sunny Container Home' },
    { anyOf: ['little black cabin'], canonical: 'Little Black Cabins' },
    { anyOf: ['air frame', 'airframe'], canonical: 'Air Frame Glamping Tents' },
    { anyOf: ['glamping cabin'], canonical: 'Glamping Cabins' },
    { anyOf: ['tiny blue one'], canonical: 'Tiny Blue One' },
    { anyOf: ['tiny blue two'], canonical: 'Tiny Blue Two' },
    { anyOf: ['family octagon'], canonical: 'Family Octagon Tent Stay' },
    { anyOf: ['bring your own tent', 'byot'], canonical: 'Bring your own tent or RV' },
    { anyOf: ['tent camping'], canonical: 'Tent Camping' }
  ],
  'Camp Monk Vasind': [
    { anyOf: ['tiny cabin by the river'], withAll: ['pet'], canonical: 'Tiny Cabin by the River (Pet Friendly)' },
    { anyOf: ['tiny cabin by the river'], canonical: 'Tiny Cabin by the River' },
    { anyOf: ['bring your own tent', 'byot'], canonical: 'Bring Your Own Tent (BYOT) by the river' },
    { anyOf: ['family octagon'], canonical: 'Family Octagon Tent @ Vasind' },
    { anyOf: ['octagon'], canonical: 'Octagon Tent' },
    { anyOf: ['sun dome'], canonical: 'Sun Domes Tents  @ Camp Monk Vasind' }
  ]
};

function normalizeAccommodation(property, rawText, unmatchedTracker) {
  if (!rawText) return rawText;
  var rules = ACCOMMODATION_ALIASES[property];
  if (!rules) return rawText; // unknown property — no alias table yet, pass through as-is
  var lower = rawText.toString().toLowerCase();
  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i];
    var anyMatch = rule.anyOf.some(function(kw) { return lower.indexOf(kw) !== -1; });
    if (!anyMatch) continue;
    var allMatch = !rule.withAll || rule.withAll.every(function(kw) { return lower.indexOf(kw) !== -1; });
    if (allMatch) return rule.canonical;
  }
  if (unmatchedTracker) unmatchedTracker.push(property + ' → "' + rawText + '"');
  return rawText; // no rule matched — pass through raw so nothing silently disappears, but flag it
}

// Canonical output schema for the "OTA Bookings" tab.
var OTA_BOOKINGS_HEADERS = [
  'SL', 'Confirmation Code', 'Platform',
  'Booking Date', 'Booking Mth',
  'Travel Start Date', 'Travel Mth',
  'Travel End Date', 'Travel End Mth',
  'No of Nights', 'Total units',
  'Campsite Name', 'Accommodation', 'Camper Name',
  'Booking Status', 'MMT Payment Status',
  'TOTAL (Gross Booking Value)', 'Commission', 'Commission %', 'GST on Commission',
  'TCS Amount', 'TDS Amount', 'Host Payout (Expected)',
  'Settled Amount Paid', 'Settlement Date', 'Settlement Payment Ref',
  'Amount Match', 'Qualified Status'
];

// Self-generated 1-based column map — derived FROM the headers array above,
// never hand-typed. See COLUMN INDEXING NOTE at top of file.
var OTA_COL = {};
(function buildOtaColMap() {
  var keys = [
    'SL', 'CONF_CODE', 'PLATFORM',
    'BOOKING_DATE', 'BOOKING_MTH',
    'TRAVEL_START', 'TRAVEL_MTH',
    'TRAVEL_END', 'TRAVEL_END_MTH',
    'NIGHTS', 'TOTAL_UNITS',
    'CAMPSITE', 'ACCOMMODATION', 'CAMPER_NAME',
    'BOOKING_STATUS', 'MMT_PAYMENT_STATUS',
    'TOTAL', 'COMMISSION', 'COMMISSION_PCT', 'GST_ON_COMMISSION',
    'TCS', 'TDS', 'HOST_PAYOUT',
    'SETTLED_AMOUNT', 'SETTLEMENT_DATE', 'SETTLEMENT_REF',
    'AMOUNT_MATCH', 'QUALIFIED_STATUS'
  ];
  if (keys.length !== OTA_BOOKINGS_HEADERS.length) {
    throw new Error('OTA_COL key count (' + keys.length + ') does not match OTA_BOOKINGS_HEADERS length (' + OTA_BOOKINGS_HEADERS.length + ') — fix the mismatch before running.');
  }
  for (var i = 0; i < keys.length; i++) {
    OTA_COL[keys[i]] = i + 1; // 1-based: matches Google's Range/PivotTable column convention
  }
})();

// MENU

// Single-button flow, matching CL Payments' "Update Now" convention — anyone
// can run this without needing to know the two-step setup/process split.
// processOtaData() already calls setupOtaTabs() internally (creates any
// missing tabs) before reconciling, so one click does everything.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('OTA Tools')
    .addItem('Update Now', 'processOtaData')
    .addSeparator()
    .addItem('Advanced: Rebuild Tabs From Scratch', 'setupOtaTabs')
    .addToUi();
}

function safeAlert(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

// TAB BOOTSTRAP

function setupOtaTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  PLATFORM_CONFIG.GoMMT.bookingsTabs.forEach(function(cfg) {
    var sh = ss.getSheetByName(cfg.tabName);
    if (!sh) {
      sh = ss.insertSheet(cfg.tabName);
      sh.getRange(1, 1, 1, MMT_BOOKING_RAW_HEADERS.length).setValues([MMT_BOOKING_RAW_HEADERS]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, MMT_BOOKING_RAW_HEADERS.length).setFontWeight('bold');
    }
  });

  var netEarnSh = ss.getSheetByName(PLATFORM_CONFIG.GoMMT.netEarningsTabName);
  if (!netEarnSh) {
    netEarnSh = ss.insertSheet(PLATFORM_CONFIG.GoMMT.netEarningsTabName);
    netEarnSh.getRange(1, 1, 1, MMT_NET_EARNINGS_RAW_HEADERS.length).setValues([MMT_NET_EARNINGS_RAW_HEADERS]);
    netEarnSh.setFrozenRows(1);
    netEarnSh.getRange(1, 1, 1, MMT_NET_EARNINGS_RAW_HEADERS.length).setFontWeight('bold');
  }

  var ctBookSh = ss.getSheetByName(PLATFORM_CONFIG.ClearTrip.bookingsTabName);
  if (!ctBookSh) {
    ctBookSh = ss.insertSheet(PLATFORM_CONFIG.ClearTrip.bookingsTabName);
    ctBookSh.getRange(1, 1, 1, CLEARTRIP_BOOKING_RAW_HEADERS.length).setValues([CLEARTRIP_BOOKING_RAW_HEADERS]);
    ctBookSh.setFrozenRows(1);
    ctBookSh.getRange(1, 1, 1, CLEARTRIP_BOOKING_RAW_HEADERS.length).setFontWeight('bold');
  }

  var ctSettleSh = ss.getSheetByName(PLATFORM_CONFIG.ClearTrip.settlementsTabName);
  if (!ctSettleSh) {
    ctSettleSh = ss.insertSheet(PLATFORM_CONFIG.ClearTrip.settlementsTabName);
    ctSettleSh.getRange(1, 1, 1, CLEARTRIP_SETTLEMENT_RAW_HEADERS.length).setValues([CLEARTRIP_SETTLEMENT_RAW_HEADERS]);
    ctSettleSh.setFrozenRows(1);
    ctSettleSh.getRange(1, 1, 1, CLEARTRIP_SETTLEMENT_RAW_HEADERS.length).setFontWeight('bold');
  }

  var manualSh = ss.getSheetByName(PLATFORM_CONFIG.Manual.bookingsTabName);
  if (!manualSh) {
    manualSh = ss.insertSheet(PLATFORM_CONFIG.Manual.bookingsTabName);
    manualSh.getRange(1, 1, 1, MANUAL_BOOKING_HEADERS.length).setValues([MANUAL_BOOKING_HEADERS]);
    manualSh.setFrozenRows(1);
    manualSh.getRange(1, 1, 1, MANUAL_BOOKING_HEADERS.length).setFontWeight('bold');
  }

  var otaSh = ss.getSheetByName(OTA_BOOKINGS_TAB);
  if (!otaSh) {
    otaSh = ss.insertSheet(OTA_BOOKINGS_TAB);
  }
  otaSh.clear();
  otaSh.getRange(1, 1, 1, OTA_BOOKINGS_HEADERS.length).setValues([OTA_BOOKINGS_HEADERS]);
  otaSh.setFrozenRows(1);
  otaSh.getRange(1, 1, 1, OTA_BOOKINGS_HEADERS.length).setFontWeight('bold');

  safeAlert('OTA tabs ready. Paste MMT booking exports into the "Raw MMT Bookings - <Property>" tabs, ' +
    'and stack net_earnings settlement files into "Raw MMT Net Earnings". ' +
    'Then run "Update Now".');
}

// HEADER RESOLUTION (by name, not position — see resolveBookingsColumns lesson
// in Sales_Pivots_Setup.gs)

function indexHeaders(headerRow) {
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    var text = (headerRow[i] === null || headerRow[i] === undefined) ? '' : headerRow[i].toString().trim();
    if (text && !(text in map)) map[text] = i; // 0-based array index, used directly against row arrays
  }
  return map;
}

// PARSE: Raw MMT Bookings - <Property>  →  aggregated-by-Parent-Booking-ID map

function parseMmtBookings(ss, unmatchedAccommodations) {
  var byParent = {};
  var seenBookingIds = {}; // guards against double-counting an accidentally re-pasted row

  PLATFORM_CONFIG.GoMMT.bookingsTabs.forEach(function(cfg) {
    var sh = ss.getSheetByName(cfg.tabName);
    if (!sh || sh.getLastRow() < 2) return;

    var data = sh.getDataRange().getValues();
    var h = indexHeaders(data[0]);

    function val(row, name) {
      var idx = h[name];
      return (idx === undefined) ? '' : row[idx];
    }
    function num(row, name) {
      var v = val(row, name);
      var n = parseFloat(v);
      return isNaN(n) ? 0 : n;
    }

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var bookingId = (val(row, 'BookingID') || '').toString().trim();
      if (!bookingId) continue;
      if (seenBookingIds[bookingId]) { if (dupeStats) dupeStats.mmtBookings++; continue; } // duplicate row, skip
      seenBookingIds[bookingId] = true;

      var vendorBookingId = (val(row, 'Vendor Booking ID') || '').toString().trim();
      var parentId = (val(row, 'Parent Booking ID (only for MultiRoom booking)') || '').toString().trim();
      if (!parentId) parentId = vendorBookingId.replace(/C\d+$/, ''); // strip room-suffix if this IS the parent
      if (!parentId) parentId = vendorBookingId || bookingId;

      if (!byParent[parentId]) {
        byParent[parentId] = {
          parentId: parentId,
          platform: PLATFORM_CONFIG.GoMMT.label,
          property: cfg.property,
          roomNames: [],
          totalUnits: 0,
          nights: 0,
          bookedOn: val(row, 'Booked On'),
          checkin: val(row, 'Check-in'),
          checkout: val(row, 'Check-out'),
          customerName: val(row, 'Customer Name'),
          bookingStatus: val(row, 'Booking Status'),
          paymentStatus: val(row, 'Payment Status'),
          hotelGrossCharges: 0,
          commission: 0,
          gstOnCommission: 0,
          tcsAmount: 0,
          tdsAmount: 0,
          hostPayoutExpected: 0
        };
      }

      var agg = byParent[parentId];
      var roomName = normalizeAccommodation(cfg.property, val(row, 'Room name (Rate Plan)'), unmatchedAccommodations);
      if (roomName && agg.roomNames.indexOf(roomName) === -1) agg.roomNames.push(roomName);
      agg.totalUnits += num(row, 'No. of Rooms');
      agg.nights = Math.max(agg.nights, num(row, 'Total Room Nights'));
      agg.hotelGrossCharges += num(row, 'Hotel Gross Charges (A+B+C+T)');
      agg.commission += num(row, 'Total Commission (D)');
      agg.gstOnCommission += num(row, 'GST on Commission');
      agg.tcsAmount += num(row, 'TCS Amount');
      agg.tdsAmount += num(row, 'TDS Amount');
      agg.hostPayoutExpected += num(row, 'Amount Payable to Hotel');
    }
  });

  Object.keys(byParent).forEach(function(pid) {
    var a = byParent[pid];
    a.commissionPct = a.hotelGrossCharges ? (a.commission / a.hotelGrossCharges * 100) : 0;
  });

  return byParent;
}

// PARSE: Raw MMT Net Earnings  →  by-Parent-Booking-ID settlement map
//
// Flat CSV — Payment Status* is always "Settled" here (this report only
// lists bookings that have cleared), and Net Earnings is the final,
// adjustment-inclusive payout figure. A multi-room booking can have more
// than one row under the same Parent Booking ID (one per room-line, same
// as the Bookings report), so this dedupes true accidental re-pastes at the
// room-line level (this report's own "Booking ID" column) while SUMMING Net
// Earnings across every row that shares a Parent Booking ID — the same
// aggregation parseMmtBookings() already does on the booking side, so a
// 2-room booking's total settled payout lines up with its total expected
// payout instead of only counting whichever room-line row happened to be
// seen last.

function parseMmtNetEarnings(ss) {
  var byBookingId = {};
  var seenRoomLineIds = {}; // dedupe guard — this report's own room-line-level "Booking ID"
  var sh = ss.getSheetByName(PLATFORM_CONFIG.GoMMT.netEarningsTabName);
  if (!sh || sh.getLastRow() < 2) return byBookingId;

  var data = sh.getDataRange().getValues();
  var h = indexHeaders(data[0]);
  function val(row, name) { var idx = h[name]; return (idx === undefined) ? '' : row[idx]; }

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var pid = (val(row, 'Parent Booking ID') || '').toString().trim();
    if (!pid) continue;

    var roomLineId = (val(row, 'Booking ID') || '').toString().trim() || pid; // fall back to parent id if this column is ever blank
    if (seenRoomLineIds[roomLineId]) continue; // true accidental re-paste of the same row — skip entirely, don't double-sum
    seenRoomLineIds[roomLineId] = true;

    var netEarnings = parseFloat(val(row, 'Net Earnings')) || 0;
    if (!byBookingId[pid]) {
      byBookingId[pid] = { amountPaid: 0, settlementDate: '', paymentRef: '', propertyName: '' };
    }
    byBookingId[pid].amountPaid += netEarnings; // sum across every room-line under this parent booking
    byBookingId[pid].settlementDate = val(row, 'Payment Date') || byBookingId[pid].settlementDate;
    byBookingId[pid].paymentRef = val(row, 'Payment Ref.') || byBookingId[pid].paymentRef;
  }
  return byBookingId;
}

// PARSE: Raw ClearTrip Bookings  →  { trip_id: {...} }
//
// ClearTrip's export has one row per event, not one row per booking — a
// single trip_id can appear with event_name = "BOOKING" AND, later, a second
// row with event_name = "REFUND" (same trip_id). Only BOOKING rows become
// their own output row here; REFUND rows for a trip not yet in this tab are
// simply not represented until this tab is refreshed with a newer export
// that includes them (this tab is the commitment/commercial-terms source,
// not the payment source — see parseClearTripSettlements() for how REFUND
// events actually get reconciled against payment).

function parseClearTripBookings(ss, unmatchedAccommodations) {
  var bookings = {};       // trip_id -> booking object (BOOKING events only)
  var seenTripIds = {};

  var sh = ss.getSheetByName(PLATFORM_CONFIG.ClearTrip.bookingsTabName);
  if (!sh || sh.getLastRow() < 2) return { bookings: bookings };

  var data = sh.getDataRange().getValues();
  var h = indexHeaders(data[0]);
  function val(row, name) { var idx = h[name]; return (idx === undefined) ? '' : row[idx]; }
  function num(row, name) { var n = parseFloat(val(row, name)); return isNaN(n) ? 0 : n; }

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var tripId = (val(row, 'trip_id') || '').toString().trim();
    if (!tripId) continue;

    var refId = (val(row, 'settlement_ref_id') || '').toString().trim();

    var eventName = (val(row, 'event_name') || '').toString().trim();
    if (eventName !== 'BOOKING') continue; // REFUND rows in this tab aren't turned into their own output row

    var dedupeKey = tripId + '|BOOKING';
    if (seenTripIds[dedupeKey]) continue;
    seenTripIds[dedupeKey] = true;

    var campsite = val(row, 'hotel_name');
    var rawAccommodation = [val(row, 'room_name'), val(row, 'rate_plan_name')].filter(function(x){return x;}).join(' - ');
    bookings[tripId] = {
      tripId: tripId,
      platform: PLATFORM_CONFIG.ClearTrip.label,
      refId: refId,
      campsite: campsite,
      accommodation: normalizeAccommodation(campsite, rawAccommodation, unmatchedAccommodations),
      camperName: val(row, 'guest_name'),
      bookingDate: val(row, 'booking_date'),
      checkin: val(row, 'check_in'),
      checkout: val(row, 'check_out'),
      nights: num(row, 'room_nights'),
      totalUnits: 1,
      bookingStatus: val(row, 'cancellation_time') ? 'Cancelled' : (val(row, 'fusion_status') || ''),
      platformPaymentStatus: val(row, 'payment_status'),
      grossCharge: num(row, 'grossCharge_amount'),
      commission: num(row, 'commission_amount'),
      gstOnCommission: num(row, 'commissionGst_amount'),
      tcsAmount: num(row, 'tcs_amount'),
      tdsAmount: num(row, 'tds_amount'),
      hostPayoutExpected: num(row, 'netPayable_amount')
    };
  }

  Object.keys(bookings).forEach(function(tid) {
    var b = bookings[tid];
    b.commissionPct = b.grossCharge ? (b.commission / b.grossCharge * 100) : 0;
  });

  return { bookings: bookings };
}

// PARSE: Raw ClearTrip Settlements  →  {trip_id: {netPayable, utrNo, paymentDate, paymentStatus}}
//
// Keyed directly by trip_id — the same key the bookings tab uses — so
// reconciliation is a straight lookup, no batch/UUID grouping needed.
//
// A trip_id can appear as TWO rows here: event_name = "BOOKING" (what was
// originally payable) and, if the guest later cancelled, a second
// event_name = "REFUND" row for the SAME trip_id. The REFUND row's own
// netPayable_amount is the trip's FINAL, actually-paid figure (0 for a full
// refund, a reduced number for a partial one) — it supersedes the BOOKING
// row's netPayable_amount, it does not add to it. Confirmed against a real
// export: summing both rows overcounts a batch's transfer_amount; taking
// only the REFUND row (when present) reproduces transfer_amount exactly.
//
// Multiple exports may get pasted underneath each other over time, so
// dedupe by trip_id + event_name and tolerate a repeated header row.

function parseClearTripSettlements(ss) {
  var byTripId = {};
  var sh = ss.getSheetByName(PLATFORM_CONFIG.ClearTrip.settlementsTabName);
  if (!sh || sh.getLastRow() < 2) return byTripId;

  var data = sh.getDataRange().getValues();
  var h = indexHeaders(data[0]);
  function val(row, name) { var idx = h[name]; return (idx === undefined) ? '' : row[idx]; }
  function num(row, name) { var n = parseFloat(val(row, name)); return isNaN(n) ? 0 : n; }

  var eventsByTrip = {}; // trip_id -> { BOOKING: {...}, REFUND: {...} }
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var tripId = (val(row, 'trip_id') || '').toString().trim();
    if (!tripId || tripId === 'trip_id') continue; // blank or a repeated pasted-in header row

    var eventName = (val(row, 'event_name') || '').toString().trim() || 'BOOKING';
    if (!eventsByTrip[tripId]) eventsByTrip[tripId] = {};
    if (eventsByTrip[tripId][eventName]) continue; // dedupe accidental re-paste of the same event row

    eventsByTrip[tripId][eventName] = {
      netPayable: num(row, 'netPayable_amount'),
      utrNo: val(row, 'utr_no'),
      paymentDate: val(row, 'payment_date'),
      paymentStatus: val(row, 'payment_status')
    };
  }

  Object.keys(eventsByTrip).forEach(function(tid) {
    var events = eventsByTrip[tid];
    byTripId[tid] = events.REFUND || events.BOOKING; // REFUND, when present, wins
  });

  return byTripId;
}

// PARSE: Raw manual OTAs  →  by-Booking-Ref map
//
// No separate booking-report/payment-report pair here — it's one hand-typed
// row per booking, updated in place once the bank transfer clears. Qualified
// purely on whether Payout Date + Payout are both filled in; there's no
// second source to cross-check against, so there's no "Mismatch" state for
// this source — just Pending Payment or Qualified.

function parseManualBookings(ss, unmatchedAccommodations) {
  var bookings = {};
  var sh = ss.getSheetByName(PLATFORM_CONFIG.Manual.bookingsTabName);
  if (!sh || sh.getLastRow() < 2) return bookings;

  var data = sh.getDataRange().getValues();
  var h = indexHeaders(data[0]);
  function val(row, name) { var idx = h[name]; return (idx === undefined) ? '' : row[idx]; }
  function num(row, name) { var n = parseFloat(val(row, name)); return isNaN(n) ? 0 : n; }
  function isFilled(v) { return v !== '' && v !== null && v !== undefined; }

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var ref = (val(row, 'Booking Ref') || '').toString().trim();
    if (!ref) continue;
    if (bookings[ref]) continue; // dedupe accidental duplicate rows

    var campsite = val(row, 'Campsite Name');
    var platform = (val(row, 'Platform') || '').toString().trim() || 'Manual';
    var customerPaid = num(row, 'Customer Paid');
    var payoutRaw = val(row, 'Payout');
    var hasPayout = isFilled(payoutRaw) && isFilled(val(row, 'Payout Date'));
    var payout = hasPayout ? (parseFloat(payoutRaw) || 0) : '';

    var checkin = val(row, 'Travel Start Date');
    var checkout = val(row, 'Travel End Date');
    var nights = (checkin instanceof Date && checkout instanceof Date)
      ? Math.round((checkout - checkin) / 86400000) : '';

    bookings[ref] = {
      platform: platform,
      campsite: campsite,
      accommodation: normalizeAccommodation(campsite, val(row, 'Accommodation'), unmatchedAccommodations),
      camperName: val(row, 'Camper Name'),
      bookingDate: val(row, 'Booking Date'),
      checkin: checkin,
      checkout: checkout,
      nights: nights,
      grossCharge: customerPaid,
      commission: hasPayout ? (customerPaid - payout) : '',
      commissionPct: (hasPayout && customerPaid) ? ((customerPaid - payout) / customerPaid * 100) : '',
      hostPayoutExpected: payout,
      settledAmount: payout,
      settlementDate: hasPayout ? val(row, 'Payout Date') : '',
      hasPayout: hasPayout
    };
  }
  return bookings;
}

// RECONCILE + BUILD "OTA Bookings" OUTPUT TAB

function processOtaData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupOtaTabs(); // idempotent — creates anything missing, clears+rebuilds only the output tab

  var MISMATCH_TOLERANCE = 1; // rupees — guards against float rounding noise
  var outRows = [];
  var outPlatforms = []; // parallel array to outRows, used for MMT/ClearTrip counts in the alert
  var outBrands = [];    // parallel array to outRows, used for Confirmation Code cell color-coding (GH/NH/ClearTrip)
  var outStatuses = [];  // parallel array to outRows, used for Qualified Status cell color-coding
  var sl = 1;
  var qualifiedCount = 0, mismatchCount = 0, pendingCount = 0;
  var unmatchedAccommodations = [];

  // MMT
  var mmtBookingsByParent = parseMmtBookings(ss, unmatchedAccommodations);
  var mmtPaymentsByBookingId = parseMmtNetEarnings(ss);
  var mmtParentIds = Object.keys(mmtBookingsByParent).sort();

  mmtParentIds.forEach(function(pid) {
    var b = mmtBookingsByParent[pid];
    var pay = mmtPaymentsByBookingId[pid];

    var settledAmount = '', settlementDate = '', settlementRef = '', amountMatch = '', qualified;

    if (pay) {
      settledAmount = pay.amountPaid;
      settlementDate = pay.settlementDate || pay.bookingDate;
      settlementRef = pay.paymentRef;
      var diff = Math.abs(pay.amountPaid - b.hostPayoutExpected);
      if (diff <= MISMATCH_TOLERANCE) {
        amountMatch = 'Match'; qualified = 'Qualified'; qualifiedCount++;
      } else {
        amountMatch = 'Mismatch (₹' + diff.toFixed(0) + ')';
        qualified = 'Qualified - Mismatch'; qualifiedCount++; mismatchCount++;
      }
    } else {
      qualified = 'Pending Payment'; pendingCount++;
    }

    // Property Name from the settlement record is authoritative when present;
    // otherwise fall back to the property this booking tab was configured for.
    var campsite = (pay && pay.propertyName) ? pay.propertyName : b.property;

    outRows.push([
      sl++, pid, b.platform,
      b.bookedOn, b.bookedOn,
      b.checkin, b.checkin,
      b.checkout, b.checkout,
      b.nights, b.totalUnits,
      campsite, b.roomNames.join(' + '), b.customerName,
      b.bookingStatus, b.paymentStatus,
      b.hotelGrossCharges, b.commission, b.commissionPct, b.gstOnCommission,
      b.tcsAmount, b.tdsAmount, b.hostPayoutExpected,
      settledAmount, settlementDate, settlementRef,
      amountMatch, qualified
    ]);
    outPlatforms.push(b.platform);
    outBrands.push(classifyMmtBrand(pid));
    outStatuses.push(qualified);
  });

  // ClearTrip
  var ctParsed = parseClearTripBookings(ss, unmatchedAccommodations);
  var ctBookings = ctParsed.bookings;
  var ctSettlements = parseClearTripSettlements(ss); // trip_id -> {netPayable, utrNo, paymentDate, paymentStatus}
  var ctTripIds = Object.keys(ctBookings).sort();

  ctTripIds.forEach(function(tid) {
    var b = ctBookings[tid];
    var settledAmount = '', settlementDate = '', settlementRef = '', amountMatch = '', qualified;

    var settlement = ctSettlements[tid];
    var isSettled = settlement && (settlement.paymentStatus || '').toString().toUpperCase() === 'SUCCESS';
    if (isSettled) {
      settledAmount = settlement.netPayable;
      settlementDate = settlement.paymentDate;
      settlementRef = settlement.utrNo;
      var diff = Math.abs(settledAmount - b.hostPayoutExpected);
      if (diff <= MISMATCH_TOLERANCE) {
        amountMatch = 'Match'; qualified = 'Qualified'; qualifiedCount++;
      } else {
        amountMatch = 'Mismatch (₹' + diff.toFixed(0) + ')';
        qualified = 'Qualified - Mismatch'; qualifiedCount++; mismatchCount++;
      }
    } else {
      qualified = 'Pending Payment'; pendingCount++;
    }

    outRows.push([
      sl++, tid, b.platform,
      b.bookingDate, b.bookingDate,
      b.checkin, b.checkin,
      b.checkout, b.checkout,
      b.nights, b.totalUnits,
      b.campsite, b.accommodation, b.camperName,
      b.bookingStatus, b.platformPaymentStatus,
      b.grossCharge, b.commission, b.commissionPct, b.gstOnCommission,
      b.tcsAmount, b.tdsAmount, b.hostPayoutExpected,
      settledAmount, settlementDate, settlementRef,
      amountMatch, qualified
    ]);
    outPlatforms.push(b.platform);
    outBrands.push('ClearTrip');
    outStatuses.push(qualified);
  });

  // Manual bank-transfer bookings (Airbnb, Stories Collective, ALIVE Booking, etc.)
  var manualBookings = parseManualBookings(ss, unmatchedAccommodations);
  var manualRefs = Object.keys(manualBookings).sort();

  manualRefs.forEach(function(ref) {
    var b = manualBookings[ref];
    var qualified = b.hasPayout ? 'Qualified' : 'Pending Payment';
    if (b.hasPayout) { qualifiedCount++; } else { pendingCount++; }

    outRows.push([
      sl++, ref, b.platform,
      b.bookingDate, b.bookingDate,
      b.checkin, b.checkin,
      b.checkout, b.checkout,
      b.nights, 1,
      b.campsite, b.accommodation, b.camperName,
      'Confirmed', '',
      b.grossCharge, b.commission, b.commissionPct, 0,
      0, 0,
      b.hostPayoutExpected,
      b.settledAmount, b.settlementDate, '',
      b.hasPayout ? 'Match' : '',
      qualified
    ]);
    outPlatforms.push(b.platform);
    outBrands.push(b.platform);
    outStatuses.push(qualified);
  });

  if (outRows.length === 0) {
    safeAlert('No OTA booking rows found. Paste data into the Raw MMT / Raw ClearTrip tabs first.');
    return;
  }

  var sh = ss.getSheetByName(OTA_BOOKINGS_TAB);
  sh.getRange(2, 1, outRows.length, OTA_BOOKINGS_HEADERS.length).setValues(outRows);

  formatOtaBookingsSheet(sh, outRows.length);
  colorCodeByBrand(sh, outBrands);
  colorCodeByStatus(sh, outStatuses);
  sh.autoResizeColumns(1, OTA_BOOKINGS_HEADERS.length);

  var msg = 'OTA data processed.\n\n' +
    'Total bookings: ' + outRows.length + ' (MMT: ' + mmtParentIds.length + ', ClearTrip: ' + ctTripIds.length +
    ', Manual: ' + manualRefs.length + ')\n' +
    'Qualified: ' + qualifiedCount + ' (of which ' + mismatchCount + ' flagged as amount mismatches — still populated, just for review)\n' +
    'Pending Payment (no settlement matched yet): ' + pendingCount;

  if (unmatchedAccommodations.length > 0) {
    var uniqueUnmatched = unmatchedAccommodations.filter(function(v, i, a) { return a.indexOf(v) === i; });
    msg += '\n\n⚠️ ' + uniqueUnmatched.length + ' accommodation name(s) didn\'t match a known alias ' +
      '(shown as raw OTA text — won\'t match CL Payments\' Master GST Data until mapped):\n' +
      uniqueUnmatched.slice(0, 10).join('\n');
  }

  safeAlert(msg);
}

// PLATFORM COLOR CODING

// GoMMT's own Confirmation Codes distinguish the underlying brand by prefix —
// "GH" = Goibibo, "NH" = MakeMyTrip (Goibibo and MakeMyTrip are the same
// GoMMT seller account, but visually distinct enough to be worth splitting).
// ClearTrip's trip_id has no such prefix, so it just gets its own color.
function classifyMmtBrand(confirmationCode) {
  var code = (confirmationCode || '').toString().trim();
  if (code.indexOf('GH') === 0) return 'Goibibo';
  if (code.indexOf('NH') === 0) return 'MakeMyTrip';
  return 'GoMMT'; // fallback for any other prefix
}

// Kept identical to CL Payments' OTA_BRAND_ROW_COLORS (same brand = same
// color everywhere). GoMMT/Stories Collective/Manual were nudged away from
// hues CL Payments already uses for its own status flags there.
var BRAND_ROW_COLORS = {
  Goibibo: '#F9CB9C',           // orange
  MakeMyTrip: '#A2C4C9',        // teal
  GoMMT: '#F0DDB8',             // tan, fallback for Goibibo/MakeMyTrip
  ClearTrip: '#D0E6F5',         // light blue
  Airbnb: '#F5C6CB',            // pink
  'Stories Collective': '#A8DDD2', // teal-green
  'ALIVE Booking': '#FCE8A6',   // yellow
  Manual: '#D6D6E8'             // lavender-grey fallback for any other manually-typed platform
};

// Only the Confirmation Code cell gets the brand tint (full-row color was
// too flashy). GH/NH/ClearTrip already sort into contiguous blocks, so this
// colors runs in one call each rather than one setBackground per row.
function colorCodeByBrand(sh, brands) {
  var i = 0;
  while (i < brands.length) {
    var color = BRAND_ROW_COLORS[brands[i]];
    var runStart = i;
    while (i < brands.length && brands[i] === brands[runStart]) i++;
    var runLength = i - runStart;
    if (color) {
      sh.getRange(runStart + 2, OTA_COL.CONF_CODE, runLength, 1).setBackground(color);
    }
  }
}

// Payment-processed vs pending is the more important signal at a glance, so
// it gets a strong, distinct color on the Qualified Status cell itself
// (Amount Match gets the same treatment) rather than tinting the whole row —
// that would fight with the platform tint above. Green = settled, amber =
// still waiting, orange = settled but flagged for review.
var STATUS_CELL_COLORS = {
  'Qualified': '#B7E1CD',
  'Qualified - Mismatch': '#F9CB9C',
  'Pending Payment': '#FFF2CC'
};

function colorCodeByStatus(sh, statuses) {
  var n = statuses.length;
  var statusColors = [];
  var matchColors = [];
  for (var i = 0; i < n; i++) {
    var isMismatch = statuses[i].indexOf('Mismatch') !== -1;
    var color = STATUS_CELL_COLORS[statuses[i]] || null;
    statusColors.push([color]);
    matchColors.push([isMismatch ? STATUS_CELL_COLORS['Qualified - Mismatch'] : null]);
  }
  sh.getRange(2, OTA_COL.QUALIFIED_STATUS, n, 1).setBackgrounds(statusColors);
  sh.getRange(2, OTA_COL.AMOUNT_MATCH, n, 1).setBackgrounds(matchColors);
  sh.getRange(2, OTA_COL.QUALIFIED_STATUS, n, 1).setFontWeight('bold');
}

function formatOtaBookingsSheet(sh, numDataRows) {
  if (numDataRows < 1) return;

  var currencyCols = [
    OTA_COL.TOTAL, OTA_COL.COMMISSION, OTA_COL.GST_ON_COMMISSION,
    OTA_COL.TCS, OTA_COL.TDS, OTA_COL.HOST_PAYOUT, OTA_COL.SETTLED_AMOUNT
  ];
  var pctCols = [OTA_COL.COMMISSION_PCT];
  var dateCols = [OTA_COL.BOOKING_DATE, OTA_COL.TRAVEL_START, OTA_COL.TRAVEL_END, OTA_COL.SETTLEMENT_DATE];
  var monthCols = [OTA_COL.BOOKING_MTH, OTA_COL.TRAVEL_MTH, OTA_COL.TRAVEL_END_MTH];

  currencyCols.forEach(function(col) {
    sh.getRange(2, col, numDataRows, 1).setNumberFormat('₹#,##0');
  });
  pctCols.forEach(function(col) {
    sh.getRange(2, col, numDataRows, 1).setNumberFormat('0.0"%"');
  });
  dateCols.forEach(function(col) {
    sh.getRange(2, col, numDataRows, 1).setNumberFormat('dd-mmm-yy');
  });
  monthCols.forEach(function(col) {
    sh.getRange(2, col, numDataRows, 1).setNumberFormat('mmm-yy');
  });
}
