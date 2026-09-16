# Camp Monk — Payment & Sales Automation

A suite of Google Apps Script tools (plus a standalone Python fallback) that automate booking-payment reconciliation and sales reporting for [Camp Monk](https://www.campmonk.com), a glamping/campsite booking platform. Built to replace manual spreadsheet reconciliation across direct bookings, Razorpay payments, and OTA (MakeMyTrip/Goibibo) settlements.

> **Note:** This is a portfolio showcase of internal automation work. Company-specific identifiers (spreadsheet IDs, sheet URLs) have been removed and are configured via [Script Properties](https://developers.google.com/apps-script/guides/properties) instead of being hardcoded. Sample data files are not included.

## What it does

| Script | Purpose |
|---|---|
| [`CL_Payment_Setup.gs`](CL_Payment_Setup.gs) | Core reconciliation engine. Merges raw property + experience booking data into one payments sheet, matches each booking against Razorpay payments, calculates commission/GST/TDS and the net payout, flags mismatches, and exports lead lists (InstaPay, Experience, Event, Test bookings) for follow-up. |
| [`OTA_Payment_Setup.gs`](OTA_Payment_Setup.gs) | Reconciles OTA bookings (MakeMyTrip, Goibibo, etc.) against their settlement reports, which arrive separately and often months later. Tracks a booking as "qualified" only once it's matched to an actual settlement. |
| [`Sales_Pivots_Setup.gs`](Sales_Pivots_Setup.gs) | Builds 11 live, self-updating pivot-table dashboards (revenue trends, occupancy mix, cancellations & refunds, promo effectiveness, booking lead time, and more) directly on top of the reconciled bookings data. |
| [`cl_payment_automation.py`](cl_payment_automation.py) | Standalone CLI that rebuilds the CL Payment workbook from a raw admin booking export and a Razorpay export — useful for one-off runs outside of Apps Script. |

All three `.gs` scripts are designed to be idempotent: re-running them rebuilds/updates in place rather than duplicating data, so they're safe to trigger repeatedly (manually or on a schedule).

## Tech stack

- **Google Apps Script** (V8 runtime) — Sheets API, `PropertiesService`, time-driven triggers
- **Python 3** + [`openpyxl`](https://openpyxl.readthedocs.io/) for the offline CLI variant

## Setup

### Apps Script tools

1. Open the target Google Sheet → **Extensions → Apps Script**.
2. Paste in the relevant `.gs` file and save.
3. Run the corresponding setup function once (e.g. `setupCLPayments()`) to grant permissions — this adds a custom menu to the sheet.
4. Under **Project Settings → Script Properties**, add the IDs the script needs at runtime, e.g.:
   - `MAIN_SPREADSHEET_ID` — the ID of the main CL Payments spreadsheet (required by `CL_Payment_Setup.gs` when running from a trigger or a different context).
   - `INSTAPAY_SHEET_ID`, `OTA_SHEET_ID`, `SALES_SHEET_ID` — created automatically the first time each export runs, or set manually to point at existing sheets.
5. Day-to-day use is via the custom menu each script adds (e.g. **CL Automation → Update Now**).

### Python CLI

```bash
pip install -r requirements.txt

python cl_payment_automation.py \
  --admin    "admin-booking-report.csv" \
  --razorpay "razorpay-export.xlsx" \
  --template "CL_Payment_Template.xlsx" \
  --output   "CL_Payment_Output.xlsx"
```

## Repo layout

```
CL_Payment_Setup.gs       # Main payment reconciliation + lead export engine
OTA_Payment_Setup.gs      # OTA booking ↔ settlement reconciliation
Sales_Pivots_Setup.gs     # Sales dashboard / pivot builder
cl_payment_automation.py  # Offline Python equivalent of the CL Payment rebuild
requirements.txt          # Python dependencies
```
