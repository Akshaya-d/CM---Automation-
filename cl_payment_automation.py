#!/usr/bin/env python3
"""
Camp Monk CL Payment Automation
================================
Populates a new CL Payment Sheet from:
  1. Admin booking report  (CSV or XLSX)
  2. Razorpay payments export (XLSX)

Usage:
  python cl_payment_automation.py \
    --admin   "properties-booking[2026-06-12 - 2026-06-13].csv" \
    --razorpay "payments - 12 Jun 26 - 13 Jun 26.xlsx" \
    --template "Camp_Monk_CL_Payment_Automation_Test.xlsx" \
    --output  "CL_Payment_Output.xlsx"

All four arguments are required.  The template is used for its lookup
tables (GST PAN CM COMM sheet) and formatting; Bookings + Razorpay
sheets are fully replaced with fresh data + live Excel formulas.
"""

import argparse
import csv
import os
import sys

try:
    import openpyxl
    from openpyxl.utils import get_column_letter, column_index_from_string
except ImportError:
    sys.exit("openpyxl not found.  Run:  pip install openpyxl --break-system-packages")


# ─────────────────────────────────────────────────────────────────────────────
# Column mapping: Admin report column name → Bookings sheet column name
# Edit the LEFT side if your admin export uses different header names.
# ─────────────────────────────────────────────────────────────────────────────
ADMIN_TO_BOOKINGS = {
    "Booking Date":           "Booking Date",
    "Booking Month":          "Booking Mth",
    "Travel Start Date":      "Travel Start Date",
    "Travel Start Mth":       "Travel Mth",
    "Travel End Date":        "Travel End Date",
    "Travel End Mth":         "Travel End Mth",
    "Booking ID":             "Booking ID",
    "Property":               "Campsite Name",
    "Original property Name": "Original campsite Name",
    "Accomodation":           "Accomodation",
    "Camp Owner Name":        "Camp Owner Name",
    "Camper":                 "Camper Name",
    "Camper phone":           "Camper phone",
    "Camper email":           "Camper email",
    "Amount":                 "TOTAL",
    "Promocode":              "Promocode Name",
    "Promocode discount":     "Promocode discount",
    "Flat discount":          "Flat discount",
    "Referral discount":      "Referral discount",
    "Weakdays discount":      "Weekdays discount",
    "Wallet amount used":     "Wallet amount used",
    "GST":                    "GST",
    "Total amount":           "GRAND TOTAL",
    "Total units":            "Total units",
    "Insta units":            "Insta units",
    "Total adults":           "Total Adults",
    "Total kids":             "Total Kids",
    "Extra adults":           "Extra Adults",
    "Extra kids":             "Extra Kids",
    "Total guests":           "Total guests",
    "Total pets":             "Total Pets",
    "No of nights":           "No of Nights",
    "Status":                 "Status",
}


# ─────────────────────────────────────────────────────────────────────────────
# Razorpay raw column names in the order they appear in the export.
# These map to columns F onward in the Razorpay sheet.
# Edit if your Razorpay export has different column names or order.
# ─────────────────────────────────────────────────────────────────────────────
RAZORPAY_SHEET_HEADERS = [
    # ── Derived columns (A–E) ── written as formulas, listed here for reference
    "Booking ID",           # A  =RIGHT(U{r},6)
    "Vlookup",              # B  =VLOOKUP(A{r},Bookings!H:H,1,FALSE)
    "Razorpay fee",         # C  =SUM(AD{r}+AE{r})
    "Paid to CM",           # D  =G{r}-C{r}
    "Comm %",               # E  =C{r}/G{r}
    # ── Raw Razorpay export columns (F onward) ──
    "id",                   # F
    "amount",               # G
    "Refund ID",            # H
    "Refunded\nAmount",     # I
    "Refund/Cancellation\nDate",  # J
    "currency",             # K
    "status",               # L
    "order_id",             # M
    "invoice_id",           # N
    "international",        # O
    "method",               # P
    "amount_refunded",      # Q
    "amount_transferred",   # R
    "refund_status",        # S
    "captured",             # T
    "description",          # U  ← Booking ID is extracted from here
    "card_id",              # V
    "card",                 # W
    "bank",                 # X
    "wallet",               # Y
    "vpa",                  # Z
    "email",                # AA
    "contact",              # AB
    "notes",                # AC
    "fee",                  # AD  ← used in Razorpay fee formula
    "tax",                  # AE  ← used in Razorpay fee formula
    "error_code",           # AF
    "error_description",    # AG
    "created_at",           # AH
    "card_type",            # AI
    "card_network",         # AJ
]

# Column index (1-based) where raw Razorpay data starts in the Razorpay sheet
RAZORPAY_RAW_START_COL = 6   # F


# ─────────────────────────────────────────────────────────────────────────────
# Formula builders
# ─────────────────────────────────────────────────────────────────────────────

def bookings_formulas(r: int) -> dict:
    """
    Return {column_letter: formula} for all formula-driven columns in the
    Bookings sheet at Excel row r (r=2 is the first data row).
    """
    return {
        # Sr No
        "A":  1 if r == 2 else f"=+A{r-1}+1",

        # ── Cross-reference: Razorpay sheet ──────────────────────────────
        # Col AI (35): Status Match — is this Booking ID found in Razorpay?
        "AI": f'=VLOOKUP(H{r},Razorpay!A:A,1,FALSE)',

        # Col AJ (36): Do TOTAL + GST equal GRAND TOTAL?
        "AJ": f'=(P{r}+W{r})=X{r}',

        # ── Cross-reference: GST PAN CM COMM lookup table ────────────────
        # Col AK (37): Is camp GST-registered?
        "AK": f"=VLOOKUP(I{r},'GST PAN CM COMM'!A:B,2,FALSE)",

        # Col AL (38): Camp PAN number
        "AL": f"=VLOOKUP(I{r},'GST PAN CM COMM'!H:I,2,FALSE)",

        # Col AM (39): Instamojo/Razorpay charge
        "AM": f'=VLOOKUP(H{r},Razorpay!A:E,3,FALSE)',

        # Col AN (40): Transfer to CM (net of gateway fee)
        "AN": f'=VLOOKUP(H{r},Razorpay!A:E,4,FALSE)',

        # Col AO (41): Does CampMonk have GST for this camp?
        "AO": f'=IFERROR(VLOOKUP(I{r},\'GST PAN CM COMM\'!A:C,3,0),"No")',

        # Col AP (42): CM commission %
        "AP": f"=VLOOKUP(I{r},'GST PAN CM COMM'!S:T,2,FALSE)",

        # ── Commission calculations ───────────────────────────────────────
        # Col AQ (43): CM Commission amount (excl. GST)
        "AQ": f'=P{r}*AP{r}',

        # Col AR (44): CM GST on commission
        "AR": f'=IF(AO{r}="Yes",(AQ{r}*18%),0)+IF(AO{r}="No",W{r})',

        # Col AS (45): CL Transfer before TDS
        "AS": f'=X{r}-AQ{r}-AR{r}',

        # Col AT (46): TDS deducted (2% if CL Has GST = "no")
        "AT": f'=IF(AO{r}="no",(AS{r}*2%))',

        # Col AU (47): CL Transfer after TDS
        "AU": f'=AS{r}-AT{r}',

        # ── Reconciliation ────────────────────────────────────────────────
        # Col AY (51): Match Total (Insta Charge + Insta Transfer)
        "AY": f'=AM{r}+AN{r}',

        # Col AZ (52): Does GRAND TOTAL match Match Total?
        "AZ": f'=X{r}=AY{r}',

        # ── Auto-status columns ───────────────────────────────────────────
        # Col BB (54): Payment found in Razorpay?
        "BB": f'=IF(COUNTIF(Razorpay!A:A,H{r})>0,"Found","Missing")',

        # Col BC (55): Amount match (within ₹2 tolerance)?
        "BC": f'=IF(ABS((P{r}+W{r})-X{r})<=2,"MATCH","CHECK")',

        # Col BD (56): Wallet flag
        "BD": f'=IF(V{r}>0,"Wallet Used","")',

        # Col BE (57): Final Status (combined)
        "BE": (
            f'=IF(BB{r}="Missing","Payment Missing",'
            f'IF(V{r}>0,"Wallet Used - Verify",'
            f'IF(ABS(X{r}-AY{r})<=2,"MATCH","CHECK")))'
        ),
    }


def razorpay_formulas(r: int) -> dict:
    """
    Return {column_letter: formula} for the 5 derived columns in the
    Razorpay sheet at Excel row r.

    Column U (description) contains text like "Accomodation booking - CM28158".
    RIGHT(U{r},6) extracts the last 6 characters.
    ⚠ If Booking IDs grow longer (e.g. CM28158 = 7 chars), change 6 → 7 here.
    """
    return {
        "A": f'=IFERROR(MID(U{r},FIND("CM",U{r}),10),RIGHT(U{r},6))',  # Booking ID — handles variable length
        "B": f'=VLOOKUP(A{r},Bookings!H:H,1,FALSE)',                    # Lookup in Bookings
        "C": f'=SUM(AD{r}+AE{r})',                                      # Total gateway fee
        "D": f'=G{r}-C{r}',                                             # Net paid to CM
        "E": f'=C{r}/G{r}',                                             # Fee %
    }


# ─────────────────────────────────────────────────────────────────────────────
# Readers
# ─────────────────────────────────────────────────────────────────────────────

def read_admin(path: str) -> list[dict]:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        with open(path, encoding="utf-8-sig") as f:
            return list(csv.DictReader(f))
    else:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        headers = [c.value for c in ws[1]]
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if any(v is not None for v in row):
                rows.append(dict(zip(headers, row)))
        return rows


def read_razorpay(path: str) -> tuple[list[str], list[dict]]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if any(v is not None for v in row):
            rows.append(dict(zip(headers, row)))
    return headers, rows


# ─────────────────────────────────────────────────────────────────────────────
# Sheet writers
# ─────────────────────────────────────────────────────────────────────────────

def write_bookings(ws, admin_rows: list[dict]):
    """Clear existing data rows and repopulate with admin data + formulas."""
    # Build column-name → letter map from existing header row
    header_to_col = {}
    for cell in ws[1]:
        if cell.value:
            header_to_col[cell.value] = get_column_letter(cell.column)

    # Unmerge any merged cells, then clear all data rows (keep row 1 = headers)
    for merge in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merge))
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.value = None

    skipped = 0
    for i, arow in enumerate(admin_rows):
        r = i + 2  # Excel row number (row 1 = header)
        formulas = bookings_formulas(r)

        # Sr No
        ws[f'A{r}'] = formulas["A"]

        # Data columns from admin report
        for admin_col, book_col in ADMIN_TO_BOOKINGS.items():
            val = arow.get(admin_col)
            if val is None:
                continue
            if book_col not in header_to_col:
                skipped += 1
                continue
            col_ltr = header_to_col[book_col]
            # Convert numeric strings
            if isinstance(val, str):
                val = val.strip()
                try:
                    val = int(val) if val.isdigit() else float(val)
                except ValueError:
                    pass
            ws[f'{col_ltr}{r}'] = val

        # Formula columns (skip A — already written)
        for col_ltr, formula in formulas.items():
            if col_ltr == "A":
                continue
            ws[f'{col_ltr}{r}'] = formula

    if skipped:
        print(f"  ⚠  {skipped} admin column(s) not found in Bookings header — check ADMIN_TO_BOOKINGS mapping.")


def write_razorpay(ws, rp_export_headers: list[str], rp_rows: list[dict]):
    """
    Clear existing data rows and repopulate with Razorpay data + formulas.

    The first 5 columns (A–E) are formula-derived.
    Raw Razorpay data is written starting at column F, using the actual
    export column names as the sheet header (row 1, cols F onward).
    """
    # Unmerge any merged cells, then clear all data rows (keep row 1 = headers)
    for merge in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merge))
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.value = None

    # Rewrite header row for raw columns (F onward) to match current export
    for j, hdr in enumerate(rp_export_headers):
        ws.cell(row=1, column=RAZORPAY_RAW_START_COL + j, value=hdr)

    for i, rrow in enumerate(rp_rows):
        r = i + 2
        # Write formula columns A–E
        for col_ltr, formula in razorpay_formulas(r).items():
            ws[f'{col_ltr}{r}'] = formula
        # Write raw data starting at column F
        for j, hdr in enumerate(rp_export_headers):
            ws.cell(row=r, column=RAZORPAY_RAW_START_COL + j, value=rrow.get(hdr))


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Camp Monk CL Payment Sheet Automation"
    )
    parser.add_argument("--admin",     required=True, help="Admin booking report (CSV or XLSX)")
    parser.add_argument("--razorpay",  required=True, help="Razorpay payments export (XLSX)")
    parser.add_argument("--template",  required=True, help="Existing CL Payment Sheet (XLSX) — used as template")
    parser.add_argument("--output",    required=True, help="Output XLSX file path")
    args = parser.parse_args()

    for path in [args.admin, args.razorpay, args.template]:
        if not os.path.exists(path):
            sys.exit(f"File not found: {path}")

    # ── Load template, keep only Bookings + Razorpay sheets ─────────────
    print(f"Loading template: {os.path.basename(args.template)}")
    wb = openpyxl.load_workbook(args.template)

    KEEP_SHEETS = {"Bookings", "Razorpay", "GST PAN CM COMM"}
    for name in [s for s in wb.sheetnames if s not in KEEP_SHEETS]:
        del wb[name]

    # ── Bookings sheet ────────────────────────────────────────────────────
    print(f"Reading admin report: {os.path.basename(args.admin)}")
    admin_rows = read_admin(args.admin)
    print(f"  → {len(admin_rows)} booking rows")

    ws_bookings = wb["Bookings"]
    write_bookings(ws_bookings, admin_rows)
    print(f"  ✓ Bookings sheet written")

    # ── Razorpay sheet ────────────────────────────────────────────────────
    print(f"Reading Razorpay export: {os.path.basename(args.razorpay)}")
    rp_headers, rp_rows = read_razorpay(args.razorpay)
    print(f"  → {len(rp_rows)} payment rows")

    ws_razorpay = wb["Razorpay"]
    write_razorpay(ws_razorpay, rp_headers, rp_rows)
    print(f"  ✓ Razorpay sheet written")

    # ── Save output ────────────────────────────────────────────────────────
    wb.save(args.output)
    print(f"\n✅ Done!  Output saved to: {args.output}")
    print(
        "\nOpen the file in Excel and press Ctrl+Alt+F9 (Windows) or "
        "Cmd+Option+F9 (Mac) to force-recalculate all formulas."
    )


if __name__ == "__main__":
    main()
