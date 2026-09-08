import pandas as pd
from pathlib import Path

# ============================================================
# GST INVOICE AUTOMATION - STEP 1
# Prepare FY 2025-26 Invoice List
# ============================================================

print("=" * 60)
print("GST INVOICE AUTOMATION")
print("=" * 60)

# ------------------------------------------------------------
# 1. DEFINE PROJECT FOLDERS
# ------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

# ------------------------------------------------------------
# 2. FIND GSTZEN EXCEL FILE
# ------------------------------------------------------------

excel_files = list(INPUT_DIR.glob("*.xlsx"))

if not excel_files:
    print("\nERROR: No Excel file found in the input folder.")
    print("Please put the GSTZen Excel file inside:")
    print(INPUT_DIR)
    exit()

# Use the first Excel file found
INPUT_FILE = excel_files[0]

print("\nInput file:")
print(INPUT_FILE)

# ------------------------------------------------------------
# 3. READ GSTZEN EXCEL
# ------------------------------------------------------------

print("\nReading GSTZen Excel...")

try:
    df = pd.read_excel(INPUT_FILE)
except Exception as e:
    print("\nERROR while reading Excel:")
    print(e)
    exit()

print(f"Total records found: {len(df):,}")

# ------------------------------------------------------------
# 4. CHECK REQUIRED COLUMNS
# ------------------------------------------------------------

required_columns = [
    "Invoice Date",
    "Invoice Number",
    "Invoice PDF URL"
]

missing_columns = [
    col for col in required_columns
    if col not in df.columns
]

if missing_columns:
    print("\nERROR: Required columns are missing:")
    for col in missing_columns:
        print("-", col)

    print("\nColumns available in your Excel:")
    for col in df.columns:
        print("-", col)

    exit()

# ------------------------------------------------------------
# 5. CONVERT INVOICE DATE
# ------------------------------------------------------------

print("\nConverting invoice dates...")

df["Invoice Date"] = pd.to_datetime(
    df["Invoice Date"],
    errors="coerce"
)

invalid_dates = df["Invoice Date"].isna().sum()

print(f"Invalid/missing invoice dates: {invalid_dates:,}")

# ------------------------------------------------------------
# 6. FILTER FINANCIAL YEAR 2025-26
# ------------------------------------------------------------

start_date = pd.Timestamp("2025-04-01")
end_date = pd.Timestamp("2026-03-31")

print("\nFiltering Financial Year:")
print("Start:", start_date.strftime("%d-%b-%Y"))
print("End  :", end_date.strftime("%d-%b-%Y"))

fy = df[
    (df["Invoice Date"] >= start_date) &
    (df["Invoice Date"] <= end_date)
].copy()

print(f"\nFY 2025-26 invoices: {len(fy):,}")

# ------------------------------------------------------------
# 7. CHECK PDF URL
# ------------------------------------------------------------

print("\nChecking Invoice PDF URLs...")

fy["Invoice PDF URL"] = fy["Invoice PDF URL"].astype("string")

pdf_available = (
    fy["Invoice PDF URL"].notna()
    & (fy["Invoice PDF URL"].str.strip() != "")
)

fy["PDF Available"] = pdf_available

available_count = pdf_available.sum()
missing_count = (~pdf_available).sum()

print(f"PDF URLs available: {available_count:,}")
print(f"PDF URLs missing  : {missing_count:,}")

# ------------------------------------------------------------
# 8. CHECK DUPLICATE INVOICE NUMBERS
# ------------------------------------------------------------

print("\nChecking duplicate invoice numbers...")

duplicate_mask = fy["Invoice Number"].duplicated(
    keep=False
)

duplicate_count = duplicate_mask.sum()

print(f"Duplicate invoice records: {duplicate_count:,}")

# ------------------------------------------------------------
# 9. ADD MONTH INFORMATION
# ------------------------------------------------------------

fy["Invoice Month"] = fy["Invoice Date"].dt.strftime("%b-%Y")

# ------------------------------------------------------------
# 10. SORT DATA
# ------------------------------------------------------------

fy = fy.sort_values(
    by=["Invoice Date", "Invoice Number"]
)

# ------------------------------------------------------------
# 11. CREATE OUTPUT FOLDER
# ------------------------------------------------------------

OUTPUT_DIR.mkdir(
    parents=True,
    exist_ok=True
)

OUTPUT_FILE = OUTPUT_DIR / "FY_2025_26_invoice_list.xlsx"

# ------------------------------------------------------------
# 12. SAVE PREPARED FILE
# ------------------------------------------------------------

print("\nSaving prepared invoice list...")

fy.to_excel(
    OUTPUT_FILE,
    index=False
)

# ------------------------------------------------------------
# 13. CREATE MISSING PDF REPORT
# ------------------------------------------------------------

missing_pdf = fy[
    ~fy["PDF Available"]
].copy()

if len(missing_pdf) > 0:

    missing_pdf_file = (
        OUTPUT_DIR /
        "Missing_PDF_URLs_FY_2025_26.xlsx"
    )

    missing_pdf.to_excel(
        missing_pdf_file,
        index=False
    )

    print(
        "\nMissing PDF report created:"
    )

    print(missing_pdf_file)

# ------------------------------------------------------------
# 14. FINAL SUMMARY
# ------------------------------------------------------------

print("\n")
print("=" * 60)
print("STEP 1 COMPLETED")
print("=" * 60)

print(f"Total GSTZen records       : {len(df):,}")
print(f"FY 2025-26 invoices        : {len(fy):,}")
print(f"PDF URLs available         : {available_count:,}")
print(f"PDF URLs missing           : {missing_count:,}")
print(f"Duplicate invoice records : {duplicate_count:,}")

print("\nPrepared file:")
print(OUTPUT_FILE)

print("\nNext step will be PDF download testing.")
print("DO NOT start bulk PDF downloading yet.")

print("=" * 60)