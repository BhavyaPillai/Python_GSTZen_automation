import pandas as pd
from pathlib import Path
import requests

# ============================================================
# GST INVOICE AUTOMATION - STEP 2
# Test One Invoice PDF URL
# ============================================================

print("=" * 60)
print("GST PDF DOWNLOAD TEST")
print("=" * 60)

# Project directory
BASE_DIR = Path(__file__).resolve().parent.parent

# Prepared invoice file
INPUT_FILE = (
    BASE_DIR
    / "output"
    / "FY_2025_26_invoice_list.xlsx"
)

# Test download location
TEST_FOLDER = BASE_DIR / "output" / "PDF_TEST"

TEST_FOLDER.mkdir(
    parents=True,
    exist_ok=True
)

# ------------------------------------------------------------
# READ PREPARED FILE
# ------------------------------------------------------------

print("\nReading invoice list...")

df = pd.read_excel(INPUT_FILE)

print(f"Total FY invoices: {len(df):,}")

# ------------------------------------------------------------
# FIND FIRST VALID PDF URL
# ------------------------------------------------------------

valid_urls = df[
    df["Invoice PDF URL"].notna()
    & (df["Invoice PDF URL"].astype(str).str.strip() != "")
].copy()

if valid_urls.empty:
    print("\nERROR: No PDF URL found.")
    exit()

# Take first invoice
invoice = valid_urls.iloc[0]

invoice_number = str(
    invoice["Invoice Number"]
)

pdf_url = str(
    invoice["Invoice PDF URL"]
)

print("\nTest Invoice:")
print("Invoice Number:", invoice_number)

print("\nPDF URL:")
print(pdf_url)

# ------------------------------------------------------------
# DOWNLOAD TEST
# ------------------------------------------------------------

print("\nTrying to download PDF...")

try:

    response = requests.get(
        pdf_url,
        timeout=30,
        allow_redirects=True
    )

    print("\nHTTP Status Code:", response.status_code)

    print(
        "Content-Type:",
        response.headers.get("Content-Type")
    )

    print(
        "Downloaded Size:",
        len(response.content),
        "bytes"
    )

    # --------------------------------------------------------
    # SAVE RESPONSE
    # --------------------------------------------------------

    output_file = (
        TEST_FOLDER
        / f"{invoice_number}_TEST.pdf"
    )

    with open(
        output_file,
        "wb"
    ) as f:

        f.write(response.content)

    print("\nFile saved:")
    print(output_file)

    # --------------------------------------------------------
    # CHECK WHETHER IT IS ACTUALLY A PDF
    # --------------------------------------------------------

    if response.content.startswith(b"%PDF"):

        print("\nSUCCESS!")
        print("The URL returned a valid PDF.")

    else:

        print("\nWARNING!")
        print(
            "The response is NOT a normal PDF."
        )

        print(
            "\nFirst 100 characters:"
        )

        print(
            response.text[:100]
            if response.text
            else "No text available"
        )

except Exception as e:

    print("\nDOWNLOAD ERROR:")
    print(type(e).__name__)
    print(e)

print("\n")
print("=" * 60)
print("TEST COMPLETED")
print("=" * 60)