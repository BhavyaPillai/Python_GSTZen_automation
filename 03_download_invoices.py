import pandas as pd
import requests
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import time


# ============================================================
# GST INVOICE AUTOMATION
# FAST + RESUMABLE PDF DOWNLOADER
# ============================================================

print("=" * 75)
print("GST INVOICE AUTOMATION")
print("FAST + RESUMABLE PDF DOWNLOADER")
print("=" * 75)


# ============================================================
# 1. SETTINGS
# ============================================================

# ------------------------------------------------------------
# TEST MODE
# ------------------------------------------------------------
# True  = limited testing
# False = process ALL invoices
# ------------------------------------------------------------

TEST_MODE = False

# Number of invoices during testing
TEST_LIMIT = 100


# ------------------------------------------------------------
# PARALLEL DOWNLOADS
# ------------------------------------------------------------
# 5 is a safe starting point.
#
# If 100 invoices work with 0 or very few failures,
# this can later be increased to 8 or 10.
# ------------------------------------------------------------

MAX_WORKERS = 5


# ------------------------------------------------------------
# RETRIES
# ------------------------------------------------------------

MAX_RETRIES = 3


# ------------------------------------------------------------
# TIMEOUT
# ------------------------------------------------------------

REQUEST_TIMEOUT = 60


# ------------------------------------------------------------
# SAVE LOG AFTER THIS MANY COMPLETED DOWNLOADS
# ------------------------------------------------------------

SAVE_EVERY = 100


# ------------------------------------------------------------
# DELAY AFTER EACH DOWNLOAD
# ------------------------------------------------------------
# Parallel downloads already make the process fast.
# Keep a tiny delay to avoid hitting the server too aggressively.
# ------------------------------------------------------------

DOWNLOAD_DELAY = 0.1


# ============================================================
# 2. PROJECT PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent.parent


INPUT_FILE = (
    BASE_DIR
    / "output"
    / "FY_2025_26_invoice_list.xlsx"
)


OUTPUT_DIR = (
    BASE_DIR
    / "output"
    / "FY_2025_26"
)


LOG_DIR = (
    BASE_DIR
    / "logs"
)


LOG_FILE = (
    LOG_DIR
    / "Invoice_Download_Log.xlsx"
)


FAILED_FILE = (
    LOG_DIR
    / "Failed_Invoices.xlsx"
)


# Create folders if they don't exist

OUTPUT_DIR.mkdir(
    parents=True,
    exist_ok=True
)

LOG_DIR.mkdir(
    parents=True,
    exist_ok=True
)


# ============================================================
# 3. CHECK INPUT FILE
# ============================================================

if not INPUT_FILE.exists():

    print("\nERROR")
    print("-" * 75)

    print("Invoice list was not found:")

    print(INPUT_FILE)

    print("\nRun Step 1 first.")

    exit()


# ============================================================
# 4. LOAD INVOICE LIST
# ============================================================

print("\nReading invoice list...")

try:

    df = pd.read_excel(
        INPUT_FILE
    )

except Exception as e:

    print("\nERROR reading Excel:")
    print(e)

    exit()


print(
    f"Total records in invoice list: {len(df):,}"
)


# ============================================================
# 5. REQUIRED COLUMNS
# ============================================================

required_columns = [
    "Invoice Number",
    "Invoice Date",
    "Invoice PDF URL"
]


missing_columns = [
    column
    for column in required_columns
    if column not in df.columns
]


if missing_columns:

    print("\nERROR - Required columns missing:")

    for column in missing_columns:
        print(f"  - {column}")

    print("\nAvailable columns:")

    for column in df.columns:
        print(f"  - {column}")

    exit()


# ============================================================
# 6. CLEAN DATA
# ============================================================

df["Invoice Date"] = pd.to_datetime(
    df["Invoice Date"],
    errors="coerce"
)


df["Invoice Number"] = (
    df["Invoice Number"]
    .astype("string")
    .str.strip()
)


df["Invoice PDF URL"] = (
    df["Invoice PDF URL"]
    .astype("string")
    .str.strip()
)


# ============================================================
# 7. REMOVE INVALID RECORDS
# ============================================================

download_df = df[
    df["Invoice Number"].notna()
    &
    df["Invoice Date"].notna()
    &
    df["Invoice PDF URL"].notna()
    &
    (df["Invoice PDF URL"] != "")
].copy()


print(
    f"Records with valid PDF URLs: "
    f"{len(download_df):,}"
)


# ============================================================
# 8. REMOVE DUPLICATES
# ============================================================

before_duplicates = len(download_df)


download_df = (
    download_df
    .drop_duplicates(
        subset=[
            "Invoice Number",
            "Invoice PDF URL"
        ],
        keep="first"
    )
)


duplicates_removed = (
    before_duplicates
    - len(download_df)
)


print(
    f"Duplicate records removed: "
    f"{duplicates_removed:,}"
)


# ============================================================
# 9. TEST / PRODUCTION MODE
# ============================================================

if TEST_MODE:

    download_df = (
        download_df
        .head(TEST_LIMIT)
    )

    print("\n")
    print("=" * 75)
    print("TEST MODE")
    print("=" * 75)

    print(
        f"Invoices to process: "
        f"{len(download_df):,}"
    )

else:

    print("\n")
    print("=" * 75)
    print("PRODUCTION MODE")
    print("=" * 75)

    print(
        f"Invoices to process: "
        f"{len(download_df):,}"
    )


# ============================================================
# 10. LOAD EXISTING LOG
# ============================================================

if LOG_FILE.exists():

    try:

        existing_log = pd.read_excel(
            LOG_FILE
        )

        print(
            f"\nExisting log found: "
            f"{len(existing_log):,} records"
        )

    except Exception as e:

        print(
            "\nWarning: Existing log could not be read."
        )

        print(e)

        existing_log = pd.DataFrame()

else:

    existing_log = pd.DataFrame()

    print(
        "\nNo existing download log found."
    )


# ============================================================
# 11. HELPER - SAFE FILE NAME
# ============================================================

def safe_filename(value):

    value = str(value)

    invalid_characters = [
        "<",
        ">",
        ":",
        '"',
        "/",
        "\\",
        "|",
        "?",
        "*"
    ]

    for character in invalid_characters:

        value = value.replace(
            character,
            "_"
        )

    return value


# ============================================================
# 12. CHECK VALID PDF
# ============================================================

def is_valid_pdf(file_path):

    try:

        if not file_path.exists():

            return False

        # Very small files are unlikely to be valid invoices
        if file_path.stat().st_size < 1000:

            return False

        with open(
            file_path,
            "rb"
        ) as file:

            header = file.read(5)

        return header == b"%PDF-"

    except Exception:

        return False


# ============================================================
# 13. DOWNLOAD ONE INVOICE
# ============================================================

def download_invoice(
    invoice_number,
    invoice_date,
    pdf_url
):

    invoice_number = str(
        invoice_number
    )

    # --------------------------------------------------------
    # MONTH-WISE FOLDER
    # --------------------------------------------------------

    month_folder = (
        OUTPUT_DIR
        / invoice_date.strftime("%b-%Y")
    )


    month_folder.mkdir(
        parents=True,
        exist_ok=True
    )


    # --------------------------------------------------------
    # FILE NAME
    # --------------------------------------------------------

    clean_invoice_number = (
        safe_filename(
            invoice_number
        )
    )


    output_file = (
        month_folder
        / f"{clean_invoice_number}.pdf"
    )


    # --------------------------------------------------------
    # ALREADY DOWNLOADED?
    # --------------------------------------------------------

    if is_valid_pdf(
        output_file
    ):

        return {

            "Invoice Number":
                invoice_number,

            "Invoice Date":
                invoice_date,

            "Invoice PDF URL":
                pdf_url,

            "Status":
                "SKIPPED_EXISTING",

            "Attempts":
                0,

            "File Path":
                str(output_file),

            "Downloaded At":
                datetime.now(),

            "Error":
                ""
        }


    # --------------------------------------------------------
    # RETRIES
    # --------------------------------------------------------

    last_error = ""


    for attempt in range(
        1,
        MAX_RETRIES + 1
    ):

        temp_file = (
            output_file.with_suffix(
                ".pdf.part"
            )
        )


        try:

            # ------------------------------------------------
            # REQUEST
            # ------------------------------------------------

            response = requests.get(

                pdf_url,

                timeout=REQUEST_TIMEOUT,

                allow_redirects=True,

                headers={

                    "User-Agent":
                        "Mozilla/5.0 "
                        "(Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 "
                        "(KHTML, like Gecko) "
                        "Chrome/131.0 Safari/537.36",

                    "Accept":
                        "application/pdf,"
                        "application/octet-stream,"
                        "*/*"
                }

            )


            # ------------------------------------------------
            # HTTP STATUS
            # ------------------------------------------------

            if response.status_code != 200:

                last_error = (
                    f"HTTP "
                    f"{response.status_code}"
                )

                raise Exception(
                    last_error
                )


            # ------------------------------------------------
            # CHECK PDF CONTENT
            # ------------------------------------------------

            content = response.content


            if not content.startswith(
                b"%PDF-"
            ):

                last_error = (
                    "Response is not a valid PDF"
                )

                raise Exception(
                    last_error
                )


            # ------------------------------------------------
            # SAVE TEMP FILE FIRST
            # ------------------------------------------------

            with open(
                temp_file,
                "wb"
            ) as file:

                file.write(
                    content
                )


            # ------------------------------------------------
            # VERIFY TEMP FILE
            # ------------------------------------------------

            if not is_valid_pdf(
                temp_file
            ):

                last_error = (
                    "PDF validation failed"
                )

                temp_file.unlink(
                    missing_ok=True
                )

                raise Exception(
                    last_error
                )


            # ------------------------------------------------
            # RENAME TEMP → FINAL
            # ------------------------------------------------

            temp_file.replace(
                output_file
            )


            # ------------------------------------------------
            # SUCCESS
            # ------------------------------------------------

            time.sleep(
                DOWNLOAD_DELAY
            )


            return {

                "Invoice Number":
                    invoice_number,

                "Invoice Date":
                    invoice_date,

                "Invoice PDF URL":
                    pdf_url,

                "Status":
                    "SUCCESS",

                "Attempts":
                    attempt,

                "File Path":
                    str(output_file),

                "Downloaded At":
                    datetime.now(),

                "Error":
                    ""
            }


        except Exception as e:

            last_error = str(e)


            temp_file.unlink(
                missing_ok=True
            )


            # ------------------------------------------------
            # RETRY DELAY
            # ------------------------------------------------

            if attempt < MAX_RETRIES:

                wait_time = (
                    2 ** attempt
                )

                time.sleep(
                    wait_time
                )


    # ========================================================
    # FAILED
    # ========================================================

    return {

        "Invoice Number":
            invoice_number,

        "Invoice Date":
            invoice_date,

        "Invoice PDF URL":
            pdf_url,

        "Status":
            "FAILED",

        "Attempts":
            MAX_RETRIES,

        "File Path":
            "",

        "Downloaded At":
            datetime.now(),

        "Error":
            last_error
    }


# ============================================================
# 14. PREPARE TASKS
# ============================================================

tasks = []


for _, row in download_df.iterrows():

    tasks.append({

        "Invoice Number":
            str(row["Invoice Number"]),

        "Invoice Date":
            row["Invoice Date"],

        "Invoice PDF URL":
            str(row["Invoice PDF URL"])

    })


# ============================================================
# 15. START FAST DOWNLOAD
# ============================================================

print("\n")
print("=" * 75)
print("STARTING PARALLEL DOWNLOAD")
print("=" * 75)

print(
    f"Parallel workers : {MAX_WORKERS}"
)

print(
    f"Retry attempts   : {MAX_RETRIES}"
)

print(
    f"Total invoices   : {len(tasks):,}"
)

print(
    "\nYou can safely stop this program."
)

print(
    "Run it again later to resume."
)

print(
    "=" * 75
)


start_time = datetime.now()


results = []


completed = 0
successful = 0
failed = 0
skipped = 0


# Lock used for console output

print_lock = threading.Lock()


# ============================================================
# 16. PARALLEL EXECUTION
# ============================================================

with ThreadPoolExecutor(
    max_workers=MAX_WORKERS
) as executor:


    future_map = {}


    for task in tasks:

        future = executor.submit(

            download_invoice,

            task["Invoice Number"],

            task["Invoice Date"],

            task["Invoice PDF URL"]

        )

        future_map[
            future
        ] = task


    # --------------------------------------------------------
    # PROCESS COMPLETED DOWNLOADS
    # --------------------------------------------------------

    for future in as_completed(
        future_map
    ):

        task = future_map[
            future
        ]


        try:

            result = future.result()


        except Exception as e:

            result = {

                "Invoice Number":
                    task["Invoice Number"],

                "Invoice Date":
                    task["Invoice Date"],

                "Invoice PDF URL":
                    task["Invoice PDF URL"],

                "Status":
                    "FAILED",

                "Attempts":
                    MAX_RETRIES,

                "File Path":
                    "",

                "Downloaded At":
                    datetime.now(),

                "Error":
                    str(e)
            }


        results.append(
            result
        )


        completed += 1


        if result["Status"] == "SUCCESS":

            successful += 1

            status_text = "SUCCESS"


        elif (
            result["Status"]
            == "SKIPPED_EXISTING"
        ):

            skipped += 1

            status_text = "SKIPPED"


        else:

            failed += 1

            status_text = "FAILED"


        # ----------------------------------------------------
        # PROGRESS
        # ----------------------------------------------------

        with print_lock:

            print(

                f"[{completed:,}/"
                f"{len(tasks):,}] "
                f"{result['Invoice Number']} "
                f"-> "
                f"{status_text}"

            )


        # ----------------------------------------------------
        # SAVE PROGRESS
        # ----------------------------------------------------

        if (
            completed % SAVE_EVERY == 0
        ):

            current_df = pd.DataFrame(
                results
            )


            if not existing_log.empty:

                combined_log = pd.concat(

                    [
                        existing_log,
                        current_df
                    ],

                    ignore_index=True

                )

            else:

                combined_log = current_df


            combined_log = (
                combined_log
                .drop_duplicates(

                    subset=[

                        "Invoice Number",
                        "Invoice PDF URL"

                    ],

                    keep="last"

                )
            )


            combined_log.to_excel(

                LOG_FILE,

                index=False

            )


            print("\n")
            print(
                f"Progress saved: "
                f"{completed:,} completed"
            )
            print("\n")


# ============================================================
# 17. FINAL LOG
# ============================================================

new_results_df = pd.DataFrame(
    results
)


if not existing_log.empty:

    final_log = pd.concat(

        [
            existing_log,
            new_results_df
        ],

        ignore_index=True

    )

else:

    final_log = new_results_df


# Remove duplicate invoice records

if not final_log.empty:

    final_log = (
        final_log
        .drop_duplicates(

            subset=[

                "Invoice Number",
                "Invoice PDF URL"

            ],

            keep="last"

        )
    )


# Save final log

final_log.to_excel(

    LOG_FILE,

    index=False

)


# ============================================================
# 18. FAILED REPORT
# ============================================================

if not final_log.empty:

    failed_df = final_log[
        final_log["Status"]
        == "FAILED"
    ].copy()


    if not failed_df.empty:

        failed_df.to_excel(

            FAILED_FILE,

            index=False

        )


# ============================================================
# 19. FINAL STATISTICS
# ============================================================

end_time = datetime.now()


elapsed_time = (
    end_time
    - start_time
)


# Count from final log

if not final_log.empty:

    total_success = (
        final_log["Status"]
        == "SUCCESS"
    ).sum()


    total_skipped = (
        final_log["Status"]
        == "SKIPPED_EXISTING"
    ).sum()


    total_failed = (
        final_log["Status"]
        == "FAILED"
    ).sum()

else:

    total_success = 0
    total_skipped = 0
    total_failed = 0


# ============================================================
# 20. FINAL REPORT
# ============================================================

print("\n")
print("=" * 75)
print("DOWNLOAD PROCESS COMPLETED")
print("=" * 75)


print(
    f"Successful PDFs       : "
    f"{total_success:,}"
)


print(
    f"Existing PDFs skipped : "
    f"{total_skipped:,}"
)


print(
    f"Failed PDFs            : "
    f"{total_failed:,}"
)


print(
    f"Processing time        : "
    f"{elapsed_time}"
)


print("\n")
print("PDF location:")

print(
    OUTPUT_DIR
)


print("\n")
print("Download log:")

print(
    LOG_FILE
)


if total_failed > 0:

    print("\n")
    print("Failed invoice report:")

    print(
        FAILED_FILE
    )


print("\n")
print("=" * 75)
print("AUTOMATION FINISHED")
print("=" * 75)