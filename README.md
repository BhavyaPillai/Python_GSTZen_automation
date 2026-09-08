## GST Invoice Download Automation Pipeline

A **Python-based automation pipeline** for downloading GST invoice PDFs in bulk from GSTZen and organizing them into a structured, month-wise folder system.

## 🚀 Overview

This project replaces repetitive manual invoice downloading with an automated workflow using Python.

```text
GSTZen Excel Export
        ↓
Data Cleaning & FY Filtering
        ↓
PDF URL Validation
        ↓
Parallel PDF Download
        ↓
PDF Validation
        ↓
Month-wise Organization
        ↓
Download & Error Logs
```

## ✨ Key Features

* 📊 Excel data processing using **Pandas**
* 🧹 Data cleaning and duplicate handling
* 📅 Financial-year filtering
* 🔗 PDF URL validation
* ⚡ Parallel downloading using `ThreadPoolExecutor`
* 🔄 Automatic retry for failed downloads
* ♻️ Resume capability by skipping existing PDFs
* 📄 PDF validation before saving
* 📁 Automatic month-wise folder creation
* 📝 Success and failure logging
* ❌ Separate failed-invoice report

## 🛠️ Tech Stack

**Python · Pandas · Requests · ThreadPoolExecutor · Excel**

## 📂 Project Structure

```text
GST-INVOICE-AUTOMATION/
│
├── scripts/
│   ├── 01_prepare_invoice_list.py
│   ├── 02_test_pdf.py
│   └── 03_download_invoices.py
│
├── output/
│   └── FY_2025_26/
│       ├── Apr-2025/
│       ├── May-2025/
│       ├── ...
│       └── Mar-2026/
│
├── logs/
│   ├── Invoice_Download_Log.xlsx
│   └── Failed_Invoices.xlsx
│
├── requirements.txt
└── README.md
```

## 🔄 How It Works

1. **Prepare Data** – Read the GSTZen Excel export and clean/filter the invoice records.
2. **Validate URLs** – Check whether invoice URLs return valid PDF files.
3. **Download** – Download multiple invoices simultaneously using parallel processing.
4. **Validate & Store** – Validate each PDF and save it according to the invoice month.
5. **Resume & Retry** – Skip already downloaded files and retry temporary failures.
6. **Log Results** – Maintain download status and failed-invoice reports.

## 📈 Test Result

The pipeline was tested with **100 invoices**:

* **100 successful downloads**
* **0 failed downloads**
* **20 existing PDFs skipped**
* Processing time: **~53 seconds**

## 🎯 Objective

To convert a repetitive manual invoice-download activity into a **reliable, resumable, and scalable Python automation workflow** with structured document storage and tracking.

> **Note:** Use only with authorized business data. Do not commit credentials, tokens, or confidential invoice information to GitHub.
