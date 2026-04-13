# atm-receivables-reconciliation
Automated ATM receivables GL reconciliation system built with Google Apps Script — replacing a manual Excel process for a Nigerian bank branch

# 🏧 ATM Receivables GL Reconciliation

> Automated reconciliation of 5 ATM receivable GL accounts for a Nigerian bank branch —
> replacing a multi-hour manual Excel process with a one-click solution.

---

## The Problem

Every reconciliation period, an operations analyst had to manually:
- Export 5 separate GL Activity Reports from the core banking system
- Open each one, copy transactions row by row into a master Excel workbook
- Extract the RRN (transaction reference) from complex description strings by hand
- Recalculate proof balances and check for differences
- The process took **2–3 hours** and was prone to copy-paste errors

## The Solution

Two automation layers, each independently usable:

| Layer | Tool | Best For |
|---|---|---|
| Python script | pandas + openpyxl | Running locally, CI pipelines, BigQuery integration |
| Google Apps Script | Google Sheets + Drive API | Browser-only, no installations required |

Both layers perform the same logic: read the GL files, extract RRNs, append transactions,
recalculate proof balances, and output a structured summary.

## Results

| Metric | Before | After |
|---|---|---|
| Time per reconciliation | 2–3 hours | < 5 minutes |
| Manual steps | ~200 | 4 (drop files → open sheet → click run → verify) |
| Error risk | High (manual copy-paste) | Near-zero (formula-verified, duplicate-protected) |
| GL accounts handled | 5 (one by one) | 5 simultaneously |

All 5 GL accounts reconciled to **zero difference** on first automated run.

---

## GL Accounts Covered

| Sheet | GL Account | Description |
|---|---|---|
| 16436-119110010 | 119110010 | ISW ATM Settlement Receivable |
| 16484-119110038 | 119110038 | MC Domestic ATM Receivable |
| 16533-119130021 | 119130021 | Appzone ATM Withdrawal Settlement |
| 16459-119110026 | 119110026 | VISA V-Pay Settlement Receivable |
| 119110093 | 119110093 | AFRIGO Financial Services Receivable |

---

## How It Works

### RRN Extraction
The core challenge was extracting the transaction reference number (RRN) from
inconsistent description strings produced by the core banking system. The solution
handles 6 distinct formats:
