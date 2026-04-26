---
name: amazon-order-tax-summary
description: >
  Converts an Amazon Seller Central order report (TSV) into a formatted Excel
  workbook with a detailed tab (all rows + Sales Including Tax column) and a
  state-by-state tax summary tab (all 52 US entries + international groups).
  Use this skill whenever the user wants to generate an Amazon order tax
  summary, process an Amazon order report, or create a state sales summary
  from Amazon data — even if they just say "run my Amazon report" or
  "make the tax summary".
---

# Amazon Order Tax Summary

Generates an Excel tax summary from a raw Amazon Seller Central order report.
Works for any Amazon seller. No coding required — just provide the input file.

---

## Step 1: Get the input file

Check whether the user has already provided a file path or uploaded a file.

If not, ask:

> Please provide your Amazon order report file. You can:
> - Paste the full file path (e.g. `/Users/you/Downloads/Amazon order report 2026 Jan.txt`)
> - Or tell me where you saved it after downloading from Seller Central
>
> To download the report: **Amazon Seller Central → Reports → Fulfillment → All Orders**
> Set date range type to **Order Date**, select the full month, click **Request Report**.

Wait for the user to provide the path before continuing.

---

## Step 2: Locate the script

Check if `amazon_tax_summary.py` is already available:

```bash
# Check common locations
ls amazon_tax_summary.py 2>/dev/null || \
ls ~/amazon-seller-tax-summary/amazon_tax_summary.py 2>/dev/null || \
echo "NOT_FOUND"
```

**If found:** use that path.

**If NOT_FOUND:** download it from GitHub into a temp location:

```bash
curl -fsSL \
  "https://raw.githubusercontent.com/kaimai/amazon-seller-tax-summary/main/amazon_tax_summary.py" \
  -o /tmp/amazon_tax_summary.py
SCRIPT=/tmp/amazon_tax_summary.py
```

Confirm download succeeded (non-zero file size) before continuing.

---

## Step 3: Check dependencies

```bash
python3 -c "import pandas, openpyxl" 2>&1
```

If the import fails, install:

```bash
pip3 install pandas openpyxl
```

---

## Step 4: Determine output destination

Ask the user:

> Where should I save the output Excel file?
>
> A) Same folder as the input file (recommended)
> B) Somewhere else — tell me the path

If A (or no preference): the script writes the output next to the input automatically — no extra flag needed.

If B: note the desired path; after running, move the file there.

---

## Step 5: Check for an existing quarterly workbook

Ask the user:

> Is this a standalone month, or should I add it to an existing quarterly
> workbook (e.g. `Amazon order tax summary - 2026 Q1.xlsx`)?
>
> A) Standalone — create a new file for this month
> B) Add to an existing quarterly file — provide the path

If B: collect the path to the existing workbook; use `--append-to` flag in Step 6.

---

## Step 6: Run the script

**Standalone (no existing workbook):**

```bash
python3 "$SCRIPT" "<INPUT_FILE_PATH>"
```

**Append to existing quarterly workbook:**

```bash
python3 "$SCRIPT" "<INPUT_FILE_PATH>" --append-to "<QUARTERLY_FILE_PATH>"
```

Capture stdout and check for errors. Expected output looks like:

```
Detected period: 202601
Loaded 48 rows, 37 columns
Wrote '202601 detailed' tab: 48 data rows
Wrote '202601 summary' tab: 52 US states
Saved: /path/to/Amazon order tax summary - 202601.xlsx
```

If the script exits with an error, show the error message and stop.

---

## Step 7: Report results

Tell the user:

- The full path to the output file
- How many rows were processed
- Which tabs were created
- A brief summary of the totals (read the Grand Total from the summary tab):

```python
import openpyxl
wb = openpyxl.load_workbook("<OUTPUT_PATH>")
ws = wb["<YYYYMM> summary"]
for row in ws.iter_rows(values_only=True):
    if row[0] == "Grand Total":
        print(f"Grand Total: ${row[2]:,.2f}")
        break
```

Example response to user:

> Done! Your tax summary is saved to:
> `/Users/you/Downloads/Amazon order tax summary - 202601.xlsx`
>
> - **202601 detailed** — 48 rows (all orders including cancelled)
> - **202601 summary** — 52 US state/territory rows + international groups
> - **Grand Total: $18,277.88**
>
> Open the file to review. The summary tab shows every US state — states with
> no orders show $0.00.

---

## Notes

- Full state names (e.g. `MARYLAND`, `CALIFORNIA`) are automatically normalized to 2-letter abbreviations.
- Cancelled orders appear in the detailed tab with $0.00 but do not affect summary totals.
- Puerto Rico (PR) is included in the US section of the summary.
- International orders are grouped by country code above the US section.
- Running the same month twice on the same workbook safely overwrites that month's tabs.
