# Amazon Order Tax Summary

Converts a raw Amazon order report (TSV) into a formatted Excel workbook with a detailed tab and a state-by-state tax summary tab.

## Generating the Input File

1. Log in to **Amazon Seller Central**
2. Go to **Reports → Fulfillment → All Orders**
3. Set date range type to **Order Date**
4. Select the event range to cover the **whole month** (e.g. Jan 1 – Jan 31)
5. Click **Request Report**, then download when ready
6. Rename the file to match the pattern: `Amazon order report YYYY Mon.txt` (e.g. `Amazon order report 2026 Jan.txt`)

## Usage

### Single month — creates a new file

```bash
python amazon_tax_summary.py "Amazon order report 2026 Jan.txt"
```

Output: `Amazon order tax summary - 202601.xlsx` in the same directory as the input.

### Append a month to an existing workbook (e.g. a quarterly file)

```bash
python amazon_tax_summary.py "Amazon order report 2026 Feb.txt" \
    --append-to "Amazon order tax summary - 2026 Q1.xlsx"
```

Running the same month twice safely overwrites the existing tabs for that period.

## Output Format

Each processed month adds two tabs to the workbook:

| Tab | Contents |
|-----|----------|
| `YYYYMM detailed` | All rows from the raw report, plus a computed **Sales Including Tax** column (`item-price + item-tax`) |
| `YYYYMM summary` | Sales Including Tax aggregated by `ship-country` / `ship-state`, with all 52 US entries shown (50 states + DC + Puerto Rico), country subtotals, and a grand total |

States with zero sales still appear in the summary with a `0.00` value.

## Notes

- Full state names in the raw report (e.g. `MARYLAND`, `CALIFORNIA`) are automatically converted to 2-letter abbreviations before aggregation.
- Cancelled orders are included in the detailed tab (with `0.00` sales) to preserve a complete audit trail, but contribute nothing to the summary totals.
- International orders are grouped by country above the US section in the summary tab.

## Requirements

```
pandas
openpyxl
```

Install with:

```bash
pip install pandas openpyxl
```
