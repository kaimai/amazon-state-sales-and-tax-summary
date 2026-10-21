# Non-USD Currency and Blank-Country Handling Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Stop non-USD orders from being summed into the USD Grand Total, and stop empty blank-country rows from producing a nameless " Total" row in the summary tab.

**Problem (observed in a real monthly report):**
- An order shipped to MX with `currency=MXN` was added to the Grand Total as if its peso amount were USD, inflating the Grand Total.
- An order with blank ship-country/currency/prices made the summary show a nameless `" Total"` row with 0.00.

**Architecture:** All changes live in `build_summary_rows()` in `amazon_state_sales_and_tax_summary.py`. Key decision: the rule keys off the `currency` column, not country — the existing fixture's Canada order is `USD` and must keep counting toward the Grand Total.
- `currency` blank or `USD` → treated as USD (current behaviour, unchanged).
- Any other currency → grouped as `"<country> (<currency>)"` (e.g. `MX (MXN)`), shown with its own subtotal in native currency, **excluded from Grand Total**. No FX conversion (out of scope; would need an external rate source).
- Rows with blank ship-country AND zero Sales Including Tax AND zero item-tax → dropped from the summary (they still appear in the detailed tab). Blank-country rows that do carry money are labelled `(blank)` so nothing is silently hidden.
- Missing `currency` column (synthetic test inputs) → treated as all-USD.

**Tech Stack:** Python 3, pandas, openpyxl, unittest/pytest.

**Success criteria:**
1. New tests pass through the real consumer path (`ats.main` → xlsx → `read_summary_tab`).
2. All 19 existing tests still pass (January fixture output unchanged).
3. Regenerated real report: Grand Total equals US Total plus any USD international groups; the MXN order appears under `MX (MXN) Total`; no `" Total"` row.

---

### Task 1: Failing tests

**Files:**
- Modify: `tests/test_integration.py` (append new class before `class TestAppendMode`)

- [ ] **Step 1: Write the failing tests**

```python
class TestCurrencyAndBlankCountry(unittest.TestCase):
    """Non-USD orders stay out of the Grand Total; empty blank-country rows are dropped."""

    def _run(self, rows):
        import tempfile, pandas as pd
        tmp = tempfile.mkdtemp()
        p = Path(tmp) / "Amazon order report 2026 Jul.txt"
        pd.DataFrame(rows).to_csv(p, sep='\t', index=False)
        ats.main(str(p))
        out = Path(tmp) / "Amazon state sales and tax summary - 202607.xlsx"
        return read_summary_tab(out, "202607 summary")

    US_ROW = {"amazon-order-id": "2", "currency": "USD", "item-price": "100.00", "item-tax": "8.00",
              "ship-state": "CA", "ship-country": "US"}

    def test_non_usd_grouped_with_currency_label(self):
        actual = self._run([{"amazon-order-id": "1", "currency": "MXN", "item-price": "5999.46",
                             "item-tax": "", "ship-state": "DISTRITO FEDERAL", "ship-country": "MX"},
                            self.US_ROW])
        mx_total = next(sit for c, s, sit, tax in actual if c == "MX (MXN) Total")
        self.assertAlmostEqual(mx_total, 5999.46, places=2)

    def test_non_usd_excluded_from_grand_total(self):
        actual = self._run([{"amazon-order-id": "1", "currency": "MXN", "item-price": "5999.46",
                             "item-tax": "", "ship-state": "DISTRITO FEDERAL", "ship-country": "MX"},
                            self.US_ROW])
        grand_sit, grand_tax = next((sit, tax) for c, s, sit, tax in actual if c == "Grand Total")
        self.assertAlmostEqual(grand_sit, 108.00, places=2)
        self.assertAlmostEqual(grand_tax, 8.00, places=2)

    def test_empty_blank_country_row_dropped(self):
        actual = self._run([{"amazon-order-id": "S01-1", "quantity": "2"}, self.US_ROW])
        countries = {c for c, s, sit, tax in actual}
        self.assertEqual(countries, {"US", "US Total", "Grand Total"})
```

- [ ] **Step 2: Run to verify they fail**

Run: `python3 -m pytest tests/test_integration.py::TestCurrencyAndBlankCountry -v`
Expected: 3 FAIL (`StopIteration` for `MX (MXN) Total`; Grand Total 6107.46 ≠ 108.00; countries include `" Total"`).

- [ ] **Step 3: Commit**

```bash
git add tests/test_integration.py
git commit -m "test: cover non-USD currency and empty blank-country rows in summary"
```

### Task 2: Implement in `build_summary_rows()`

**Files:**
- Modify: `amazon_state_sales_and_tax_summary.py` — `build_summary_rows()`

- [ ] **Step 1: Add currency normalization, empty-row drop, and labelling** — insert directly after the two `.str.strip()` lines for `ship-country` / `ship-state`, before the PR mask:

```python
    if 'currency' not in work.columns:
        work['currency'] = ''
    work['currency'] = work['currency'].fillna('').str.strip().str.upper()

    # Drop blank-country rows that carry no money (e.g. orders with no ship data)
    empty_mask = ((work['ship-country'] == '')
                  & (work['Sales Including Tax'] == 0) & (work['item-tax'] == 0))
    work = work[~empty_mask].copy()
    work.loc[work['ship-country'] == '', 'ship-country'] = '(blank)'

    # Non-USD orders are reported in their own currency and kept out of the Grand Total
    foreign_mask = ~work['currency'].isin(['', 'USD'])
    work.loc[foreign_mask, 'ship-country'] = (
        work.loc[foreign_mask, 'ship-country'] + ' (' + work.loc[foreign_mask, 'currency'] + ')'
    )
```

- [ ] **Step 2: Exclude foreign groups from the Grand Total** — in the international loop, replace

```python
        rows.append((f"{country} Total", '', country_sit, country_tax))
        grand_sit += country_sit
        grand_tax += country_tax
```

with

```python
        rows.append((f"{country} Total", '', country_sit, country_tax))
        if not foreign_mask.loc[country_df.index].any():
            grand_sit += country_sit
            grand_tax += country_tax
```

- [ ] **Step 3: Run full suite**

Run: `python3 -m pytest -q`
Expected: `22 passed`

- [ ] **Step 4: Commit**

```bash
git add amazon_state_sales_and_tax_summary.py
git commit -m "fix: keep non-USD orders out of Grand Total and drop empty blank-country rows"
```

### Task 3: README note

**Files:**
- Modify: `README.md` — `## Notes` section (append two bullets)

- [ ] **Step 1: Add bullets**

```markdown
- Orders in a currency other than USD (e.g. `MXN`) are grouped as `<country> (<currency>)` with a subtotal in that currency, and are **not** included in the Grand Total. No currency conversion is performed.
- Rows with no ship-country and no amounts are omitted from the summary (they remain in the detailed tab).
```

- [ ] **Step 2: Commit**

```bash
git add README.md
git commit -m "docs: document non-USD and blank-country summary handling"
```

### Task 4: Regenerate and verify the affected quarterly report

- [ ] **Step 1: Rebuild** the quarterly workbook from its three monthly reports (first month creates the file, rename it to the quarter name, then `--append-to` the other two). Note: `mv` may be aliased to `mv -i`; use `command mv -f` when overwriting.

```bash
python3 amazon_state_sales_and_tax_summary.py "<dir>/Amazon order report YYYY Mon1.txt" && command mv -f "<dir>/Amazon state sales and tax summary - YYYYMM.xlsx" "<quarter.xlsx>"
python3 amazon_state_sales_and_tax_summary.py "<dir>/Amazon order report YYYY Mon2.txt" --append-to "<quarter.xlsx>"
python3 amazon_state_sales_and_tax_summary.py "<dir>/Amazon order report YYYY Mon3.txt" --append-to "<quarter.xlsx>"
```

- [ ] **Step 2: Verify** totals rows: the month with the MXN order shows `MX (MXN) Total` and its Grand Total equals its US Total; no `" Total"` row; the other two months' totals are unchanged from the pre-fix run.
