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
| `YYYYMM summary` | Sales Including Tax and Sales Tax aggregated by `ship-country` / `ship-state`, with all 52 US entries shown (50 states + DC + Puerto Rico), country subtotals, and a grand total |

States with zero sales still appear in the summary with a `0.00` value.

## Notes

- Full state names in the raw report (e.g. `MARYLAND`, `CALIFORNIA`) are automatically converted to 2-letter abbreviations before aggregation.
- Cancelled orders are included in the detailed tab (with `0.00` sales) to preserve a complete audit trail, but contribute nothing to the summary totals.
- International orders are grouped by country above the US section in the summary tab.

## Using with Claude Code (no coding required)

Claude Code can run this tool for you automatically using the included skill. There are two ways to use Claude Code — pick whichever fits you.

### Option A: Claude Code on Claude Desktop

Claude Desktop has a built-in **Claude Code** tab (a full coding assistant with filesystem access). No terminal required.

1. Download and install [Claude Desktop](https://claude.ai/download)
2. Open Claude Desktop and click the **Claude Code** tab
3. Open the folder where your order report is saved as a project
4. Install the skill — paste this into the Claude Code chat:

   ```bash
   mkdir -p .claude/skills/amazon-order-tax-summary && curl -fsSL "https://raw.githubusercontent.com/kaimai/amazon-seller-tax-summary/main/skill/SKILL.md" -o .claude/skills/amazon-order-tax-summary/SKILL.md
   ```

5. Then just say:

   > Run my Amazon order tax summary for January 2026.

Claude will ask for your file path, run the script, and tell you where the output was saved.

### Option B: Claude Code in the terminal (CLI)

1. Install Claude Code:

   ```bash
   npm install -g @anthropic-ai/claude-code
   ```

2. Open a terminal in the folder where your order report is saved:

   ```bash
   cd ~/Downloads
   claude
   ```

3. Install the skill — paste this into the Claude Code session:

   ```bash
   mkdir -p .claude/skills/amazon-order-tax-summary && curl -fsSL "https://raw.githubusercontent.com/kaimai/amazon-seller-tax-summary/main/skill/SKILL.md" -o .claude/skills/amazon-order-tax-summary/SKILL.md
   ```

4. Then say:

   > Run my Amazon order tax summary for January 2026.

## Using with an AI Chatbot (Claude, ChatGPT, etc.)

You can ask an AI assistant to run this tool on your behalf — no command line needed.

1. Upload your Amazon order report file to the chat
2. Also upload `amazon_tax_summary.py` (download it from this repo), **or** paste the raw script URL so the AI can fetch it itself:
   ```
   https://raw.githubusercontent.com/kaimai/amazon-seller-tax-summary/main/amazon_tax_summary.py
   ```
3. Paste this prompt:

> I've uploaded my Amazon Seller Central order report (TSV) and the script `amazon_tax_summary.py`. Please run the script on my order report and give me back the Excel file with a detailed tab and a state-by-state sales summary tab.

The AI will install any missing dependencies, run the script, and return the output Excel — all in one step.

If you use **Claude Code** with filesystem access, you can point it directly at local files instead of uploading:

> Download `amazon_tax_summary.py` from `https://raw.githubusercontent.com/kaimai/amazon-seller-tax-summary/main/amazon_tax_summary.py`, then run it on `/path/to/Amazon order report 2026 Jan.txt`.

## Requirements

```
pandas
openpyxl
```

Install with:

```bash
pip install pandas openpyxl
```
