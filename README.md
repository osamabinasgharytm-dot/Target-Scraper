# Target.com Product Scraper (v1)

On-demand Python pipeline that:

1. Reads DPCIs from `dpcis.csv`
2. Resolves each DPCI → Target TCIN via Redsky search
3. Scrapes product details + paginated reviews
4. Cleans reviews, runs **VADER** sentiment, extracts one-word remarks
5. Writes a combined Excel workbook with `Products` and `Reviews` sheets

Google Sheets sync and daily scheduling are **deferred** (not in v1).

## Setup

```powershell
cd path\to\Scrapper
python -m venv .venv
# .\.venv\Scripts\Activate.ps1
.venv\Scripts\Activate
.venv\Scripts\Deactivate
pip install -r requirements.txt
```

Optional: copy `.env.example` to `.env` and adjust delays / review page size.

## Run

From the project root (with the venv active):

```powershell
# Smoke test — first 5 DPCIs
python -m src.main --limit 5

# Fetch up to 100 reviews per DPCI (without limiting DPCI count)
python -m src.main --review-limit 100

# Single DPCI
python -m src.main --dpci 062-18-0006

# Selected DPCIs — export current-month reviews to the combined workbook
python -m src.main --dpcis 060-15-0108,097-02-2055

# All DPCIs from dpcis.csv — export current-month reviews
python -m src.main --current-month

# Selected DPCIs — fetch up to 2,550 reviews each, then export August–September 2026
python -m src.main --dpcis 060-15-0108,097-02-2055 --review-limit 2550 --from-month 2026-08 --to-month 2026-09

# All DPCIs from dpcis.csv — fetch up to 2,550 reviews each, export August–September 2026
python -m src.main --review-limit 2550 --from-month 2026-08 --to-month 2026-09

# Full file (~2,200 DPCIs; several hours with default delay)
python -m src.main
```

### Outputs

| File | Contents |
|------|----------|
| `output/reviews_range.xlsx` | Reviews for specified DPCIs and inclusive month range |
| `output/<Current Month> Target Data.xlsx` | Combined workbook with `Products` and `Reviews` sheets |
| `output/run.log` | Run log |

Separate product and review CSV files are disabled. Product summaries and selected reviews are written to the combined workbook.

### CLI flags

| Flag | Description |
|------|-------------|
| `--csv PATH` | Input CSV (default `dpcis.csv`) |
| `--limit N` | Process only first N DPCIs |
| `--review-limit N` | Maximum reviews fetched per DPCI (default: 50) |
| `--dpci ###-##-####` | Process one DPCI |
| `--dpcis DPCI1,DPCI2,...` | Process a comma-separated list of specific DPCIs |
| `--from-month YYYY-MM` | First month to include in the Excel review export; without `--dpci`/`--dpcis`, process DPCIs from `--csv` |
| `--to-month YYYY-MM` | Last month to include in the Excel review export (inclusive) |
| `--excel-output PATH` | Excel output path for a month-range export |
| `--workbook-output PATH` | Output path for the combined workbook (default: `output/<Current Month> Target Data.xlsx`) |
| `--save-reviews` | Deprecated; review rows are included in the combined workbook automatically |
| `--current-month` | Read DPCIs from the input CSV and export only current-month reviews |

## Products sheet columns

| Column | Source |
|--------|--------|
| DPCI | Input / API |
| Item Description | Product title |
| Program | Matching `PROGRAM` value from `DPCI LIBRARY.xlsx` |
| Item Type | Target item type / category breadcrumb |
| Brand | Product brand |
| Avg. Rating | Mean of scraped review stars (fallback: API summary) |
| Positive % / Negative % / Neutral % | VADER review labels |
| Average Satisfaction % | Mean of review-level satisfaction percentages |
| Top Remark | Most common one-word keyword across reviews |
| Status | Matching `STATUS` value from `DPCI LIBRARY.xlsx` |
| TCIN | Target item number |

Some DPCIs may return `not_found` if the item is no longer searchable on Target.com.
For a month-range export of all DPCIs in `dpcis.csv`, pass `--from-month` and `--to-month`; use `--csv PATH` to select a different input file. To limit the export to chosen products, also pass `--dpcis` (or one `--dpci`). Both month endpoints are included. Matching reviews are written to Excel.
Reviews are fetched newest-first, up to `--review-limit` per DPCI (default `50`). Increase the limit when exporting older months; the Target API supports at most 2,550 reviews per product.
The combined workbook's `Reviews` sheet includes the VADER sentiment, normalized score (-1 to +1), and individual satisfaction percentage (0 to 100). Only negative reviews are matched case-insensitively against whole words and contiguous phrases in `DEFECTS LIST.xlsx`; matching `Reason`, source `Sub Category` (shown as `Defect` in the review sheet), and `Main Category` are written in adjacent columns after `Review Text`. Multiple matches are joined with ` | `. Positive and neutral reviews show `-` in all three defect columns. The `Products` sheet includes mean satisfaction per product; positive/negative/neutral percentages are based on the VADER labels.

## Power BI

1. Open Power BI Desktop → **Get data** → **Excel workbook**
2. Select `output/<Current Month> Target Data.xlsx`
3. Import the `Products` sheet and refresh after each scraper run.

Suggested visuals (from the flow diagram):

- Cards: Total Products, Avg. Rating, Positive %, Negative %
- Donut: sentiment distribution (use Positive/Negative/Neutral % columns, or unpivot)
- Bar chart: Top Remark

Scrape success and failure counts are recorded in `output/run.log`; the product sheet's `Status` column contains the DPCI Library status.

## Configuration

See `.env.example`:

- `DPCI_LIBRARY_XLSX` (default `DPCI LIBRARY.xlsx` in the project root) — Excel lookup keyed by the `DPCI` column; `PROGRAM` and `STATUS` populate matching product output fields
- `DEFECTS_LIBRARY_XLSX` (default `DEFECTS LIST.xlsx` in the project root) — review phrase lookup with `Reason`, `Sub Category`, and `Main Category` columns
- `REVIEW_LIMIT` (default `50`) — maximum total reviews fetched per DPCI; can be overridden with `--review-limit N`.
- `REQUEST_DELAY_SECONDS` (default `1.2`) — be polite to Target
- `REVIEW_PAGE_SIZE` (default `50`; Target API maximum) — number of reviews requested per API page. Target allows page indexes 0–50, so the API can return at most 2,550 reviews per DPCI.
- `STORE_ID` / location knobs for pricing context

## Deferred (post-v1)

- Google Sheets API append + history
- Windows Task Scheduler daily loop
- Paid proxies / browser automation if Target blocks plain requests

## Notes

- Review exports include the VADER compound `Sentiment Score` (-1 to +1) and `Satisfaction Percentage` (0 to 100); 5-star/4.5+ reviews are forced positive as in the supplied sentiment script.
- Uses Target’s public JSON APIs (Redsky + r2d2), the same endpoints the website calls.
- API keys are public web keys bootstrapped from Target’s homepage when possible.
- Use responsibly and respect Target’s terms of use and rate limits.
