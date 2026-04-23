# GSA QBR — TAE Pilot Performance Automation

Automates the population of **Slide 7 (TAE Pilot Performance)** in the GSA Quarterly Business Review deck.

Instead of manually copying numbers from SQL query results into PowerPoint every quarter, this notebook runs the queries, formats the numbers, and writes them directly into the slide table — all in one place.

---

## How It Works

The notebook has 3 steps:

1. **Connect to Darwin** — establishes a Trino session via `linkedin.lisql`
2. **Run 2 SQL queries** — `pilot_master` and `renewal_core_master` pull all the data pre-calculated, including NAMER, EMEAL, and Total GSA rows (via `GROUPING SETS`). Timeframes (Current Quarter, Prior Quarter, YoY) are auto-detected from `current_date`, so nothing needs to change between quarters.
3. **Populate the PowerPoint** — reads the two DataFrames, formats each number for display (`0.19` → `19%`, `13400` → `$13.4`, etc.), and writes them into the correct cells of the slide table while preserving the original fonts, colors, and alignment.

The end result is a populated `.pptx` file ready to be presented.

---

## Metrics Covered

The slide table has 7 metric groups, each with 3 columns (Prior Quarter, Current QTD, YoY):

| # | Metric | Source Query | Value Format | YoY Format |
|---|--------|-------------|-------------|-----------|
| 1 | # of Paid Pilots | `pilot_master` | `1,706` | `-28%` |
| 2 | Pilot ASPs ($ in K) | `pilot_master` | `$13.4` | `+6%` |
| 3 | Avg. Discount | `pilot_master` | `19%` | `-3ppt` |
| 4 | Avg. Duration (Months) | `pilot_master` | `5.4` | `+1%` |
| 5 | Pilot Bookings % of TAE | `pilot_master` | `17%` | `-13ppt` |
| 6 | Pilot Renewal Conversion % | `renewal_core_master` | `52%` | `-15ppt` |
| 7 | Pilot Renewal RIG | `renewal_core_master` | `0.69` | `-4ppt` |

Rows: **NAMER**, **EMEAL**, **Total GSA** (3 regions × 7 metrics × 3 columns = 63 cells total).

---

## Setup

### Prerequisites

- VS Code connected to Darwin (remote Jupyter kernel)
- `python-pptx` installed in the environment:
  ```
  pip install python-pptx
  ```

### File Structure

```
QBRautomation/
├── QBR_Slide_Automate.ipynb       # The automation notebook
├── GSA_QBR_Template_Slide7.pptx   # Template (slide 7 only)
└── README.md
```

### Configuration

In the notebook's **Step 1** cell, update the paths to match your environment:

```python
BASE_DIR = Path('.')  # same folder as the notebook
TEMPLATE = BASE_DIR / 'GSA_QBR_Template_Slide7.pptx'
OUTPUT   = BASE_DIR / 'GSA_QBR_Slide7_populated.pptx'
```

The slide and table identifiers are also configured there:

```python
SLIDE_INDEX = 0          # 0 if template has only slide 7, 6 if full deck
TABLE_NAME  = 'Table 11' # Shape name of the table in the slide
```

---

## Usage

1. Open the notebook in VS Code with a Darwin kernel
2. Run all cells
3. The populated deck is saved to `OUTPUT` path

The SQL queries auto-detect the current quarter from `current_date`, so **no manual date changes are needed between quarters**. Just run it.

---

## How the PowerPoint Mapping Works

The code doesn't "see" the slide visually. It treats the `.pptx` as a data structure:

- **Table identification**: finds the table by its shape name (`Table 11`)
- **Row identification**: reads column 0 of the table to find which row is NAMER, EMEAL, or Total GSA
- **Column identification**: the table has 22 columns. Column 0 is the label, then groups of 3 columns per metric (prior, current, yoy). The mapping is defined in `CELL_MAP`
- **Format preservation**: when writing text into a cell, the code saves the existing font, size, color, and alignment, writes the new text, then re-applies the saved formatting

If the slide template changes (rows added, columns moved, table renamed), the mapping in `CELL_MAP` and `ROWS` may need to be updated.

---

## Architecture

```
Darwin (Trino/SQL)          Python (notebook)              PowerPoint
┌─────────────────┐         ┌──────────────────┐          ┌───────────────┐
│  pilot_master    │────────▶│  Format numbers  │─────────▶│  Slide 7      │
│  (metrics 1-5)   │         │  (fmt_pct, etc.) │          │  Table 11     │
├─────────────────┤         │                  │          │               │
│  renewal_core    │────────▶│  Write to cells  │          │  63 cells     │
│  (metrics 6-7)   │         │  (set_cell)      │          │  populated    │
└─────────────────┘         └──────────────────┘          └───────────────┘
```

The SQL queries do all the heavy lifting (aggregation, YoY calculation, Total GSA via GROUPING SETS). Python only formats and writes.

---

## Division of Work

| Area | Owner | Description |
|------|-------|-------------|
| SQL queries | Igor | Query logic, parametrization, GROUPING SETS, new metrics |
| Python / PPTX | João | Number formatting, slide population, template mapping |

---

## Inspired By

This project follows the same pattern as the MBR Automation by Samuel Sorensen, adapted for the GSA Pilot data and simplified to use fewer, more complete SQL queries.
