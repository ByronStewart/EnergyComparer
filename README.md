# Energy Made Easy Scraper

Compare electricity plans from the Australian government's [Energy Made Easy](https://www.energymadeeasy.gov.au/) website. Enter your postcode, and the scraper fetches all available plans and exports them to a formatted Excel spreadsheet with multiple comparison views and an interactive cost calculator.

Covers postcodes in **NSW, QLD, SA, TAS, and the ACT**.

---

## Quick Start

### Windows (easiest)

Double-click **`run_scraper_enhanced.bat`**. It will create the virtual environment, install dependencies, and launch the scraper automatically.

### Manual Setup

Requires **Python 3.10+**.

```bash
# 1. Clone the repo
git clone <repo-url>
cd SolarComparerOpus

# 2. Create a virtual environment
python -m venv venv

# 3. Activate it
# Windows
venv\Scripts\activate
# macOS / Linux
source venv/bin/activate

# 4. Install dependencies
pip install -r requirements.txt

# 5. Run the scraper
python scraper_enhanced.py
```

The scraper will prompt you for a postcode, then fetch plans and produce an Excel file in the current directory.

---

## Usage Examples

```bash
# Interactive mode - prompts for everything
python scraper_enhanced.py

# Specify a postcode directly
python scraper_enhanced.py 2000

# Multi-distributor postcode - pick which distributor
python scraper_enhanced.py 2850

# Skip the distributor prompt by passing the ID
python scraper_enhanced.py 2850 --dist 13

# Fetch plans for ALL distributors at once
python scraper_enhanced.py 2850 --dist all

# Include controlled load plans (off-peak hot water, pool pump, etc.)
python scraper_enhanced.py 4075 --controlled-load

# Disable all filtering to see every plan from the API
python scraper_enhanced.py 4075 --no-filter

# Gas plans or business plans
python scraper_enhanced.py 2000 --fuel gas
python scraper_enhanced.py 2000 --type business
```

### All Arguments

| Argument | Default | Description |
|---|---|---|
| `postcode` | prompted | 4-digit Australian postcode |
| `--fuel` | `electricity` | `electricity` or `gas` |
| `--type` | `residential` | `residential` or `business` |
| `--dist` | interactive | Distributor ID (e.g. `13`) or `all` |
| `--controlled-load` | off | Include plans with controlled load circuits |
| `--no-filter` | off | Disable all plan filtering |

---

## Output

The scraper produces an Excel file named `energy_plans_{postcode}_{fuel}_{timestamp}.xlsx` with these sheets:

| Sheet | What it contains |
|---|---|
| **Summary** | Search parameters, date, plan count, breakdown by retailer/distributor |
| **All Plans** | Every plan returned, with 28 data columns per plan |
| **Single Rate Plans** | Flat-rate plans only |
| **Time of Use Plans** | Peak/off-peak/shoulder plans only |
| **Best Solar FIT** | Plans with solar feed-in tariffs, sorted highest first |
| **Cheapest Plans** | Top 50 cheapest plans by estimated yearly cost |
| **Plan Calculator** | Interactive calculator - enter your usage and solar export to compare costs |

### Plan Calculator

The Plan Calculator sheet lets you model actual costs based on your household. All cells use live Excel formulas, so changing any input recalculates everything instantly.

**Inputs** (yellow cells):
- **Daily Usage (kWh)** - your electricity consumption per day
- **Daily Solar Export (kWh)** - how much solar you export to the grid
- **Usage Profile** - how usage splits across peak/off-peak (dropdown: Flat, Slight Peak, Heavy Peak, Off-Peak Heavy, Battery Optimised)
- **Controlled Load** - toggle Yes/No if you have a CL circuit
- **Controlled Load Usage (kWh/day)** - daily kWh on your CL circuit

---

## Plan Filtering

The API returns **all** plans for a postcode, including plans most residential customers can't use. By default, the scraper filters these out to match the Energy Made Easy website's behaviour:

- **Demand charge plans** are excluded (require a demand meter most homes don't have)
- **Controlled load plans** are excluded unless you pass `--controlled-load`

Use `--no-filter` to get everything unfiltered.

---

## Distributor Selection

Some postcodes are served by multiple electricity distributors (e.g. boundary postcodes). The scraper will detect this and let you choose. If you're unsure which distributor serves your address, check your electricity bill for the "network provider" or "DNSP".

---

## Building a Standalone Executable

A PyInstaller spec file is included to build a single `.exe` for Windows:

```bash
pip install pyinstaller pywin32
pyinstaller EnergyPlanScraper.spec
```

The executable will be created in the `dist/` folder.

---

## Notes

- All rates are shown **including 10% GST** (matching the Energy Made Easy website). Solar feed-in tariffs are GST exempt.
- Estimated annual costs come from the Energy Made Easy API's benchmark usage profiles.
- The scraper uses the same public API the website calls. No browser automation or HTML scraping is involved.
- NMI-based personalised results are not supported (the NMI API requires browser-session authentication).
