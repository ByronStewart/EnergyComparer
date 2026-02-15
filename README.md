# Energy Made Easy Scraper

A Python command-line tool that scrapes energy plan data from the Australian government's [Energy Made Easy](https://www.energymadeeasy.gov.au/) website and exports it to a formatted Excel spreadsheet for comparison.

Covers postcodes in **NSW, QLD, SA, TAS, and the ACT**.

## Requirements

- Python 3.10+

## Setup

1. Clone or download this repository.

2. Create a virtual environment and install dependencies:

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate

pip install -r requirements.txt
```

## Quick Start (Windows)

Double-click **`run_scraper_enhanced.bat`** to run the scraper with distributor selection.

The batch file will create the virtual environment and install dependencies automatically if needed.

## Plan Filtering

The Energy Made Easy API returns **all** plans for a postcode, including plans that most residential customers can't use. For example, postcode 4075 returns 1,283 plans from the API, but the Energy Made Easy website only shows 315 of them by default.

The scraper replicates the **exact same client-side filtering** that the website applies:

1. **Demand charge plans are excluded** - These plans require a demand meter/tariff that standard residential customers don't have. They're designed for customers with specific network tariff arrangements (e.g. Energex demand tariff codes 3970, 3750, etc.).

2. **Controlled load plans are excluded by default** - These plans include pricing for a separate controlled load circuit (off-peak hot water, pool pump, floor heating). If you have a controlled load circuit, use `--controlled-load` to include them.

### Filtering Options

| Flag | Effect |
|---|---|
| *(default)* | Excludes demand charge plans and controlled load plans |
| `--controlled-load` | Includes controlled load plans (for customers with CL circuits) |
| `--no-filter` | Disables all filtering; returns every plan from the API |

In interactive mode (no `--controlled-load` or `--no-filter` flag), the scraper will ask whether you have a controlled load circuit.

## Usage

```bash
# Interactive mode - prompts for postcode and distributor
python scraper_enhanced.py

# Postcode with a single distributor (behaves like basic scraper)
python scraper_enhanced.py 2000

# Multi-distributor postcode - interactive distributor selection
python scraper_enhanced.py 2850

# Specify a distributor ID directly (skip the prompt)
python scraper_enhanced.py 2850 --dist 13

# Fetch plans for ALL distributors at once
python scraper_enhanced.py 2850 --dist all

# Combine with other options
python scraper_enhanced.py 2850 --dist 13 --fuel electricity --type residential
```

### Arguments

| Argument | Required | Default | Description |
|---|---|---|---|
| `postcode` | No (prompted if omitted) | - | 4-digit Australian postcode |
| `--fuel` | No | `electricity` | Fuel type: `electricity` or `gas` |
| `--type` | No | `residential` | Customer type: `residential` or `business` |
| `--dist` | No | Interactive | Distributor ID (e.g. `13`) or `all` for every distributor |
| `--controlled-load` | No | Off | Include plans with controlled load circuits |
| `--no-filter` | No | Off | Disable all plan filtering |

### How Distributor Selection Works

Some postcodes sit on the boundary between two electricity distributors. For example, postcode **2850 (Mudgee)** is split between **Endeavour** (ID 13) and **Essential Energy Far West** (ID 39). Different distributors serve different streets, so the available plans and pricing depend on which distributor serves your address.

When you enter a boundary postcode, the scraper will:

1. Query the API for all distributors registered for that postcode
2. Probe each distributor to check which ones actually have plans available
3. Show you the options with plan counts
4. Let you pick one, or fetch all at once

If your postcode has only one distributor, it proceeds automatically.

**How to find your distributor:** Check your electricity bill - it will list your distributor (also called your "network provider" or "DNSP"). Common NSW distributors include Ausgrid, Endeavour Energy, and Essential Energy.

### Example (Multi-Distributor)

```
$ python scraper_enhanced.py 2850

Step 2: Determining electricity distributor(s)...

  Checking available electricity distributors...
  Found 4 distributors. Checking plan availability...
    [1] Endeavour (ID: 13) - 1176 plans
    [ ] Essential Energy Standard (ID: 40) - no plans (skipped)
    [ ] Essential Energy (ID: 41) - no plans (skipped)
    [2] Essential Energy Far West (ID: 39) - 1478 plans

  This postcode has 2 active distributors.
  Your electricity distributor depends on your street address.
  Check your electricity bill or contact your provider if unsure.

  Select a distributor:
    [1] Endeavour (1176 plans)
    [2] Essential Energy Far West (1478 plans)
    [A] Fetch ALL distributors (2654 plans total)

  Enter your choice (number or A): 1
```

## Output

The scraper produces an Excel file named `energy_plans_{postcode}_{fuel}_{timestamp}.xlsx` in the current directory. The distributor name is included in the filename when a specific distributor is selected.

### Spreadsheet Sheets

| Sheet | Description |
|---|---|
| **Summary** | Search parameters, date, total plan count, breakdown by retailer (and by distributor if multiple) |
| **All Plans** | Every plan returned for the postcode, unfiltered |
| **Single Rate Plans** | Plans with a flat (single) usage rate only |
| **Time of Use Plans** | Plans with peak/off-peak/shoulder pricing |
| **Best Solar FIT** | All plans with a solar feed-in tariff, sorted highest first |
| **Cheapest Plans** | Top 50 cheapest plans by estimated yearly cost (medium usage, with discounts) |
| **Plan Calculator** | Interactive cost calculator - enter your daily usage and solar export to compare plans (see below) |

### Plan Calculator Sheet

The **Plan Calculator** sheet lets you model the actual daily and monthly cost of each plan based on your household's usage and solar generation. All calculations use live Excel formulas, so changing any input instantly recalculates every plan.

**User Inputs** (yellow cells at the top of the sheet):

| Cell | Input | Default | Description |
|---|---|---|---|
| B4 | Daily Usage (kWh) | 20 | Your estimated daily electricity consumption |
| B5 | Daily Solar Export (kWh) | 10 | How much solar you expect to export to the grid per day |
| B6 | Usage Profile | Flat Usage | How your usage is split across peak/off-peak periods (TOU plans only) |

**Usage Profile Presets** (dropdown in B6):

| Profile | Peak / Off-Peak | Use Case |
|---|---|---|
| Flat Usage | 50% / 50% | Even usage throughout the day |
| Slight Peak | 60% / 40% | Slightly more daytime usage |
| Heavy Peak | 75% / 25% | Most usage during peak hours |
| Off-Peak Heavy | 30% / 70% | Night owl or shifted usage pattern |
| Battery Optimised | 10% / 90% | Battery stores solar, discharges at peak |

**Calculated Columns** (per plan):

| Column | Description |
|---|---|
| Plan Name | Full plan name |
| Retailer | Energy company name |
| Tariff Type | SR (single rate) or TOU (time of use) |
| Plan URL | Clickable link to the plan on Energy Made Easy |
| Supply (c/day) | Daily supply charge |
| Usage Rate (c/kWh) | The usage rate applied (single rate plans only) |
| Peak Rate (c/kWh) | Peak rate (TOU plans only) |
| Off-Peak Rate (c/kWh) | Off-peak rate (TOU plans only) |
| Solar FIT first tier (c/kWh) | First-tier solar feed-in tariff rate (or the flat rate if no tiers) |
| Solar FIT thereafter (c/kWh) | Second-tier rate applied after the first-tier volume cap is reached |
| Solar FIT Details | Text description of the full FIT structure |
| Peak % | Percentage of usage at peak rate (from selected profile; 100% for single rate) |
| Off-Peak % | Percentage of usage at off-peak rate (from selected profile; 0% for single rate) |
| Usage Cost/day (c) | Daily usage cost in cents |
| Solar Credit/day (c) | Daily solar export credit in cents (tiered calculation) |
| Net Cost/day (c) | Supply + Usage - Solar Credit, in cents |
| Net Cost/month ($) | Net daily cost multiplied by 30.44 days, converted to dollars |

**How it works:**

- **Single rate plans:** `Usage Cost = Daily Usage * Usage Rate`
- **TOU plans:** `Usage Cost = Daily Usage * (Peak% * Peak Rate + Off-Peak% * Off-Peak Rate)`
- **Solar credit (flat FIT):** `Solar Credit = Daily Export * FIT Rate`
- **Solar credit (tiered FIT):** `Solar Credit = MIN(Export, Cap) * Tier1 Rate + MAX(Export - Cap, 0) * Tier2 Rate`
  - Example: 10c/kWh first 8kWh + 3c/kWh thereafter with 10kWh export = `MIN(10,8)*10 + MAX(10-8,0)*3 = 86c`
- **Net cost:** `Supply Charge + Usage Cost - Solar Credit`

Plans are sorted by supply charge (ascending) as a neutral default. Change the input cells to model different scenarios and use Excel's built-in sort/filter on the table to find the best plan for your situation.

### Columns

Each plan row contains 28 data columns. All numeric columns are stored as actual numbers in Excel so you can sort and filter on them directly.

| Column | Type | Description |
|---|---|---|
| Plan ID | Text | Unique identifier from Energy Made Easy |
| Plan Name | Text | Full plan name |
| Retailer | Text | Energy company name |
| Distributor | Text | Electricity distributor/network |
| Plan URL | Link | Clickable link to the plan's detail page on Energy Made Easy |
| Tariff Type | Text | `SR` (single rate), `TOU` (time of use), `SRCL`/`TOUCL` (with controlled load) |
| Pricing Model | Text | `SR` or `TOU` |
| Contract Term | Text | Lock-in period (e.g. "No lock-in", "1 year") |
| Benefit Period | Text | How long any benefit/discount lasts |
| Supply Charge (c/day) | Number | Daily supply charge in cents, inc. GST |
| Usage Rate Min (c/kWh) | Number | Lowest usage rate in cents per kWh, inc. GST |
| Usage Rate Max (c/kWh) | Number | Highest usage rate in cents per kWh, inc. GST |
| Peak Rate (c/kWh) | Number | Peak rate for TOU plans, inc. GST (blank for single rate) |
| Off-Peak Rate (c/kWh) | Number | Off-peak rate for TOU plans, inc. GST (blank for single rate) |
| Solar FIT Min (c/kWh) | Number | Lowest solar feed-in tariff rate (GST exempt) |
| Solar FIT Max (c/kWh) | Number | Highest solar feed-in tariff rate (GST exempt) |
| Solar FIT Details | Text | Tiered feed-in tariff breakdown (e.g. "5c/kWh first 10kWh/day; 1c/kWh") |
| Controlled Load | Text | Controlled load rates and supply charges, if applicable |
| Discounts | Text | Available discounts and their percentages |
| Fees | Text | Connection, disconnection, late payment, and other fees |
| Payment Options | Text | Accepted payment methods (Direct Debit, BPay, Credit Card, etc.) |
| Meter Types | Text | Compatible meter types (Basic, Smart, etc.) |
| Est. Cost/Year (Low Usage) | Number | Estimated annual cost for a 1-person household, with discounts |
| Est. Cost/Year (Medium Usage) | Number | Estimated annual cost for a 2-3 person household, with discounts |
| Est. Cost/Year (High Usage) | Number | Estimated annual cost for a 4+ person household, with discounts |
| Est. Cost/Year (Low, No Disc.) | Number | Estimated annual cost, low usage, without discounts |
| Est. Cost/Year (Medium, No Disc.) | Number | Estimated annual cost, medium usage, without discounts |
| Est. Cost/Year (High, No Disc.) | Number | Estimated annual cost, high usage, without discounts |

## Notes

- All supply charges and usage rates are displayed **including 10% GST**, matching the values shown on the Energy Made Easy website.
- Solar feed-in tariffs are **GST exempt** and shown as-is.
- Estimated annual costs are provided by the Energy Made Easy API based on benchmark usage profiles for the postcode's region.
- The scraper uses the same public API that the Energy Made Easy website itself calls. No browser automation or HTML parsing is required.
- Plan filtering matches the Energy Made Easy website's default behaviour: demand charge plans and controlled load plans are excluded unless explicitly included. This was reverse-engineered from the website's client-side JavaScript filtering logic.

## NMI Support

NMI (National Meter Identifier) based personalised results are **not currently supported**. The Energy Made Easy NMI API requires browser-session authentication that cannot be replicated via simple HTTP requests. If NMI support is needed in the future, it would require browser automation (e.g. Playwright or Selenium).
