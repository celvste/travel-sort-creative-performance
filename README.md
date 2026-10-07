# applovin-ua

Weekly AppLovin creative-performance pull for Travel Sort UA, plus a workbook-style dashboard.

## Files

| File | What it does |
|---|---|
| `applovin_week.py` | Pulls asset-level spend, impressions and clicks from AppLovin's Asset Reporting API and writes one weekly tab per campaign to `weeks/<week>/`, in the master-workbook format. |
| `build_dashboard.py` | Bundles the master workbooks, weekly pulls and Total sheet exports into `dashboard.html`. |
| `dashboard_template.html` | The dashboard page: campaign first, then a Total tab and one tab per week, with week-over-week changes and a Total / Video / Playable toggle. |
| `config.example.json` | Template for `config.json`. |

## Setup

```bash
pip install openpyxl
cp config.example.json config.json   # then add your report key and campaign IDs
```

The report key can also come from the `APPLOVIN_REPORT_KEY` environment variable. Run `python3 applovin_week.py --list-campaigns` to find campaign IDs.

## Weekly run

```bash
python3 applovin_week.py      # last completed Wed–Tue week, every campaign
python3 build_dashboard.py    # rebuild dashboard.html
```

AppLovin data is final after 06:00 UTC, and the API only reaches back 45 days. See the docstring at the top of `applovin_week.py` for all options.

## Local folders (not in the repo)

These hold spend data and are listed in `.gitignore`:

- `workbooks/`: master workbooks (`[Travel Sort] D28 <campaign> Applovin Creative Performance by week.xlsx`). Each one's weekly tabs feed the dashboard, and its Total tab is rebuilt from them. Master_Data is only checked against the weekly tabs.
- `weeks/`: output of `applovin_week.py`. Used for any week a master workbook doesn't have yet.
- `totals/`: Total sheet CSV exports. They give `applovin_week.py` its Added Dates, and give the dashboard its Total tab when a campaign has no master workbook.

## Added Date

Added Date is per campaign: when the creative was added to that campaign, not when it was uploaded to AppLovin. When a creative is missing from the Total sheet, `applovin_week.py` fills in the first day it served in that campaign (blue cell). That's the first delivery date, so it can be later than the actual add date.
