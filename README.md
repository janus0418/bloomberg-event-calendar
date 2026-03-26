# Bloomberg Market Event Calendar

A dark-themed, interactive Jupyter notebook calendar for tracking economic releases, large-cap earnings, commodity futures timing, market holidays, and high-priority macro events. Built for Bloomberg's BQuant environment with hedge fund-grade UX.

## Features

- **Economic Calendar** - Central bank decisions, employment data, CPI, GDP, PMI, and 600+ key releases across 15+ countries
- **Earnings Tracker** - Top 250 market-cap companies with expected report dates
- **Commodity Futures** - Options/futures expiry schedule with delivery dates and live pricing
- **Market Holidays** - Embedded closure schedule for US, UK, EU, Asia-Pacific, and LatAm markets
- **Key Event Rules** - Auto-highlights high-impact releases (FOMC, NFP, CPI, etc.) via embedded importance rules

### Hedge Fund UX Enhancements

- **Next Key Event Countdown** - Prominent T-minus countdown to the next important release
- **Beat/Miss Indicators** - Instant visual for actual vs expected (green beat / red miss chips)
- **Quick Filters** - One-click: Today, This Week, Next 7 Days, Key Events Only, Econ Only
- **Calendar Heatmap** - Day cells colored by event density with priority markers
- **Priority Scoring** - 1-5 star system based on importance rules, flags, and data availability
- **Category Color Coding** - Teal=Economic, Purple=Earnings, Pink=Commodity, Green=Holiday

## Requirements

- **Bloomberg BQuant** notebook environment (for live data via BQL)
- Python 3.8+
- `ipywidgets`
- `pandas`
- `IPython`

### Optional (for export features)
- `python-docx` (Word export)
- `openpyxl` (Excel export)

## Usage

### In Bloomberg BQuant

Paste the entire `bloomberg_event_calendar.py` into a single notebook cell and run it.

The app will:
1. Render the full interactive calendar UI
2. Look for existing `df_ECO` and `df_EQ_Earnings` DataFrames in the notebook scope
3. If found, load them automatically
4. If not, click **Refresh Data** to pull live data via BQL

### Standalone (without Bloomberg)

The UI renders fully without Bloomberg. To pre-populate with sample data:

```python
import pandas as pd

# Cell 1: Create empty DataFrames (or load from CSV)
df_ECO = pd.DataFrame()
df_EQ_Earnings = pd.DataFrame()

# Cell 2: Paste bloomberg_event_calendar.py
```

The following features work without BQL:
- Full UI rendering and navigation
- Embedded holiday schedule
- Hardcoded commodity expiry schedule
- State persistence (flags, notes, important markers)
- All filtering, searching, and calendar interactions

## Architecture

The app uses a monkey-patch architecture with 7 progressive enhancement layers:

1. **Base** - Core `BloombergEventCalendarApp` class with BQL integration
2. **Patched** - Enhanced event rendering, badges, and field display
3. **Alarm** - Flagged-event alarm monitor with JavaScript countdown
4. **Fast Explorer** - Lazy-rendered tabular event browser
5. **Stable Explorer** - Fixed-width table with pagination
6. **Text Explorer** - Text-based tab explorer
7. **Enhanced/Custom** - Custom event form, recovery mode, hedge fund UX

## Data Sources

| Source | Provider | Update Method |
|--------|----------|--------------|
| Economic releases | BQL `calendar(type='ECONOMIC_RELEASES')` | Live refresh |
| Central bank decisions | BQL `calendar(type='central_banks')` | Live refresh |
| Earnings dates | BQL `expected_report_dt()` | Live refresh |
| Commodity futures | BQL `univ.options/futures` | Live refresh |
| Market holidays | Embedded TSV schedule | Static (update annually) |
| Key event rules | Embedded importance table | Static |
| Commodity expiry | Embedded hardcoded schedule | Static |

## State Persistence

Annotations (flags, notes, important markers) auto-save to `event_annotations.csv` in the notebook's working directory. State is preserved across notebook restarts.

## License

MIT
