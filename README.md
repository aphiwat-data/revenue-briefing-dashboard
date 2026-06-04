<div align="center">

# Daily Revenue Briefing Dashboard

### A production-grade Streamlit application for hospitality revenue management

[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.57-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://streamlit.io/)
[![Pandas](https://img.shields.io/badge/Pandas-2.x-150458?style=for-the-badge&logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![Plotly](https://img.shields.io/badge/Plotly-5.x-3F4F75?style=for-the-badge&logo=plotly&logoColor=white)](https://plotly.com/)
[![OpenPyXL](https://img.shields.io/badge/openpyxl-Excel%20I%2FO-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://openpyxl.readthedocs.io/)

**Turns 30+ daily Duetto D4cast exports into a 5-page operational dashboard with a single-click morning briefing.**

[Architecture](#architecture)  [Features](#features)  [Tech Stack](#tech-stack)  [Project Structure](#project-structure)

</div>

---

## The Problem

Hotel revenue management teams spent **20–30 minutes every morning** manually:
- Opening each property's Duetto export individually
- Calculating variances (Forecast vs Budget, OTB vs STLY, …) in Excel
- Switching between 7+ workbooks to compare hotels
- Re-formatting tables for the morning briefing meeting

This dashboard collapses that whole workflow into **5 minutes** with auto-loaded files, color-coded variance tables, and a one-click Excel briefing export.

---

## Impact

| Metric | Before | After |
| :--- | :--- | :--- |
| Morning briefing prep time | ~30 min | **< 5 min** |
| Files to open daily | 7+ workbooks | **1 dashboard** |
| Variance calculations | Manual in Excel | **Auto-computed** |
| Cross-hotel comparisons | Eyeballing | **Sortable leaderboard** |
| Morning briefing deck | Manual copy-paste | **1-click 8-sheet Excel** |

---

## Tech Stack

<table>
<tr>
<td><strong>Frontend / UI</strong></td>
<td>Streamlit 1.57+ (page navigation via <code>st.pills</code>, sticky-column dataframes, slot-reservation pattern)</td>
</tr>
<tr>
<td><strong>Data Engine</strong></td>
<td>Pandas (long-format reshaping, drop-duplicates snapshots, multi-index pivoting, groupby aggregations)</td>
</tr>
<tr>
<td><strong>Visualization</strong></td>
<td>Plotly Graph Objects (combo charts, dual Y-axis, secondary axes, spline + filled area, custom hovertemplates)</td>
</tr>
<tr>
<td><strong>File I/O</strong></td>
<td>openpyxl (multi-sheet Excel writer, in-memory ZIP expansion, CSV / XLS / XLSX parsing)</td>
</tr>
<tr>
<td><strong>Styling</strong></td>
<td>Pandas Styler (color-coded variance cells), custom CSS (sticky columns, brand theming)</td>
</tr>
<tr>
<td><strong>Deployment</strong></td>
<td>Streamlit Community Cloud (auto-redeploy on git push)</td>
</tr>
<tr>
<td><strong>Language</strong></td>
<td>Python 3.10+ (type-hinted where it matters)</td>
</tr>
</table>

---

## Features

### Five Pages, One Source of Truth

| Page | What it does |
| :--- | :--- |
| **Overview** | At-a-glance KPI snapshot · Compare chart · Variance Pivot · Same-Time Pace · Historical Finals · Forecast Movement |
| **Budget Review** | Priority table sorted by worst variance — green/red/yellow color-coded with sticky context columns |
| **Leaderboard** | 7 hotels × 10 metrics, ranked per stay month, with gold/silver/bronze tints for top 3 and blue highlights for selected hotels |
| **Trend** | Dual-Y combo chart (bars + line + baselines) with `vs Forecast` / `vs Budget` toggle · Daily % momentum |
| **Export** | One-click **8-sheet Excel briefing** + quick CSV exports, with auto-named files (`g5_daily_briefing_<month>.xlsx`) |

### Engineering Highlights

- **Smart File Role Detection** — Auto-classifies uploads as Today / Yesterday / Last-7D / 1st-of-Month without configuration
- **Look-back Logic** — `drop_duplicates(["Hotel","Stay Month","Metric","Reference"], keep="last")` pattern lets past stay months pull from their own latest report, enabling month-over-month analysis
- **Generic Variance Styler** — Any column matching `" VS "` is auto-color-coded — add new variance columns without touching the styler
- **Variance Pivot with 7 columns** — `Today VS BUD`, `Duetto VS BUD`, `Today VS Duetto`, `Today VS STLY/ST2Y/ST3Y`, `Duetto VS Final LY/2Y/3Y`
- **Sticky Columns** — Stay Month + Metric pinned left when the variance pivot scrolls horizontally
- **Status Badge** — Top-right freshness indicator showing report date and file count, with auto-detection of today's data

---

## Architecture

```text
                    ┌─────────────────────────────────────────────┐
   Duetto Reports   │  Upload Mode  ── ZIP expansion (in memory)  │
   (CSV/XLSX)  ──→  │  Folder Mode  ── Path scanning              │
                    └─────────────────────────────────────────────┘
                                       │
                                       ▼
                    ┌─────────────────────────────────────────────┐
                    │  File Catalog  ── auto-detect report date   │
                    │                    from filename patterns   │
                    └─────────────────────────────────────────────┘
                                       │
                                       ▼
                    ┌─────────────────────────────────────────────┐
                    │  Role Selection  ── Today/Yesterday/7D/1st  │
                    │                     + look-back month merge │
                    └─────────────────────────────────────────────┘
                                       │
                                       ▼
                    ┌─────────────────────────────────────────────┐
                    │  Parser  ── pattern-match column names →    │
                    │             canonical (Today, Duetto, …)    │
                    └─────────────────────────────────────────────┘
                                       │
                                       ▼
                    ┌─────────────────────────────────────────────┐
                    │  metric_long DataFrame                      │
                    │  (Hotel × Stay Month × Metric × Reference)  │
                    └─────────────────────────────────────────────┘
                                       │
        ┌──────────────────┬──────────────────┬──────────────────┐
        ▼                  ▼                  ▼                  ▼
  ┌─────────────┐  ┌────────────────┐  ┌─────────────┐  ┌─────────────┐
  │  Overview   │  │ Budget Review  │  │ Leaderboard │  │   Export    │
  │  (Pivot +   │  │   (Priority    │  │ (Mega rank  │  │  (8-sheet   │
  │   KPIs)     │  │    table)      │  │  table)     │  │   Excel)    │
  └─────────────┘  └────────────────┘  └─────────────┘  └─────────────┘
```

### Key Design Patterns

- **Layered modular structure** — `core/` (constants & helpers) → `data/` (ingestion) → `domain/` (business logic) → `features/` (page rendering) → `ui/` (styling); `app.py` is a thin orchestrator
- **Long-format core** — All filters and aggregations operate on a uniform `(Hotel, Stay Month, Metric, Reference, Value)` long-format DataFrame
- **Latest-snapshot pattern** — `drop_duplicates(..., keep="last")` over a sorted Report Date is applied at every render to support look-back across months
- **Display-layer rename** — Internal column names stay stable for styling logic; user-facing labels rename at `st.dataframe` time via `.rename()` + parallel styler
- **Generic detection** — Variance styler auto-discovers VS-pattern columns; no per-column wiring needed
- **Slot-reservation pattern** — Title bar uses `st.empty().container()` to reserve top-of-page space and fill it later once sidebar has computed the data

---

## Project Structure

```text
.
├── app.py                          # Thin orchestrator (~1,000 LOC) — main UI flow only
├── assets/
│   └── logo.png                    # Brand logo (base64-encoded at startup)
├── requirements.txt
├── README.md
└── src/
    ├── core/                       # Constants, formatters, shared helpers
    │   ├── assets.py
    │   ├── constants.py            # REFERENCE_PATTERNS, METRIC_PATTERNS, METRIC_ORDER
    │   ├── helpers.py              # fmt_*, safe_delta, month_sort_key, date helpers
    │   └── page.py
    ├── data/                       # File ingestion + schema normalization
    │   └── ingest.py               # Catalog builders, role detection, parser pipeline
    ├── domain/                     # Pure business logic (no Streamlit calls)
    │   ├── helpers.py              # Stay-month filter, budget variance, status mapping
    │   ├── aggregations.py         # build_movement_summary, build_pace_summary, …
    │   ├── pivot.py                # build_latest_pivot_table, build_variance_pivot_table
    │   ├── budget.py               # build_budget_review + summary view + movement
    │   └── briefing.py             # Portfolio snapshot, hotel scorecard, 8-sheet builder
    ├── services/                   # External integrations
    │   └── excel_export.py         # to_excel_bytes (multi-sheet writer)
    ├── ui/                         # Streamlit UI primitives
    │   ├── styles.py               # CUSTOM_CSS + inject_styles
    │   ├── sidebar.py              # Sidebar helpers
    │   └── dataframe_stylers.py    # Pandas Styler functions for variance tables
    └── features/                   # One module per page
        ├── overview.py             # Variance Pivot, comparison charts, snapshot KPIs
        ├── leaderboard.py          # Mega Leaderboard (10-column ranking)
        ├── budget_review.py        # Priority budget table
        ├── trend.py                # Trend Comparison combo + momentum charts
        ├── movement.py             # Forecast Movement table
        └── shared.py               # Compact tables, pace cards (shared by pages)
```

### Module Responsibility Matrix

| Layer | Reads from | Writes to | Streamlit calls? |
| :--- | :--- | :--- | :--- |
| `core/` | — | — | No |
| `data/` | Files | `metric_long` DataFrame | No |
| `domain/` | `metric_long` | Derived DataFrames | No |
| `services/` | DataFrames | Excel bytes | No |
| `ui/` | Styling rules | CSS / Stylers | No |
| `features/` | All of the above | Streamlit components | **Yes** |
| `app.py` | `features/`, `ui/` | Final rendered app | **Yes** (orchestration) |

This separation means **domain logic is testable in isolation** without Streamlit, and **adding a new page is a single new module** in `features/`.

---

## Skills Demonstrated

| Domain | Skills |
| :--- | :--- |
| **Python** | Pandas (multi-index, pivot_table, groupby, merge), Plotly graph_objects + subplots, OpenPyXL multi-sheet writer, BytesIO streaming, ZIP file handling |
| **Data Engineering** | Schema normalization across heterogeneous Duetto exports, idempotent file role detection, deduplication strategies (`keep="last"`), look-back data pattern |
| **Software Architecture** | Layered modular package (`core`/`data`/`domain`/`services`/`ui`/`features`), separation of pure logic from Streamlit calls, slot-reservation pattern, thin-orchestrator main module |
| **Visualization** | Combo charts with secondary Y-axis, custom hover templates, color-coded cell styling via Pandas Styler, sticky-column dataframes |
| **Domain (Hospitality)** | Revenue management metrics (OTB / Forecast / Budget / STLY / Final), pickup analysis, variance attribution, pace benchmarking |
| **Engineering Practice** | Dead-code elimination (removed ~1,700 lines of legacy code), bug post-mortems with root-cause docs, naming conventions for export safety (avoiding non-ASCII symbols), comprehensive README |

---

## Quick Start

### Local Run

```bash
git clone <repo-url>
cd <project>
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

The app opens at `http://localhost:8501`.

### Streamlit Cloud

- Entrypoint: `app.py`
- Python: `3.10+`
- Requirements: `requirements.txt`

Push to `main` → Streamlit Cloud auto-redeploys.

### Using the Dashboard

1. **Upload mode** (cloud): drag-and-drop daily Duetto exports (CSV / XLSX / ZIP)
2. **Folder mode** (local): point at the daily report folder
3. Pick **Properties** and **Stay Month range** at the top
4. Switch pages via the pills navigation
5. Export the **Daily Briefing Excel** from the Export tab

---

## Naming Convention

| Term | Meaning |
| :--- | :--- |
| **Today** | On-the-book (OTB) — what's already booked |
| **Duetto** | Forecast — Duetto's predicted landing |
| **Budget** / **BUD** | Locked target for the period |
| **VS** | Comparison operator — export-safe (no `▲` symbols that break CSV/Excel imports) |
| **STLY / ST2Y / ST3Y** | Same-time last year / 2 years ago / 3 years ago |
| **Final LY / 2Y / 3Y** | Historical full-period actuals |

### Variance Formula

For every `X VS Y` column:

```python
variance_pct = (X - Y) / abs(Y) * 100
```

- Green = positive (ahead of target / pace)
- Red = negative (behind)
- Yellow = flat

---

## Code Quality

| Metric | Value |
| :--- | :--- |
| Total project LOC | ~4,600 |
| `app.py` LOC | ~1,000 (thin orchestrator) |
| Number of Python modules | 21 |
| Largest single module | `features/trend.py` (~770 LOC) |
| Top-level functions in `app.py` | 1 (`main()`) — everything else lives in modules |
| Test coverage | Roadmap — `pytest` planned |
| Compile-clean | `python -m py_compile` passes on every commit |

---

## Troubleshooting

| Symptom | Likely cause |
| :--- | :--- |
| "No Forecast / Duetto columns found" | Source files don't match `REFERENCE_PATTERNS` — extend the pattern list |
| Empty Stay Month dropdown | Uploaded files filtered out — check filenames and date parsing |
| Past stay month shows "No data" | Previous month's report files not in catalog — upload them too |
| Variance cell uncolored | Column name missing `" VS "` — the styler won't auto-detect it |
| `git push` rejected (non-fast-forward) | Use `git push --force-with-lease origin main` after verifying local has the right code |

---

## License

This is a portfolio project. The code is shared for demonstration purposes.
Refer to the LICENSE file for terms.

---

<div align="center">

**Built with Python, Pandas, Streamlit, and Plotly.**

</div>
