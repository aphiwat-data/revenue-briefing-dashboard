# G5 D4cast Revenue Dashboard

Daily revenue briefing dashboard for ATMIND Group hotels.
Parses Duetto D4cast exports (CSV / Excel) and turns them into a one-screen
operations view: snapshot KPIs, variance pivot, leaderboard, trend, and
one-click Excel briefing for the morning meeting.

---

## Quick start

### Local
```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

### Streamlit Cloud
- Entrypoint: `app.py`
- Python: 3.10+
- Dependencies: `requirements.txt`

Once deployed, the dashboard auto-detects the latest report on every refresh
and renders the briefing view.

---

## Data input

Drop daily Duetto D4cast exports — the parser accepts:

| Extension | Notes |
|---|---|
| `.csv` | Title/Generated header rows expected at top |
| `.xlsx` / `.xls` | Auto-detects header row by looking for "Hotel" + "Stay Month" |
| `.zip` | Expanded in memory; nested CSV/XLSX extracted automatically |

The catalog auto-detects the **report date** from the filename
(patterns supported: `YYYY-MM-DD`, `YYYYMMDD`, `YYYY_MM_DD`, `D-M-YYYY`).

### Two source modes

| Mode | Use when |
|---|---|
| **Folder** | Local presentation — point to the folder of daily exports |
| **Upload** | Cloud / share — drag-and-drop multiple files or a single ZIP |

> For Streamlit Cloud, use **Upload mode** (local paths like `G:\My Drive\...`
> are not visible on the cloud server).

---

## Page structure

```
┌────────────────────────────────────────────────────────────────────┐
│ Revenue Briefing               ● TODAY'S DATA Mon, 01 Jun 2026     │
│                                  Report month Jun, 2026 · 35 files │
├────────────────────────────────────────────────────────────────────┤
│ [ Overview ] Budget Review  Leaderboard  Trend  Export             │
├────────────────────────────────────────────────────────────────────┤
│ PROPERTIES        STAY MONTH FROM       STAY MONTH TO              │
│ [multi-select]    [May, 2026 ▾]         [Aug, 2026 ▾]              │
├────────────────────────────────────────────────────────────────────┤
│ ...page content...                                                 │
└────────────────────────────────────────────────────────────────────┘
```

### Pages

| Page | What it shows |
|---|---|
| **Overview** | Compare chart · Forecast vs Budget KPI · Variance Pivot · Same-Time Pace · Historical Final · Forecast Movement |
| **Budget Review** | Priority budget table — worst variance first, color-coded |
| **Leaderboard** | All hotels ranked per stay month with 10 columns (ADR/Occ/Rev × Today/Duetto + variances). Selected hotels highlighted blue. |
| **Trend** | Trend Comparison combo chart (bars + line, dual Y) · Forecast trend by stay month · Daily % change momentum · Hotel momentum bubble |
| **Export** | 1-click Daily Briefing Excel (8 sheets) + quick CSVs |

### Filters

- **Properties** — multi-select; default = Altera properties
- **Stay Month From / To** — default = report month − 1 → report month + 2 (4-month window)
- Filters live on the main page (not in the sidebar) so they apply to every tab

### Look-back

The dashboard auto-loads the previous month's report files in addition to
the current month's. Each (Hotel × Stay Month × Metric × Reference)
combination uses its **latest available report row** — so when looking at
a past stay month, the data comes from that month's last daily report,
not the current month's.

---

## Variance pivot (Overview)

Single table with raw values **and** color-coded variance % for every key
comparison:

```
Hotel | Stay Month | Metric
| Today | Budget | Today VS BUD
| Duetto | Duetto VS BUD | Today VS Duetto
| STLY | Today VS STLY | ST2Y | Today VS ST2Y | ST3Y | Today VS ST3Y
| Final LY | Duetto VS Final LY | Final 2Y | Duetto VS Final 2Y | Final 3Y | Duetto VS Final 3Y
```

### Naming convention

| Term | Meaning |
|---|---|
| **Today** | On-the-book (OTB) — what's already booked |
| **Duetto** | Forecast — Duetto's predicted landing |
| **Budget** / **BUD** | Locked target for the period |
| **VS** | Comparison symbol — export-safe (no `▲` arrows that break CSV/Excel) |
| **STLY / ST2Y / ST3Y** | Same-time last year / 2 years ago / 3 years ago |
| **Final LY / 2Y / 3Y** | Historical full-period actuals |

### Variance formula

For every `X VS Y` column:

```
variance % = (X − Y) / |Y| × 100
```

Cells are color-coded:
- 🟢 green = positive variance (above target / ahead of pace)
- 🔴 red = negative variance (below target / behind pace)
- 🟡 yellow = zero / flat

### Important: each `X VS STLY/ST2Y/ST3Y` uses its OWN year

`Today VS STLY` compares Today against **STLY only** (not the highest of
STLY/ST2Y/ST3Y). This avoids sign-flip surprises when prior years had
unusually strong / weak comparable periods.

If you need to see the best of all three benchmarks, look at all three
columns side-by-side rather than relying on a hidden `max()`.

---

## Mega Leaderboard

Per stay month (CM / M+1 / M+2 / …) — all hotels ranked side-by-side.

| Column | Source |
|---|---|
| ADR 1st | 1st-of-month report (anchor) |
| ADR Today | Today's OTB ADR |
| ADR Duetto | Duetto forecast ADR |
| Occ 1st / Today / Duetto | Same pattern, Occupancy |
| ADR Today VS BUD % | Today's ADR vs ADR Budget |
| ADR Duetto VS BUD % | Forecast ADR vs ADR Budget |
| Rev Today VS BUD (K) | Today's Rev vs Budget (thousands) |
| Rev Duetto VS BUD (K) | Forecast Rev vs Budget (thousands) |

- Top 3 ranks tinted gold / silver / bronze
- Selected hotels (from the Properties filter) highlighted in blue
- Default sort: **Rev Duetto VS BUD (K)** (biggest revenue gap first)

---

## Trend Comparison combo chart

Dual-Y combo chart, Duetto reference style:

| Series | Visual | Axis |
|---|---|---|
| Bars metric (default = Rev) | Filled green area + spline line | Left Y |
| Line metric (default = ADR) | Smooth amber spline | Right Y |
| Bars baseline | Dashed teal | Left Y |
| Line baseline | Dotted amber | Right Y |

Pill toggle for the comparison mode:
- **Forecast VS Budget** (default) — Duetto values vs Budget
- **OTB VS Budget** — Today's on-the-book values vs Budget

> A single stay month renders as enlarged markers (the area+spline can't draw
> a curve from one point). Select 2+ months to see the full trend curve.

---

## Daily Briefing Excel (Export)

One click produces `g5_daily_briefing_<month>.xlsx` with 8 sheets in
morning-meeting reading order:

1. **Portfolio Snapshot** — per-metric totals + variance %
2. **Hotel Scorecard** — one row per hotel
3. **Variance Pivot** — full pivot with all variance columns
4. **Same-Time Pace** — OTB vs STLY / ST2Y / ST3Y
5. **Historical Final** — Forecast vs Final LY / 2Y / 3Y
6. **Duetto Movement** — vs 1-day / 7-day / first-of-month
7. **Hotel Momentum** — daily forecast pickup per hotel
8. **Role Selection** — file → role mapping for transparency

---

## File structure

```
app.py             ← single-file Streamlit dashboard
assets/logo.png    ← brand logo (base64-encoded at startup)
requirements.txt   ← Python dependencies
README.md          ← this file
```

### Key dependencies
```text
streamlit
pandas
openpyxl
plotly
```

---

## Architecture notes

| Concern | Approach |
|---|---|
| Single-file | Easier deploy + hot-reload (no module imports for non-engineers) |
| Long-format `metric_long` | Wide → long once at load, all downstream filters are uniform |
| Auto-detect role files | `select_role_files()` infers Today / Yesterday / 7D / 1st Month |
| Look-back | `drop_duplicates(["Hotel","Stay Month","Metric","Reference"], keep="last")` per render |
| Display rename | Internal column names stay stable; display headers rename at `st.dataframe` time |
| Generic styler | Variance cells auto-detected by `" VS " in column_name` — add a new variance column and styling just works |
| Page nav via `st.pills` | Free layout: nav at top, filters below, content below — `st.tabs` can't do this |

---

## Bug-prevention checklist

When adding a new variance column:

1. Name it `<X> VS <Y>` where `<Y>` matches the actual reference used in
   the formula. No hidden `max()` / `min()` / `best`.
2. Use `safe_pct(num, denom)` — handles NaN and divide-by-zero.
3. Aggregate correctly:
   - **Rev / Rooms** → `sum`
   - **ADR / Occ** → `mean`
4. Add the column to the order list in `build_variance_pivot_table()`.
5. The styler picks it up automatically via the `" VS "` detector.

---

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| "No Forecast / Duetto columns found" | Source files don't match `REFERENCE_PATTERNS` — extend the pattern list in `app.py` |
| Empty Stay Month dropdown | All uploaded files filtered out — check filenames and date parsing |
| Past stay month shows "No data" | Previous month's report files not in the catalog — upload them too |
| Variance cell uncolored | Column name missing `" VS "` — the styler won't auto-detect it |
| `git push` rejected (non-fast-forward) | Remote was overwritten by accident — use `git push --force-with-lease origin main` after verifying local has the right code |
