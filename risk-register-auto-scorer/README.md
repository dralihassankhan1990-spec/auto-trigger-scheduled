# Risk Register Auto-Scorer

**by Muhammad Bilal** · Category: Auto-Trigger / Scheduled Automation · 100% local — no APIs, no network, no cloud, no credentials.

## The problem

Risk registers go stale. A team writes 15 risks at kickoff, scores them once by gut feel, and never revisits them — so the steering meeting debates the wrong risks. Manual scoring in a spreadsheet is inconsistent (is that risk a 12 or a 16?), and nobody has time to redraw the heatmap every week.

## The solution

A Python script that reads your risk register (`risk-register.csv`), computes exposure for every risk, assigns RAG bands, and produces a ready-to-present weekly pack:

- **`output/risk-report-YYYY-MM-DD.html`** — summary cards, Top-5 risks table, embedded heatmap, full scored register
- **`output/risk-heatmap.png`** — 5×5 Probability × Impact heatmap with risk IDs plotted in their cells
- **`output/scored-register-YYYY-MM-DD.xlsx`** — every risk with computed exposure + band, filtered and colour-coded

### Scoring rules (configurable at the top of `risk_scorer.py`)

```
Exposure = Probability (1–5) × Impact (1–5)
Red    = Exposure ≥ 15   → act now, steering attention
Amber  = Exposure 8–14   → actively managed
Green  = Exposure < 8    → watch
```

## How it works

```
risk-register.csv ──► risk_scorer.py ──► HTML report + heatmap PNG + scored XLSX
```

1. `load_risks()` reads and validates the CSV (probability/impact must be integers 1–5, or the run fails with the row number).
2. `band_for()` assigns the RAG band from the exposure score.
3. `draw_heatmap()` renders the 5×5 heatmap with matplotlib (headless `Agg` backend — safe in scheduled tasks) and plots each risk ID in its cell.
4. `write_html()` / `write_xlsx()` render the dated report pack. The heatmap is embedded in the HTML as base64, so the report is a single self-contained file.

Everything runs offline with the Python standard library + openpyxl + matplotlib.

## Run it

```powershell
# one-time setup
pip install -r requirements.txt

# manual run (uses risk-register.csv in this folder)
python risk_scorer.py

# or: double-click run_risk_scorer.bat
# or: point it at a different file / folder
python risk_scorer.py --input my_risks.csv --outdir ./out
```

## Schedule it

**Windows Task Scheduler (recommended):** open Task Scheduler → *Import Task* → select `task-scheduler.xml` → edit the `<Command>` path to your folder. Runs **every Monday at 09:00**.

**Linux / macOS cron** — run `crontab -e` and add:

```cron
0 9 * * 1 cd /path/to/risk-register-auto-scorer && python3 risk_scorer.py >> risk-scorer-run.log 2>&1
```

## Input format — `risk-register.csv`

| Column | Example | Notes |
|---|---|---|
| `risk_id` | RSK-01 | Unique ID |
| `risk` | Key vendor fails to deliver… | Free text |
| `category` | Supplier | Any grouping |
| `owner` | Procurement Lead | Risk owner |
| `probability` | 3 | Integer 1–5 |
| `impact` | 5 | Integer 1–5 |
| `mitigation` | Dual-source critical instruments… | Current mitigation |

The bundled `risk-register.csv` is **fictional demo data**: 15 realistic project risks. Replace it with your real register (same columns) — no code changes needed.

## Configuration

All knobs live in the `CONFIG` dict at the top of `risk_scorer.py`: RAG thresholds, `top_n` featured risks, input/output paths, heatmap filename.

## Scaling

- Works for 15 risks or 500 — the heatmap groups multiple risk IDs per cell.
- Re-run weekly: the dated reports give you a week-by-week audit trail of how exposure moved.
- Feed the Top-5 table straight into your steering deck — it's designed to be screenshot-ready.

## Files

| File | Purpose |
|---|---|
| `risk_scorer.py` | The automation |
| `risk-register.csv` | Sample input (15 fictional risks) |
| `requirements.txt` | Pinned dependencies |
| `run_risk_scorer.bat` | One-click Windows runner |
| `task-scheduler.xml` | Importable Windows Task Scheduler task (Mondays 09:00) |
| `TEST-REPORT.md` | Verified run evidence |
| `risk-register-auto-scorer-demo.mp4` | Demo video |
| `output/` | Generated reports + heatmap |

## Limitations (honest)

- Scores are only as good as the probability/impact ratings a human enters — the tool computes consistently, it does not judge.
- Exposure is a simple P×I product; it does not model correlations, cascading risks, or velocity (how fast a risk is approaching).
- Reads a CSV snapshot — it does not sync live with a PM tool. Re-export the register before each run.
