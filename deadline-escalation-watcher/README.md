# Deadline Escalation Watcher

**by Muhammad Bilal** · Category: Auto-Trigger / Scheduled Automation · 100% local — no APIs, no network, no cloud, no credentials.

## The problem

Project managers discover overdue work too late. Tasks slip quietly: a UAT sign-off sits overdue for two weeks, a milestone-blocking task stalls, and nobody sees the pattern until a steering meeting turns into a post-mortem. Manually scanning a task list every morning is slow, inconsistent, and the first thing that gets skipped on a busy day.

## The solution

A Python script that reads your task register (`tasks.csv`) every morning, classifies every open task by urgency, and produces a ready-to-share escalation pack:

- **`output/escalation-report-YYYY-MM-DD.html`** — styled, severity-coloured report with summary cards
- **`output/escalation-report-YYYY-MM-DD.xlsx`** — same data in Excel, filtered and colour-coded
- **`output/escalation-log.csv`** — append-only history, so you can trend escalations over time

### Severity rules (configurable at the top of `escalation_watcher.py`)

| Severity | Rule |
|---|---|
| **Critical** | Overdue > 7 days **OR** blocks a milestone |
| **High** | Overdue ≤ 7 days |
| **Medium** | Due within 48 hours |
| **Watch** | Due within 7 days |

Completed / cancelled / closed tasks are ignored. The log is **idempotent**: re-running on the same day adds zero duplicate entries (keyed on task ID + date).

## How it works

```
tasks.csv ──► escalation_watcher.py ──► HTML report + XLSX report + escalation-log.csv
```

1. `load_tasks()` reads and validates the CSV (missing columns or bad dates fail loudly with the row number).
2. `classify()` assigns each open task a severity bucket.
3. `write_html()` / `write_xlsx()` render the dated reports.
4. `append_log()` appends only new (task_id, date) pairs.

Everything runs offline with the Python standard library + openpyxl.

## Run it

```powershell
# one-time setup
pip install -r requirements.txt

# manual run (uses tasks.csv in this folder)
python escalation_watcher.py

# or: double-click run_escalation.bat
# or: point it at a different file / folder
python escalation_watcher.py --input my_tasks.csv --outdir ./out
```

## Schedule it

**Windows Task Scheduler (recommended):** open Task Scheduler → *Import Task* → select `task-scheduler.xml` → edit the `<Command>` path to your folder. Runs **daily at 07:00**.

**Linux / macOS cron** — run `crontab -e` and add:

```cron
0 7 * * * cd /path/to/deadline-escalation-watcher && python3 escalation_watcher.py >> escalation-watcher-run.log 2>&1
```

## Input format — `tasks.csv`

| Column | Example | Notes |
|---|---|---|
| `task_id` | ATL-101 | Unique ID |
| `project` | Project Atlas — LIMS Rollout | Any project name |
| `task_name` | Finalize UAT test cases | Free text |
| `owner` | Ayesha K. | Responsible person |
| `due_date` | 2026-10-02 | `YYYY-MM-DD` |
| `status` | In Progress | Done/Completed/Cancelled/Closed are ignored |
| `blocks_milestone` | Yes / No | Yes ⇒ overdue becomes Critical |

The bundled `tasks.csv` is **fictional demo data**: 30 tasks across 3 projects. Replace it with your real register (same columns) — no code changes needed.

## Configuration

All knobs live in the `CONFIG` dict at the top of `escalation_watcher.py`: thresholds (`overdue_critical_days`, `due_soon_hours`, `due_this_week_days`), ignored statuses, input/output paths.

## Scaling

- 30 tasks or 30,000 — the script is O(n) with no per-row overhead worth worrying about.
- Point `--input` at an export from your PM tool (Jira / Asana / MS Project CSV) as long as the columns match.
- The log gives you a time series: count Critical escalations per week to show whether delivery health is improving.

## Files

| File | Purpose |
|---|---|
| `escalation_watcher.py` | The automation |
| `tasks.csv` | Sample input (30 fictional tasks) |
| `requirements.txt` | Pinned dependency (`openpyxl==3.1.5`) |
| `run_escalation.bat` | One-click Windows runner |
| `task-scheduler.xml` | Importable Windows Task Scheduler task (daily 07:00) |
| `TEST-REPORT.md` | Verified run evidence |
| `deadline-escalation-watcher-demo.mp4` | Demo video |
| `output/` | Generated reports + log |

## Limitations (honest)

- Reads a CSV snapshot — it does not sync live with Jira/Asana/MS Project. Re-export the register (or point the script at a regularly exported file) before each run.
- Severity is rule-based, not judgement: a task can be "Critical" by rule and still be genuinely low-impact — the PM should review, not blindly forward.
- Dates use the machine's local date; run it on the machine whose timezone matches the team.
