# 05 — Auto-Trigger / Scheduled Automations

**by Muhammad Bilal** · 100% local — no APIs, no network, no cloud, no credentials. All demo data fictional.

Automations that run themselves on a schedule, so the project team gets its morning brief and weekly risk picture without anyone remembering to produce them.

## Automations in this category

### 1. Deadline Escalation Watcher — `deadline-escalation-watcher/`

Reads `tasks.csv` every morning and flags overdue / due-soon tasks with severity tiers (Critical / High / Medium / Watch). Produces a styled HTML report, a colour-coded Excel report, and an append-only escalation log that is **idempotent** — re-running the same day adds zero duplicates.

- **Schedule:** daily 07:00 (Windows Task Scheduler XML included; cron line in README)
- **Demo:** `deadline-escalation-watcher-demo.mp4`
- **Sample result:** 30 tasks → 25 escalations (8 Critical, 7 High, 6 Medium, 4 Watch)

### 2. Risk Register Auto-Scorer — `risk-register-auto-scorer/`

Reads `risk-register.csv`, computes exposure = Probability × Impact, assigns RAG bands (Red ≥ 15, Amber 8–14, Green < 8), draws a 5×5 heatmap with the risk IDs plotted in their cells, and produces an HTML report + scored Excel register.

- **Schedule:** every Monday 09:00 (Windows Task Scheduler XML included; cron line in README)
- **Demo:** `risk-register-auto-scorer-demo.mp4`
- **Sample result:** 15 risks → 4 Red, 9 Amber, 2 Green; top risk RSK-05 (exposure 16)

## Running them

Each folder is self-contained: source, sample data, `requirements.txt` (pinned), `run_*.bat`, `task-scheduler.xml`, `README.md`, `TEST-REPORT.md` with real run evidence, and a demo video. See each folder's README for exact commands.

## Design rules for this category

1. **Scheduled, not manual** — every automation ships with a Task Scheduler XML + cron line.
2. **Idempotent or dated** — scheduled runs never corrupt history (dedupe keys, dated filenames).
3. **Fail loudly** — bad input produces a clear error message and a non-zero exit code, never a silently wrong report.
4. **Human-readable output** — HTML + Excel reports a PM can forward without editing.
