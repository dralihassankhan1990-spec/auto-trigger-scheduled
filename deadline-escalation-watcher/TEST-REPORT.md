# TEST-REPORT — Deadline Escalation Watcher

**Tested:** 2026-10-16 · **Tester:** automated verification run (by Muhammad Bilal) · **Environment:** Python 3.12, openpyxl 3.1.5, Linux (same code runs on Windows via `run_escalation.bat`)

## Test 1 — First run (fresh output folder)

Command: `python3 escalation_watcher.py`

Real output (verbatim):

```
2026-10-16 10:05:54,734 [INFO] Loaded 30 tasks from .../deadline-escalation-watcher/tasks.csv
2026-10-16 10:05:54,925 [INFO] Appended 25 new log entries
2026-10-16 10:05:54,927 [INFO] DONE — 25 escalations (Critical=8 High=7 Medium=6 Watch=4)
2026-10-16 10:05:54,927 [INFO] HTML: .../output/escalation-report-2026-10-16.html
2026-10-16 10:05:54,928 [INFO] XLSX: .../output/escalation-report-2026-10-16.xlsx
2026-10-16 10:05:54,928 [INFO] Log: .../output/escalation-log.csv (+25 new)
```

**Result: PASS.** Exit code 0. All three output files created.

## Test 2 — Idempotency (second run, same day)

Command: `python3 escalation_watcher.py` (immediately re-run)

Real output (verbatim):

```
2026-10-16 10:05:56,563 [INFO] Loaded 30 tasks from .../deadline-escalation-watcher/tasks.csv
2026-10-16 10:05:56,772 [INFO] Log already up to date for 2026-10-16 — 0 new entries
2026-10-16 10:05:56,772 [INFO] DONE — 25 escalations (Critical=8 High=7 Medium=6 Watch=4)
```

`wc -l output/escalation-log.csv` → **26 lines** (1 header + 25 entries) after two runs — no duplicates.

**Result: PASS.** Re-running the same day does not duplicate log entries.

## Test 3 — Output content checks

| Check | Expected | Actual | Result |
|---|---|---|---|
| HTML report opens in browser, severity cards visible | 8 / 7 / 6 / 4 / 25 | 8 Critical, 7 High, 6 Medium, 4 Watch, 25 total (screenshot verified) | PASS |
| HTML contains all escalated task IDs | 25 rows | 25 rows, sorted Critical → Watch | PASS |
| XLSX opens, 25 data rows + header | 26 rows used in sheet | Verified via openpyxl: header + 25 rows | PASS |
| Done/Cancelled tasks excluded | BCN-210 (Done), CDR-310 (Cancelled) absent | Both absent from report and log | PASS |
| Future tasks (>7 days) excluded | ATL-110, BCN-209, CDR-309 absent | All absent | PASS |
| Milestone-blocking overdue task → Critical | ATL-106 (overdue 1 day, blocks milestone) = Critical | Critical | PASS |
| Overdue ≤7 days, not blocking → High | CDR-303 (overdue 7 days) = High | High | PASS |
| Overdue >7 days → Critical even if not blocking | CDR-302 (overdue 12 days) = Critical | Critical | PASS |

## Test 4 — Bug found and fixed during testing

**Bug:** first run failed with `FAILED: Colors must be aRGB hex values` — openpyxl requires 8-digit aRGB colours, and the code passed 6-digit hex.
**Fix:** added `argb()` helper (prepends `FF` alpha) and applied it to every fill/font colour, including the header fill.
**Re-test:** both runs above passed after the fix. The fix is in the shipped code.

## Test 5 — Error handling (negative tests)

| Test | Command / input | Behaviour | Result |
|---|---|---|---|
| Missing input file | `--input no-such-file.csv` | `FAILED: Task register not found: .../no-such-file.csv`, exit 1, no traceback | PASS (executed) |
| Bad date | due_date `2026/10/02` | `FAILED: Row 2 (task X-1): bad due_date '2026/10/02' — expected YYYY-MM-DD`, exit 1 | PASS (executed) |
| Missing column | CSV without `owner` | `FAILED: tasks.csv missing columns: ['owner']`, exit 1 | PASS (executed) |

## Files produced by the verified run

```
output/
├── escalation-report-2026-10-16.html   (9,155 bytes)
├── escalation-report-2026-10-16.xlsx   (7,430 bytes)
└── escalation-log.csv                  (3,799 bytes, 26 lines)
```

## Verdict

✅ **All tests PASS.** The automation executes end-to-end without errors, produces all three outputs, and is idempotent.
