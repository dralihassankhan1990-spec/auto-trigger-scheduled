# TEST-REPORT — Risk Register Auto-Scorer

**Tested:** 2026-10-16 · **Tester:** automated verification run (by Muhammad Bilal) · **Environment:** Python 3.12, openpyxl 3.1.5, matplotlib 3.6.3, numpy 1.26.4, Linux (same code runs on Windows via `run_risk_scorer.bat`)

## Test 1 — Full run

Command: `python3 risk_scorer.py`

Real output (verbatim):

```
2026-10-16 10:06:12,157 [INFO] Loaded 15 risks from .../risk-register-auto-scorer/risk-register.csv
2026-10-16 10:06:14,304 [INFO] DONE — 15 risks (Red=4 Amber=9 Green=2)
2026-10-16 10:06:14,307 [INFO] Heatmap: .../output/risk-heatmap.png
2026-10-16 10:06:14,310 [INFO] HTML: .../output/risk-report-2026-10-16.html
2026-10-16 10:06:14,310 [INFO] XLSX: .../output/scored-register-2026-10-16.xlsx
```

**Result: PASS.** Exit code 0. All three outputs created.

## Test 2 — Scoring correctness (hand-verified against the CSV)

| Risk | P × I | Exposure | Expected band | Actual band | Result |
|---|---|---|---|---|---|
| RSK-05 Key member resigns | 4 × 4 | 16 | Red (≥15) | Red | PASS |
| RSK-01 Vendor fails to deliver | 3 × 5 | 15 | Red | Red | PASS |
| RSK-02 Scope creep | 5 × 3 | 15 | Red | Red | PASS |
| RSK-03 Data migration errors | 3 × 5 | 15 | Red | Red | PASS |
| RSK-06 Cold-chain power outage | 3 × 4 | 12 | Amber (8–14) | Amber | PASS |
| RSK-15 Extreme weather | 2 × 2 | 4 | Green (<8) | Green | PASS |

Boundary checks: exposure 15 → Red (not Amber), exposure 7 → Green, exposure 8 → Amber — verified by the band thresholds in code (`Red 15–25, Amber 8–14, Green 1–7`).

## Test 3 — Output content checks

| Check | Expected | Actual | Result |
|---|---|---|---|
| HTML report opens, summary cards | 4 Red / 9 Amber / 2 Green / 15 total | Matches (screenshot verified) | PASS |
| Top-5 table | RSK-05 (16) first, then the three 15s | Matches; sorted by exposure desc | PASS |
| Heatmap PNG | 5×5 grid, risk IDs in cells, RSK-01/RSK-03 in the 15 cell (P3×I5) | Verified visually in report screenshot | PASS |
| Heatmap embedded in HTML | Base64 image renders without external file | Renders in browser screenshot | PASS |
| XLSX opens | 15 data rows + title + header, exposure + band columns | Verified via openpyxl: 18 rows (title + blank + header + 15 data) | PASS |

## Test 4 — Bug found and fixed during testing

**Bug:** first version failed with `FAILED: Colors must be aRGB hex values` — the tinted fill used `colour + "33"` (a 9-char string with `#`), and the header fill used a 6-digit hex.
**Fix:** added `argb(hex, alpha)` helper; tinted fills now use `argb(colour, "33")` and header fill uses `FF0E4C81`.
**Re-test:** the run above passed after the fix. The fix is in the shipped code.

## Test 5 — Error handling (negative tests)

| Test | Behaviour | Result |
|---|---|---|
| Missing input file | `FAILED: Risk register not found: .../no-such.csv`, exit 1, no traceback | PASS (executed) |
| Probability 6 | `FAILED: Row 2 (risk R-1): probability/impact must be 1–5, got 6/3`, exit 1 | PASS (executed) |
| Non-integer probability | `FAILED: Row 2 (risk R-1): probability/impact must be integers 1–5`, exit 1 | PASS (executed) |
| Missing column | `FAILED: risk-register.csv missing columns: [...]`, exit 1 (same code path as escalation watcher, verified there) | PASS |

## Files produced by the verified run

```
output/
├── risk-heatmap.png                  (65,654 bytes)
├── risk-report-2026-10-16.html      (94,926 bytes, heatmap embedded)
└── scored-register-2026-10-16.xlsx   (7,149 bytes)
```

## Verdict

✅ **All tests PASS.** The automation executes end-to-end without errors and produces all outputs with correct scoring.
