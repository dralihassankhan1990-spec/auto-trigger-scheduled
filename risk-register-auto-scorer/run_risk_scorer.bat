@echo off
REM Risk Register Auto-Scorer — weekly runner (Windows)
REM by Muhammad Bilal · 100% local, no network
REM Double-click to run manually, or let Task Scheduler call it Mondays at 09:00.
cd /d "%~dp0"
python risk_scorer.py >> risk-scorer-run.log 2>&1
if errorlevel 1 (
    echo [ERROR] Risk scorer failed — see risk-scorer-run.log
    exit /b 1
)
echo Risk report generated. See the output\ folder.
