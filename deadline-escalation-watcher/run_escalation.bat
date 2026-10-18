@echo off
REM Deadline Escalation Watcher — daily runner (Windows)
REM by Muhammad Bilal · 100% local, no network
REM Double-click to run manually, or let Task Scheduler call it at 07:00.
cd /d "%~dp0"
python escalation_watcher.py >> escalation-watcher-run.log 2>&1
if errorlevel 1 (
    echo [ERROR] Escalation watcher failed — see escalation-watcher-run.log
    exit /b 1
)
echo Escalation report generated. See the output\ folder.
