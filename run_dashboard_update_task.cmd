@echo off
schtasks /Run /TN "\Codex\Customs Export BI Dashboard Update"
if errorlevel 1 (
  echo Failed to start the scheduled task.
  pause
  exit /b 1
)
echo Dashboard update task started.
echo Check D:\Drivers\customs-export-bi\logs for the update log.
pause
