@echo off
:: ============================================================
:: BLAST SESSION MONITOR - Admin / Remote Viewer
:: ============================================================
:: Run this on the admin workstation to monitor a remote VDI
:: endpoint. Reads the JSON file produced by the user script
:: and displays a live dashboard. Does not collect counters.
::
:: CONFIGURATION: Change the WATCH_FILE value below
:: ============================================================

:: Path to the VDI endpoint's JSON file (created by blast-monitor-user.cmd)
set WATCH_FILE=\\fileserver\vdi-perf\DESKTOP01\blast_live.json

:: Dashboard port (use different ports to monitor multiple VDIs)
set PORT=8888

:: ============================================================
set PS64=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe
"%PS64%" -ExecutionPolicy Bypass -NoProfile -File "%~dp0blast-monitor.ps1" -WatchFile "%WATCH_FILE%" -Port %PORT%
pause
