@echo off
:: ============================================================
:: BLAST SESSION MONITOR - User / VDI Endpoint
:: ============================================================
:: Run this on the VDI desktop to collect Blast session metrics.
:: Metrics are stored in memory and optionally written to a
:: network share for centralized monitoring.
::
:: CONFIGURATION: Change the OUTPUT_DIR value below
:: ============================================================

:: Network share path for centralized monitoring (admin reads this)
set OUTPUT_DIR=\\fileserver\vdi-perf\%COMPUTERNAME%

:: Collection interval (seconds)
set INTERVAL=5

:: Flush interval - how often to write data to file (seconds)
set FLUSH_INTERVAL=30

:: ============================================================
:: Request admin privileges (required for Input Delay counter)
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting admin privileges...
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)

set PS64=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe
"%PS64%" -ExecutionPolicy Bypass -NoProfile -File "%~dp0blast-monitor.ps1" -IntervalSeconds %INTERVAL% -OutputDir "%OUTPUT_DIR%" -FlushIntervalSeconds %FLUSH_INTERVAL%
pause
