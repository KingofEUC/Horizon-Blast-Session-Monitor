@echo off
:: ============================================================
:: BLAST SESSION MONITOR - User / VDI Endpoint
:: ============================================================
:: Bu dosyayi kullanicinin VDI masaustunde calistirin.
:: Metrikleri toplar ve paylasimli klasore yazar.
::
:: YAPILANDIRMA: Asagidaki OUTPUT_DIR degerini degistirin
:: ============================================================

:: Veri yazilacak paylasimli klasor (admin okuyacak)
set OUTPUT_DIR=\\fileserver\vdi-perf\%COMPUTERNAME%

:: Toplama araligi (saniye)
set INTERVAL=5

:: Flush araligi - ne siklikta dosyaya yazilsin (saniye)
set FLUSH_INTERVAL=30

:: ============================================================
:: Admin yetkisi iste (Input Delay counter icin gerekli)
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting admin privileges...
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)

set PS64=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe
"%PS64%" -ExecutionPolicy Bypass -NoProfile -File "%~dp0blast-monitor.ps1" -IntervalSeconds %INTERVAL% -OutputDir "%OUTPUT_DIR%" -FlushIntervalSeconds %FLUSH_INTERVAL%
pause
