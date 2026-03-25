@echo off
:: ============================================================
:: BLAST SESSION MONITOR - Admin / Remote Viewer
:: ============================================================
:: Bu dosyayi admin makinesinde calistirin.
:: VDI endpoint'teki veriyi okur ve dashboard'da gosterir.
:: Counter toplamaz, sadece izleme yapar.
::
:: YAPILANDIRMA: Asagidaki WATCH_FILE degerini degistirin
:: ============================================================

:: Izlenecek VDI'nin JSON dosyasi (VDI tarafinda blast-monitor-user.cmd olusturur)
set WATCH_FILE=\\fileserver\vdi-perf\HOLWINVDI04\blast_live.json

:: Dashboard portu (her VDI icin farkli port kullanabilirsiniz)
set PORT=8888

:: ============================================================
set PS64=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe
"%PS64%" -ExecutionPolicy Bypass -NoProfile -File "%~dp0blast-monitor.ps1" -WatchFile "%WATCH_FILE%" -Port %PORT%
pause
