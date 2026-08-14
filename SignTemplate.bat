@echo off
setlocal

:: Get the directory of the batch file
set "batchDir=%~dp0"
set "excelFile=default_macro_pivot_template.xlsm"
set "excelPath=%batchDir%%excelFile%"
set "logFile=%batchDir%signing_log.txt"

:: Require the KSP password to be supplied via environment variable, not hardcoded
if not defined DIGICERT_KSP_PASSWORD (
    echo ERROR: DIGICERT_KSP_PASSWORD environment variable is not set.
    echo Set it before running this script, e.g.:
    echo   set DIGICERT_KSP_PASSWORD=your-password
    exit /b 1
)

:: Run signtool and capture output
echo Signing started at %DATE% %TIME% > "%logFile%"
signtool_32.exe sign /csp "DigiCert Signing Manager KSP" /kc key_1272019481 ^
  /f "C:\Program Files\DigiCert\DigiCert Keylocker Tools\orbit_analytics_inc_1272019481_New.p7b" ^
  /p "%DIGICERT_KSP_PASSWORD%" ^
  /sha1 699E762289B0C51419206A0CA3A2591E6BBC1FD2 ^
  /tr http://timestamp.digicert.com /td SHA256 /v /debug /fd SHA256 "%excelPath%" >> "%logFile%" 2>&1

echo Signing completed at %DATE% %TIME% >> "%logFile%"
echo Log written to: %logFile%

exit
