@echo off
cd /D "%~dp0"
echo Unregistering SIFPreviewHandler.dll...
regsvr32.exe /u SIFPreviewHandler.dll

echo Restarting prevhost.exe...
taskkill /f /im prevhost.exe
start "" /wait "%SystemRoot%\System32\prevhost.exe"

echo DLL unregistration and prevhost.exe restart complete.
pause
