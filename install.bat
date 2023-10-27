@echo off
cd /D "%~dp0"
echo Registering SIFPreviewHandler.dll...
regsvr32.exe SIFPreviewHandler.dll

echo Restarting explorer.exe...
taskkill /f /im explorer.exe
start "" explorer.exe

echo DLL registration and explorer.exe restart complete.
ECHO ***********************************************************
ECHO * NOTICE: RESTARTING EXPLORER MAY BRIEFLY DISPLAY A BLACK *
ECHO *   SCREEN; THIS SHOULD NOT TAKE MORE THAN A FEW SECONDS. *
ECHO ***********************************************************

pause
