@echo off
title Zee AI for Word - launcher
cd /d "%~dp0"
set PORT=3200
set LOG=start-log.txt
echo ===== Zee AI for Word start %date% %time% ===== > %LOG%

echo [1/4] Stopping any previous Zee AI for Word server on port %PORT%...
rem Only stops a server started from THIS folder - never another app that happens to use the port.
powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort %PORT% -State Listen -ErrorAction SilentlyContinue | ForEach-Object { $p = Get-CimInstance Win32_Process -Filter ('ProcessId=' + $_.OwningProcess); if ($p.CommandLine -match 'Word_AddIn') { Stop-Process -Id $_.OwningProcess -Force; 'stopped old server pid ' + $_.OwningProcess } else { 'PORT %PORT% USED BY ANOTHER APP: ' + $p.CommandLine } }" >> %LOG% 2>&1
findstr /c:"USED BY ANOTHER APP" %LOG% >NUL
if not errorlevel 1 (
  echo [X] Port %PORT% is used by another program - see start-log.txt. Close it and run this again.
  pause
  exit /b 1
)
timeout /t 2 >NUL

rem The localhost certificate lasts 30 days - renew it when needed.
call npx --no office-addin-dev-certs verify >NUL 2>NUL
if errorlevel 1 (
  echo Renewing the localhost HTTPS certificate - click YES if Windows asks.
  call npx --no office-addin-dev-certs install >> %LOG% 2>&1
)

echo [2/4] Type-checking, testing and building (about 30 seconds)...
call npm run build >> %LOG% 2>&1
if errorlevel 1 (
  echo [X] Build failed - see start-log.txt
  echo BUILD FAILED >> %LOG%
  pause
  exit /b 1
)
echo build OK >> %LOG%

echo [3/4] Starting the server on https://localhost:%PORT% ...
start "Zee AI for Word server - keep this open" cmd /k "npx --no vite preview"
set TRIES=0
:waitloop
set CODE=000
for /f %%c in ('curl.exe -k -s -o NUL -w "%%{http_code}" https://localhost:%PORT%/taskpane.html') do set CODE=%%c
if "%CODE%"=="200" goto served
set /a TRIES+=1
if %TRIES% GEQ 30 goto failed
timeout /t 2 >NUL
goto waitloop
:failed
echo [X] Server did not start. See the "Zee AI for Word server" window. >> %LOG%
echo [X] Server did not start. Check the "Zee AI for Word server" window.
pause
exit /b 1
:served
echo server OK (HTTP %CODE%) >> %LOG%

echo [4/4] Registering the add-in and opening Word...
call npx --no office-addin-dev-settings debugging manifest.xml --disable >> %LOG% 2>&1
call npx --no office-addin-dev-settings register manifest.xml >> %LOG% 2>&1
call npx --no office-addin-dev-settings sideload manifest.xml desktop >> %LOG% 2>&1
echo sideload exit %errorlevel% >> %LOG%
echo.
echo Done. In Word: Home tab - Zee AI.
echo Keep the "Zee AI for Word server" window open while you use the add-in.
timeout /t 10
