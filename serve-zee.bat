@echo off
title Zee AI for Word server - keep this open
cd /d "%~dp0"
rem Local server (https://localhost:3200) serving the built files in dist\.
set CODE=000
for /f %%c in ('curl.exe -k -s -o NUL -w "%%{http_code}" https://localhost:3200/taskpane.html') do set CODE=%%c
if "%CODE%"=="200" (
  echo Zee AI for Word server is already running on https://localhost:3200
  timeout /t 5
  exit /b 0
)
rem Renew the localhost certificate if it expired (it lasts 30 days).
call npx --no office-addin-dev-certs verify >NUL 2>NUL
if errorlevel 1 (
  echo Renewing the localhost HTTPS certificate - click YES if Windows asks.
  call npx --no office-addin-dev-certs install
)
if not exist dist\taskpane.html call npm run build
echo Zee AI for Word server starting on https://localhost:3200 - keep this window open (you can minimize it).
call npx --no vite preview
pause
