@echo off
title Zee AI for Word - install on this computer
echo ============================================================
echo   Zee AI for Word - install the add-in on this computer
echo   (no Node.js or code needed - the add-in runs from GitHub)
echo ============================================================
echo.
rem Change GHUSER if your GitHub username is different.
set GHUSER=zeexx99
set DIR=%USERPROFILE%\ZeeAIWord
set URL=https://%GHUSER%.github.io/zee-ai-word/manifest.xml
if not exist "%DIR%" mkdir "%DIR%"

echo [1/3] Downloading the add-in manifest...
curl.exe -s -f -o "%DIR%\manifest.xml" "%URL%"
if errorlevel 1 (
  echo [X] Could not download %URL%
  echo     Check the internet connection, or copy manifest.prod.xml from the main PC
  echo     to %DIR%\manifest.xml and run this file again.
  if not exist "%DIR%\manifest.xml" goto end
)
echo [OK] Saved to %DIR%\manifest.xml
echo.

echo [2/3] Close ALL Word windows now, then press any key.
pause >NUL
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v ZeeAIWord /t REG_SZ /d "%DIR%\manifest.xml" /f >NUL
if errorlevel 1 (
  echo [X] Could not register the add-in. Ask IT whether add-ins are blocked by policy.
  goto end
)
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >NUL 2>NUL
echo [OK] Registered.
echo.
echo [3/3] Open Word - Home tab - Zee AI.
echo      First time: open Settings (gear) and enter your Gemini API key.
echo.
echo To uninstall later:  reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v ZeeAIWord /f
:end
echo.
pause
