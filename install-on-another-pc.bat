@echo off
title Zee AI for Word - install on this computer
rem ============================================================
rem  Installs (or removes) the hosted Zee AI for Word add-in on THIS computer.
rem  No Node.js, Git or code needed - the add-in runs from GitHub Pages.
rem  Copy only this file to the other PC and double-click it.
rem  To uninstall later: run it again and choose U.
rem ============================================================
setlocal
set GHUSER=zeexx99
set SITE=https://%GHUSER%.github.io/zee-ai-word
set DIR=%USERPROFILE%\ZeeAIWord
set KEY=HKCU\Software\Microsoft\Office\16.0\WEF\Developer
set NAME=ZeeAIWord

echo ============================================================
echo   Zee AI for Word - install on this computer
echo ============================================================
echo.
echo   I = install or update     U = uninstall     X = exit
choice /c IUX /n /m "  Your choice [I/U/X]: "
if errorlevel 3 goto end
if errorlevel 2 goto uninstall

echo.
echo [1/4] Checking that the add-in is online...
set CODE=000
for /f %%c in ('curl.exe -s -o NUL -w "%%{http_code}" "%SITE%/taskpane.html"') do set CODE=%%c
if not "%CODE%"=="200" (
  echo [X] %SITE% is not reachable - HTTP %CODE%.
  echo     Check the internet connection or a company proxy/firewall, then run this again.
  goto end
)
echo [OK] Online.

echo [2/4] Downloading the add-in manifest...
if not exist "%DIR%" mkdir "%DIR%"
curl.exe -s -f -o "%DIR%\manifest.xml" "%SITE%/manifest.xml"
if errorlevel 1 (
  echo [X] Could not download %SITE%/manifest.xml
  goto end
)
findstr /c:"<OfficeApp" "%DIR%\manifest.xml" >NUL
if errorlevel 1 (
  echo [X] The downloaded file is not an add-in manifest. Try again later.
  goto end
)
echo [OK] Saved to %DIR%\manifest.xml

echo [3/4] Close ALL Word windows now (save your work first), then press any key.
pause >NUL
tasklist /fi "imagename eq WINWORD.EXE" | find /i "WINWORD.EXE" >NUL
if not errorlevel 1 (
  echo     Word is still open - close it, then press any key.
  pause >NUL
)
reg add "%KEY%" /v %NAME% /t REG_SZ /d "%DIR%\manifest.xml" /f >NUL
if errorlevel 1 (
  echo [X] Could not register the add-in. Your IT policy may block add-ins - ask IT.
  goto end
)
rem Clear Office's add-in cache so the newest version loads.
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >NUL 2>NUL
echo [OK] Registered.

echo [4/4] Done.
echo.
echo   Open Word - Home tab - Zee AI.
echo   First time: open Settings (gear icon) and paste your Gemini API key
echo   (free at https://aistudio.google.com/apikey), then click Test connection.
echo   Updates are automatic - new versions load from GitHub. If an old version
echo   ever sticks, run this file again.
goto end

:uninstall
echo.
echo Close ALL Word windows, then press any key to remove Zee AI for Word.
pause >NUL
reg delete "%KEY%" /v %NAME% /f >NUL 2>NUL
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >NUL 2>NUL
if exist "%DIR%" rmdir /s /q "%DIR%" >NUL 2>NUL
echo [OK] Zee AI for Word was removed from this computer.

:end
echo.
pause
endlocal
