@echo off
cd /d "%~dp0"
title PictoPDF

if not exist "node_modules\.bin\vinext.cmd" goto missing

echo Starting PictoPDF...
start "" powershell.exe -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 4; Start-Process 'http://localhost:3000'"
set "WRANGLER_LOG_PATH=.wrangler/wrangler.log"
call "node_modules\.bin\vinext.cmd" dev
goto done

:missing
echo Preparing PictoPDF for first use...
call npm install --ignore-scripts --no-audit --no-fund
if errorlevel 1 goto error
call "%~f0"
exit /b

:error
echo Setup failed. Please install Node.js, then try again.
pause
exit /b 1

:done
echo.
echo PictoPDF has stopped.
pause
