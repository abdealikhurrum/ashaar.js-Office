@echo off
REM Ashaar Poetry - Windows installer launcher.
REM Double-click this file. It runs the PowerShell installer (which will ask
REM for administrator rights to share the add-in folder).
echo Starting the Ashaar Poetry installer...
powershell -NoProfile -ExecutionPolicy Bypass -Command "iwr 'https://abdealikhurrum.github.io/ashaar.js-Office/install/Install-Ashaar.ps1' -UseBasicParsing | iex"
echo.
echo If a window flashed and closed, right-click Install-Ashaar.ps1 and choose "Run with PowerShell".
pause
