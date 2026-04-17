@echo off
cd /d "%~dp0"
echo Looking for Rscript...
set RSCRIPT=

:: Method 1 - Already in PATH
where Rscript >nul 2>&1
if %ERRORLEVEL% == 0 (
    set RSCRIPT=Rscript
    goto :run
)

:: Method 2 - Registry 64-bit
for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\R-core\R" /v "InstallPath" 2^>nul') do set RSCRIPT="%%b\bin\Rscript.exe"
if defined RSCRIPT goto :verify

:: Method 3 - Registry 32-bit
for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\WOW6432Node\R-core\R" /v "InstallPath" 2^>nul') do set RSCRIPT="%%b\bin\Rscript.exe"
if defined RSCRIPT goto :verify

:: Method 4 - Registry current user
for /f "tokens=2*" %%a in ('reg query "HKCU\SOFTWARE\R-core\R" /v "InstallPath" 2^>nul') do set RSCRIPT="%%b\bin\Rscript.exe"
if defined RSCRIPT goto :verify

:: Method 5 - RStudio's own R (reads RStudio config)
for /f "delims=" %%i in ('dir /b /s "%LOCALAPPDATA%\RStudio\rstudio-prefs.json" 2^>nul') do set PREFS=%%i
if defined PREFS (
    for /f "tokens=2 delims=:," %%a in ('findstr /i "r_home" "%PREFS%"') do (
        set R_HOME=%%a
        set R_HOME=!R_HOME:"=!
        set R_HOME=!R_HOME: =!
        set RSCRIPT="!R_HOME!\bin\Rscript.exe"
    )
)
if defined RSCRIPT goto :verify

:: Method 6 - Scan Program Files for any R version
for /d %%i in ("C:\Program Files\R\R-*") do set RSCRIPT="%%i\bin\Rscript.exe"
if defined RSCRIPT goto :verify

for /d %%i in ("C:\Program Files (x86)\R\R-*") do set RSCRIPT="%%i\bin\Rscript.exe"
if defined RSCRIPT goto :verify

:: Method 7 - RStudio default R locations
for /d %%i in ("C:\Program Files\RStudio\resources\app\bin\quarto\bin\tools\*") do set RSCRIPT="%%i\Rscript.exe"
if defined RSCRIPT goto :verify

:: Method 8 - Scan user AppData (RStudio sometimes installs R here)
for /d %%i in ("%LOCALAPPDATA%\Programs\R\R-*") do set RSCRIPT="%%i\bin\Rscript.exe"
if defined RSCRIPT goto :verify

:: Method 9 - Portable R bundled with app
if exist "%~dp0R\bin\Rscript.exe" (
    set RSCRIPT="%~dp0R\bin\Rscript.exe"
    goto :verify
)

:: Method 10 - PowerShell deep search as last resort
echo Trying deep search via PowerShell...
for /f "delims=" %%i in ('powershell -command "Get-ChildItem 'C:\' -Recurse -ErrorAction SilentlyContinue -Filter 'Rscript.exe' | Select-Object -First 1 -ExpandProperty FullName"') do set RSCRIPT="%%i"
if defined RSCRIPT goto :verify

:: Nothing found
echo.
echo ERROR: R not found on this machine.
echo Please install R from https://cran.r-project.org
echo Or install RStudio from https://posit.co/download/rstudio-desktop/
echo.
pause
exit /b 1

:verify
:: Verify the found path actually exists
if not exist %RSCRIPT% (
    set RSCRIPT=
    echo Path not valid, continuing search...
    goto :notfound
)

:run
echo Found: %RSCRIPT%
echo.
echo Checking packages...
%RSCRIPT% -e ^
  "pkgs <- c('shiny','DT','dplyr','readxl','stringi','writexl','tibble','svDialogs'); ^
   missing <- pkgs[!pkgs %%in%% installed.packages()[,'Package']]; ^
   if(length(missing)){ ^
     message('Installing: ', paste(missing, collapse=', ')); ^
     install.packages(missing, repos='https://cloud.r-project.org') ^
   } else { message('All packages OK') }"

echo.
echo Launching app...
%RSCRIPT% -e "shiny::runApp('app.R', launch.browser = TRUE)"
pause
exit /b 0

:notfound
echo.
echo ERROR: R not found on this machine.
echo Please install R from https://cran.r-project.org
echo.
pause
exit /b 1