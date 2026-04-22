@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ================================================
echo  Lam sach WHONET phien giai - iSharp AMS
echo ================================================
echo.
echo Dang tim Rscript...
set RSCRIPT=

:: Phuong phap 1 - Rscript co san trong PATH
where Rscript >nul 2>&1
if %ERRORLEVEL% == 0 (
    set RSCRIPT=Rscript
    goto :run
)

:: Phuong phap 2 - Registry 64-bit
for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\R-core\R" /v "InstallPath" 2^>nul') do set RSCRIPT="%%b\bin\Rscript.exe"
if defined RSCRIPT goto :verify

:: Phuong phap 3 - Registry 32-bit
for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\WOW6432Node\R-core\R" /v "InstallPath" 2^>nul') do set RSCRIPT="%%b\bin\Rscript.exe"
if defined RSCRIPT goto :verify

:: Phuong phap 4 - Registry current user
for /f "tokens=2*" %%a in ('reg query "HKCU\SOFTWARE\R-core\R" /v "InstallPath" 2^>nul') do set RSCRIPT="%%b\bin\Rscript.exe"
if defined RSCRIPT goto :verify

:: Phuong phap 5 - Doc cau hinh RStudio de tim R
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

:: Phuong phap 6 - Quet Program Files
for /d %%i in ("C:\Program Files\R\R-*") do set RSCRIPT="%%i\bin\Rscript.exe"
if defined RSCRIPT goto :verify

for /d %%i in ("C:\Program Files (x86)\R\R-*") do set RSCRIPT="%%i\bin\Rscript.exe"
if defined RSCRIPT goto :verify

:: Phuong phap 7 - Thu muc RStudio mac dinh
for /d %%i in ("C:\Program Files\RStudio\resources\app\bin\quarto\bin\tools\*") do set RSCRIPT="%%i\Rscript.exe"
if defined RSCRIPT goto :verify

:: Phuong phap 8 - AppData nguoi dung
for /d %%i in ("%LOCALAPPDATA%\Programs\R\R-*") do set RSCRIPT="%%i\bin\Rscript.exe"
if defined RSCRIPT goto :verify

:: Phuong phap 9 - R di dong kem theo app
if exist "%~dp0R\bin\Rscript.exe" (
    set RSCRIPT="%~dp0R\bin\Rscript.exe"
    goto :verify
)

:: Phuong phap 10 - Tim quet sau qua PowerShell (cham, chi dung khi can)
echo Dang tim kiem sau qua PowerShell...
for /f "delims=" %%i in ('powershell -command "Get-ChildItem 'C:\' -Recurse -ErrorAction SilentlyContinue -Filter 'Rscript.exe' | Select-Object -First 1 -ExpandProperty FullName"') do set RSCRIPT="%%i"
if defined RSCRIPT goto :verify

:: Khong tim thay R
echo.
echo LOI: Khong tim thay R tren may tinh nay.
echo Vui long cai dat R tai: https://cran.r-project.org
echo Hoac RStudio tai: https://posit.co/download/rstudio-desktop/
echo.
pause
exit /b 1

:verify
if not exist %RSCRIPT% (
    set RSCRIPT=
    echo Duong dan khong hop le, tiep tuc tim kiem...
    goto :notfound
)

:run
echo Tim thay: %RSCRIPT%
echo.

:: Kiem tra file renv.lock
if not exist "%~dp0renv.lock" (
    echo LOI: Khong tim thay renv.lock trong thu muc du an.
    echo Vui long dam bao da clone day du repository.
    echo.
    pause
    exit /b 1
)

:: Kiem tra renv/activate.R
if not exist "%~dp0renv\activate.R" (
    echo LOI: Khong tim thay renv\activate.R.
    echo Vui long dam bao da clone day du repository.
    echo.
    pause
    exit /b 1
)

echo Buoc 1/2 - Khoi phuc goi R tu renv.lock...
echo (Lan chay dau tien co the mat vai phut - vui long cho)
echo.
%RSCRIPT% --vanilla -e "source('renv/activate.R'); renv::restore(prompt = FALSE)"
if %ERRORLEVEL% neq 0 (
    echo.
    echo LOI: Khoi phuc goi that bai.
    echo Kiem tra thong bao loi phia tren va chay lai.
    echo.
    pause
    exit /b 1
)

echo.
echo Buoc 2/2 - Khoi chay ung dung Lam sach WHONET...
echo.
%RSCRIPT% -e "shiny::runApp('app_whonet.R', launch.browser = TRUE)"
pause
exit /b 0

:notfound
echo.
echo LOI: Khong tim thay R tren may tinh nay.
echo Vui long cai dat R tai: https://cran.r-project.org
echo.
pause
exit /b 1
