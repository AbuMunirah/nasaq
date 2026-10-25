@echo off
rem Installs Nasaq Classic (the VBA .ppam edition) for the current Windows user, for
rem PowerPoint versions that can't run the web add-in (e.g. Office LTSC 2021, 2019, 2016).
rem No admin rights needed. Run "install-windows-classic.cmd /uninstall" to remove it.
rem
rem The add-in goes into PowerPoint's user add-ins folder, which Office trusts by default,
rem so it loads without a macro prompt unless company policy disables trusted locations.
setlocal

set "PPAM_URL=https://abumunirah.github.io/nasaq/install/Nasaq.ppam"
set "ADDINS_DIR=%APPDATA%\Microsoft\AddIns"
set "PPAM=%ADDINS_DIR%\Nasaq.ppam"
set "ADDIN_KEY=HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\Nasaq"
set "TRUST_KEY=HKCU\Software\Microsoft\Office\16.0\PowerPoint\Security\Trusted Locations"
set "TRUST_POLICY=HKCU\Software\Policies\Microsoft\Office\16.0\PowerPoint\Security\Trusted Locations"

tasklist /fi "imagename eq POWERPNT.EXE" | find /i "POWERPNT.EXE" >nul
if not errorlevel 1 (
    echo Please close PowerPoint, then run this installer again.
    goto done
)

if /i "%~1"=="/uninstall" goto uninstall

echo Installing Nasaq Classic...

if not exist "%ADDINS_DIR%" mkdir "%ADDINS_DIR%"
curl.exe -fsSL -o "%PPAM%.tmp" "%PPAM_URL%"
if errorlevel 1 goto download_failed
for %%F in ("%PPAM%.tmp") do if %%~zF LSS 10000 goto download_failed
move /y "%PPAM%.tmp" "%PPAM%" >nul

rem Load the add-in automatically every time PowerPoint starts.
reg add "%ADDIN_KEY%" /v Path /t REG_SZ /d "%PPAM%" /f >nul
reg add "%ADDIN_KEY%" /v AutoLoad /t REG_DWORD /d 1 /f >nul
if errorlevel 1 goto register_failed

echo.
echo Nasaq Classic is installed. Open PowerPoint and look for the Nasaq tab.

rem Warn if PowerPoint may still ask about macros.
reg query "%TRUST_POLICY%" /v AllLocationsDisabled 2>nul | find "0x1" >nul
if not errorlevel 1 (
    echo.
    echo Note: your organization has disabled trusted locations, so PowerPoint may
    echo ask to enable macros or block the add-in. Ask your IT team to allow it.
    goto done
)
reg query "%TRUST_KEY%" /s /f "Microsoft\AddIns" >nul 2>&1
if errorlevel 1 (
    echo.
    echo Note: PowerPoint's add-ins folder is not a trusted location on this PC, so
    echo PowerPoint may ask to enable macros. Ask your IT team if it's blocked.
)
goto done

:uninstall
reg delete "%ADDIN_KEY%" /f >nul 2>&1
if exist "%PPAM%" del "%PPAM%"
echo Nasaq Classic was removed.
goto done

:download_failed
if exist "%PPAM%.tmp" del "%PPAM%.tmp"
echo Could not download Nasaq. Check your internet connection and try again.
goto done

:register_failed
echo Could not register Nasaq with PowerPoint.
goto done

:done
echo.
pause
