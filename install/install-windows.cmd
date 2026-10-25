@echo off
rem Installs the Nasaq PowerPoint add-in for the current Windows user (no admin rights needed).
rem Usage: double-click this file. Run "install-windows.cmd /uninstall" to remove Nasaq.
setlocal

set "MANIFEST_URL=https://abumunirah.github.io/nasaq/manifest.xml"
set "ADDIN_ID=9f4a790a-ef14-41ef-8b0d-6c6ee8a906b2"
set "INSTALL_DIR=%LOCALAPPDATA%\Nasaq"
set "MANIFEST=%INSTALL_DIR%\manifest.xml"
set "REG_KEY=HKCU\Software\Microsoft\Office\16.0\Wef\Developer"

if /i "%~1"=="/uninstall" goto uninstall

echo Installing Nasaq...

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
curl.exe -fsSL -o "%MANIFEST%.tmp" "%MANIFEST_URL%"
if errorlevel 1 goto download_failed
findstr /c:"<OfficeApp" "%MANIFEST%.tmp" >nul || goto download_failed
move /y "%MANIFEST%.tmp" "%MANIFEST%" >nul

rem Register the manifest the same way Office's own developer tools sideload add-ins.
reg add "%REG_KEY%" /v "%ADDIN_ID%" /t REG_SZ /d "%MANIFEST%" /f >nul
if errorlevel 1 goto register_failed

echo.
echo Nasaq is installed.
echo   1. Close PowerPoint completely and open it again.
echo   2. Choose Home ^> Add-ins ^> My Add-ins ^> Developer Add-ins ^> Nasaq.
goto done

:uninstall
reg delete "%REG_KEY%" /v "%ADDIN_ID%" /f >nul 2>&1
if exist "%INSTALL_DIR%" rmdir /s /q "%INSTALL_DIR%"
echo Nasaq was removed. Restart PowerPoint.
goto done

:download_failed
if exist "%MANIFEST%.tmp" del "%MANIFEST%.tmp"
echo Could not download Nasaq. Check your internet connection and try again.
goto done

:register_failed
echo Could not register Nasaq with Office.
goto done

:done
echo.
pause
