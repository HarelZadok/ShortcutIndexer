@echo off
setlocal EnableDelayedExpansion

:: Check for administrator privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo This uninstaller requires administrator privileges.
    echo Please run as administrator.
    pause
    exit /b 1
)

echo ==========================================
echo ShortcutIndexer Uninstaller
echo ==========================================
echo.

set "INSTALL_DIR=%ProgramFiles%\ShortcutIndexer"

echo Stopping Windows Explorer to release file locks...
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 3 /nobreak >nul

echo Unregistering shell extension...
set "REGASM_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
if not exist "%REGASM_PATH%" (
    set "REGASM_PATH=C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
)

if exist "%REGASM_PATH%" (
    if exist "%INSTALL_DIR%\ShortcutIndexerHandler.dll" (
        "%REGASM_PATH%" "%INSTALL_DIR%\ShortcutIndexerHandler.dll" /u /silent
        if %errorLevel% neq 0 (
            echo WARNING: Failed to unregister shell extension with RegAsm.
        ) else (
            echo Shell extension unregistered successfully.
        )
    )
) else (
    echo WARNING: RegAsm.exe not found. Manual registry cleanup will be performed.
)

echo Cleaning up registry...
reg delete "HKCR\*\shellex\ContextMenuHandlers\ShortcutIndexer" /f 2>nul
reg delete "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\ShortcutIndexer" /f 2>nul

echo Restarting Windows Explorer to apply changes...
taskkill /f /im explorer.exe >nul 2>&1
start explorer.exe

echo Removing files...
rmdir /S /Q "%INSTALL_DIR%"

echo.
echo ==========================================
echo ShortcutIndexer has been uninstalled successfully.
echo ==========================================
echo.
pause
