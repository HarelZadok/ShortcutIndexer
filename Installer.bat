@echo off
setlocal EnableDelayedExpansion

:: Check for administrator privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo This installer requires administrator privileges.
    echo Please run as administrator or with sudo.
    pause
    exit /b 1
)

echo ==========================================
echo ShortcutIndexer Installer
echo ==========================================
echo.

:: Set installation directory
set "INSTALL_DIR=%ProgramFiles%\ShortcutIndexer"
set "TEMP_BUILD=%TEMP%\ShortcutIndexer_Build"

echo Creating installation directory...
if not exist "%INSTALL_DIR%" (
    mkdir "%INSTALL_DIR%"
)

echo Creating temporary build directory...
if exist "%TEMP_BUILD%" (
    rmdir /s /q "%TEMP_BUILD%"
)
mkdir "%TEMP_BUILD%"

:: Use the .NET Framework compiler directly
set "CSC_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC_PATH%" (
    set "CSC_PATH=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)

if not exist "%CSC_PATH%" (
    echo ERROR: .NET Framework compiler not found.
    echo Please install .NET Framework 4.5 or later.
    pause
    exit /b 1
)

echo Found .NET Framework compiler: %CSC_PATH%

:: Compile the main application
echo.
echo Compiling ShortcutIndexer.exe...
"%CSC_PATH%" /target:winexe /out:"%TEMP_BUILD%\ShortcutIndexer.exe" /reference:System.dll /reference:System.Windows.Forms.dll /reference:System.Drawing.dll "%~dp0ShortcutIndexer.cs"
if %errorLevel% neq 0 (
    echo ERROR: Failed to compile ShortcutIndexer.exe
    pause
    exit /b 1
)

:: Compile the shell extension handler
echo Compiling ShortcutIndexerHandler.dll...
"%CSC_PATH%" /target:library /out:"%TEMP_BUILD%\ShortcutIndexerHandler.dll" /reference:System.dll /reference:System.Windows.Forms.dll "%~dp0ShortcutIndexerHandler.cs"
if %errorLevel% neq 0 (
    echo ERROR: Failed to compile ShortcutIndexerHandler.dll
    pause
    exit /b 1
)

:: Copy files to installation directory
echo.
echo Installing files...
copy "%TEMP_BUILD%\ShortcutIndexer.exe" "%INSTALL_DIR%\" >nul
copy "%TEMP_BUILD%\ShortcutIndexerHandler.dll" "%INSTALL_DIR%\" >nul

:: Register the COM component
echo Registering shell extension...
set "REGASM_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
if not exist "%REGASM_PATH%" (
    set "REGASM_PATH=C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
)

if exist "%REGASM_PATH%" (
    "%REGASM_PATH%" "%INSTALL_DIR%\ShortcutIndexerHandler.dll" /codebase /tlb
    if %errorLevel% neq 0 (
        echo WARNING: Failed to register shell extension with RegAsm.
    ) else (
        echo Shell extension registered successfully with RegAsm.
    )
) else (
    echo WARNING: RegAsm.exe not found. Shell extension may not work properly.
)

:: Create uninstaller
echo Creating uninstaller...
copy "%~dp0Uninstaller.bat" "%INSTALL_DIR%\Uninstall.bat" >nul
if %errorLevel% neq 0 (
    echo ERROR: Failed to create uninstaller.
) else (
    :: Check if -noappwiz argument is present
    echo %* | findstr /i "\-noappwiz" >nul
    if %errorLevel% neq 0 (
        :: Add to Windows Add/Remove Programs (only if -noappwiz is NOT present)
        echo Adding to Add/Remove Programs...
        reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\ShortcutIndexer" /v "DisplayName" /t REG_SZ /d "ShortcutIndexer" /f >nul
        reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\ShortcutIndexer" /v "UninstallString" /t REG_SZ /d "\"%INSTALL_DIR%\Uninstall.bat\"" /f >nul
        reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\ShortcutIndexer" /v "DisplayVersion" /t REG_SZ /d "1.10" /f >nul
        reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\ShortcutIndexer" /v "Publisher" /t REG_SZ /d "ShortcutIndexer" /f >nul
        reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\ShortcutIndexer" /v "InstallLocation" /t REG_SZ /d "%INSTALL_DIR%" /f >nul
    ) else (
        echo Skipping Add/Remove Programs registration (-noappwiz specified)...
    )
)

:: Clean up temporary files
echo Cleaning up temporary files...
rmdir /s /q "%TEMP_BUILD%" 2>nul

echo.
echo ==========================================
echo Installation completed successfully!
echo ==========================================
echo.
echo ShortcutIndexer has been installed to:
echo %INSTALL_DIR%
echo.
echo You can now right-click any file to see the
echo "Shortcut Indexer" context menu.
echo.
echo To uninstall, run: %INSTALL_DIR%\Uninstall.bat
echo or use Add/Remove Programs in Windows Settings.
echo.
pause