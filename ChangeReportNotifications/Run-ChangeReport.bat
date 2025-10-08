@echo off
REM Batch file wrapper for Change Report Email Notifications
REM Designed for Windows Task Scheduler integration
REM
REM Usage:
REM   Run-ChangeReport.bat [normal|test|force] [config_path] [report_date]
REM
REM Parameters:
REM   Mode: normal (default), test, force
REM   Config Path: Optional path to config file (default: .\config\config.json)
REM   Report Date: Optional date in YYYY-MM-DD format (default: current date)
REM
REM Examples:
REM   Run-ChangeReport.bat
REM   Run-ChangeReport.bat test
REM   Run-ChangeReport.bat normal "C:\Config\prod.json"
REM   Run-ChangeReport.bat normal ".\config\config.json" "2024-01-15"

setlocal enabledelayedexpansion

REM Set default values
set "MODE=%~1"
set "CONFIG_PATH=%~2"
set "REPORT_DATE=%~3"

REM Default mode to normal if not specified
if "%MODE%"=="" set "MODE=normal"

REM Get script directory
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM Set log file for batch execution
set "BATCH_LOG=%SCRIPT_DIR%logs\batch_execution_%date:~-4,4%%date:~-10,2%%date:~-7,2%_%time:~0,2%%time:~3,2%%time:~6,2%.log"

REM Create logs directory if it doesn't exist
if not exist "%SCRIPT_DIR%logs" mkdir "%SCRIPT_DIR%logs"

REM Log batch execution start
echo [%date% %time%] Starting Change Report batch execution >> "%BATCH_LOG%"
echo [%date% %time%] Mode: %MODE% >> "%BATCH_LOG%"
echo [%date% %time%] Config Path: %CONFIG_PATH% >> "%BATCH_LOG%"
echo [%date% %time%] Report Date: %REPORT_DATE% >> "%BATCH_LOG%"
echo [%date% %time%] Working Directory: %SCRIPT_DIR% >> "%BATCH_LOG%"

REM Build PowerShell command based on parameters
set "PS_COMMAND=powershell.exe -ExecutionPolicy Bypass -NoProfile -File ""%SCRIPT_DIR%Send-ChangeReport.ps1"""

REM Add parameters based on mode and arguments
if /i "%MODE%"=="test" (
    set "PS_COMMAND=!PS_COMMAND! -TestMode"
    echo [%date% %time%] Running in TEST mode >> "%BATCH_LOG%"
)

if /i "%MODE%"=="force" (
    set "PS_COMMAND=!PS_COMMAND! -Force"
    echo [%date% %time%] Running in FORCE mode >> "%BATCH_LOG%"
)

if not "%CONFIG_PATH%"=="" (
    set "PS_COMMAND=!PS_COMMAND! -ConfigPath ""%CONFIG_PATH%"""
    echo [%date% %time%] Using config path: %CONFIG_PATH% >> "%BATCH_LOG%"
)

if not "%REPORT_DATE%"=="" (
    set "PS_COMMAND=!PS_COMMAND! -ReportDate ""%REPORT_DATE%"""
    echo [%date% %time%] Using report date: %REPORT_DATE% >> "%BATCH_LOG%"
)

REM Add verbose logging for scheduled execution
set "PS_COMMAND=!PS_COMMAND! -Verbose"

echo [%date% %time%] Executing PowerShell command: !PS_COMMAND! >> "%BATCH_LOG%"

REM Execute PowerShell script and capture exit code
!PS_COMMAND!
set "EXIT_CODE=%ERRORLEVEL%"

REM Log execution result
echo [%date% %time%] PowerShell script completed with exit code: %EXIT_CODE% >> "%BATCH_LOG%"

REM Set appropriate exit messages and codes for Task Scheduler
if %EXIT_CODE% equ 0 (
    echo [%date% %time%] SUCCESS: Change Report process completed successfully >> "%BATCH_LOG%"
    echo Change Report Email Notification completed successfully
) else (
    echo [%date% %time%] ERROR: Change Report process failed with exit code %EXIT_CODE% >> "%BATCH_LOG%"
    echo ERROR: Change Report Email Notification failed - Check logs for details
)

REM Clean up old batch log files (keep last 30 days)
forfiles /p "%SCRIPT_DIR%logs" /m "batch_execution_*.log" /d -30 /c "cmd /c del @path" 2>nul

echo [%date% %time%] Batch execution completed >> "%BATCH_LOG%"

REM Exit with the same code as PowerShell script for monitoring
exit /b %EXIT_CODE%