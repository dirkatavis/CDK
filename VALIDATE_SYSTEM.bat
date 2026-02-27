@echo off
REM ============================================================================
REM CDK Grand Validation - Master Health Check
REM This script runs the complete system validation suite including infrastructure,
REM application logic, and stress tests.
REM ============================================================================

echo =============================================================================
echo INITIALIZING CDK VALIDATION...
echo =============================================================================

cscript.exe //nologo tests\run_validation_tests.vbs

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] System Unhealthy. Review the red [FAIL] flags above.
    echo.
    if "%1" neq "/silent" pause
    exit /b 1
)

echo.
echo [SUCCESS] All systems functional.
echo.
if "%1" neq "/silent" pause
exit /b 0
