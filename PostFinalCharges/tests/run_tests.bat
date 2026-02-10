@echo off
REM Test runner that ensures correct working directory
cd /d "%~dp0"
echo Running tests from: %CD%
cscript run_all_tests.vbs
pause