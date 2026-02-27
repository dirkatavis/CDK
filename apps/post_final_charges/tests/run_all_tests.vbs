'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** RunAllTests
' **DATE CREATED:** 2025-11-19
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Comprehensive test runner for the MockBzhao testing suite.
' Runs all available mock tests and reports results.
'-----------------------------------------------------------------------------------

Option Explicit

Dim fso, shell, testResults, totalTests, passedTests
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
testResults = ""
totalTests = 0
passedTests = 0

Sub RunTest(testName, command)
    totalTests = totalTests + 1
    WScript.Echo "Running " & testName & "..."

    Dim exec, exitCode, output, errors
    Set exec = shell.Exec(command)
    Do While exec.Status = 0
        WScript.Sleep 100
    Loop
    exitCode = exec.ExitCode

    ' Capture output and error streams
    output = ""
    errors = ""
    If Not exec.StdOut.AtEndOfStream Then output = exec.StdOut.ReadAll
    If Not exec.StdErr.AtEndOfStream Then errors = exec.StdErr.ReadAll

    If exitCode = 0 Then
        passedTests = passedTests + 1
        testResults = testResults & "[PASS] " & testName & ": PASSED" & vbCrLf
        WScript.Echo "[PASS] PASSED"
    Else
        testResults = testResults & "[FAIL] " & testName & ": FAILED (Exit code: " & exitCode & ")" & vbCrLf
        WScript.Echo "[FAIL] FAILED (Exit code: " & exitCode & ")"
        If Len(errors) > 0 Then
            WScript.Echo "  Error: " & Trim(errors)
        End If
        If Len(output) > 0 And InStr(output, "Error") > 0 Then
            WScript.Echo "  Output: " & Left(Trim(output), 200) & "..."
        End If
    End If
End Sub

Sub Main()
    WScript.Echo "=== PostFinalCharges Testing Suite ==="
    WScript.Echo ""

    ' Change working directory to current script folder
    Dim scriptDir
    scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
    shell.CurrentDirectory = scriptDir

    ' Test 1: Standalone Mock Test
    RunTest "Standalone Mock Test", "cscript.exe test_mock_bzhao.vbs"

    ' Test 2: Integration Test
    RunTest "Integration Test", "cscript.exe test_integration.vbs"

    ' Test 3: Default Value Detection Unit Tests (NEW BUGFIX)
    RunTest "Default Value Detection Tests", "cscript.exe test_default_value_detection.vbs"

    ' Test 4: Default Value Integration Tests (NEW BUGFIX)
    RunTest "Default Value Integration Tests", "cscript.exe test_default_value_integration.vbs"

    ' Test 5: OPEN Status Tests (NEW FEATURE)
    RunTest "OPEN Status Tests", "cscript.exe test_open_status.vbs"

    ' Test 6: Pattern Issue Reproduction Test (TDD)
    RunTest "Pattern Issue Reproduction Test", "cscript.exe test_reproduce_pattern_issue.vbs"

    ' Test 7: InStr Bug Reproduction Test (Production Issue)
    RunTest "InStr Bug Reproduction Test", "cscript.exe test_reproduce_instr_bug.vbs"

    ' Test 8: Bug Fix Verification Test (Production Fix)
    RunTest "Bug Fix Verification Test", "cscript.exe test_verify_bug_fix.vbs"

    ' Test 9: Comprehensive InStr Bug Scan
    RunTest "Comprehensive InStr Bug Scan", "cscript.exe test_scan_all_instr_bugs.vbs"

    ' Test 10: All InStr Fixes Verification
    RunTest "All InStr Fixes Verification", "cscript.exe test_verify_all_instr_fixes.vbs"

    ' Summary
    WScript.Echo ""
    WScript.Echo "=== Test Results Summary ==="
    WScript.Echo testResults
    WScript.Echo "Total Tests: " & totalTests
    WScript.Echo "Passed: " & passedTests
    WScript.Echo "Failed: " & (totalTests - passedTests)

    If passedTests = totalTests Then
        WScript.Echo ""
        WScript.Echo "SUCCESS: ALL TESTS PASSED! MockBzhao is working correctly."
    Else
        WScript.Echo ""
        WScript.Echo "WARNING: Some tests failed. Check the output above for details."
        WScript.Quit 1
    End If
End Sub

Main()