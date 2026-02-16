'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** RunDefaultValueTests
' **DATE CREATED:** 2025-12-29
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Test runner for all default value detection tests.
' Runs both unit tests and integration tests for the bugfix.
'-----------------------------------------------------------------------------------

Option Explicit

Sub RunAllTests()
    WScript.Echo "Default Value Detection Test Runner"
    WScript.Echo "==================================="
    WScript.Echo "Running all tests for the default value detection bugfix..."
    WScript.Echo ""

    Dim shell, exec, testResults(1), i
    Set shell = CreateObject("WScript.Shell")
    
    ' Test 1: Unit tests for the HasDefaultValueInPrompt function
    WScript.Echo "1. Running unit tests..."
    WScript.Echo String(40, "-")
    Set exec = shell.Exec("cscript.exe test_default_value_detection.vbs")
    
    ' Wait for completion and capture exit code
    Do While exec.Status = 0
        WScript.Sleep 100
    Loop
    
    testResults(0) = exec.ExitCode
    If exec.ExitCode = 0 Then
        WScript.Echo "Unit tests: PASSED"
    Else
        WScript.Echo "Unit tests: FAILED (Exit code: " & exec.ExitCode & ")"
    End If
    WScript.Echo ""
    
    ' Test 2: Integration tests with MockBzhao
    WScript.Echo "2. Running integration tests..."
    WScript.Echo String(40, "-")
    Set exec = shell.Exec("cscript.exe test_default_value_integration.vbs")
    
    ' Wait for completion and capture exit code
    Do While exec.Status = 0
        WScript.Sleep 100
    Loop
    
    testResults(1) = exec.ExitCode
    If exec.ExitCode = 0 Then
        WScript.Echo "Integration tests: PASSED"
    Else
        WScript.Echo "Integration tests: FAILED (Exit code: " & exec.ExitCode & ")"
    End If
    WScript.Echo ""
    
    ' Summary
    WScript.Echo "Test Summary"
    WScript.Echo "============"
    Dim totalTests, passedTests
    totalTests = 2
    passedTests = 0
    
    For i = 0 To UBound(testResults)
        If testResults(i) = 0 Then passedTests = passedTests + 1
    Next
    
    WScript.Echo "Total tests: " & totalTests
    WScript.Echo "Passed: " & passedTests
    WScript.Echo "Failed: " & (totalTests - passedTests)
    
    If passedTests = totalTests Then
        WScript.Echo ""
        WScript.Echo "SUCCESS: All tests passed! The default value detection bugfix is working correctly."
    Else
        WScript.Echo ""
        WScript.Echo "FAILURE: Some tests failed. Please review the output above."
    End If
End Sub

Sub ShowTestInstructions()
    WScript.Echo "How to run these tests:"
    WScript.Echo ""
    WScript.Echo "1. Manual test runs:"
    WScript.Echo "   cd tests"
    WScript.Echo "   cscript test_default_value_detection.vbs"
    WScript.Echo "   cscript test_default_value_integration.vbs"
    WScript.Echo ""
    WScript.Echo "2. Run all tests:"
    WScript.Echo "   cd tests"
    WScript.Echo "   cscript run_default_value_tests.vbs"
    WScript.Echo ""
    WScript.Echo "3. Integration with existing test suite:"
    WScript.Echo "   Add these tests to run_all_tests.vbs in the parent directory"
End Sub

' Main entry point
Sub Main()
    Dim args
    Set args = WScript.Arguments
    
    If args.Count > 0 And args(0) = "--help" Then
        ShowTestInstructions()
    Else
        RunAllTests()
    End If
End Sub

Main()