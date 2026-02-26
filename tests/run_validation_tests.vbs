' ============================================================================
' CDK Validation Test Suite Runner
' Purpose: Run all validation tests (positive and negative)
' Usage: cscript.exe run_validation_tests.vbs
' ============================================================================

Option Explicit

Dim g_fso
Dim g_shell
Dim g_overallPassed
Dim g_overallFailed

Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_shell = CreateObject("WScript.Shell")
g_overallPassed = 0
g_overallFailed = 0

WScript.Echo vbNewLine & "=" & String(76, "=")
WScript.Echo "CDK VALIDATION TEST SUITE RUNNER"
WScript.Echo "=" & String(76, "=") & vbNewLine

' Determine repo root
Dim repoRoot
On Error Resume Next
repoRoot = g_shell.Environment("USER")("CDK_BASE")
On Error GoTo 0

If repoRoot = "" Then
    WScript.Echo "ERROR: CDK_BASE environment variable not set"
    WScript.Echo "Please run: tools\setup_cdk_base.vbs"
    WScript.Quit 1
End If

' Preflight reset: normalize mutable state from prior runs
RunTest "Preflight Reset (normalize test state)", "test_reset_state.vbs"

' Run positive tests
RunTest "Positive Tests (all dependencies present)", "test_validation_positive.vbs"

' Run negative tests
RunTest "Negative Tests (missing dependencies)", "test_validation_negative.vbs"

' Run reorg baseline contract tests
RunTest "Reorg Contract: Entrypoints", "test_reorg_contract_entrypoints.vbs"
RunTest "Reorg Contract: Config Paths", "test_reorg_contract_config_paths.vbs"

' Print overall summary
PrintOverallSummary()

If g_overallFailed > 0 Then
    WScript.Quit 1
End If

' ============================================================================
' Helper Functions
' ============================================================================

Sub RunTest(testName, scriptFile)
    WScript.Echo ""
    WScript.Echo "-" & String(74, "-")
    WScript.Echo "Running: " & testName
    WScript.Echo "-" & String(74, "-")
    
    ' Tests are in the same directory as this runner script (tests/)
    Dim testsPath
    testsPath = g_fso.GetParentFolderName(WScript.ScriptFullName)
    
    Dim scriptPath
    scriptPath = g_fso.BuildPath(testsPath, scriptFile)
    
    If Not g_fso.FileExists(scriptPath) Then
        WScript.Echo "ERROR: Test script not found: " & scriptPath
        g_overallFailed = g_overallFailed + 1
        Exit Sub
    End If
    
    ' Run the test script
    Dim cmd
    Dim exitCode
    cmd = "cscript.exe " & Chr(34) & scriptPath & Chr(34)
    
    On Error Resume Next
    exitCode = g_shell.Run(cmd, 1, True)
    On Error GoTo 0
    
    If exitCode = 0 Then
        g_overallPassed = g_overallPassed + 1
    Else
        g_overallFailed = g_overallFailed + 1
    End If
End Sub

Sub PrintOverallSummary()
    WScript.Echo ""
    WScript.Echo "=" & String(76, "=")
    WScript.Echo "OVERALL TEST RESULTS"
    WScript.Echo "=" & String(76, "=")
    WScript.Echo "  ✓ Test suites passed: " & g_overallPassed
    WScript.Echo "  ✗ Test suites failed: " & g_overallFailed
    
    If g_overallFailed = 0 Then
        WScript.Echo ""
        WScript.Echo "SUCCESS: All validation tests passed!"
        WScript.Echo "The validation system is working correctly."
    Else
        WScript.Echo ""
        WScript.Echo "FAILURE: Some test suites failed."
        WScript.Echo "Review output above for details."
    End If
    
    WScript.Echo "=" & String(76, "=") & vbNewLine
End Sub
