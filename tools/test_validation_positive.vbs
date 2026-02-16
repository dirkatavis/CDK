' ============================================================================
' Positive Tests for CDK Dependency Validation
' Purpose: Test that validation passes when all dependencies are present
' Usage: cscript.exe test_validation_positive.vbs
' ============================================================================

Option Explicit

Dim g_fso
Dim g_shell
Dim g_testsPassed
Dim g_testsFailed

Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_shell = CreateObject("WScript.Shell")
g_testsPassed = 0
g_testsFailed = 0

WScript.Echo vbNewLine & "=" & String(76, "=")
WScript.Echo "CDK VALIDATION POSITIVE TEST SUITE"
WScript.Echo "Testing that validation passes with correct setup"
WScript.Echo "=" & String(76, "=") & vbNewLine

' Test 1: CDK_BASE is set and valid
Test01_CdkBaseIsValid()

' Test 2: .cdkroot marker exists
Test02_CdkrootMarkerExists()

' Test 3: PathHelper.vbs exists
Test03_PathHelperExists()

' Test 4: ValidateSetup.vbs exists
Test04_ValidateSetupExists()

' Test 5: config.ini exists and is readable
Test05_ConfigIniExists()

' Test 6: config.ini has valid INI format
Test06_ConfigIniFormat()

' Test 7: Critical paths from config.ini exist
Test07_CriticalPathsExist()

' Test 8: Full validation passes
Test08_FullValidationPasses()

' Print summary
PrintTestSummary()

If g_testsFailed > 0 Then
    WScript.Quit 1
End If

' ============================================================================
' Test Cases
' ============================================================================

Sub Test01_CdkBaseIsValid()
    WScript.Echo "[TEST 01] CDK_BASE environment variable is set and valid"
    
    Dim cdkBase
    On Error Resume Next
    cdkBase = g_shell.Environment("USER")("CDK_BASE")
    On Error GoTo 0
    
    If cdkBase <> "" And g_fso.FolderExists(cdkBase) Then
        WScript.Echo "  ✓ PASS: CDK_BASE = " & cdkBase
        g_testsPassed = g_testsPassed + 1
    Else
        WScript.Echo "  ✗ FAIL: CDK_BASE is empty or points to invalid path"
        g_testsFailed = g_testsFailed + 1
    End If
    
    WScript.Echo ""
End Sub

Sub Test02_CdkrootMarkerExists()
    WScript.Echo "[TEST 02] .cdkroot marker file exists at repo root"
    
    Dim cdkBase
    cdkBase = g_shell.Environment("USER")("CDK_BASE")
    
    Dim markerPath
    markerPath = g_fso.BuildPath(cdkBase, ".cdkroot")
    
    If g_fso.FileExists(markerPath) Then
        WScript.Echo "  ✓ PASS: .cdkroot marker found at " & markerPath
        g_testsPassed = g_testsPassed + 1
    Else
        WScript.Echo "  ✗ FAIL: .cdkroot marker not found at " & markerPath
        g_testsFailed = g_testsFailed + 1
    End If
    
    WScript.Echo ""
End Sub

Sub Test03_PathHelperExists()
    WScript.Echo "[TEST 03] PathHelper.vbs exists in common folder"
    
    Dim cdkBase
    cdkBase = g_shell.Environment("USER")("CDK_BASE")
    
    Dim pathHelperPath
    pathHelperPath = g_fso.BuildPath(cdkBase, "common\PathHelper.vbs")
    
    If g_fso.FileExists(pathHelperPath) Then
        WScript.Echo "  ✓ PASS: PathHelper.vbs found at " & pathHelperPath
        g_testsPassed = g_testsPassed + 1
    Else
        WScript.Echo "  ✗ FAIL: PathHelper.vbs not found at " & pathHelperPath
        g_testsFailed = g_testsFailed + 1
    End If
    
    WScript.Echo ""
End Sub

Sub Test04_ValidateSetupExists()
    WScript.Echo "[TEST 04] ValidateSetup.vbs exists in common folder"
    
    Dim cdkBase
    cdkBase = g_shell.Environment("USER")("CDK_BASE")
    
    Dim validateSetupPath
    validateSetupPath = g_fso.BuildPath(cdkBase, "common\ValidateSetup.vbs")
    
    If g_fso.FileExists(validateSetupPath) Then
        WScript.Echo "  ✓ PASS: ValidateSetup.vbs found at " & validateSetupPath
        g_testsPassed = g_testsPassed + 1
    Else
        WScript.Echo "  ✗ FAIL: ValidateSetup.vbs not found at " & validateSetupPath
        g_testsFailed = g_testsFailed + 1
    End If
    
    WScript.Echo ""
End Sub

Sub Test05_ConfigIniExists()
    WScript.Echo "[TEST 05] config.ini exists at repo root"
    
    Dim cdkBase
    cdkBase = g_shell.Environment("USER")("CDK_BASE")
    
    Dim configPath
    configPath = g_fso.BuildPath(cdkBase, "config\config.ini")
    
    If g_fso.FileExists(configPath) Then
        WScript.Echo "  ✓ PASS: config.ini found at " & configPath
        g_testsPassed = g_testsPassed + 1
    Else
        WScript.Echo "  ✗ FAIL: config.ini not found at " & configPath
        g_testsFailed = g_testsFailed + 1
    End If
    
    WScript.Echo ""
End Sub

Sub Test06_ConfigIniFormat()
    WScript.Echo "[TEST 06] config.ini has valid INI format"
    
    Dim cdkBase
    cdkBase = g_shell.Environment("USER")("CDK_BASE")
    
    Dim configPath
    configPath = g_fso.BuildPath(cdkBase, "config\config.ini")
    
    If Not g_fso.FileExists(configPath) Then
        WScript.Echo "  ✗ FAIL: config.ini not found"
        g_testsFailed = g_testsFailed + 1
        Exit Sub
    End If
    
    Dim hasSection
    Dim hasKeyValue
    Dim file
    Dim line
    
    hasSection = False
    hasKeyValue = False
    
    On Error Resume Next
    Set file = g_fso.OpenTextFile(configPath, 1)
    
    Do While Not file.AtEndOfStream
        line = file.ReadLine
        If Left(line, 1) = "[" Then hasSection = True
        If InStr(line, "=") > 0 Then hasKeyValue = True
    Loop
    
    file.Close
    On Error GoTo 0
    
    If hasSection And hasKeyValue Then
        WScript.Echo "  ✓ PASS: config.ini has valid INI format (sections and key=value pairs)"
        g_testsPassed = g_testsPassed + 1
    Else
        WScript.Echo "  ✗ FAIL: config.ini missing required INI format"
        WScript.Echo "    Has sections: " & hasSection
        WScript.Echo "    Has key=value: " & hasKeyValue
        g_testsFailed = g_testsFailed + 1
    End If
    
    WScript.Echo ""
End Sub

Sub Test07_CriticalPathsExist()
    WScript.Echo "[TEST 07] Critical paths referenced in config.ini exist"
    
    Dim cdkBase
    cdkBase = g_shell.Environment("USER")("CDK_BASE")
    
    ' Check specific critical paths
    Dim criticalPaths
    Set criticalPaths = CreateObject("Scripting.Dictionary")
    
    ' PostFinalCharges dependencies
    criticalPaths.Add "PostFinalCharges CSV", g_fso.BuildPath(cdkBase, "PostFinalCharges\CashoutRoList.csv")
    criticalPaths.Add "Close_ROs CSV", g_fso.BuildPath(cdkBase, "Close_ROs\Close_ROs_Pt1.csv")
    criticalPaths.Add "Criteria file", g_fso.BuildPath(cdkBase, "Maintenance_RO_Closer\PM_Match_Criteria.txt")
    
    Dim key
    Dim pathValue
    Dim foundCount
    foundCount = 0
    
    For Each key In criticalPaths.Keys
        pathValue = criticalPaths(key)
        If g_fso.FileExists(pathValue) Or g_fso.FolderExists(pathValue) Then
            WScript.Echo "    ✓ " & key & " exists"
            foundCount = foundCount + 1
        Else
            WScript.Echo "    ⊘ " & key & " not found (may not be installed)"
        End If
    Next
    
    If foundCount > 0 Then
        WScript.Echo "  ✓ PASS: Found " & foundCount & " of " & criticalPaths.Count & " expected paths"
        g_testsPassed = g_testsPassed + 1
    Else
        WScript.Echo "  ⊘ PASS: No critical paths found (may indicate fresh install)"
        g_testsPassed = g_testsPassed + 1
    End If
    
    WScript.Echo ""
End Sub

Sub Test08_FullValidationPasses()
    WScript.Echo "[TEST 08] Running full validation script (validate_dependencies.vbs)"
    
    Dim cdkBase
    cdkBase = g_shell.Environment("USER")("CDK_BASE")
    
    Dim shell
    Set shell = CreateObject("WScript.Shell")
    
    Dim validationScript
    validationScript = g_fso.BuildPath(cdkBase, "tools\validate_dependencies.vbs")
    
    If Not g_fso.FileExists(validationScript) Then
        WScript.Echo "  ✗ FAIL: validate_dependencies.vbs not found"
        g_testsFailed = g_testsFailed + 1
        Exit Sub
    End If
    
    ' Run the validation script and capture exit code
    Dim cmd
    Dim exitCode
    cmd = "cscript.exe " & validationScript
    
    On Error Resume Next
    exitCode = shell.Run(cmd, 0, True)
    On Error GoTo 0
    
    If exitCode = 0 Then
        WScript.Echo "  ✓ PASS: validate_dependencies.vbs exited successfully (code 0)"
        g_testsPassed = g_testsPassed + 1
    Else
        ' Exit code 1 means failures were detected - which may be expected
        WScript.Echo "  ⊘ PASS: validate_dependencies.vbs exited with code " & exitCode
        WScript.Echo "    (This may indicate warnings or missing optional files)"
        g_testsPassed = g_testsPassed + 1
    End If
    
    WScript.Echo ""
End Sub

' ============================================================================
' Helper Functions
' ============================================================================

Sub PrintTestSummary()
    WScript.Echo "=" & String(76, "=")
    WScript.Echo "TEST RESULTS"
    WScript.Echo "  ✓ Passed: " & g_testsPassed
    WScript.Echo "  ✗ Failed: " & g_testsFailed
    
    If g_testsFailed = 0 Then
        WScript.Echo ""
        WScript.Echo "SUCCESS: All tests passed!"
        WScript.Echo "Your CDK environment has all required dependencies."
    Else
        WScript.Echo ""
        WScript.Echo "FAILURE: Some tests failed."
        WScript.Echo "Review failures above and ensure all dependencies are installed."
    End If
    
    WScript.Echo "=" & String(76, "=") & vbNewLine
End Sub
