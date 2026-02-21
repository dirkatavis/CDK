' ============================================================================
' CDK Dependency Validator
' Purpose: Pre-flight check for CDK automation system setup
' Usage: cscript.exe validate_dependencies.vbs
' ============================================================================

Option Explicit

Dim g_fso
Dim g_shell
Dim g_repoRoot
Dim g_failureCount
Dim g_warningCount

Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_shell = CreateObject("WScript.Shell")
g_failureCount = 0
g_warningCount = 0

' Colors for console output
Const COLOR_DEFAULT = 0
Const COLOR_GREEN = 2
Const COLOR_RED = 4
Const COLOR_YELLOW = 6

' ============================================================================
' Main Execution
' ============================================================================

WScript.Echo vbNewLine & "=" & String(76, "=")
WScript.Echo "CDK DEPENDENCY VALIDATOR"
WScript.Echo "=" & String(76, "=") & vbNewLine

' Check 1: CDK_BASE environment variable
CheckCdkBaseVariable()

' Check 2: .cdkroot marker file
CheckCdkRootMarker()

' Check 3: PathHelper.vbs
CheckPathHelper()

' Check 4: config.ini
CheckConfigIni()

' Check 5: All paths referenced in config.ini
If g_repoRoot <> "" Then
    CheckConfigPaths()
End If

' Output summary
PrintSummary()

If g_failureCount > 0 Then
    WScript.Quit 1
End If

WScript.Quit 0

' ============================================================================
' Check Functions
' ============================================================================

Sub CheckCdkBaseVariable()
    WScript.Echo "[CHECK 1/5] CDK_BASE Environment Variable"
    
    On Error Resume Next
    Dim cdkBase
    cdkBase = g_shell.Environment("USER")("CDK_BASE")
    On Error GoTo 0
    
    If cdkBase = "" Or IsNull(cdkBase) Then
        AddFailure "CDK_BASE environment variable not set"
        WScript.Echo "  ✗ FAILURE: CDK_BASE environment variable not found"
        WScript.Echo "  Remediation: Run tools\setup_cdk_base.vbs to set it up"
        g_failureCount = g_failureCount + 1
    Else
        If g_fso.FolderExists(cdkBase) Then
            WScript.Echo "  ✓ PASS: CDK_BASE = " & cdkBase
            g_repoRoot = cdkBase
        Else
            AddFailure "CDK_BASE is set but points to non-existent folder: " & cdkBase
            WScript.Echo "  ✗ FAILURE: CDK_BASE points to non-existent folder"
            WScript.Echo "  Path: " & cdkBase
            WScript.Echo "  Remediation: Update CDK_BASE to point to valid CDK repo root"
            g_failureCount = g_failureCount + 1
        End If
    End If
    
    WScript.Echo ""
End Sub

Sub CheckCdkRootMarker()
    WScript.Echo "[CHECK 2/5] Repository Root Marker (.cdkroot)"
    
    If g_repoRoot = "" Then
        AddWarning "Cannot check .cdkroot (CDK_BASE not set)"
        WScript.Echo "  ⊘ SKIP: CDK_BASE not set, skipping marker check"
        WScript.Echo ""
        Exit Sub
    End If
    
    Dim markerPath
    markerPath = g_fso.BuildPath(g_repoRoot, ".cdkroot")
    
    If g_fso.FileExists(markerPath) Then
        WScript.Echo "  ✓ PASS: .cdkroot marker found at repo root"
    Else
        AddWarning "Repository root marker not found"
        WScript.Echo "  ⊘ WARNING: .cdkroot marker not found at " & g_repoRoot
        WScript.Echo "  This file should exist to identify the repo root"
        g_warningCount = g_warningCount + 1
    End If
    
    WScript.Echo ""
End Sub

Sub CheckPathHelper()
    WScript.Echo "[CHECK 3/5] PathHelper.vbs"
    
    If g_repoRoot = "" Then
        AddWarning "Cannot check PathHelper (CDK_BASE not set)"
        WScript.Echo "  ⊘ SKIP: CDK_BASE not set, skipping PathHelper check"
        WScript.Echo ""
        Exit Sub
    End If
    
    Dim pathHelperPath
    pathHelperPath = g_fso.BuildPath(g_repoRoot, "framework\PathHelper.vbs")
    
    If g_fso.FileExists(pathHelperPath) Then
        WScript.Echo "  ✓ PASS: PathHelper.vbs found"
        WScript.Echo "  Path: " & pathHelperPath
    Else
        AddFailure "PathHelper.vbs not found at " & pathHelperPath
        WScript.Echo "  ✗ FAILURE: PathHelper.vbs not found"
        WScript.Echo "  Expected: " & pathHelperPath
        WScript.Echo "  Remediation: Ensure framework\PathHelper.vbs exists in repo root"
        g_failureCount = g_failureCount + 1
    End If
    
    WScript.Echo ""
End Sub

Sub CheckConfigIni()
    WScript.Echo "[CHECK 4/5] config.ini"
    
    If g_repoRoot = "" Then
        AddWarning "Cannot check config.ini (CDK_BASE not set)"
        WScript.Echo "  ⊘ SKIP: CDK_BASE not set, skipping config.ini check"
        WScript.Echo ""
        Exit Sub
    End If
    
    Dim configPath
    configPath = g_fso.BuildPath(g_repoRoot, "config\config.ini")
    
    If g_fso.FileExists(configPath) Then
        WScript.Echo "  ✓ PASS: config.ini found at config/ subfolder"
        WScript.Echo "  Path: " & configPath
        
        ' Validate it's parseable
        If ValidateIniFormat(configPath) Then
            WScript.Echo "  ✓ PASS: config.ini format is valid"
        Else
            AddWarning "config.ini format may be invalid"
            WScript.Echo "  ⊘ WARNING: config.ini format validation inconclusive"
            g_warningCount = g_warningCount + 1
        End If
    Else
        AddFailure "config.ini not found at " & configPath
        WScript.Echo "  ✗ FAILURE: config.ini not found at repo root"
        WScript.Echo "  Expected: " & configPath
        WScript.Echo "  Remediation: Ensure config.ini exists at repo root"
        g_failureCount = g_failureCount + 1
    End If
    
    WScript.Echo ""
End Sub

Sub CheckConfigPaths()
    WScript.Echo "[CHECK 5/5] Config File Path References"
    
    Dim configPath
    configPath = g_fso.BuildPath(g_repoRoot, "config\config.ini")
    
    If Not g_fso.FileExists(configPath) Then
        WScript.Echo "  ⊘ SKIP: config.ini not found, cannot validate paths"
        WScript.Echo ""
        Exit Sub
    End If
    
    Dim pathsToCheck
    Set pathsToCheck = CreateObject("Scripting.Dictionary")
    
    ' Read all path values from config.ini
    ReadConfigPaths configPath, pathsToCheck
    
    If pathsToCheck.Count = 0 Then
        WScript.Echo "  ⊘ WARNING: No path entries found in config.ini"
        g_warningCount = g_warningCount + 1
    Else
        WScript.Echo "  Checking " & pathsToCheck.Count & " configured paths..."
        
        Dim key
        Dim pathValue
        Dim fullPath
        Dim pathExists
        Dim checkCount
        checkCount = 0
        
        For Each key In pathsToCheck.Keys
            pathValue = pathsToCheck(key)
            fullPath = g_fso.BuildPath(g_repoRoot, pathValue)
            
            ' Skip if path contains forward slashes (metadata like StartSequenceNumber)
            If InStr(pathValue, "/") > 0 Or Not ContainsPathSeparator(pathValue) Then
                ' Skip non-path values
            Else
                checkCount = checkCount + 1
                
                If g_fso.FileExists(fullPath) Or g_fso.FolderExists(fullPath) Then
                    WScript.Echo "    ✓ " & key & " exists"
                Else
                    ' For files/folders that don't need to exist yet (like logs), just warn
                    If InStr(LCase(key), "log") > 0 Or _
                       InStr(LCase(key), "output") > 0 Or _
                       InStr(LCase(key), "debugmarker") > 0 Then
                        WScript.Echo "    ⊘ " & key & " (will be created: " & pathValue & ")"
                    Else
                        AddWarning "Path not found: " & key & " = " & pathValue
                        WScript.Echo "    ✗ " & key & " NOT FOUND"
                        WScript.Echo "      Path: " & fullPath
                        g_warningCount = g_warningCount + 1
                    End If
                End If
            End If
        Next
        
        If checkCount = 0 Then
            WScript.Echo "  ⊘ No verifiable paths found in config.ini"
        End If
    End If
    
    WScript.Echo ""
End Sub

' ============================================================================
' Helper Functions
' ============================================================================

Sub ReadConfigPaths(configPath, pathsDict)
    Dim file
    Dim line
    Dim parts
    Dim key
    Dim value
    
    On Error Resume Next
    Set file = g_fso.OpenTextFile(configPath, 1) ' ForReading
    
    Do While Not file.AtEndOfStream
        line = file.ReadLine
        
        ' Skip empty lines and comments
        If line <> "" And Left(line, 1) <> ";" And Left(line, 1) <> "#" Then
            ' Skip section headers
            If Left(line, 1) <> "[" Then
                ' Parse key=value
                If InStr(line, "=") > 0 Then
                    parts = Split(line, "=", 2)
                    If UBound(parts) >= 1 Then
                        key = Trim(parts(0))
                        value = Trim(parts(1))
                        If key <> "" And value <> "" Then
                            pathsDict.Add key, value
                        End If
                    End If
                End If
            End If
        End If
    Loop
    
    file.Close
    On Error GoTo 0
End Sub

Function ValidateIniFormat(configPath)
    ' Simple format check: look for section headers and key=value pairs
    Dim file
    Dim line
    Dim hasSections
    Dim hasKeyValues
    
    On Error Resume Next
    Set file = g_fso.OpenTextFile(configPath, 1)
    
    hasSections = False
    hasKeyValues = False
    
    Do While Not file.AtEndOfStream And Not (hasSections And hasKeyValues)
        line = file.ReadLine
        If Left(line, 1) = "[" Then hasSections = True
        If InStr(line, "=") > 0 Then hasKeyValues = True
    Loop
    
    file.Close
    On Error GoTo 0
    
    ValidateIniFormat = (hasSections And hasKeyValues)
End Function

Function ContainsPathSeparator(value)
    ContainsPathSeparator = (InStr(value, "\") > 0)
End Function

Sub AddFailure(msg)
    ' For future use - currently we track count directly
End Sub

Sub AddWarning(msg)
    ' For future use - currently we track count directly
End Sub

Sub PrintSummary()
    WScript.Echo "=" & String(76, "=")
    
    If g_failureCount = 0 And g_warningCount = 0 Then
        WScript.Echo "✓ ALL CHECKS PASSED"
        WScript.Echo ""
        WScript.Echo "Your CDK environment is ready to run scripts."
        WScript.Echo "Next: Run PostFinalCharges.vbs or other automation scripts."
    ElseIf g_failureCount = 0 Then
        WScript.Echo "⊘ " & g_warningCount & " WARNING(S) DETECTED"
        WScript.Echo ""
        WScript.Echo "Your CDK environment will likely work, but review warnings above."
    Else
        WScript.Echo "✗ " & g_failureCount & " FAILURE(S) DETECTED"
        If g_warningCount > 0 Then
            WScript.Echo "⊘ + " & g_warningCount & " WARNING(S)"
        End If
        WScript.Echo ""
        WScript.Echo "Please address failures above before running scripts."
    End If
    
    WScript.Echo "=" & String(76, "=") & vbNewLine
End Sub
