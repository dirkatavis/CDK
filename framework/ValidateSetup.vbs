' ============================================================================
' Common Validation Routines
' Purpose: Shared validation logic for all CDK scripts
' Usage: Include this in any script that needs dependency validation
' NOTE: Works in both WScript and BlueZone contexts (no WScript.Echo)
' ============================================================================

Option Explicit

' ============================================================================
' ValidateScriptDependencies
' Checks if all required dependencies are present before script execution
' Returns: True if all checks pass, False otherwise
' Works in: Both standalone (WScript) and BlueZone contexts
' ============================================================================

Function ValidateScriptDependencies()
    Dim fso
    Dim shell
    Dim repoRoot
    Dim failures
    Dim warnings
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    
    failures = 0
    warnings = 0
    
    ' Check 1: Get repo root from CDK_BASE
    On Error Resume Next
    repoRoot = shell.Environment("USER")("CDK_BASE")
    On Error GoTo 0
    
    If repoRoot = "" Or IsNull(repoRoot) Then
        Call SafeOutput("ERROR: CDK_BASE environment variable not set.")
        Call SafeOutput("Please run: cscript.exe tools\setup_cdk_base.vbs")
        failures = failures + 1
        ValidateScriptDependencies = False
        Exit Function
    End If
    
    If Not fso.FolderExists(repoRoot) Then
        Call SafeOutput("ERROR: CDK_BASE points to non-existent folder: " & repoRoot)
        failures = failures + 1
        ValidateScriptDependencies = False
        Exit Function
    End If
    
    ' Check 2: Verify .cdkroot marker exists
    Dim markerPath
    markerPath = fso.BuildPath(repoRoot, ".cdkroot")
    If Not fso.FileExists(markerPath) Then
        Call SafeOutput("WARNING: Repository marker (.cdkroot) not found at " & repoRoot)
        warnings = warnings + 1
    End If
    
    ' Check 3: Verify PathHelper.vbs exists
    Dim pathHelperPath
    pathHelperPath = fso.BuildPath(repoRoot, "common\PathHelper.vbs")
    If Not fso.FileExists(pathHelperPath) Then
        Call SafeOutput("ERROR: PathHelper.vbs not found at " & pathHelperPath)
        failures = failures + 1
        ValidateScriptDependencies = False
        Exit Function
    End If
    
    ' Check 4: Verify config.ini exists
    Dim configPath
    configPath = fso.BuildPath(repoRoot, "config\config.ini")
    If Not fso.FileExists(configPath) Then
        Call SafeOutput("ERROR: config.ini not found at " & configPath)
        failures = failures + 1
        ValidateScriptDependencies = False
        Exit Function
    End If
    
    If failures > 0 Then
        ValidateScriptDependencies = False
    Else
        If warnings > 0 Then
            Call SafeOutput("Dependencies validated with " & warnings & " warning(s).")
        Else
            Call SafeOutput("All dependencies validated successfully.")
        End If
        ValidateScriptDependencies = True
    End If
End Function

' ============================================================================
' SafeOutput
' Outputs message in both WScript and BlueZone contexts
' In WScript: Uses WScript.Echo
' In BlueZone: Uses existing LogInfo if available, else silent
' ============================================================================

Sub SafeOutput(msg)
    ' Try WScript first (standalone context)
    On Error Resume Next
    WScript.Echo msg
    On Error GoTo 0
    
    ' If we're in BlueZone and LogInfo exists, try that too
    On Error Resume Next
    If g_CurrentCriticality >= 0 Then
        LogInfo msg, "Validation"
    End If
    On Error GoTo 0
End Sub

' ============================================================================
' GetRepoRootSafe
' Returns the repo root or empty string if not available
' ============================================================================

Function GetRepoRootSafe()
    Dim shell
    Dim repoRoot
    
    Set shell = CreateObject("WScript.Shell")
    
    On Error Resume Next
    repoRoot = shell.Environment("USER")("CDK_BASE")
    On Error GoTo 0
    
    If repoRoot = "" Or IsNull(repoRoot) Then
        GetRepoRootSafe = ""
    Else
        GetRepoRootSafe = repoRoot
    End If
End Function

' ============================================================================
' MustHaveValidDependencies
' Enforces strict validation; terminates/aborts if any check fails
' Used by critical scripts (PostFinalCharges, etc.)
' Works in: Both standalone (WScript.Quit) and BlueZone (sets abort flag)
' ============================================================================

Sub MustHaveValidDependencies()
    If Not ValidateScriptDependencies() Then
        Call SafeOutput("")
        Call SafeOutput("FATAL: Required dependencies not available.")
        Call SafeOutput("Run tools\validate_dependencies.vbs for detailed diagnostics.")
        
        ' Try to exit as standalone VBScript
        On Error Resume Next
        WScript.Quit 1
        On Error GoTo 0
        
        ' If WScript.Quit failed, we're in BlueZone - set abort flag
        On Error Resume Next
        g_ShouldAbort = True
        g_AbortReason = "Dependency validation failed"
        On Error GoTo 0
    End If
End Sub
