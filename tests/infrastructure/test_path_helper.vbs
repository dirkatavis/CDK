Option Explicit

' ==============================================================================
' test_path_helper.vbs - Validate PathHelper configuration
' ==============================================================================
' This script tests the centralized path configuration system.
' Run this from BlueZone to verify the setup works correctly.
' ==============================================================================

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
' Use string literal to avoid conflict with PathHelper's BASE_ENV_VAR constant
Dim envVarName: envVarName = "CDK_BASE"
Dim isSilent: isSilent = HasArg("silent")

' Load PathHelper module using CDK_BASE
Dim sh: Set sh = CreateObject("WScript.Shell")
Dim basePath: basePath = sh.Environment("USER")(envVarName)

If basePath = "" Or Not fso.FolderExists(basePath) Then
    If isSilent Then
        WScript.Echo "ERROR: Invalid or missing CDK_BASE. Value: " & basePath
    Else
        MsgBox "ERROR: Invalid or missing CDK_BASE" & vbCrLf & "Value: " & basePath, vbCritical, "Test Failed"
    End If
    WScript.Quit 1
End If

Dim helperPath: helperPath = fso.BuildPath(basePath, "framework\PathHelper.vbs")
If Not fso.FileExists(helperPath) Then
    If isSilent Then
        WScript.Echo "ERROR: Cannot find PathHelper.vbs. Looked at: " & helperPath
    Else
        MsgBox "ERROR: Cannot find PathHelper.vbs" & vbCrLf & "Looked at: " & helperPath, vbCritical, "Test Failed"
    End If
    WScript.Quit 1
End If

ExecuteGlobal fso.OpenTextFile(helperPath).ReadAll

' Test repo root discovery
Dim repoRoot: repoRoot = GetRepoRoot()

' Test config path reading
Dim csvPath, logPath
On Error Resume Next
csvPath = GetConfigPath("Prepare_Close_Pt1", "CSV")
logPath = GetConfigPath("Prepare_Close_Pt1", "Log")
On Error GoTo 0

' Build report
Dim report
report = "CDK Path Helper Test Results" & vbCrLf
report = report & String(50, "=") & vbCrLf & vbCrLf
report = report & "Repo Root: " & repoRoot & vbCrLf
report = report & "  .cdkroot exists: " & fso.FileExists(fso.BuildPath(repoRoot, ".cdkroot")) & vbCrLf
report = report & "  config.ini exists: " & fso.FileExists(fso.BuildPath(repoRoot, "config\config.ini")) & vbCrLf
report = report & vbCrLf
report = report & "Sample Config Paths:" & vbCrLf
report = report & "  Prepare_Close_Pt1 CSV: " & csvPath & vbCrLf
report = report & "  Prepare_Close_Pt1 Log: " & logPath & vbCrLf
report = report & vbCrLf

If csvPath <> "" And logPath <> "" Then
    report = report & "Status: SUCCESS - Path Helper is working!" & vbCrLf
    ' Show result
    If isSilent Then
        WScript.Echo report
    Else
        MsgBox report, vbInformation, "Path Helper Test"
    End If
    WScript.Quit 0
Else
    report = report & "Status: FAILED - Could not read config paths" & vbCrLf
    ' Show result
    If isSilent Then
        WScript.Echo report
    Else
        MsgBox report, vbCritical, "Path Helper Test"
    End If
    WScript.Quit 1
End If

Function HasArg(argName)
    Dim i, candidate
    HasArg = False
    For i = 0 To WScript.Arguments.Count - 1
        candidate = LCase(Trim(WScript.Arguments(i)))
        If candidate = "/" & LCase(argName) Or candidate = "-" & LCase(argName) Then
            HasArg = True
            Exit Function
        End If
    Next
End Function