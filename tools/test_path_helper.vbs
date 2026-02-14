Option Explicit

' ==============================================================================
' test_path_helper.vbs - Validate PathHelper configuration
' ==============================================================================
' This script tests the centralized path configuration system.
' Run this from BlueZone to verify the setup works correctly.
' ==============================================================================

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Const BASE_ENV_VAR = "CDK_BASE"

' Load PathHelper module using CDK_BASE
Dim sh: Set sh = CreateObject("WScript.Shell")
Dim basePath: basePath = sh.Environment("USER")(BASE_ENV_VAR)

If basePath = "" Or Not fso.FolderExists(basePath) Then
    MsgBox "ERROR: Invalid or missing CDK_BASE" & vbCrLf & "Value: " & basePath, vbCritical, "Test Failed"
    End
End If

Dim helperPath: helperPath = fso.BuildPath(basePath, "common\PathHelper.vbs")
If Not fso.FileExists(helperPath) Then
    MsgBox "ERROR: Cannot find PathHelper.vbs" & vbCrLf & "Looked at: " & helperPath, vbCritical, "Test Failed"
    End
End If

ExecuteGlobal fso.OpenTextFile(helperPath).ReadAll

' Test repo root discovery
Dim repoRoot: repoRoot = GetRepoRoot()

' Test config path reading
Dim csvPath, logPath
On Error Resume Next
csvPath = GetConfigPath("Close_ROs_Pt1", "CSV")
logPath = GetConfigPath("Close_ROs_Pt1", "Log")
On Error GoTo 0

' Build report
Dim report
report = "CDK Path Helper Test Results" & vbCrLf
report = report & String(50, "=") & vbCrLf & vbCrLf
report = report & "Repo Root: " & repoRoot & vbCrLf
report = report & "  .cdkroot exists: " & fso.FileExists(fso.BuildPath(repoRoot, ".cdkroot")) & vbCrLf
report = report & "  config.ini exists: " & fso.FileExists(fso.BuildPath(repoRoot, "config.ini")) & vbCrLf
report = report & vbCrLf
report = report & "Sample Config Paths:" & vbCrLf
report = report & "  Close_ROs_Pt1 CSV: " & csvPath & vbCrLf
report = report & "  Close_ROs_Pt1 Log: " & logPath & vbCrLf
report = report & vbCrLf

If csvPath <> "" And logPath <> "" Then
    report = report & "Status: SUCCESS - Path Helper is working!" & vbCrLf
Else
    report = report & "Status: FAILED - Could not read config paths" & vbCrLf
End If

' Write to temp folder
Dim outputPath: outputPath = fso.BuildPath(fso.GetSpecialFolder(2), "cdk_path_test_result.txt")
Dim outFile: Set outFile = fso.OpenTextFile(outputPath, 2, True)
outFile.Write report
outFile.Close

' Show result
MsgBox report & vbCrLf & "Full report written to:" & vbCrLf & outputPath, vbInformation, "Path Helper Test"
