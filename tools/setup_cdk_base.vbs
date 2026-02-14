Option Explicit

' =====================================================================
' setup_cdk_base.vbs - One-time setup for CDK_BASE environment variable
' =====================================================================
' This script:
' 1) Finds the repo root by locating .cdkroot relative to this script
' 2) Sets CDK_BASE for the current user (HKCU Environment)
' =====================================================================

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim sh: Set sh = CreateObject("WScript.Shell")

Dim scriptPath: scriptPath = WScript.ScriptFullName
Dim scriptDir: scriptDir = fso.GetParentFolderName(scriptPath)

Dim repoRoot: repoRoot = FindRepoRoot(scriptDir)
If repoRoot = "" Then
    MsgBox "ERROR: Could not find .cdkroot. Ensure you run this from inside the CDK folder.", vbCritical, "CDK Setup"
    WScript.Quit 1
End If

' Write to user environment
On Error Resume Next
sh.Environment("USER")("CDK_BASE") = repoRoot
If Err.Number <> 0 Then
    MsgBox "ERROR: Failed to set CDK_BASE. " & Err.Description, vbCritical, "CDK Setup"
    WScript.Quit 1
End If
On Error GoTo 0

MsgBox "CDK_BASE set to:" & vbCrLf & repoRoot & vbCrLf & vbCrLf & _
       "Please restart BlueZone so it picks up the new variable.", vbInformation, "CDK Setup"

Function FindRepoRoot(startDir)
    Dim current: current = startDir
    Dim i: i = 0
    Do While i < 10
        If fso.FileExists(fso.BuildPath(current, ".cdkroot")) Then
            FindRepoRoot = current
            Exit Function
        End If
        Dim parent: parent = fso.GetParentFolderName(current)
        If parent = "" Or parent = current Then Exit Do
        current = parent
        i = i + 1
    Loop
    FindRepoRoot = ""
End Function
