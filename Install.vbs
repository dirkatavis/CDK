Option Explicit

' ==============================================================================
' Install.vbs - Automated Deployment Verification Wrapper
' ==============================================================================
' This script runs the three core verification steps in sequence:
' 1. setup_cdk_base.vbs
' 2. validate_dependencies.vbs
' 3. test_path_helper.vbs
' ==============================================================================

Dim sh: Set sh = CreateObject("WScript.Shell")
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim scriptDir: scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' Ensure this installer always runs in console host so WScript.Echo does not
' appear as many pop-up dialogs when user double-clicks the file.
If InStr(LCase(WScript.FullName), "wscript.exe") > 0 Then
    Dim relaunchCmd
    relaunchCmd = "cscript.exe //NoLogo """ & WScript.ScriptFullName & """"
    sh.Run relaunchCmd, 1, True
    WScript.Quit 0
End If

Sub RunStep(relativeScriptPath, description)
    WScript.Echo ">>> RUNNING STEP: " & description & " (" & relativeScriptPath & ")..."
    
    Dim fullScriptPath: fullScriptPath = fso.BuildPath(scriptDir, relativeScriptPath)
    If Not fso.FileExists(fullScriptPath) Then
        WScript.Echo "!!! FAILED: " & description & " script not found: " & fullScriptPath
        MsgBox "Deployment verification failed: script not found for step: " & description & vbCrLf & fullScriptPath, vbCritical, "Deployment Error"
        WScript.Quit 1
    End If

    Dim command: command = "cscript.exe //NoLogo """ & fullScriptPath & """ /silent"
    Dim exitCode: exitCode = sh.Run(command, 1, True)
    
    If exitCode <> 0 Then
        WScript.Echo "!!! FAILED: " & description & " returned error code " & exitCode
        MsgBox "Deployment verification failed at step: " & description, vbCritical, "Deployment Error"
        WScript.Quit 1
    End If
    
    WScript.Echo ">>> SUCCESS: " & description & " completed." & vbNewLine
End Sub

' Start sequence
WScript.Echo "CDK AUTOMATION - FULL DEPLOYMENT VERIFICATION"
WScript.Echo String(50, "=") & vbNewLine

RunStep "tools\setup_cdk_base.vbs",              "Environment Initialization"
RunStep "tools\validate_dependencies.vbs",       "Dependency Validation"
RunStep "tests\infrastructure\test_path_helper.vbs", "Path Configuration Test"

WScript.Echo String(50, "=")
WScript.Echo "FULL DEPLOYMENT VERIFIED SUCCESSFULLY!"
MsgBox "All deployment steps passed successfully!", vbInformation, "Verification Complete"
