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

Sub RunStep(scriptName, description)
    WScript.Echo ">>> RUNNING STEP: " & description & " (" & scriptName & ")..."
    
    Dim command: command = "cscript.exe //NoLogo """ & fso.BuildPath(fso.BuildPath(scriptDir, "tools"), scriptName) & """"
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

RunStep "setup_cdk_base.vbs",      "Environment Initialization"
RunStep "validate_dependencies.vbs", "Dependency Validation"
RunStep "test_path_helper.vbs",     "Path Configuration Test"

WScript.Echo String(50, "=")
WScript.Echo "FULL DEPLOYMENT VERIFIED SUCCESSFULLY!"
MsgBox "All deployment steps passed successfully!", vbInformation, "Verification Complete"
