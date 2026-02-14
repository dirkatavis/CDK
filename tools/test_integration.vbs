'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestMainScriptWithMock
' **DATE CREATED:** 2025-11-19
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Integration test that runs the main PostFinalCharges script with MockBzhao.
' Tests basic RO processing workflow without requiring BlueZone.
'-----------------------------------------------------------------------------------

Option Explicit

' Set test mode environment variable
Sub SetTestEnvironment()
    Dim shell
    Set shell = CreateObject("WScript.Shell")
    shell.Environment("PROCESS")("PFC_TEST_MODE") = "true"
    shell.Environment("PROCESS")("PFC_LOG_LEVEL") = "DEBUG"
End Sub

Sub TestMainScript()
    WScript.Echo "Testing main script with MockBzhao..."
    
    ' Set environment for test mode
    SetTestEnvironment()
    
    ' Run the main script (this will use the mock)
    ' Note: In a real test, we'd capture output, but for now just run it
    Dim mainScript
    mainScript = "../PostFinalCharges.vbs"
    
    Dim shell
    Set shell = CreateObject("WScript.Shell")
    
    ' Run the script - this should work with the mock
    Dim exec
    Set exec = shell.Exec("cscript.exe " & mainScript)
    
    ' Wait for completion
    Do While exec.Status = 0
        WScript.Sleep 100
    Loop
    
    WScript.Echo "Main script test completed. Exit code: " & exec.ExitCode
    
    If exec.ExitCode = 0 Then
        WScript.Echo "SUCCESS: Main script ran with mock!"
    Else
        WScript.Echo "FAILED: Main script failed with exit code " & exec.ExitCode
    End If
End Sub

' Run the test
TestMainScript