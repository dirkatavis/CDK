Option Explicit

' Host compatibility helpers to run under WSH (cscript/wscript) and non-WSH hosts (BlueZone)
' Provides: Host_SafeScriptFolder() and Host_Quit(exitCode)

Dim g_hc_fso: Set g_hc_fso = CreateObject("Scripting.FileSystemObject")

Function Host_SafeScriptFolder()
    On Error Resume Next
    Dim sh, envBase
    Set sh = CreateObject("WScript.Shell")
    envBase = ""
    If Not sh Is Nothing Then
        envBase = sh.Environment("USER")("CDK_BASE")
    End If

    If envBase <> "" And g_hc_fso.FolderExists(envBase) And g_hc_fso.FileExists(g_hc_fso.BuildPath(envBase, ".cdkroot")) Then
        Host_SafeScriptFolder = envBase
        If Right(Host_SafeScriptFolder, 1) <> Chr(92) Then Host_SafeScriptFolder = Host_SafeScriptFolder & Chr(92)
        Exit Function
    End If

    ' Fallback to current working directory
    Host_SafeScriptFolder = g_hc_fso.GetAbsolutePathName(".")
    If Right(Host_SafeScriptFolder, 1) <> Chr(92) Then Host_SafeScriptFolder = Host_SafeScriptFolder & Chr(92)
End Function

Sub Host_Quit(Optional exitCode)
    ' Use a raised error to terminate execution; avoids referencing WScript in non-WSH hosts
    If IsNumeric(exitCode) Then
        Err.Raise 9999, "Host_Quit", "Terminating script. ExitCode=" & CStr(exitCode)
    Else
        Err.Raise 9999, "Host_Quit", "Terminating script."
    End If
End Sub
