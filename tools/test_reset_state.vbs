' ============================================================================
' CDK Test State Reset
' Purpose: Normalize mutable test state before running validation suites
' Usage: cscript.exe test_reset_state.vbs
' ============================================================================

Option Explicit

Dim g_fso
Dim g_shell
Dim g_repoRoot
Dim g_failures

Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_shell = CreateObject("WScript.Shell")
g_failures = 0

g_repoRoot = g_shell.Environment("USER")("CDK_BASE")
If Len(g_repoRoot) = 0 Then
    WScript.Echo "FAIL: CDK_BASE environment variable is not set"
    WScript.Quit 1
End If

If Not g_fso.FolderExists(g_repoRoot) Then
    WScript.Echo "FAIL: CDK_BASE does not exist: " & g_repoRoot
    WScript.Quit 1
End If

WScript.Echo "Resetting test state under: " & g_repoRoot

' Restore known backup pairs created by negative tests
RestoreBackupPair ".cdkroot", ".cdkroot.backup"
RestoreBackupPair "common\PathHelper.vbs", "common\PathHelper.vbs.backup"
RestoreBackupPair "config\config.ini", "config\config.ini.backup"

' Remove stale backup artifacts if both original and backup exist
DeleteIfExists ".cdkroot.backup"
DeleteIfExists "common\PathHelper.vbs.backup"
DeleteIfExists "config\config.ini.backup"

' Ensure config.ini still has basic INI shape
ValidateConfigBasic

If g_failures > 0 Then
    WScript.Echo ""
    WScript.Echo "Test reset FAILED with " & g_failures & " issue(s)."
    WScript.Quit 1
End If

WScript.Echo "Test reset PASSED"
WScript.Quit 0

Sub RestoreBackupPair(originalRelPath, backupRelPath)
    Dim originalPath, backupPath
    originalPath = g_fso.BuildPath(g_repoRoot, originalRelPath)
    backupPath = g_fso.BuildPath(g_repoRoot, backupRelPath)

    If g_fso.FileExists(backupPath) Then
        If g_fso.FileExists(originalPath) Then
            WScript.Echo "INFO: Both original and backup exist; keeping original, removing backup: " & backupRelPath
            On Error Resume Next
            g_fso.DeleteFile backupPath, True
            If Err.Number <> 0 Then
                WScript.Echo "FAIL: Could not delete stale backup " & backupRelPath & " -> " & Err.Description
                g_failures = g_failures + 1
                Err.Clear
            End If
            On Error GoTo 0
        Else
            WScript.Echo "INFO: Restoring backup " & backupRelPath & " -> " & originalRelPath
            On Error Resume Next
            g_fso.MoveFile backupPath, originalPath
            If Err.Number <> 0 Then
                WScript.Echo "FAIL: Could not restore backup " & backupRelPath & " -> " & Err.Description
                g_failures = g_failures + 1
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End If
End Sub

Sub DeleteIfExists(relPath)
    Dim fullPath
    fullPath = g_fso.BuildPath(g_repoRoot, relPath)
    If g_fso.FileExists(fullPath) Then
        On Error Resume Next
        g_fso.DeleteFile fullPath, True
        If Err.Number <> 0 Then
            WScript.Echo "FAIL: Could not delete stale file " & relPath & " -> " & Err.Description
            g_failures = g_failures + 1
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Sub

Sub ValidateConfigBasic()
    Dim cfgPath, ts, line, hasSection, hasKeyValue
    cfgPath = g_fso.BuildPath(g_repoRoot, "config\config.ini")

    If Not g_fso.FileExists(cfgPath) Then
        WScript.Echo "FAIL: config.ini missing after reset"
        g_failures = g_failures + 1
        Exit Sub
    End If

    hasSection = False
    hasKeyValue = False

    On Error Resume Next
    Set ts = g_fso.OpenTextFile(cfgPath, 1, False)
    If Err.Number <> 0 Then
        WScript.Echo "FAIL: Could not open config.ini -> " & Err.Description
        g_failures = g_failures + 1
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If Len(line) > 0 Then
            If Left(line, 1) = "[" And Right(line, 1) = "]" Then hasSection = True
            If InStr(line, "=") > 0 Then hasKeyValue = True
        End If
    Loop
    ts.Close

    If Not hasSection Or Not hasKeyValue Then
        WScript.Echo "FAIL: config.ini appears malformed after reset"
        g_failures = g_failures + 1
    End If
End Sub
