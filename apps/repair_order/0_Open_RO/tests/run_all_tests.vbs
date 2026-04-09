' ============================================================================
' run_all_tests.vbs  (Open_RO test suite runner)
' Usage: cscript.exe //nologo run_all_tests.vbs
' ============================================================================
Option Explicit

Dim fso, shell
Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

Dim scriptDir
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = scriptDir

WScript.Echo "=== Open_RO Test Suite ==="
WScript.Echo ""

RunTestFile "test_log_level_parser_contract.vbs"
RunTestFile "test_ro_scrape_timing_regression.vbs"

' ----------------------------------------------------------------------------
Sub RunTestFile(filename)
    Dim scriptPath : scriptPath = fso.BuildPath(scriptDir, filename)
    If Not fso.FileExists(scriptPath) Then
        WScript.Echo "[FAIL] Not found: " & filename
        WScript.Quit 1
    End If

    Dim cmd  : cmd  = "cscript.exe //nologo """ & scriptPath & """"
    Dim exec : Set exec = shell.Exec(cmd)

    Do While exec.Status = 0
        Do While Not exec.StdOut.AtEndOfStream
            WScript.Echo exec.StdOut.ReadLine()
        Loop
        WScript.Sleep 100
    Loop
    Do While Not exec.StdOut.AtEndOfStream
        WScript.Echo exec.StdOut.ReadLine()
    Loop

    If exec.ExitCode <> 0 Then WScript.Quit 1
End Sub
