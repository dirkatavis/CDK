Option Explicit

' Generic prompt-action tests for Maintenance_RO_Closer.
' Add new occasional prompt scenarios by appending a RunPromptCase call.

Dim g_Pass, g_Fail, g_fso, g_shell, g_repoRoot
Dim g_re

g_Pass = 0
g_Fail = 0

Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_shell = CreateObject("WScript.Shell")
g_repoRoot = g_shell.Environment("USER")("CDK_BASE")

If g_repoRoot = "" Then
    WScript.Echo "FAIL: CDK_BASE is not set"
    WScript.Quit 1
End If

Sub AssertEqual(ByVal label, ByVal expected, ByVal actual)
    If CStr(expected) = CStr(actual) Then
        g_Pass = g_Pass + 1
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "FAIL: " & label & " | expected=[" & expected & "] actual=[" & actual & "]"
    End If
End Sub

Sub AssertBool(ByVal label, ByVal expected, ByVal actual)
    If CBool(expected) = CBool(actual) Then
        g_Pass = g_Pass + 1
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "FAIL: " & label & " | expected=[" & expected & "] actual=[" & actual & "]"
    End If
End Sub

Sub RunPromptCase(ByVal label, ByVal promptText, ByVal expectedMatched, ByVal expectedResponse)
    Dim actualMatched, actualResponse, logLevel, logMessage

    actualResponse = ""
    logLevel = ""
    logMessage = ""
    actualMatched = GetReviewPromptAction(g_re, promptText, "A", actualResponse, logLevel, logMessage)

    AssertBool label & " (matched)", expectedMatched, actualMatched
    If expectedMatched Then
        AssertEqual label & " (response)", expectedResponse, actualResponse
    End If
End Sub

Dim scriptPath, fileContent, scriptStream
scriptPath = g_fso.BuildPath(g_repoRoot, "apps\maintenance_ro_closer\Maintenance_RO_Closer.vbs")
Set scriptStream = g_fso.OpenTextFile(scriptPath)
fileContent = scriptStream.ReadAll
scriptStream.Close
fileContent = Replace(fileContent, "Set g_bzhao = CreateObject(""BZWhll.WhllObj"")", "Set g_bzhao = Nothing")
fileContent = Replace(fileContent, vbCrLf & "' Execute" & vbCrLf & "RunAutomation", vbCrLf & "' Execute disabled during tests")
ExecuteGlobal fileContent

Set g_re = CreateObject("VBScript.RegExp")
g_re.IgnoreCase = True
g_re.Global = False

' Table-style cases. Add new occasional prompts here without changing test harness logic.
RunPromptCase "Comeback prompt", "Is this a comeback (Y/N)...", True, "Y"
RunPromptCase "Technician prompt no default", "TECHNICIAN?", True, "99"
RunPromptCase "Technician prompt with default", "TECHNICIAN (12)?", True, ""
RunPromptCase "Operation code prompt no default", "OPERATION CODE FOR LINE A, L1?", True, "I"
RunPromptCase "Operation code prompt with default", "OPERATION CODE FOR LINE A, L1 (I)?", True, ""
RunPromptCase "Actual hours no default", "ACTUAL HOURS?", True, "0"
RunPromptCase "Actual hours with default", "ACTUAL HOURS (3)?", True, ""
RunPromptCase "Actual hours with decimal default", "ACTUAL HOURS (3.5)?", True, ""
RunPromptCase "Sold hours no default", "SOLD HOURS?", True, "0"
RunPromptCase "Sold hours with default", "SOLD HOURS (23): SOLD HOURS?", True, ""
RunPromptCase "Sold hours with decimal default", "SOLD HOURS (2.0)?", True, ""
RunPromptCase "Add labor operation", "ADD A LABOR OPERATION (N)?", True, ""
RunPromptCase "Unknown prompt should not match", "SOME COMPLETELY NEW PROMPT", False, ""

If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All " & g_Pass & " prompt action tests passed."
Else
    WScript.Echo "FAIL: " & g_Fail & " prompt action test(s) failed."
    WScript.Quit 1
End If
