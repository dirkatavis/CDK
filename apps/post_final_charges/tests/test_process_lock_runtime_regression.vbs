'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestProcessLockRuntimeRegression
' **DATE CREATED:** 2026-04-15
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Regression guard for RO process-lock handling:
' Ensures Main() detects "Process is locked by" on the main prompt line,
' sends Enter to recover, waits for COMMAND:, and records a skip result.
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_fso, g_scriptPath, g_content
Set g_fso = CreateObject("Scripting.FileSystemObject")
g_scriptPath = "../PostFinalCharges.vbs"

If Not g_fso.FileExists(g_scriptPath) Then
    WScript.Echo "[FAIL] PostFinalCharges.vbs not found at: " & g_scriptPath
    WScript.Quit 1
End If

Dim ts
Set ts = g_fso.OpenTextFile(g_scriptPath, 1)
g_content = ts.ReadAll
ts.Close

Dim failures
failures = 0

Sub AssertContains(label, needle)
    If InStr(1, g_content, needle, vbTextCompare) > 0 Then
        WScript.Echo "[PASS] " & label
    Else
        WScript.Echo "[FAIL] " & label & " (missing: " & needle & ")"
        failures = failures + 1
    End If
End Sub

Function IndexOf(needle)
    IndexOf = InStr(1, g_content, needle, vbTextCompare)
End Function

Sub AssertOrder(label, firstNeedle, secondNeedle)
    Dim i1, i2
    i1 = IndexOf(firstNeedle)
    i2 = IndexOf(secondNeedle)

    If i1 > 0 And i2 > 0 And i1 < i2 Then
        WScript.Echo "[PASS] " & label
    Else
        WScript.Echo "[FAIL] " & label & " (expected order not found)"
        failures = failures + 1
    End If
End Sub

WScript.Echo "Process Lock Runtime Regression Test"
WScript.Echo "===================================="

AssertContains "Main reads line 23 prompt text", "mainPromptText = GetScreenLine(MainPromptLine)"
AssertContains "Main checks process lock text", "Process is locked by"
AssertContains "Main sends Enter for lock recovery", "Call FastKey(""<Enter>"")"
AssertContains "Main waits for COMMAND after lock recovery", "Call WaitForPrompt(""COMMAND:"""
AssertContains "Main includes lock recovery wait label", "Process Lock Recovery"
AssertContains "Main records lock skip result", "lastRoResult = ""Skipped - Process locked by another user"""
AssertOrder "Process lock result assignment occurs before VEHID skip result", "lastRoResult = ""Skipped - Process locked by another user""", "lastRoResult = ""Skipped - VEHID not on file"""

WScript.Echo ""
If failures = 0 Then
    WScript.Echo "SUCCESS: Process-lock runtime wiring is present."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & failures & " regression checks failed."
    WScript.Quit 1
End If
