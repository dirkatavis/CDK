'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestWchSkipCounterRuntimeRegression
' **DATE CREATED:** 2026-04-01
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Regression guard for WCH skip behavior in PostFinalCharges runtime logic.
' Verifies WCH detection, counter increment, per-RO result labeling, and summary line.
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

WScript.Echo "WCH Skip Counter Runtime Regression Test"
WScript.Echo "========================================="

AssertContains "Global skip counter is declared", "Dim g_SkipWarrantyCount"
AssertContains "Counter is reset each run", "g_SkipWarrantyCount = 0"
AssertContains "Main checks for WCH labor type", "If IsTextPresent(""WCH"") Then"
AssertContains "Counter increments when WCH is detected", "g_SkipWarrantyCount = g_SkipWarrantyCount + 1"
AssertContains "Per-RO result marks WCH skip", "lastRoResult = ""Skipped - WCH labor type"""
AssertContains "Summary includes WCH skip count", "Skips - Warranty (WCH): "
AssertOrder "WCH increment occurs before per-RO skip result", "g_SkipWarrantyCount = g_SkipWarrantyCount + 1", "lastRoResult = ""Skipped - WCH labor type"""

WScript.Echo ""
If failures = 0 Then
    WScript.Echo "SUCCESS: WCH skip counter runtime wiring is present."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & failures & " regression checks failed."
    WScript.Quit 1
End If
