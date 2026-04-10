'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestPartsChargedGateRuntimeRegression
' **DATE CREATED:** 2026-04-09
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Regression guard for HasPartsCharged() gate in PostFinalCharges runtime logic.
' The WCH skip gate was removed; ROs are now allowed or skipped based solely on
' whether at least one P-line carries a non-zero SALE AMT.
' Verifies function presence, guard placement in Closeout_Ro, and skip result label.
'
' NOTE: Also confirms the old WCH unconditional skip is gone.
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

Sub AssertNotContains(label, needle)
    If InStr(1, g_content, needle, vbTextCompare) = 0 Then
        WScript.Echo "[PASS] " & label
    Else
        WScript.Echo "[FAIL] " & label & " (should be absent: " & needle & ")"
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

WScript.Echo "Parts-Charged Gate Runtime Regression Test"
WScript.Echo "==========================================="

' HasPartsCharged function is present
AssertContains "HasPartsCharged function is declared", "Function HasPartsCharged()"
AssertContains "HasPartsCharged scans P-line indicator", "Mid(buf, 6, 1) = ""P"""
AssertContains "HasPartsCharged reads SALE AMT column", "Mid(buf, 70, 11)"
AssertContains "HasPartsCharged returns True on positive amount", "HasPartsCharged = True"

' Guard is wired into Closeout_Ro before status routing
AssertContains "Closeout_Ro calls HasPartsCharged", "If Not HasPartsCharged() Then"
AssertContains "Skip result label is correct", "lastRoResult = ""Skipped - No parts charged"""

' Guard fires before FC/F commands (guard appears before Closeout_ReadyToPost)
AssertOrder "Parts guard precedes READY TO POST closeout", _
    "If Not HasPartsCharged() Then", "Call Closeout_ReadyToPost()"

' Old WCH unconditional skip is gone
AssertNotContains "WCH skip gate removed", "lastRoResult = ""Skipped - WCH labor type"""
AssertNotContains "g_SkipWarrantyCount removed", "Dim g_SkipWarrantyCount"
AssertNotContains "Warranty summary line removed", "Skips - Warranty (WCH):"

WScript.Echo ""
If failures = 0 Then
    WScript.Echo "SUCCESS: Parts-charged gate runtime wiring is correct."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & failures & " regression checks failed."
    WScript.Quit 1
End If
