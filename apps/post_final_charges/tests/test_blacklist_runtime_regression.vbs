'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestBlacklistRuntimeRegression
' **DATE CREATED:** 2026-03-12
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Regression guard for production runtime behavior:
' Ensures PostFinalCharges.vbs actively loads blacklist_terms from config,
' checks for matched blacklist term in Main(), and marks RO as skipped.
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

WScript.Echo "Blacklist Runtime Regression Test"
WScript.Echo "================================"

AssertContains "Loads blacklist_terms from config", "GetIniSetting(""PostFinalCharges"", ""blacklist_terms"", """")"
AssertContains "Has blacklist matcher function", "Function GetMatchedBlacklistTerm(blacklistTermsCsv)"
AssertContains "Main calls blacklist matcher", "matchedBlacklistTerm = GetMatchedBlacklistTerm(g_BlacklistTermsRaw)"
AssertContains "Main sets blacklisted skip result", "lastRoResult = ""Skipped - Blacklisted term: "" & matchedBlacklistTerm"
AssertContains "Main exits blacklisted path", "Exit Sub"
AssertOrder "Blacklist check occurs before trigger detection", "matchedBlacklistTerm = GetMatchedBlacklistTerm(g_BlacklistTermsRaw)", "trigger = FindTrigger()"

WScript.Echo ""
If failures = 0 Then
    WScript.Echo "SUCCESS: Blacklist runtime wiring is present."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & failures & " regression checks failed."
    WScript.Quit 1
End If
