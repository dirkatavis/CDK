'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestSkipRoListRuntimeRegression
' **DATE CREATED:** 2026-03-17
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Regression guard for SkipRoList runtime behavior:
' Ensures PostFinalCharges.vbs loads SkipRoList from config,
' builds lookup, checks RO in Main(), and sets skip result.
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

WScript.Echo "SkipRoList Runtime Regression Test"
WScript.Echo "================================="

AssertContains "Loads SkipRoList from config", "g_SkipRoListRaw = GetIniSetting("
AssertContains "Has skip-list loader", "Function LoadSkipRoLookup(skipRoListCsvPaths, ByRef lookupDict)"
AssertContains "Has skip-list checker", "Function ShouldSkipRo(roValue)"
AssertContains "Main checks configured skip list", "If ShouldSkipRo(currentRODisplay) Then"
AssertContains "Main sets configured skip result", "lastRoResult = ""Skipped - Configured RO skip list"""
AssertContains "Main sends E for configured skip", "Call FastText(""E"")"
AssertOrder "Configured skip check occurs after RO screen ready gate", "RO Screen Ready", "If ShouldSkipRo(currentRODisplay) Then"

WScript.Echo ""
If failures = 0 Then
    WScript.Echo "SUCCESS: SkipRoList runtime wiring is present."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & failures & " regression checks failed."
    WScript.Quit 1
End If
