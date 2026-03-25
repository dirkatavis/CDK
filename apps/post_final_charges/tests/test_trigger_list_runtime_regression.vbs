'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestTriggerListRuntimeRegression
' **DATE CREATED:** 2026-03-25
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Regression guard for TriggerList runtime behavior:
' Ensures PostFinalCharges.vbs loads TriggerList from config,
' builds a trigger array, and FindTrigger scans the loaded entries.
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_fso, g_scriptPath, g_content, g_configPath, g_configContent
Set g_fso = CreateObject("Scripting.FileSystemObject")
g_scriptPath = "../PostFinalCharges.vbs"
g_configPath = "../../../config/config.ini"

If Not g_fso.FileExists(g_scriptPath) Then
    WScript.Echo "[FAIL] PostFinalCharges.vbs not found at: " & g_scriptPath
    WScript.Quit 1
End If

If Not g_fso.FileExists(g_configPath) Then
    WScript.Echo "[FAIL] config.ini not found at: " & g_configPath
    WScript.Quit 1
End If

Dim ts
Set ts = g_fso.OpenTextFile(g_scriptPath, 1)
g_content = ts.ReadAll
ts.Close

Set ts = g_fso.OpenTextFile(g_configPath, 1)
g_configContent = ts.ReadAll
ts.Close

Dim failures
failures = 0

Sub AssertContains(label, haystack, needle)
    If InStr(1, haystack, needle, vbTextCompare) > 0 Then
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

WScript.Echo "TriggerList Runtime Regression Test"
WScript.Echo "==================================="

AssertContains "Config defines TriggerList", g_configContent, "TriggerList=apps\post_final_charges\CloseoutTriggers.csv"
AssertContains "Loads TriggerList helper", g_content, "Function LoadCloseoutTriggers(triggerListPath, ByRef triggerArray)"
AssertContains "InitializeConfig loads TriggerList from config", g_content, "LoadCloseoutTriggers(GetConfigPath(""PostFinalCharges"", ""TriggerList""), g_CloseoutTriggers)"
AssertContains "FindTrigger iterates loaded trigger array", g_content, "For i = LBound(g_CloseoutTriggers) To UBound(g_CloseoutTriggers)"
AssertContains "FindTrigger uses loaded candidate", g_content, "candidate = g_CloseoutTriggers(i)"
AssertOrder "TriggerList load occurs before FindTrigger definition", "Function LoadCloseoutTriggers(triggerListPath, ByRef triggerArray)", "Function FindTrigger()"

WScript.Echo ""
If failures = 0 Then
    WScript.Echo "SUCCESS: TriggerList runtime wiring is present."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & failures & " regression checks failed."
    WScript.Quit 1
End If