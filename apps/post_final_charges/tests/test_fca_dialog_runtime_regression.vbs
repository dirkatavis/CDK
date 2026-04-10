'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestFcaDialogRuntimeRegression
' **DATE CREATED:** 2026-04-10
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Regression guard for the FCA warranty dialog handler in PostFinalCharges.vbs.
' Verifies that all three handler functions are declared, the detection string
' is present, config keys are read, and the FCA check is wired before
' ProcessPromptSequence(fnlPrompts) in ProcessLinesSequentially.
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

WScript.Echo "FCA Dialog Runtime Regression Test"
WScript.Echo "==================================="

' All handler functions/subs are declared
AssertContains "IsWchLine function declared", "Function IsWchLine("
AssertContains "ExtractPartNumberForFca function declared", "Function ExtractPartNumberForFca()"
AssertContains "CreateFcaPromptDictionary function declared", "Function CreateFcaPromptDictionary("
AssertContains "HandleFcaDialog sub declared", "Sub HandleFcaDialog("

' Detection string is present (used in IsTextPresent calls)
AssertContains "FCA dialog detection string present", "FCA GLOBAL CLAIMS INFORMATION"

' Config keys are read
AssertContains "FcaDialogEnabled feature flag is read", "FcaDialogEnabled"
AssertContains "FcaConditionCode config key is read", "FcaConditionCode"
AssertContains "FcaCausalLop config key is read", "FcaCausalLop"
AssertContains "FcaCalEmissions config key is read", "FcaCalEmissions"

' Disabled-path sets correct skip result
AssertContains "Disabled handler sets skip result", "Skipped - FCA dialog handler not yet configured"

' IsWchLine uses correct LTYPE column position
AssertContains "IsWchLine reads LTYPE at col 50", "Mid(buf, 50, 6)"

' Part number extraction uses correct column positions
AssertContains "Part number extracted from col 9", "Mid(buf, 9, 20)"

' Missing part number triggers abort flag
AssertContains "Missing part number sets abort flag", "Flagged - Missing part number for FCA dialog"

' Proactive WCH detection precedes FNL command in ProcessLinesSequentially
AssertOrder "WCH proactive detection precedes FNL send in ProcessLinesSequentially", _
    "IsWchLine(lineLetterChar)", _
    "Call WaitForPrompt(""COMMAND:"", ""FNL "" & lineLetterChar"

' FCA dialog check is wired in ProcessLinesSequentially before ProcessPromptSequence(fnlPrompts)
AssertOrder "FCA detection precedes fnlPrompts sequence in ProcessLinesSequentially", _
    "IsTextPresent(""FCA GLOBAL CLAIMS INFORMATION"")", _
    "Call ProcessPromptSequence(fnlPrompts)"

WScript.Echo ""
If failures = 0 Then
    WScript.Echo "SUCCESS: FCA dialog runtime wiring is correct."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & failures & " regression checks failed."
    WScript.Quit 1
End If
