'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestSummaryAccountingRuntimeRegression
' **DATE CREATED:** 2026-04-15
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Runtime guard that ensures the end-of-run MsgBox includes explicit accounting
' lines so users can reconcile totals against reviewed count.
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

WScript.Echo "Summary Accounting Runtime Regression Test"
WScript.Echo "=========================================="

AssertContains "Other outcome counter is declared", "Dim g_SummaryOtherOutcomeCount"
AssertContains "Result-accounting helper exists", "Function IsResultRepresentedInSummary(resultText)"
AssertContains "Process loop increments Other Outcomes", "g_SummaryOtherOutcomeCount = g_SummaryOtherOutcomeCount + 1"
AssertContains "Summary includes Other Outcomes line", "Other Outcomes: "
AssertContains "Summary includes Accounted Total line", "Accounted Total: "
AssertContains "Summary includes closed count line", "Closed (already): "
AssertContains "Summary includes not-on-file count line", "Not On File: "
AssertContains "Summary includes no-closeout-text count line", "Skipped - No closeout text: "
AssertContains "Summary includes no-parts-charged count line", "Skipped - No parts charged: "
AssertContains "Summary includes grouped Other Outcome details", "Other Outcome Breakdown:"
AssertContains "Summary includes raw Other Outcome details", "Other Outcome Raw Results:"
AssertContains "Older attempted line marked as subset", "Older ROs Attempted (subset): "

WScript.Echo ""
If failures = 0 Then
    WScript.Echo "SUCCESS: Summary accounting runtime wiring is correct."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & failures & " regression checks failed."
    WScript.Quit 1
End If
