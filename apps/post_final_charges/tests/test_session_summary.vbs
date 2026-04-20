'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestSessionSummary
' **DATE CREATED:** 2026-04-20
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Unit tests for BuildSessionSummary(). Sets global counters to known values,
' calls BuildSessionSummary(), and asserts the returned string contains the
' correct lines with correct counts. No BlueZone, no MsgBox, no config needed.
'
' Test cases:
'   1) Frequent counters appear with correct values
'   2) Accounted Total equals the sum of all counters
'   3) Misc line shows correct collapsed total
'   4) Misc detail lines absent when all misc counters are zero
'   5) Misc detail lines appear (indented) when a misc counter is non-zero
'   6) Collapsed counters do not appear at top level when zero
'   7) Unsupported warranty ltype line is present and counted correctly
'   8) Other Outcome Breakdown absent when g_SummaryOtherOutcomeCount is zero
'   9) Other Outcome Breakdown present when g_SummaryOtherOutcomeCount > 0
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_Pass, g_Fail
g_Pass = 0
g_Fail = 0

' ---- Globals required by BuildSessionSummary and helpers ----
Dim g_ReviewedROCount
Dim g_FiledROCount
Dim g_SkipConfiguredCount
Dim g_SkipTechCodeCount
Dim g_SkipBlacklistCount
Dim g_SkipStatusOpenCount
Dim g_SkipStatusPreassignedCount
Dim g_SkipStatusOtherCount
Dim g_ClosedRoCount
Dim g_NotOnFileRoCount
Dim g_SkipVehidNotOnFileCount
Dim g_SkipNoCloseoutTextCount
Dim g_SkipNoPartsChargedCount
Dim g_SkipUnsupportedWarrantyCount
Dim g_LeftOpenManualCount
Dim g_ErrorInMainCount
Dim g_NoResultRecordedCount
Dim g_SummaryOtherOutcomeCount
Dim g_SummaryOtherOutcomeBreakdown
Dim g_SummaryOtherOutcomeRawBreakdown
Dim g_OlderRoAttemptCount

Sub ResetCounters()
    g_ReviewedROCount = 0
    g_FiledROCount = 0
    g_SkipConfiguredCount = 0
    g_SkipTechCodeCount = 0
    g_SkipBlacklistCount = 0
    g_SkipStatusOpenCount = 0
    g_SkipStatusPreassignedCount = 0
    g_SkipStatusOtherCount = 0
    g_ClosedRoCount = 0
    g_NotOnFileRoCount = 0
    g_SkipVehidNotOnFileCount = 0
    g_SkipNoCloseoutTextCount = 0
    g_SkipNoPartsChargedCount = 0
    g_SkipUnsupportedWarrantyCount = 0
    g_LeftOpenManualCount = 0
    g_ErrorInMainCount = 0
    g_NoResultRecordedCount = 0
    g_SummaryOtherOutcomeCount = 0
    g_OlderRoAttemptCount = 0
    Set g_SummaryOtherOutcomeBreakdown = CreateObject("Scripting.Dictionary")
    Set g_SummaryOtherOutcomeRawBreakdown = CreateObject("Scripting.Dictionary")
End Sub

' ---- BuildSessionSummary and helpers (copy-pasted from PostFinalCharges.vbs) ----

Function BuildSessionSummary()
    Dim accountedTotal, summaryText, otherOutcomeDetails
    Dim miscTotal, miscDetail

    ' Infrequent counters collapsed into a single Misc line.
    ' Individual lines shown below Misc only when their value is non-zero.
    miscTotal = g_SkipConfiguredCount + _
        g_SkipBlacklistCount + _
        g_ClosedRoCount + _
        g_NotOnFileRoCount + _
        g_SkipVehidNotOnFileCount + _
        g_SkipNoCloseoutTextCount + _
        g_LeftOpenManualCount + _
        g_ErrorInMainCount + _
        g_NoResultRecordedCount + _
        g_SummaryOtherOutcomeCount + _
        g_OlderRoAttemptCount

    accountedTotal = g_FiledROCount + _
        g_SkipTechCodeCount + _
        g_SkipStatusOpenCount + _
        g_SkipStatusPreassignedCount + _
        g_SkipStatusOtherCount + _
        g_SkipNoPartsChargedCount + _
        g_SkipUnsupportedWarrantyCount + _
        miscTotal

    summaryText = "DONE" & vbCrLf & _
        "ROs Reviewed: " & g_ReviewedROCount & vbCrLf & _
        "ROs Posted: " & g_FiledROCount & vbCrLf & _
        "Skips - Non-compliant tech code: " & g_SkipTechCodeCount & vbCrLf & _
        "Skips - Open: " & g_SkipStatusOpenCount & vbCrLf & _
        "Skips - Pre-Assigned: " & g_SkipStatusPreassignedCount & vbCrLf & _
        "Skips - Other Statuses: " & g_SkipStatusOtherCount & vbCrLf & _
        "Skipped - No parts charged: " & g_SkipNoPartsChargedCount & vbCrLf & _
        "Skipped - Unsupported warranty ltype: " & g_SkipUnsupportedWarrantyCount & vbCrLf & _
        "Misc: " & miscTotal & vbCrLf & _
        "Accounted Total: " & accountedTotal

    ' Expand Misc breakdown — only non-zero lines shown
    miscDetail = ""
    If g_SkipConfiguredCount > 0 Then miscDetail = miscDetail & "  Skips - Specific ROs: " & g_SkipConfiguredCount & vbCrLf
    If g_SkipBlacklistCount > 0 Then miscDetail = miscDetail & "  Skips - Other Terms: " & g_SkipBlacklistCount & vbCrLf
    If g_ClosedRoCount > 0 Then miscDetail = miscDetail & "  Closed (already): " & g_ClosedRoCount & vbCrLf
    If g_NotOnFileRoCount > 0 Then miscDetail = miscDetail & "  Not On File: " & g_NotOnFileRoCount & vbCrLf
    If g_SkipVehidNotOnFileCount > 0 Then miscDetail = miscDetail & "  Skipped - VEHID not on file: " & g_SkipVehidNotOnFileCount & vbCrLf
    If g_SkipNoCloseoutTextCount > 0 Then miscDetail = miscDetail & "  Skipped - No closeout text: " & g_SkipNoCloseoutTextCount & vbCrLf
    If g_LeftOpenManualCount > 0 Then miscDetail = miscDetail & "  Left Open for manual closing: " & g_LeftOpenManualCount & vbCrLf
    If g_ErrorInMainCount > 0 Then miscDetail = miscDetail & "  Errors in Main: " & g_ErrorInMainCount & vbCrLf
    If g_NoResultRecordedCount > 0 Then miscDetail = miscDetail & "  No result recorded: " & g_NoResultRecordedCount & vbCrLf
    If g_SummaryOtherOutcomeCount > 0 Then miscDetail = miscDetail & "  Other Outcomes: " & g_SummaryOtherOutcomeCount & vbCrLf
    If g_OlderRoAttemptCount > 0 Then miscDetail = miscDetail & "  Older ROs Attempted (subset): " & g_OlderRoAttemptCount & vbCrLf

    If Len(miscDetail) > 0 Then
        summaryText = summaryText & vbCrLf & Left(miscDetail, Len(miscDetail) - Len(vbCrLf))
    End If

    If g_SummaryOtherOutcomeCount > 0 Then
        otherOutcomeDetails = BuildOtherOutcomeBreakdown(8)
        If Len(Trim(CStr(otherOutcomeDetails))) > 0 Then
            summaryText = summaryText & vbCrLf & "Other Outcome Breakdown:" & vbCrLf & otherOutcomeDetails
        End If

        otherOutcomeDetails = BuildOtherOutcomeRawBreakdown(12)
        If Len(Trim(CStr(otherOutcomeDetails))) > 0 Then
            summaryText = summaryText & vbCrLf & "Other Outcome Raw Results:" & vbCrLf & otherOutcomeDetails
        End If
    End If

    BuildSessionSummary = summaryText
End Function

Function BuildOtherOutcomeBreakdown(maxLines)
    Dim key, countValue, linesAdded, hiddenCategories, output
    If maxLines <= 0 Then maxLines = 8
    output = ""
    linesAdded = 0
    hiddenCategories = 0
    If Not IsObject(g_SummaryOtherOutcomeBreakdown) Then
        BuildOtherOutcomeBreakdown = ""
        Exit Function
    End If
    For Each key In g_SummaryOtherOutcomeBreakdown.Keys
        If linesAdded < maxLines Then
            countValue = CLng(g_SummaryOtherOutcomeBreakdown(key))
            output = output & "  - " & CStr(key) & ": " & CStr(countValue) & vbCrLf
            linesAdded = linesAdded + 1
        Else
            hiddenCategories = hiddenCategories + 1
        End If
    Next
    If hiddenCategories > 0 Then
        output = output & "  - (+" & hiddenCategories & " more categories)"
    ElseIf Len(output) > 0 Then
        output = Left(output, Len(output) - Len(vbCrLf))
    End If
    BuildOtherOutcomeBreakdown = output
End Function

Function BuildOtherOutcomeRawBreakdown(maxLines)
    Dim key, countValue, linesAdded, hiddenCategories, output
    If maxLines <= 0 Then maxLines = 12
    output = ""
    linesAdded = 0
    hiddenCategories = 0
    If Not IsObject(g_SummaryOtherOutcomeRawBreakdown) Then
        BuildOtherOutcomeRawBreakdown = ""
        Exit Function
    End If
    For Each key In g_SummaryOtherOutcomeRawBreakdown.Keys
        If linesAdded < maxLines Then
            countValue = CLng(g_SummaryOtherOutcomeRawBreakdown(key))
            output = output & "  - " & CStr(key) & ": " & CStr(countValue) & vbCrLf
            linesAdded = linesAdded + 1
        Else
            hiddenCategories = hiddenCategories + 1
        End If
    Next
    If hiddenCategories > 0 Then
        output = output & "  - (+" & hiddenCategories & " more raw results)"
    ElseIf Len(output) > 0 Then
        output = Left(output, Len(output) - Len(vbCrLf))
    End If
    BuildOtherOutcomeRawBreakdown = output
End Function

' ---- Assert helpers ----
Sub AssertContains(ByVal label, ByVal haystack, ByVal needle)
    If InStr(1, haystack, needle, vbTextCompare) > 0 Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (missing: '" & needle & "')"
    End If
End Sub

Sub AssertAbsent(ByVal label, ByVal haystack, ByVal needle)
    If InStr(1, haystack, needle, vbTextCompare) = 0 Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (should be absent: '" & needle & "')"
    End If
End Sub

Sub AssertEqual(ByVal label, ByVal expected, ByVal actual)
    If CStr(expected) = CStr(actual) Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (expected='" & expected & "' actual='" & actual & "')"
    End If
End Sub

' ============================
' Tests
' ============================
WScript.Echo "Session Summary Tests"
WScript.Echo "====================="

' --- Test 1: Frequent counters appear with correct values ---
Call ResetCounters()
g_ReviewedROCount = 111
g_FiledROCount = 59
g_SkipTechCodeCount = 2
g_SkipStatusOpenCount = 1
g_SkipStatusOtherCount = 1
g_SkipNoPartsChargedCount = 38
g_SkipUnsupportedWarrantyCount = 10

Dim summary1
summary1 = BuildSessionSummary()

AssertContains "Summary starts with DONE", summary1, "DONE"
AssertContains "ROs Reviewed line correct", summary1, "ROs Reviewed: 111"
AssertContains "ROs Posted line correct", summary1, "ROs Posted: 59"
AssertContains "Non-compliant tech code line correct", summary1, "Skips - Non-compliant tech code: 2"
AssertContains "Skips Open line correct", summary1, "Skips - Open: 1"
AssertContains "Skips Other Statuses line correct", summary1, "Skips - Other Statuses: 1"
AssertContains "No parts charged line correct", summary1, "Skipped - No parts charged: 38"
AssertContains "Unsupported warranty ltype line present", summary1, "Skipped - Unsupported warranty ltype: 10"

' --- Test 2: Accounted Total equals sum of frequent + misc (all misc zero) ---
' Sum: 59 + 2 + 1 + 0 + 1 + 38 + 10 + 0 (misc) = 111
AssertContains "Accounted Total equals sum", summary1, "Accounted Total: 111"

' --- Test 3: Misc line shows zero when all misc counters are zero ---
AssertContains "Misc line present with zero", summary1, "Misc: 0"

' --- Test 4: Misc detail lines absent when all misc counters are zero ---
AssertAbsent "Skips - Specific ROs detail absent when zero", summary1, "  Skips - Specific ROs:"
AssertAbsent "Skips - Other Terms detail absent when zero", summary1, "  Skips - Other Terms:"
AssertAbsent "Closed (already) detail absent when zero", summary1, "  Closed (already):"
AssertAbsent "Not On File detail absent when zero", summary1, "  Not On File:"
AssertAbsent "Errors in Main detail absent when zero", summary1, "  Errors in Main:"

' --- Test 5: Misc detail lines appear (indented) when a misc counter is non-zero ---
Call ResetCounters()
g_ReviewedROCount = 10
g_FiledROCount = 5
g_SkipConfiguredCount = 2
g_ClosedRoCount = 1
g_ErrorInMainCount = 1
g_OlderRoAttemptCount = 3

Dim summary2
summary2 = BuildSessionSummary()

' miscTotal = 2 + 1 + 1 + 3 = 7; accountedTotal = 5 + 7 = 12
AssertContains "Misc line shows collapsed total", summary2, "Misc: 7"
AssertContains "Accounted Total includes misc", summary2, "Accounted Total: 12"
AssertContains "Skips - Specific ROs detail appears when non-zero", summary2, "  Skips - Specific ROs: 2"
AssertContains "Closed (already) detail appears when non-zero", summary2, "  Closed (already): 1"
AssertContains "Errors in Main detail appears when non-zero", summary2, "  Errors in Main: 1"
AssertContains "Older ROs detail appears when non-zero", summary2, "  Older ROs Attempted (subset): 3"
AssertAbsent "Not On File detail absent when zero even with other misc non-zero", summary2, "  Not On File:"

' --- Test 7: Unsupported warranty ltype present even when zero ---
Call ResetCounters()
Dim summary3
summary3 = BuildSessionSummary()
AssertContains "Unsupported warranty ltype line present even when zero", summary3, "Skipped - Unsupported warranty ltype: 0"
AssertContains "Zero counters still produce Accounted Total: 0", summary3, "Accounted Total: 0"
AssertContains "Zero counters produce Misc: 0", summary3, "Misc: 0"

' --- Test 8: Other Outcomes section absent when g_SummaryOtherOutcomeCount is zero ---
AssertAbsent "Other Outcome Breakdown absent when count is zero", summary3, "Other Outcome Breakdown:"
AssertAbsent "Other Outcome Raw Results absent when count is zero", summary3, "Other Outcome Raw Results:"

' --- Test 9: Other Outcome Breakdown present when g_SummaryOtherOutcomeCount > 0 ---
Call ResetCounters()
g_ReviewedROCount = 5
g_FiledROCount = 3
g_SummaryOtherOutcomeCount = 2
g_SummaryOtherOutcomeBreakdown.Add "Skipped - FCA dialog handler not configured", 2
g_SummaryOtherOutcomeRawBreakdown.Add "Skipped - FCA dialog handler not configured for ltype XY", 2

Dim summary4
summary4 = BuildSessionSummary()

' g_SummaryOtherOutcomeCount contributes to miscTotal, which feeds accountedTotal
' miscTotal = 2; accountedTotal = 3 + 2 = 5
AssertContains "Misc line includes Other Outcomes count", summary4, "Misc: 2"
AssertContains "Accounted Total includes Other Outcomes via misc", summary4, "Accounted Total: 5"
AssertContains "Other Outcomes indented detail appears", summary4, "  Other Outcomes: 2"
AssertContains "Other Outcome Breakdown section appears", summary4, "Other Outcome Breakdown:"
AssertContains "Other Outcome Breakdown entry appears", summary4, "Skipped - FCA dialog handler not configured: 2"
AssertContains "Other Outcome Raw Results section appears", summary4, "Other Outcome Raw Results:"

WScript.Echo ""
If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All " & g_Pass & " session summary tests passed."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & g_Fail & " test(s) failed."
    WScript.Quit 1
End If
