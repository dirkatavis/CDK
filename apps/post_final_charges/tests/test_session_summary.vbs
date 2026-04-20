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
'   1) All named counters appear with correct values
'   2) Accounted Total equals the sum of all counters
'   3) Other Outcomes section absent when count is zero
'   4) Other Outcome Breakdown and Raw Results appear when count > 0
'   5) Unsupported warranty ltype line is present and counted correctly
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

    accountedTotal = g_FiledROCount + _
        g_SkipConfiguredCount + _
        g_SkipTechCodeCount + _
        g_SkipBlacklistCount + _
        g_SkipStatusOpenCount + _
        g_SkipStatusPreassignedCount + _
        g_SkipStatusOtherCount + _
        g_ClosedRoCount + _
        g_NotOnFileRoCount + _
        g_SkipVehidNotOnFileCount + _
        g_SkipNoCloseoutTextCount + _
        g_SkipNoPartsChargedCount + _
        g_SkipUnsupportedWarrantyCount + _
        g_LeftOpenManualCount + _
        g_ErrorInMainCount + _
        g_NoResultRecordedCount + _
        g_SummaryOtherOutcomeCount

    summaryText = "DONE" & vbCrLf & _
        "ROs Reviewed: " & g_ReviewedROCount & vbCrLf & _
        "ROs Posted: " & g_FiledROCount & vbCrLf & _
        "Skips - Specific ROs: " & g_SkipConfiguredCount & vbCrLf & _
        "Skips - Non-compliant tech code: " & g_SkipTechCodeCount & vbCrLf & _
        "Skips - Other Terms: " & g_SkipBlacklistCount & vbCrLf & _
        "Skips - Open: " & g_SkipStatusOpenCount & vbCrLf & _
        "Skips - Pre-Assigned: " & g_SkipStatusPreassignedCount & vbCrLf & _
        "Skips - Other Statuses: " & g_SkipStatusOtherCount & vbCrLf & _
        "Closed (already): " & g_ClosedRoCount & vbCrLf & _
        "Not On File: " & g_NotOnFileRoCount & vbCrLf & _
        "Skipped - VEHID not on file: " & g_SkipVehidNotOnFileCount & vbCrLf & _
        "Skipped - No closeout text: " & g_SkipNoCloseoutTextCount & vbCrLf & _
        "Skipped - No parts charged: " & g_SkipNoPartsChargedCount & vbCrLf & _
        "Skipped - Unsupported warranty ltype: " & g_SkipUnsupportedWarrantyCount & vbCrLf & _
        "Left Open for manual closing: " & g_LeftOpenManualCount & vbCrLf & _
        "Errors in Main: " & g_ErrorInMainCount & vbCrLf & _
        "No result recorded: " & g_NoResultRecordedCount & vbCrLf & _
        "Other Outcomes: " & g_SummaryOtherOutcomeCount & vbCrLf & _
        "Accounted Total: " & accountedTotal & vbCrLf & _
        "Older ROs Attempted (subset): " & g_OlderRoAttemptCount

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

' --- Test 1: All named counters appear with correct values ---
Call ResetCounters()
g_ReviewedROCount = 111
g_FiledROCount = 59
g_SkipConfiguredCount = 0
g_SkipTechCodeCount = 2
g_SkipBlacklistCount = 0
g_SkipStatusOpenCount = 1
g_SkipStatusPreassignedCount = 0
g_SkipStatusOtherCount = 1
g_ClosedRoCount = 0
g_NotOnFileRoCount = 0
g_SkipVehidNotOnFileCount = 0
g_SkipNoCloseoutTextCount = 0
g_SkipNoPartsChargedCount = 38
g_SkipUnsupportedWarrantyCount = 10
g_LeftOpenManualCount = 0
g_ErrorInMainCount = 0
g_NoResultRecordedCount = 0
g_SummaryOtherOutcomeCount = 0
g_OlderRoAttemptCount = 0

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
AssertContains "Other Outcomes line correct", summary1, "Other Outcomes: 0"

' --- Test 2: Accounted Total equals sum of all counters ---
' Sum: 59 + 0 + 2 + 0 + 1 + 0 + 1 + 0 + 0 + 0 + 0 + 38 + 10 + 0 + 0 + 0 + 0 = 111
AssertContains "Accounted Total equals sum", summary1, "Accounted Total: 111"

' --- Test 3: Other Outcomes section absent when count is zero ---
AssertAbsent "Other Outcome Breakdown absent when count is zero", summary1, "Other Outcome Breakdown:"
AssertAbsent "Other Outcome Raw Results absent when count is zero", summary1, "Other Outcome Raw Results:"

' --- Test 4: Other Outcome Breakdown appears when count > 0 ---
Call ResetCounters()
g_ReviewedROCount = 5
g_FiledROCount = 3
g_SummaryOtherOutcomeCount = 2
g_SummaryOtherOutcomeBreakdown.Add "Skipped - FCA dialog handler not configured", 2
g_SummaryOtherOutcomeRawBreakdown.Add "Skipped - FCA dialog handler not configured for ltype XY", 2

Dim summary2
summary2 = BuildSessionSummary()

AssertContains "Other Outcomes line shows count", summary2, "Other Outcomes: 2"
AssertContains "Other Outcome Breakdown section appears", summary2, "Other Outcome Breakdown:"
AssertContains "Other Outcome Breakdown entry appears", summary2, "Skipped - FCA dialog handler not configured: 2"
AssertContains "Other Outcome Raw Results section appears", summary2, "Other Outcome Raw Results:"

' --- Test 5: Accounted Total includes Other Outcomes count ---
' Sum: 3 posted + 2 other = 5
AssertContains "Accounted Total includes Other Outcomes", summary2, "Accounted Total: 5"

' --- Test 6: Zero unsupported warranty ltype shows as 0 (not absent) ---
Call ResetCounters()
Dim summary3
summary3 = BuildSessionSummary()
AssertContains "Unsupported warranty ltype line present even when zero", summary3, "Skipped - Unsupported warranty ltype: 0"
AssertContains "Zero counters still produce Accounted Total: 0", summary3, "Accounted Total: 0"

WScript.Echo ""
If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All " & g_Pass & " session summary tests passed."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & g_Fail & " test(s) failed."
    WScript.Quit 1
End If
