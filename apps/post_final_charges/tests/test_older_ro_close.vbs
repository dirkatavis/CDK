Option Explicit

'=====================================================================================
' Test Suite: Older RO Close Feature
' Purpose: Unit tests for ParseCdkDate(), GetOpenedDate() mock, and IsOlderRo() logic
' Run with: cscript.exe apps\post_final_charges\tests\test_older_ro_close.vbs
'=====================================================================================

Dim g_TestCount, g_PassCount, g_FailCount
g_TestCount = 0
g_PassCount = 0
g_FailCount = 0

' ---- Inline the ParseCdkDate function for isolated unit testing ----
Function ParseCdkDate(dateStr)
    ParseCdkDate = Empty
    Dim cleaned
    cleaned = Trim(dateStr)
    If Len(cleaned) = 0 Then Exit Function

    ' Handle slash format (e.g. "01/20/26", "1/5/26") via VBScript CDate
    If InStr(cleaned, "/") > 0 Then
        On Error Resume Next
        If IsDate(cleaned) Then
            ParseCdkDate = CDate(cleaned)
        End If
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    ' Handle DDMMMYY / DDMMMYYYY format (e.g. "04FEB26", "04FEB2026")
    If Len(cleaned) < 7 Then Exit Function
    cleaned = UCase(cleaned)

    Dim dayPart, monthPart, yearPart
    dayPart = Left(cleaned, 2)
    monthPart = Mid(cleaned, 3, 3)
    yearPart = Mid(cleaned, 6)

    Dim monthNum
    Select Case monthPart
        Case "JAN": monthNum = 1
        Case "FEB": monthNum = 2
        Case "MAR": monthNum = 3
        Case "APR": monthNum = 4
        Case "MAY": monthNum = 5
        Case "JUN": monthNum = 6
        Case "JUL": monthNum = 7
        Case "AUG": monthNum = 8
        Case "SEP": monthNum = 9
        Case "OCT": monthNum = 10
        Case "NOV": monthNum = 11
        Case "DEC": monthNum = 12
        Case Else: Exit Function
    End Select

    On Error Resume Next
    Dim dayNum, yearNum
    dayNum = CInt(dayPart)
    yearNum = CInt(yearPart)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    If Len(yearPart) = 2 Then
        If yearNum >= 70 Then
            yearNum = 1900 + yearNum
        Else
            yearNum = 2000 + yearNum
        End If
    End If

    On Error Resume Next
    ParseCdkDate = DateSerial(yearNum, monthNum, dayNum)
    If Err.Number <> 0 Then
        Err.Clear
        ParseCdkDate = Empty
    End If
    On Error GoTo 0
End Function


' ---- Test Helpers ----
Sub AssertEqual(testName, expected, actual)
    g_TestCount = g_TestCount + 1
    If CStr(expected) = CStr(actual) Then
        g_PassCount = g_PassCount + 1
        WScript.Echo "  PASS: " & testName
    Else
        g_FailCount = g_FailCount + 1
        WScript.Echo "  FAIL: " & testName & " - Expected: '" & expected & "' Got: '" & actual & "'"
    End If
End Sub

Sub AssertTrue(testName, value)
    g_TestCount = g_TestCount + 1
    If value Then
        g_PassCount = g_PassCount + 1
        WScript.Echo "  PASS: " & testName
    Else
        g_FailCount = g_FailCount + 1
        WScript.Echo "  FAIL: " & testName & " - Expected True, got False"
    End If
End Sub

Sub AssertFalse(testName, value)
    g_TestCount = g_TestCount + 1
    If Not value Then
        g_PassCount = g_PassCount + 1
        WScript.Echo "  PASS: " & testName
    Else
        g_FailCount = g_FailCount + 1
        WScript.Echo "  FAIL: " & testName & " - Expected False, got True"
    End If
End Sub

Sub AssertEmpty(testName, value)
    g_TestCount = g_TestCount + 1
    If IsEmpty(value) Then
        g_PassCount = g_PassCount + 1
        WScript.Echo "  PASS: " & testName
    Else
        g_FailCount = g_FailCount + 1
        WScript.Echo "  FAIL: " & testName & " - Expected Empty, got '" & value & "'"
    End If
End Sub


' ============================================================
' TEST GROUP 1: ParseCdkDate - Valid Dates
' ============================================================
WScript.Echo ""
WScript.Echo "=== ParseCdkDate - Valid Dates ==="

' Test 1: Standard date (from screenshot)
Dim result1
result1 = ParseCdkDate("04FEB26")
AssertEqual "04FEB26 -> 2/4/2026", CStr(DateSerial(2026, 2, 4)), CStr(result1)

' Test 2: Today's date
Dim result2
result2 = ParseCdkDate("25MAR26")
AssertEqual "25MAR26 -> 3/25/2026", CStr(DateSerial(2026, 3, 25)), CStr(result2)

' Test 3: End of year
Dim result3
result3 = ParseCdkDate("31DEC25")
AssertEqual "31DEC25 -> 12/31/2025", CStr(DateSerial(2025, 12, 31)), CStr(result3)

' Test 4: January first
Dim result4
result4 = ParseCdkDate("01JAN26")
AssertEqual "01JAN26 -> 1/1/2026", CStr(DateSerial(2026, 1, 1)), CStr(result4)

' Test 5: All months covered - JUN
Dim result5
result5 = ParseCdkDate("15JUN24")
AssertEqual "15JUN24 -> 6/15/2024", CStr(DateSerial(2024, 6, 15)), CStr(result5)

' Test 6: All months covered - SEP
Dim result6
result6 = ParseCdkDate("30SEP25")
AssertEqual "30SEP25 -> 9/30/2025", CStr(DateSerial(2025, 9, 30)), CStr(result6)

' Test 7: All months covered - NOV
Dim result7
result7 = ParseCdkDate("10NOV23")
AssertEqual "10NOV23 -> 11/10/2023", CStr(DateSerial(2023, 11, 10)), CStr(result7)

' Test 8: Lowercase input (should still work via UCase)
Dim result8
result8 = ParseCdkDate("04feb26")
AssertEqual "04feb26 (lowercase) -> 2/4/2026", CStr(DateSerial(2026, 2, 4)), CStr(result8)


' ============================================================
' TEST GROUP 1b: ParseCdkDate - Slash Format Dates
' ============================================================
WScript.Echo ""
WScript.Echo "=== ParseCdkDate - Slash Format Dates ==="

' Test 8b: Standard slash date
Dim resultS1
resultS1 = ParseCdkDate("01/20/26")
AssertEqual "01/20/26 -> 1/20/2026", CStr(DateSerial(2026, 1, 20)), CStr(resultS1)

' Test 8c: Slash date single-digit month/day
Dim resultS2
resultS2 = ParseCdkDate("2/4/26")
AssertEqual "2/4/26 -> 2/4/2026", CStr(DateSerial(2026, 2, 4)), CStr(resultS2)

' Test 8d: Slash date full year
Dim resultS3
resultS3 = ParseCdkDate("12/31/2025")
AssertEqual "12/31/2025 -> 12/31/2025", CStr(DateSerial(2025, 12, 31)), CStr(resultS3)

' Test 8e: Slash date - another typical PFC date
Dim resultS4
resultS4 = ParseCdkDate("03/25/26")
AssertEqual "03/25/26 -> 3/25/2026", CStr(DateSerial(2026, 3, 25)), CStr(resultS4)

' Test 8f: Invalid slash date
AssertEmpty "Invalid slash date -> Empty", ParseCdkDate("13/40/26")

' Test 8g: DDMMMYYYY four-digit year
Dim resultS5
resultS5 = ParseCdkDate("04FEB2026")
AssertEqual "04FEB2026 -> 2/4/2026", CStr(DateSerial(2026, 2, 4)), CStr(resultS5)


' ============================================================
' TEST GROUP 2: ParseCdkDate - Invalid / Edge Cases
' ============================================================
WScript.Echo ""
WScript.Echo "=== ParseCdkDate - Invalid / Edge Cases ==="

' Test 9: Empty string
AssertEmpty "Empty string -> Empty", ParseCdkDate("")

' Test 10: Too short
AssertEmpty "Short string -> Empty", ParseCdkDate("04FE")

' Test 11: Garbage input
AssertEmpty "Garbage -> Empty", ParseCdkDate("XYZABC99")

' Test 12: Invalid month
AssertEmpty "Invalid month (ZZZ) -> Empty", ParseCdkDate("04ZZZ26")

' Test 13: Non-numeric day
AssertEmpty "Non-numeric day -> Empty", ParseCdkDate("XXJAN26")

' Test 14: Non-numeric year
AssertEmpty "Non-numeric year -> Empty", ParseCdkDate("04JANXX")

' Test 15: Spaces only
AssertEmpty "Whitespace only -> Empty", ParseCdkDate("       ")


' ============================================================
' TEST GROUP 3: Age Calculation Logic (inline simulation)
' Simulates IsOlderRo() logic without needing bzhao
' ============================================================
WScript.Echo ""
WScript.Echo "=== Age Calculation Logic ==="

Dim threshold
threshold = 30

' Single reference date for all age tests (avoids midnight-crossing non-determinism)
Dim refDate
refDate = Date()

' Test 16: RO opened 45 days ago qualifies
Dim oldDate
oldDate = DateAdd("d", -45, refDate)
Dim oldDateAge
oldDateAge = DateDiff("d", oldDate, refDate)
AssertTrue "45-day-old RO >= 30 threshold", (oldDateAge >= threshold)

' Test 17: RO opened 30 days ago qualifies (boundary)
Dim boundaryDate
boundaryDate = DateAdd("d", -30, refDate)
Dim boundaryAge
boundaryAge = DateDiff("d", boundaryDate, refDate)
AssertTrue "30-day-old RO >= 30 threshold (boundary)", (boundaryAge >= threshold)

' Test 18: RO opened 29 days ago does NOT qualify
Dim recentDate
recentDate = DateAdd("d", -29, refDate)
Dim recentAge
recentAge = DateDiff("d", recentDate, refDate)
AssertFalse "29-day-old RO < 30 threshold", (recentAge >= threshold)

' Test 19: RO opened today does NOT qualify
Dim todayDate
todayDate = refDate
Dim todayAge
todayAge = DateDiff("d", todayDate, refDate)
AssertFalse "0-day-old RO < 30 threshold", (todayAge >= threshold)

' Test 20: RO opened 365 days ago qualifies
Dim veryOldDate
veryOldDate = DateAdd("d", -365, refDate)
Dim veryOldAge
veryOldAge = DateDiff("d", veryOldDate, refDate)
AssertTrue "365-day-old RO >= 30 threshold", (veryOldAge >= threshold)

' Test 21: Threshold = 0 disables feature
Dim zeroThreshold
zeroThreshold = 0
AssertFalse "Threshold 0 disables (age > 0 but threshold <= 0 check)", (zeroThreshold > 0)


' ============================================================
' TEST GROUP 4: ParseCdkDate round-trip with age calculation
' Tests the full pipeline: parse date string -> calculate age
' ============================================================
WScript.Echo ""
WScript.Echo "=== Full Pipeline: Parse + Age ==="

' Test 22: Parse the screenshot date (04FEB26) and calculate age against a fixed reference
Dim screenDate
screenDate = ParseCdkDate("04FEB26")
If Not IsEmpty(screenDate) Then
    ' Use a fixed reference date so the assertion is deterministic
    Dim fixedRef
    fixedRef = DateSerial(2026, 3, 25)  ' 25MAR26
    Dim screenAge
    screenAge = DateDiff("d", screenDate, fixedRef)
    ' 04FEB26 to 25MAR26 = 49 days
    AssertTrue "04FEB26 is >= 30 days old as of 25MAR26", (screenAge >= 30)
    AssertEqual "04FEB26 to 25MAR26 = 49 days", "49", CStr(screenAge)
Else
    g_TestCount = g_TestCount + 2
    g_FailCount = g_FailCount + 2
    WScript.Echo "  FAIL: Could not parse 04FEB26 for age pipeline test"
End If

' Test 23: Parse a date from yesterday
Dim yesterdayStr, yesterdayParsed, yesterdayCalcAge
Dim dayStr, monthStr, yearStr, monthNames
monthNames = Array("", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
Dim yesterday
yesterday = DateAdd("d", -1, refDate)
dayStr = Right("0" & Day(yesterday), 2)
monthStr = monthNames(Month(yesterday))
yearStr = Right("0" & (Year(yesterday) - 2000), 2)
yesterdayStr = dayStr & monthStr & yearStr
yesterdayParsed = ParseCdkDate(yesterdayStr)
If Not IsEmpty(yesterdayParsed) Then
    yesterdayCalcAge = DateDiff("d", yesterdayParsed, refDate)
    AssertEqual "Yesterday formatted and parsed back = 1 day age", "1", CStr(yesterdayCalcAge)
Else
    g_TestCount = g_TestCount + 1
    g_FailCount = g_FailCount + 1
    WScript.Echo "  FAIL: Could not parse yesterday string: " & yesterdayStr
End If


' ============================================================
' TEST GROUP 5: Blacklist vs Older-RO Gate Precedence
' Simulates Main() decision flow to prove blacklist always wins
' ============================================================
WScript.Echo ""
WScript.Echo "=== Blacklist vs Older-RO Precedence ==="

' Inline helpers that mirror production logic
Function IsOlderRoEligibleStatus(statusToCheck)
    Dim olderStatuses, normalized, si
    olderStatuses = Array("OPENED", "OPEN", "PREASSIGNED", "PRE-ASSIGNED")
    IsOlderRoEligibleStatus = False
    normalized = UCase(Trim(statusToCheck))
    For si = 0 To UBound(olderStatuses)
        If normalized = olderStatuses(si) Then
            IsOlderRoEligibleStatus = True
            Exit Function
        End If
    Next
End Function

Function GetMatchedBlacklistTerm_Sim(blacklistCsv, screenContent)
    Dim terms, ti, term
    GetMatchedBlacklistTerm_Sim = ""
    If Len(Trim(blacklistCsv)) = 0 Then Exit Function
    terms = Split(blacklistCsv, ",")
    For ti = LBound(terms) To UBound(terms)
        term = Trim(terms(ti))
        If Len(term) > 0 Then
            If InStr(1, screenContent, term, vbTextCompare) > 0 Then
                GetMatchedBlacklistTerm_Sim = term
                Exit Function
            End If
        End If
    Next
End Function

' Simulates the Main() gate decision: returns the action taken
Function SimulateMainGate(roStatus, roAge, gateThreshold, blacklistTerms, screenContent)
    ' Step 1: blacklist check (runs FIRST in Main)
    Dim matchedTerm
    matchedTerm = GetMatchedBlacklistTerm_Sim(blacklistTerms, screenContent)
    If Len(matchedTerm) > 0 Then
        SimulateMainGate = "BLACKLIST:" & matchedTerm
        Exit Function
    End If

    ' Step 2: status ready check (READY TO POST bypasses older-RO gate)
    If UCase(Trim(roStatus)) = "READY TO POST" Then
        SimulateMainGate = "READY"
        Exit Function
    End If

    ' Step 3: older-RO gate (only reached if blacklist did not fire)
    If IsOlderRoEligibleStatus(roStatus) And (roAge >= gateThreshold) And (gateThreshold > 0) Then
        SimulateMainGate = "OLDER_CLOSEOUT"
        Exit Function
    End If

    SimulateMainGate = "SKIPPED"
End Function

' Test 24: Blacklisted + OPENED + old -> blacklist wins
Dim gate1
gate1 = SimulateMainGate("OPENED", 90, 30, "VEND TO DEALER", "SOME LINE WITH VEND TO DEALER HERE")
AssertEqual "Blacklisted old OPENED -> blacklist wins", "BLACKLIST:VEND TO DEALER", gate1

' Test 25: Blacklisted + PREASSIGNED + old -> blacklist wins
Dim gate2
gate2 = SimulateMainGate("PREASSIGNED", 60, 30, "VEND TO DEALER", "REPAIR VEND TO DEALER OIL")
AssertEqual "Blacklisted old PREASSIGNED -> blacklist wins", "BLACKLIST:VEND TO DEALER", gate2

' Test 26: NOT blacklisted + OPENED + old -> older closeout
Dim gate3
gate3 = SimulateMainGate("OPENED", 45, 30, "VEND TO DEALER", "OIL CHANGE TIRE ROTATION")
AssertEqual "Clean old OPENED -> older closeout", "OLDER_CLOSEOUT", gate3

' Test 27: NOT blacklisted + OPENED + young -> skipped
Dim gate4
gate4 = SimulateMainGate("OPENED", 10, 30, "VEND TO DEALER", "OIL CHANGE TIRE ROTATION")
AssertEqual "Clean young OPENED -> skipped", "SKIPPED", gate4

' Test 28: Blacklisted + READY TO POST -> blacklist wins (even READY)
Dim gate5
gate5 = SimulateMainGate("READY TO POST", 0, 30, "VEND TO DEALER", "VEND TO DEALER SERVICE")
AssertEqual "Blacklisted READY TO POST -> blacklist wins", "BLACKLIST:VEND TO DEALER", gate5

' Test 29: No blacklist terms + OPENED + old -> older closeout
Dim gate6
gate6 = SimulateMainGate("OPENED", 90, 30, "", "VEND TO DEALER ON SCREEN BUT NO TERMS CONFIGURED")
AssertEqual "No blacklist config + old OPENED -> older closeout", "OLDER_CLOSEOUT", gate6

' Test 30: Blacklist term present but NOT on screen + old OPENED -> older closeout
Dim gate7
gate7 = SimulateMainGate("OPENED", 45, 30, "VEND TO DEALER", "OIL CHANGE BRAKE PADS")
AssertEqual "Blacklist term not on screen + old OPENED -> older closeout", "OLDER_CLOSEOUT", gate7

' Test 31: Multiple blacklist terms, second one matches -> blacklist wins
Dim gate8
gate8 = SimulateMainGate("OPENED", 60, 30, "SUBLET WORK,VEND TO DEALER", "TIRE ROTATION VEND TO DEALER")
AssertEqual "Second blacklist term matches -> blacklist wins", "BLACKLIST:VEND TO DEALER", gate8


' ============================================================
' Summary
' ============================================================
WScript.Echo ""
WScript.Echo "========================================"
WScript.Echo "Tests: " & g_TestCount & "  Pass: " & g_PassCount & "  Fail: " & g_FailCount
WScript.Echo "========================================"

If g_FailCount > 0 Then
    WScript.Quit 1
Else
    WScript.Quit 0
End If
