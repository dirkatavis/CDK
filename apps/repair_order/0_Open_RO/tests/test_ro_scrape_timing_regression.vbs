' ============================================================================
' test_ro_scrape_timing_regression.vbs
'
' PURPOSE: Regression test confirming that GetRepairOrderEnhanced() returns ""
' when the screen has NOT yet transitioned to the "Created repair order"
' confirmation screen — and returns the correct RO number when called at the
' right time.
'
' ROOT CAUSE DOCUMENTED HERE:
'   In Main(), GetRepairOrderEnhanced() is called BEFORE
'   WaitForPrompt "Created repair order|R.O. NUMBER". On fast sessions the
'   screen has not yet transitioned, so the scrape always misses and
'   LogEntryWithRO logs "MVA: XXXXXXXX" with no " - RO: " suffix.
'   The fix is to wait for the confirmation screen FIRST, scrape, then send F3.
'
' Usage:   cscript.exe //nologo test_ro_scrape_timing_regression.vbs
' ============================================================================
Option Explicit

Dim g_pass, g_fail
g_pass = 0
g_fail = 0

WScript.Echo "=== Open_RO: RO Scrape Timing Regression ==="
WScript.Echo ""

' ============================================================================
' Minimal mock — controls what ReadScreen returns without BlueZone
' ============================================================================
Class MockBzhao
    Public ScreenContent

    Private Sub Class_Initialize()
        ScreenContent = String(24 * 80, " ")
    End Sub

    Public Sub ReadScreen(ByRef buf, length, row, col)
        ' Return a left-slice of the screen buffer from position 1
        buf = Left(ScreenContent & String(24 * 80, " "), length)
    End Sub
End Class

Dim g_bzhao
Set g_bzhao = New MockBzhao

' ============================================================================
' GetRepairOrderEnhanced — verbatim copy from Open_RO.vbs (lines 240-265)
' Kept in sync: if this copy drifts from production, tests become stale.
' ============================================================================
Function GetRepairOrderEnhanced()
    Dim screenContent, screenLength, pos, startPos, ch, roNumber
    screenLength = 24 * 80

    g_bzhao.ReadScreen screenContent, screenLength, 1, 1
    pos = InStr(1, screenContent, "Created repair order ", vbTextCompare)

    If pos > 0 Then
        startPos = pos + 21 ' Length of "Created repair order "
        roNumber = ""

        Do While startPos <= Len(screenContent)
            ch = Mid(screenContent, startPos, 1)
            If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "Z") Then
                roNumber = roNumber & ch
                startPos = startPos + 1
            Else
                Exit Do
            End If
        Loop

        GetRepairOrderEnhanced = roNumber
    Else
        GetRepairOrderEnhanced = ""
    End If
End Function

' ============================================================================
' Helpers
' ============================================================================
Sub AssertEqual(testName, expected, actual)
    If CStr(expected) = CStr(actual) Then
        g_pass = g_pass + 1
        WScript.Echo "  [PASS] " & testName
    Else
        g_fail = g_fail + 1
        WScript.Echo "  [FAIL] " & testName & _
            " — expected '" & expected & "', got '" & actual & "'"
    End If
End Sub

' Build a 1920-char padded screen buffer with text starting at position 1
Function MakeScreen(text)
    MakeScreen = Left(text & String(24 * 80, " "), 24 * 80)
End Function

' Build a 1920-char buffer with text injected at a specific character offset (1-based)
Function MakeScreenAt(text, offset)
    Dim buf
    buf = String(24 * 80, " ")
    MakeScreenAt = Left(buf, offset - 1) & text & Mid(buf, offset + Len(text))
End Function

' ============================================================================
' BLOCK 1: Screen NOT yet transitioned — premature scrape returns ""
'
' This is the failing condition in production: GetRepairOrderEnhanced() is
' called right after "O.K. TO CLOSE RO" is dismissed, before
' WaitForPrompt("Created repair order|R.O. NUMBER") has confirmed the
' screen has changed. These tests assert the function returns "" in that state.
' ============================================================================
WScript.Echo "--- Block 1: Premature scrape (screen not yet ready) ---"

g_bzhao.ScreenContent = MakeScreen("O.K. TO CLOSE RO? ")
AssertEqual _
    "Returns '' when 'O.K. TO CLOSE RO' still visible (too early)", _
    "", GetRepairOrderEnhanced()

g_bzhao.ScreenContent = MakeScreen("MILEAGE IN   :")
AssertEqual _
    "Returns '' when 'MILEAGE IN' screen still showing", _
    "", GetRepairOrderEnhanced()

g_bzhao.ScreenContent = String(24 * 80, " ")
AssertEqual _
    "Returns '' on blank/transitioning screen", _
    "", GetRepairOrderEnhanced()

' ============================================================================
' BLOCK 2: Screen has "Created repair order" — scrape should succeed
'
' These document the correct call-site: AFTER WaitForPrompt has confirmed
' "Created repair order" is visible. The fix is to reorder Main() so the
' scrape happens here, not before the WaitForPrompt call.
' ============================================================================
WScript.Echo ""
WScript.Echo "--- Block 2: Correct screen present — RO should be scraped ---"

g_bzhao.ScreenContent = MakeScreen("Created repair order 878393")
AssertEqual _
    "Scrapes numeric RO from confirmation screen", _
    "878393", GetRepairOrderEnhanced()

g_bzhao.ScreenContent = MakeScreen("Created repair order 878394 on 04/07")
AssertEqual _
    "Stops at first non-alphanumeric boundary (space after RO)", _
    "878394", GetRepairOrderEnhanced()

' RO text mid-buffer — row 6 equivalent (offset 401)
g_bzhao.ScreenContent = MakeScreenAt("Created repair order 878396", 401)
AssertEqual _
    "Finds RO when text appears mid-buffer (position 401)", _
    "878396", GetRepairOrderEnhanced()

' ============================================================================
' BLOCK 3: Edge cases for GetRepairOrderEnhanced
' ============================================================================
WScript.Echo ""
WScript.Echo "--- Block 3: Edge cases ---"

' "R.O. NUMBER" variant on screen — function does NOT handle this form;
' documents that the function only fires on "Created repair order " prefix
g_bzhao.ScreenContent = MakeScreen("R.O. NUMBER    878395  ENTER OPTION")
AssertEqual _
    "Returns '' for 'R.O. NUMBER' variant (prefix mismatch — not handled)", _
    "", GetRepairOrderEnhanced()

' RO immediately at end of buffer — no trailing characters
g_bzhao.ScreenContent = Left("Created repair order " & "878397", 24 * 80)
AssertEqual _
    "Scrapes RO at end of buffer", _
    "878397", GetRepairOrderEnhanced()

' Empty buffer edge case
g_bzhao.ScreenContent = ""
AssertEqual _
    "Returns '' on empty screen content", _
    "", GetRepairOrderEnhanced()

' ============================================================================
' Summary
' ============================================================================
WScript.Echo ""
WScript.Echo "Results: " & g_pass & " passed, " & g_fail & " failed"
If g_fail > 0 Then
    WScript.Quit 1
Else
    WScript.Quit 0
End If
