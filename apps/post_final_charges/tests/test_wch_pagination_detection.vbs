'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestWchPaginationDetection
' **DATE CREATED:** 2026-04-15
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Behavioral regression tests for pagination-aware WCH detection.
' Validates:
'   1) WCH found on page 2+ is detected.
'   2) No-WCH multi-page scans continue until END OF DISPLAY.
'   3) Page advance command sequence is N then <NumpadEnter>.
'   4) A 500ms pause is applied after each page advance.
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_Pass, g_Fail
g_Pass = 0
g_Fail = 0

Dim g_bzhao

Class FakeBzhao
    Private m_pages
    Private m_currentPage
    Private m_pendingN
    Private m_pendingB
    Private m_allowAdvance

    Private m_keys()
    Private m_keyCount

    Private m_pauses()
    Private m_pauseCount

    Private Sub Class_Initialize()
        m_currentPage = 0
        m_pendingN = False
        m_pendingB = False
        m_allowAdvance = True
        m_keyCount = 0
        m_pauseCount = 0
        ReDim m_keys(-1)
        ReDim m_pauses(-1)
    End Sub

    Public Sub SetPages(pages)
        m_pages = pages
        m_currentPage = 0
        m_pendingN = False
        m_pendingB = False
        m_allowAdvance = True
        m_keyCount = 0
        m_pauseCount = 0
        ReDim m_keys(-1)
        ReDim m_pauses(-1)
    End Sub

    Public Sub SetAllowAdvance(ByVal allowAdvance)
        m_allowAdvance = CBool(allowAdvance)
    End Sub

    Public Sub ReadScreen(ByRef outText, ByVal length, ByVal row, ByVal col)
        Dim pageBuf, pos
        pageBuf = m_pages(m_currentPage)
        pos = ((row - 1) * 80) + col
        If pos < 1 Then pos = 1
        outText = Mid(pageBuf, pos, length)
    End Sub

    Public Sub SendKey(ByVal keyText)
        m_keyCount = m_keyCount + 1
        ReDim Preserve m_keys(m_keyCount - 1)
        m_keys(m_keyCount - 1) = keyText

        If keyText = "N" Then
            m_pendingN = True
            m_pendingB = False
            Exit Sub
        End If

        If keyText = "B" Then
            m_pendingB = True
            m_pendingN = False
            Exit Sub
        End If

        If keyText = "<NumpadEnter>" Then
            If m_pendingN Then
                If m_allowAdvance Then
                    If m_currentPage < UBound(m_pages) Then m_currentPage = m_currentPage + 1
                End If
            ElseIf m_pendingB Then
                If m_currentPage > 0 Then m_currentPage = m_currentPage - 1
            End If
            m_pendingN = False
            m_pendingB = False
        End If
    End Sub

    Public Sub Pause(ByVal ms)
        m_pauseCount = m_pauseCount + 1
        ReDim Preserve m_pauses(m_pauseCount - 1)
        m_pauses(m_pauseCount - 1) = CLng(ms)
    End Sub

    Public Property Get KeyCount()
        KeyCount = m_keyCount
    End Property

    Public Property Get KeyAt(ByVal index)
        If index >= 0 And index < m_keyCount Then
            KeyAt = m_keys(index)
        Else
            KeyAt = ""
        End If
    End Property

    Public Property Get PauseCount()
        PauseCount = m_pauseCount
    End Property

    Public Property Get PauseAt(ByVal index)
        If index >= 0 And index < m_pauseCount Then
            PauseAt = m_pauses(index)
        Else
            PauseAt = -1
        End If
    End Property

    Public Property Get CurrentPageIndex()
        CurrentPageIndex = m_currentPage
    End Property

    Public Property Get HasPendingCommand()
        HasPendingCommand = (m_pendingN Or m_pendingB)
    End Property
End Class

Sub LogEvent(ByVal a, ByVal b, ByVal c, ByVal d, ByVal e, ByVal f)
    ' no-op for unit tests
End Sub

Function SetRow(ByVal pageBuf, ByVal row, ByVal rowText)
    Dim content, pos
    content = Left(rowText & String(80, " "), 80)
    pos = ((row - 1) * 80) + 1
    SetRow = Left(pageBuf, pos - 1) & content & Mid(pageBuf, pos + 80)
End Function

Function BuildPage(ByVal row22Marker, ByVal includeWch)
    Dim pageBuf
    pageBuf = String(24 * 80, " ")

    If includeWch Then
        ' Place WCH in a detail L-row region.
        pageBuf = SetRow(pageBuf, 11, "   L1 SAMPLE LABOR DESCRIPTION                    WCH")
    Else
        pageBuf = SetRow(pageBuf, 11, "   L1 SAMPLE LABOR DESCRIPTION                    C")
    End If

    pageBuf = SetRow(pageBuf, 22, row22Marker)
    BuildPage = pageBuf
End Function

Function HasWchOnAnyDetailPage()
    Dim row, buf, pageIndicator
    Dim foundWch, doneScanning
    Dim pagesAdvanced, p
    Dim preSig, postSig, preSig2, postSig2, preMarker, postMarker
    Dim maxPageAdvances

    HasWchOnAnyDetailPage = False
    pagesAdvanced = 0
    doneScanning = False
    maxPageAdvances = 50

    Do While Not doneScanning
        foundWch = False

        On Error Resume Next
        For row = 9 To 22
            buf = ""
            g_bzhao.ReadScreen buf, 80, row, 1
            If Err.Number <> 0 Then
                Err.Clear
            Else
                If InStr(1, buf, "WCH", vbTextCompare) > 0 Then
                    foundWch = True
                    Exit For
                End If
            End If
        Next
        On Error GoTo 0

        If foundWch Then
            HasWchOnAnyDetailPage = True
            doneScanning = True
        Else
            pageIndicator = ""
            On Error Resume Next
            g_bzhao.ReadScreen pageIndicator, 80, 22, 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0

            If InStr(1, pageIndicator, "(END OF DISPLAY)", vbTextCompare) > 0 Then
                doneScanning = True
            ElseIf InStr(1, pageIndicator, "(MORE ON NEXT SCREEN)", vbTextCompare) > 0 Then
                preMarker = pageIndicator
                preSig = ""
                preSig2 = ""
                On Error Resume Next
                g_bzhao.ReadScreen preSig, 80, 9, 1
                If Err.Number <> 0 Then Err.Clear
                g_bzhao.ReadScreen preSig2, 80, 10, 1
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0

                On Error Resume Next
                g_bzhao.SendKey "N"
                g_bzhao.SendKey "<NumpadEnter>"
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                g_bzhao.Pause 500

                postMarker = ""
                postSig = ""
                postSig2 = ""
                On Error Resume Next
                g_bzhao.ReadScreen postMarker, 80, 22, 1
                If Err.Number <> 0 Then Err.Clear
                g_bzhao.ReadScreen postSig, 80, 9, 1
                If Err.Number <> 0 Then Err.Clear
                g_bzhao.ReadScreen postSig2, 80, 10, 1
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0

                If postMarker = preMarker And postSig = preSig And postSig2 = preSig2 Then
                    doneScanning = True
                Else
                    pagesAdvanced = pagesAdvanced + 1
                    If pagesAdvanced >= maxPageAdvances Then
                        doneScanning = True
                    End If
                End If
            Else
                doneScanning = True
            End If
        End If
    Loop

    If pagesAdvanced > 0 Then
        For p = 1 To pagesAdvanced
            On Error Resume Next
            g_bzhao.SendKey "B"
            g_bzhao.SendKey "<NumpadEnter>"
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            g_bzhao.Pause 500
        Next
    End If
End Function

Sub AssertTrue(ByVal label, ByVal value)
    If value Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (expected True)"
    End If
End Sub

Sub AssertFalse(ByVal label, ByVal value)
    If Not value Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (expected False)"
    End If
End Sub

Sub AssertEqual(ByVal label, ByVal expected, ByVal actual)
    If CStr(expected) = CStr(actual) Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (expected='" & expected & "', actual='" & actual & "')"
    End If
End Sub

Sub AssertCommandPairs(ByVal label, ByRef mockObj)
    Dim i, cmd1, cmd2, ok
    ok = True
    If (mockObj.KeyCount Mod 2) <> 0 Then ok = False
    For i = 0 To mockObj.KeyCount - 1 Step 2
        cmd1 = mockObj.KeyAt(i)
        cmd2 = mockObj.KeyAt(i + 1)
        If Not ((cmd1 = "N" Or cmd1 = "B") And cmd2 = "<NumpadEnter>") Then
            ok = False
            Exit For
        End If
    Next
    If ok Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (commands are not paired as N/B + <NumpadEnter>)"
    End If
End Sub

Sub AssertAllPausesAre500(ByVal label, ByRef mockObj)
    Dim i, ok
    ok = True
    For i = 0 To mockObj.PauseCount - 1
        If CLng(mockObj.PauseAt(i)) <> 500 Then
            ok = False
            Exit For
        End If
    Next
    If ok Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (non-500 pause found)"
    End If
End Sub

WScript.Echo "WCH Pagination Detection Tests"
WScript.Echo "=============================="

' Test 1: Positive (WCH on page 2+)
Dim mock1, pages1
Set mock1 = New FakeBzhao
pages1 = Array( _
    BuildPage("(MORE ON NEXT SCREEN)", False), _
    BuildPage("(END OF DISPLAY)", True) _
)
mock1.SetPages pages1
Set g_bzhao = mock1
AssertTrue "Positive: detects WCH on page 2", HasWchOnAnyDetailPage()
AssertEqual "Positive: returns to page 1", 0, mock1.CurrentPageIndex
AssertEqual "Positive: one page-down and one page-up sequence", 4, mock1.KeyCount
AssertCommandPairs "Positive: N/B commands paired with Enter", mock1
AssertFalse "Positive: command buffer clear after scan", mock1.HasPendingCommand
AssertEqual "Positive: applies pause after each navigation", 2, mock1.PauseCount
AssertAllPausesAre500 "Positive: all pauses are 500ms", mock1

' Test 2: Negative multi-page (3+ pages, no WCH)
Dim mock2, pages2
Set mock2 = New FakeBzhao
pages2 = Array( _
    SetRow(BuildPage("(MORE ON NEXT SCREEN)", False), 10, "PAGE 1 MARKER"), _
    SetRow(BuildPage("(MORE ON NEXT SCREEN)", False), 10, "PAGE 2 MARKER"), _
    SetRow(BuildPage("(END OF DISPLAY)", False), 10, "PAGE 3 MARKER") _
)
mock2.SetPages pages2
Set g_bzhao = mock2
AssertFalse "Negative: no WCH across 3 pages", HasWchOnAnyDetailPage()
AssertEqual "Negative: returns to page 1", 0, mock2.CurrentPageIndex
AssertEqual "Negative: traverses until END then returns", 8, mock2.KeyCount
AssertCommandPairs "Negative: N/B commands paired with Enter", mock2
AssertFalse "Negative: command buffer clear after scan", mock2.HasPendingCommand
AssertEqual "Negative: pause called per navigation", 4, mock2.PauseCount
AssertAllPausesAre500 "Negative: all pauses are 500ms", mock2

' Test 3: Safety (MORE marker repeats but page does not advance)
Dim mock3, pages3
Set mock3 = New FakeBzhao
pages3 = Array( _
    BuildPage("(MORE ON NEXT SCREEN)", False) _
)
mock3.SetPages pages3
mock3.SetAllowAdvance False
Set g_bzhao = mock3
AssertFalse "Safety: stops when page does not advance", HasWchOnAnyDetailPage()
AssertEqual "Safety: remains on page 1", 0, mock3.CurrentPageIndex
AssertEqual "Safety: attempts one page-down sequence", 2, mock3.KeyCount
AssertCommandPairs "Safety: N/B commands paired with Enter", mock3
AssertFalse "Safety: command buffer clear after scan", mock3.HasPendingCommand
AssertEqual "Safety: pause called once for attempted advance", 1, mock3.PauseCount
AssertAllPausesAre500 "Safety: all pauses are 500ms", mock3

WScript.Echo ""
If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All " & g_Pass & " WCH pagination tests passed."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & g_Fail & " test(s) failed."
    WScript.Quit 1
End If
