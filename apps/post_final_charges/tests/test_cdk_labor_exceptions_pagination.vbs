'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestCdkLaborExceptionsPagination
' **DATE CREATED:** 2026-04-15
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Behavioral tests for exception-aware no-parts gating across paginated RO detail
' screens. Validates:
'   1) No parts + exception tech (WCH/WT/WF) proceeds.
'   2) No parts + non-exception tech skips with offending code.
'   3) Mixed exception/non-exception no-parts skips (conservative behavior).
'   4) Paging uses N/B + <NumpadEnter> pairs and returns to page 1.
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_Pass, g_Fail
g_Pass = 0
g_Fail = 0

Dim g_bzhao
Dim g_arrCDKExceptions
Dim g_arrCDKDescriptionExceptions
g_arrCDKExceptions = Array("WCH", "WT", "WF")
g_arrCDKDescriptionExceptions = Array("check and adjust")

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

Function SetColText(ByVal rowText, ByVal colNum, ByVal textValue)
    Dim base, head, tail
    base = Left(rowText & String(80, " "), 80)
    head = Left(base, colNum - 1)
    tail = Mid(base, colNum + Len(textValue))
    SetColText = Left(head & textValue & tail & String(80, " "), 80)
End Function

Function BuildHeaderRow(ByVal lineLetter)
    Dim rowText
    rowText = String(80, " ")
    rowText = SetColText(rowText, 1, lineLetter)
    BuildHeaderRow = rowText
End Function

Function BuildLRow(ByVal ltypeCode, ByVal descText)
    ' Matches real CDK L-row layout: L1 at col 4-5, description at col 7-41, LTYPE at col 50-55.
    ' Consistent with IsWchLine() (col 50-55) and GetPartsNeededLaborDesc() (col 7, 35 chars).
    Dim rowText
    rowText = String(80, " ")
    rowText = SetColText(rowText, 4, "L1")
    rowText = SetColText(rowText, 7, Left(descText & String(35, " "), 35))
    rowText = SetColText(rowText, 50, ltypeCode)
    BuildLRow = rowText
End Function

Function BuildPartRow(ByVal amountText)
    Dim rowText
    rowText = String(80, " ")
    rowText = SetColText(rowText, 6, "P1")
    rowText = SetColText(rowText, 9, "TESTPART")
    rowText = SetColText(rowText, 70, amountText)
    BuildPartRow = rowText
End Function

Function BuildPage(ByVal row22Marker, ByVal ltypeCode, ByVal lineDesc, ByVal includePartLine, ByVal partAmountText)
    Dim pageBuf
    pageBuf = String(24 * 80, " ")
    pageBuf = SetRow(pageBuf, 9, BuildHeaderRow("A"))
    pageBuf = SetRow(pageBuf, 10, BuildLRow(ltypeCode, lineDesc))
    If includePartLine Then
        pageBuf = SetRow(pageBuf, 12, BuildPartRow(partAmountText))
    End If
    pageBuf = SetRow(pageBuf, 22, row22Marker)
    BuildPage = pageBuf
End Function

Function IsCdkLaborOnlyExceptionTech(techCode)
    IsCdkLaborOnlyExceptionTech = False
    If Not IsArray(g_arrCDKExceptions) Then Exit Function

    Dim i, normalized
    normalized = UCase(Trim(CStr(techCode)))
    If Len(normalized) = 0 Then Exit Function

    For i = 0 To UBound(g_arrCDKExceptions)
        If normalized = g_arrCDKExceptions(i) Then
            IsCdkLaborOnlyExceptionTech = True
            Exit Function
        End If
    Next
End Function

Function IsCdkLaborOnlyExceptionDesc(descText)
    IsCdkLaborOnlyExceptionDesc = False
    If Not IsArray(g_arrCDKDescriptionExceptions) Then Exit Function

    Dim i, lowerDesc
    lowerDesc = LCase(Trim(CStr(descText)))
    If Len(lowerDesc) = 0 Then Exit Function

    For i = 0 To UBound(g_arrCDKDescriptionExceptions)
        If Len(g_arrCDKDescriptionExceptions(i)) > 0 Then
            If InStr(1, lowerDesc, g_arrCDKDescriptionExceptions(i), vbTextCompare) > 0 Then
                IsCdkLaborOnlyExceptionDesc = True
                Exit Function
            End If
        End If
    Next
End Function

Function EvaluatePartsChargedGate(ByRef skipReason)
    Dim row, buf, amtRaw, amtVal
    Dim pageIndicator
    Dim doneScanning, pagesAdvanced, p
    Dim preSig, postSig, preSig2, postSig2, preMarker, postMarker
    Dim maxPageAdvances
    Dim hasAnyPartLine, hasChargedPart
    Dim firstExceptionEvidence, firstNonExceptionTech
    Dim lTypeCode, lDesc, lHasTechEx, lHasDescEx

    EvaluatePartsChargedGate = False
    skipReason = "Skipped - No parts charged"

    doneScanning = False
    pagesAdvanced = 0
    maxPageAdvances = 50
    hasAnyPartLine = False
    hasChargedPart = False
    firstExceptionEvidence = ""
    firstNonExceptionTech = ""

    Do While Not doneScanning
        For row = 9 To 22
            buf = ""
            On Error Resume Next
            g_bzhao.ReadScreen buf, 80, row, 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0

            If Len(buf) >= 80 Then
                If Mid(buf, 6, 1) = "P" And IsNumeric(Mid(buf, 7, 1)) Then
                    hasAnyPartLine = True
                    amtRaw = Trim(Mid(buf, 70, 11))
                    amtVal = 0
                    If IsNumeric(amtRaw) Then amtVal = CDbl(amtRaw)
                    If amtVal > 0 Then
                        hasChargedPart = True
                        Exit For
                    End If
                End If
            End If

            ' L-rows carry LTYPE (col 50-55) and description (col 7-41).
            ' Matches the same layout used by IsWchLine() and GetPartsNeededLaborDesc().
            If Len(buf) >= 55 And Mid(buf, 4, 1) = "L" And IsNumeric(Mid(buf, 5, 1)) Then
                lTypeCode = UCase(Trim(Mid(buf, 50, 6)))
                lDesc = Trim(Mid(buf, 7, 35))
                lHasTechEx = (Len(lTypeCode) > 0 And IsCdkLaborOnlyExceptionTech(lTypeCode))
                lHasDescEx = IsCdkLaborOnlyExceptionDesc(lDesc)

                If lHasTechEx Or lHasDescEx Then
                    If Len(firstExceptionEvidence) = 0 Then
                        If lHasTechEx Then
                            firstExceptionEvidence = "LTYPE " & lTypeCode
                        Else
                            firstExceptionEvidence = "description """ & lDesc & """"
                        End If
                    End If
                Else
                    If Len(lTypeCode) > 0 Then
                        If Len(firstNonExceptionTech) = 0 Then firstNonExceptionTech = lTypeCode
                    End If
                End If
            End If
        Next

        If hasChargedPart Then
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

    If hasChargedPart Then
        EvaluatePartsChargedGate = True
        Exit Function
    End If

    If Not hasAnyPartLine Then
        If Len(firstNonExceptionTech) > 0 Then
            skipReason = "Skipped - No parts charged: " & firstNonExceptionTech
            Exit Function
        End If

        If Len(firstExceptionEvidence) > 0 Then
            EvaluatePartsChargedGate = True
            Exit Function
        End If
    End If

    If Len(firstNonExceptionTech) > 0 Then
        skipReason = "Skipped - No parts charged: " & firstNonExceptionTech
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

WScript.Echo "CDK Labor Exception Pagination Tests"
WScript.Echo "===================================="

Dim mock1, pages1, allow1, reason1
Set mock1 = New FakeBzhao
pages1 = Array( _
    BuildPage("(MORE ON NEXT SCREEN)", "CP", "HEADER", False, ""), _
    BuildPage("(END OF DISPLAY)", "WCH", "HEADER", False, "") _
)
mock1.SetPages pages1
Set g_bzhao = mock1
allow1 = EvaluatePartsChargedGate(reason1)
AssertFalse "Mixed no-parts (CP + WCH) skips conservatively", allow1
AssertEqual "Mixed no-parts reason includes first non-exception code", "Skipped - No parts charged: CP", reason1
AssertEqual "Mixed no-parts returns to page 1", 0, mock1.CurrentPageIndex
AssertCommandPairs "Mixed no-parts uses N/B paging with Enter", mock1

Dim mock2, pages2, allow2, reason2
Set mock2 = New FakeBzhao
pages2 = Array( _
    BuildPage("(END OF DISPLAY)", "CP", "HEADER", False, "") _
)
mock2.SetPages pages2
Set g_bzhao = mock2
allow2 = EvaluatePartsChargedGate(reason2)
AssertFalse "No-parts with CP skips", allow2
AssertEqual "No-parts CP reason includes code", "Skipped - No parts charged: CP", reason2
AssertEqual "No-parts CP requires no navigation", 0, mock2.KeyCount

Dim mock3, pages3, allow3, reason3
Set mock3 = New FakeBzhao
pages3 = Array( _
    BuildPage("(MORE ON NEXT SCREEN)", "WT", "HEADER", False, ""), _
    BuildPage("(END OF DISPLAY)", "WF", "HEADER", False, "") _
)
mock3.SetPages pages3
Set g_bzhao = mock3
allow3 = EvaluatePartsChargedGate(reason3)
AssertTrue "No-parts with exception-only tech codes proceeds", allow3
AssertEqual "No-parts exception leaves default reason untouched", "Skipped - No parts charged", reason3
AssertEqual "No-parts exception returns to page 1", 0, mock3.CurrentPageIndex
AssertCommandPairs "No-parts exception uses N/B paging with Enter", mock3

Dim mock4, pages4, allow4, reason4
Set mock4 = New FakeBzhao
pages4 = Array( _
    BuildPage("(MORE ON NEXT SCREEN)", "CP", "HEADER", True, "0.00"), _
    BuildPage("(END OF DISPLAY)", "CP", "HEADER", True, "25.00") _
)
mock4.SetPages pages4
Set g_bzhao = mock4
allow4 = EvaluatePartsChargedGate(reason4)
AssertTrue "Charged part on later page proceeds", allow4
AssertEqual "Charged-part scan still returns to page 1", 0, mock4.CurrentPageIndex
AssertCommandPairs "Charged-part scan uses N/B paging with Enter", mock4

Dim mock5, pages5, allow5, reason5
Set mock5 = New FakeBzhao
pages5 = Array( _
    BuildPage("(END OF DISPLAY)", "CP", "CHECK AND ADJUST FRONT BRAKES", False, "") _
)
mock5.SetPages pages5
Set g_bzhao = mock5
allow5 = EvaluatePartsChargedGate(reason5)
AssertTrue "No-parts with labor-only description clue proceeds", allow5
AssertEqual "Description exception scenario requires no navigation", 0, mock5.KeyCount

WScript.Echo ""
If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All " & g_Pass & " CDK labor exception tests passed."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & g_Fail & " test(s) failed."
    WScript.Quit 1
End If
