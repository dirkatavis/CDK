'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestBlockedLTypeGate
' **DATE CREATED:** 2026-04-16
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Behavioral tests for HasBlockedLTypeOnAnyPage() and the SkipLaborLTypes
' config gate. Validates:
'   1) WCH/WV/WF L-rows are detected and block the RO.
'   2) LTYPE "I" L-rows are not blocked.
'   3) Empty SkipLaborLTypes config disables the gate entirely.
'   4) Empty LTYPE on an L-row is never matched (no false block).
'   5) Multi-page ROs: blocked LTYPE on page 2 is detected; page 1 returns.
'   6) Multi-page ROs: all-clean pages return empty (not blocked).
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_Pass, g_Fail
g_Pass = 0
g_Fail = 0

Dim g_bzhao
Dim g_arrSkipLaborLTypes

'--------------------------------------------------------------------
' Fake BlueZone object — serves pre-built page buffers on ReadScreen
'--------------------------------------------------------------------
Class FakeBzhao
    Private m_pages
    Private m_currentPage
    Private m_keys()
    Private m_keyCount
    Private m_pauses()
    Private m_pauseCount

    Private Sub Class_Initialize()
        m_currentPage = 0
        m_keyCount = 0
        m_pauseCount = 0
        ReDim m_keys(0)
        ReDim m_pauses(0)
    End Sub

    Public Sub SetPages(pages)
        m_pages = pages
        m_currentPage = 0
    End Sub

    Public Sub ReadScreen(ByRef buf, ByVal length, ByVal row, ByVal col)
        Dim page
        page = m_pages(m_currentPage)
        Dim startPos
        startPos = ((row - 1) * 80) + 1
        buf = Mid(page, startPos, length)
    End Sub

    Public Sub SendKey(ByVal key)
        ReDim Preserve m_keys(m_keyCount)
        m_keys(m_keyCount) = key
        m_keyCount = m_keyCount + 1
        If key = "N" Then
            If m_currentPage < UBound(m_pages) Then m_currentPage = m_currentPage + 1
        ElseIf key = "B" Then
            If m_currentPage > 0 Then m_currentPage = m_currentPage - 1
        End If
    End Sub

    Public Sub Pause(ByVal ms)
        ReDim Preserve m_pauses(m_pauseCount)
        m_pauses(m_pauseCount) = ms
        m_pauseCount = m_pauseCount + 1
    End Sub

    Public Function GetKeys()
        GetKeys = m_keys
    End Function

    Public Property Get CurrentPage()
        CurrentPage = m_currentPage
    End Property
End Class

'--------------------------------------------------------------------
' Row/page builders
'--------------------------------------------------------------------
Function PadTo80(s)
    PadTo80 = Left(s & String(80, " "), 80)
End Function

Function BuildLRow(ltypeCode, descText)
    ' L-row: L1 at col 4-5, description at col 7-41, LTYPE at col 50-55
    Dim r : r = String(80, " ")
    Mid(r, 4, 2) = "L1"
    Mid(r, 7, Len(descText)) = Left(descText, 35)
    Mid(r, 50, Len(ltypeCode)) = Left(ltypeCode, 6)
    BuildLRow = r
End Function

Function BuildHeaderRow(lineLetter)
    Dim r : r = String(80, " ")
    Mid(r, 1, 1) = lineLetter
    BuildHeaderRow = r
End Function

Function BuildPage(row9, row10, row22marker)
    Dim buf : buf = String(24 * 80, " ")
    Dim r9pos : r9pos = 8 * 80 + 1
    Dim r10pos : r10pos = 9 * 80 + 1
    Dim r22pos : r22pos = 21 * 80 + 1
    Mid(buf, r9pos, 80) = PadTo80(row9)
    Mid(buf, r10pos, 80) = PadTo80(row10)
    Mid(buf, r22pos, 80) = PadTo80(row22marker)
    BuildPage = buf
End Function

'--------------------------------------------------------------------
' Local copy of HasBlockedLTypeOnAnyPage (mirrors PostFinalCharges.vbs)
'--------------------------------------------------------------------
Function HasBlockedLTypeOnAnyPage()
    Dim row, buf, pageIndicator
    Dim matchedLType, doneScanning
    Dim pagesAdvanced, p
    Dim preSig, postSig, preSig2, postSig2, preMarker, postMarker
    Dim maxPageAdvances, i, lTypeCode

    HasBlockedLTypeOnAnyPage = ""
    pagesAdvanced = 0
    doneScanning = False
    maxPageAdvances = 50

    Do While Not doneScanning
        matchedLType = ""

        On Error Resume Next
        For row = 9 To 22
            buf = ""
            g_bzhao.ReadScreen buf, 80, row, 1
            If Err.Number <> 0 Then
                Err.Clear
            ElseIf Len(buf) >= 55 And Mid(buf, 4, 1) = "L" And IsNumeric(Mid(buf, 5, 1)) Then
                lTypeCode = UCase(Trim(Mid(buf, 50, 6)))
                If Len(lTypeCode) > 0 And IsArray(g_arrSkipLaborLTypes) Then
                    For i = 0 To UBound(g_arrSkipLaborLTypes)
                        If Len(g_arrSkipLaborLTypes(i)) > 0 And lTypeCode = g_arrSkipLaborLTypes(i) Then
                            matchedLType = lTypeCode
                            Exit For
                        End If
                    Next
                End If
                If Len(matchedLType) > 0 Then Exit For
            End If
        Next
        On Error GoTo 0

        If Len(matchedLType) > 0 Then
            HasBlockedLTypeOnAnyPage = matchedLType
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
                preSig = "" : preSig2 = ""
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

                postMarker = "" : postSig = "" : postSig2 = ""
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
                    If pagesAdvanced >= maxPageAdvances Then doneScanning = True
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

'--------------------------------------------------------------------
' Assertions
'--------------------------------------------------------------------
Sub AssertEqual(label, expected, actual)
    If expected = actual Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (expected: """ & expected & """ got: """ & actual & """)"
    End If
End Sub

Sub AssertEmpty(label, actual)
    AssertEqual label, "", actual
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------
WScript.Echo "Blocked LTYPE Gate Tests"
WScript.Echo "========================"

Dim fake, result

' --- Test 1: WCH L-row is detected and blocks ---
g_arrSkipLaborLTypes = Array("WCH", "WV", "WF")
Set fake = New FakeBzhao
fake.SetPages Array( _
    BuildPage(BuildHeaderRow("A"), BuildLRow("WCH", "VEND TO DEALER"), "(END OF DISPLAY)") _
)
Set g_bzhao = fake
result = HasBlockedLTypeOnAnyPage()
AssertEqual "WCH L-row blocks RO", "WCH", result

' --- Test 2: WV L-row is detected ---
Set fake = New FakeBzhao
fake.SetPages Array( _
    BuildPage(BuildHeaderRow("A"), BuildLRow("WV", "SOME LABOR"), "(END OF DISPLAY)") _
)
Set g_bzhao = fake
result = HasBlockedLTypeOnAnyPage()
AssertEqual "WV L-row blocks RO", "WV", result

' --- Test 3: WF L-row is detected ---
Set fake = New FakeBzhao
fake.SetPages Array( _
    BuildPage(BuildHeaderRow("A"), BuildLRow("WF", "SOME LABOR"), "(END OF DISPLAY)") _
)
Set g_bzhao = fake
result = HasBlockedLTypeOnAnyPage()
AssertEqual "WF L-row blocks RO", "WF", result

' --- Test 4: LTYPE I is not blocked ---
Set fake = New FakeBzhao
fake.SetPages Array( _
    BuildPage(BuildHeaderRow("A"), BuildLRow("I", "VEND TO DEALER"), "(END OF DISPLAY)") _
)
Set g_bzhao = fake
result = HasBlockedLTypeOnAnyPage()
AssertEmpty "LTYPE I L-row is not blocked", result

' --- Test 5: Empty SkipLaborLTypes disables gate ---
g_arrSkipLaborLTypes = Array()
Set fake = New FakeBzhao
fake.SetPages Array( _
    BuildPage(BuildHeaderRow("A"), BuildLRow("WCH", "SOME LABOR"), "(END OF DISPLAY)") _
)
Set g_bzhao = fake
result = HasBlockedLTypeOnAnyPage()
AssertEmpty "Empty SkipLaborLTypes disables gate (WCH not blocked)", result

' --- Test 6: Empty LTYPE on L-row is not matched ---
g_arrSkipLaborLTypes = Array("WCH", "WV", "WF")
Set fake = New FakeBzhao
fake.SetPages Array( _
    BuildPage(BuildHeaderRow("A"), BuildLRow("", "SOME LABOR"), "(END OF DISPLAY)") _
)
Set g_bzhao = fake
result = HasBlockedLTypeOnAnyPage()
AssertEmpty "Blank LTYPE on L-row is not matched", result

' --- Test 7: Blocked LTYPE on page 2 is detected ---
g_arrSkipLaborLTypes = Array("WCH", "WV", "WF")
Dim page1, page2
page1 = BuildPage(BuildHeaderRow("A"), BuildLRow("I", "INTERNAL LABOR"), "(MORE ON NEXT SCREEN)")
page2 = BuildPage(BuildHeaderRow("B"), BuildLRow("WCH", "VEND TO DEALER"), "(END OF DISPLAY)")
Set fake = New FakeBzhao
fake.SetPages Array(page1, page2)
Set g_bzhao = fake
result = HasBlockedLTypeOnAnyPage()
AssertEqual "WCH on page 2 is detected", "WCH", result

' --- Test 8: Returns to page 1 after paginating ---
g_arrSkipLaborLTypes = Array("WCH", "WV", "WF")
Set fake = New FakeBzhao
fake.SetPages Array(page1, page2)
Set g_bzhao = fake
result = HasBlockedLTypeOnAnyPage()
AssertEqual "Returns to page 1 after blocked LTYPE found on page 2", 0, fake.CurrentPage

' --- Test 9: All-clean multi-page RO returns empty ---
g_arrSkipLaborLTypes = Array("WCH", "WV", "WF")
Dim cleanPage1, cleanPage2
cleanPage1 = BuildPage(BuildHeaderRow("A"), BuildLRow("I", "INTERNAL LABOR"), "(MORE ON NEXT SCREEN)")
cleanPage2 = BuildPage(BuildHeaderRow("B"), BuildLRow("I", "MORE INTERNAL"), "(END OF DISPLAY)")
Set fake = New FakeBzhao
fake.SetPages Array(cleanPage1, cleanPage2)
Set g_bzhao = fake
result = HasBlockedLTypeOnAnyPage()
AssertEmpty "All-I multi-page RO is not blocked", result

' --- Test 10: Returns to page 1 after scanning clean multi-page RO ---
Set fake = New FakeBzhao
fake.SetPages Array(cleanPage1, cleanPage2)
Set g_bzhao = fake
result = HasBlockedLTypeOnAnyPage()
AssertEqual "Returns to page 1 after clean multi-page scan", 0, fake.CurrentPage

'--------------------------------------------------------------------
' Summary
'--------------------------------------------------------------------
WScript.Echo ""
WScript.Echo "Results: " & g_Pass & " passed, " & g_Fail & " failed."
If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All blocked LTYPE gate tests passed."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & g_Fail & " test(s) failed."
    WScript.Quit 1
End If
