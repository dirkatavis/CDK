'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestWarrantyReviewFlow
' **DATE CREATED:** 2026-04-16
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Behavioral tests for IsWarrantyLine() and HandleWarrantyClaimsDialog().
' Validates:
'   1) IsWarrantyLine returns True for WCH, WV, WF (entries in WarrantyLTypes).
'   2) IsWarrantyLine returns False for LTYPE=I, blank LTYPE, and unrecognized codes.
'   3) HandleWarrantyClaimsDialog detects COMMAND: (single L-op) and sends E + Enter.
'   4) HandleWarrantyClaimsDialog detects LABOR OP: (multiple L-ops) and sends blank Enter.
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_Pass, g_Fail
g_Pass = 0
g_Fail = 0

Dim g_bzhao
Dim g_arrWarrantyLTypes

'--------------------------------------------------------------------
' Fake BlueZone for IsWarrantyLine tests
'--------------------------------------------------------------------
Class FakeBzhaoSinglePage
    Private m_page

    Public Sub SetPage(page)
        m_page = page
    End Sub

    Public Sub ReadScreen(ByRef buf, ByVal length, ByVal row, ByVal col)
        Dim startPos
        startPos = ((row - 1) * 80) + 1
        buf = Mid(m_page, startPos, length)
    End Sub
End Class

'--------------------------------------------------------------------
' Fake BlueZone for HandleWarrantyClaimsDialog tests
' Serves a stable page for detection polling; records all keys sent
'--------------------------------------------------------------------
Class FakeBzhaoDialog
    Private m_page
    Private m_keys()
    Private m_keyCount

    Private Sub Class_Initialize()
        m_keyCount = 0
        ReDim m_keys(0)
    End Sub

    Public Sub SetPage(page)
        m_page = page
    End Sub

    Public Sub ReadScreen(ByRef buf, ByVal length, ByVal row, ByVal col)
        Dim startPos
        startPos = ((row - 1) * 80) + 1
        buf = Mid(m_page, startPos, length)
    End Sub

    Public Sub SendKey(ByVal key)
        ReDim Preserve m_keys(m_keyCount)
        m_keys(m_keyCount) = key
        m_keyCount = m_keyCount + 1
    End Sub

    Public Function GetKeys()
        GetKeys = m_keys
    End Function

    Public Property Get KeyCount()
        KeyCount = m_keyCount
    End Property
End Class

'--------------------------------------------------------------------
' Stateful Fake BlueZone for CAUSE L sequencing tests
' Advances to the next page on every SendKey call, simulating the
' terminal transitioning: LABOR OP: -> CAUSE Ln: -> blank
'--------------------------------------------------------------------
Class FakeBzhaoDialogSequenced
    Private m_pages()
    Private m_pageCount
    Private m_currentPage
    Private m_keys()
    Private m_keyCount

    Private Sub Class_Initialize()
        m_pageCount = 0
        m_currentPage = 0
        m_keyCount = 0
        ReDim m_keys(0)
        ReDim m_pages(0)
    End Sub

    Public Sub AddPage(page)
        If m_pageCount > 0 Then ReDim Preserve m_pages(m_pageCount)
        m_pages(m_pageCount) = page
        m_pageCount = m_pageCount + 1
    End Sub

    Public Sub ReadScreen(ByRef buf, ByVal length, ByVal row, ByVal col)
        Dim startPos, page
        page = m_pages(m_currentPage)
        startPos = ((row - 1) * 80) + 1
        buf = Mid(page, startPos, length)
    End Sub

    Public Sub SendKey(ByVal key)
        ReDim Preserve m_keys(m_keyCount)
        m_keys(m_keyCount) = key
        m_keyCount = m_keyCount + 1
        If m_currentPage < m_pageCount - 1 Then
            m_currentPage = m_currentPage + 1
        End If
    End Sub

    Public Function GetKeys()
        GetKeys = m_keys
    End Function

    Public Property Get KeyCount()
        KeyCount = m_keyCount
    End Property
End Class

'--------------------------------------------------------------------
' Row/page builders
'--------------------------------------------------------------------
Function PadTo80(s)
    PadTo80 = Left(s & String(80, " "), 80)
End Function

Function BuildLRow(ltypeCode, descText)
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

Function BuildSinglePage(row9, row10)
    Dim buf : buf = String(24 * 80, " ")
    Mid(buf, 8 * 80 + 1, 80) = PadTo80(row9)
    Mid(buf, 9 * 80 + 1, 80) = PadTo80(row10)
    BuildSinglePage = buf
End Function

' Places prompt text at row 22 — inside the dialog box bottom row
Function BuildDialogPage(promptText)
    Dim buf : buf = String(24 * 80, " ")
    Mid(buf, 21 * 80 + 1, Len(promptText)) = Left(promptText, 80)
    BuildDialogPage = buf
End Function

'--------------------------------------------------------------------
' Stubs required by HandleWarrantyClaimsDialog
'--------------------------------------------------------------------
Sub WaitMs(ms)
    ' No-op in tests
End Sub

Sub WaitForPrompt(promptText, responseKey, sendEnter, timeoutMs, logLabel)
    If Len(responseKey) > 0 Then g_bzhao.SendKey responseKey
    If sendEnter Then g_bzhao.SendKey "<NumpadEnter>"
End Sub

Sub FastText(text)
    g_bzhao.SendKey text
End Sub

Sub FastKey(key)
    g_bzhao.SendKey key
End Sub

Sub LogInfo(msg, context)
End Sub

Sub LogWarn(msg, context)
End Sub

'--------------------------------------------------------------------
' Local copy of IsWarrantyLine (mirrors PostFinalCharges.vbs)
'--------------------------------------------------------------------
Function IsWarrantyLine(lineLetterChar)
    IsWarrantyLine = False
    Dim row, buf, inTargetLine, firstChar, lTypeCode, wi
    inTargetLine = False
    For row = 9 To 22
        buf = ""
        On Error Resume Next
        g_bzhao.ReadScreen buf, 80, row, 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Len(buf) >= 55 Then
            firstChar = Mid(buf, 1, 1)
            If firstChar >= "A" And firstChar <= "Z" Then
                If inTargetLine Then Exit For
                If firstChar = lineLetterChar Then inTargetLine = True
            End If
            If inTargetLine And Mid(buf, 4, 1) = "L" And IsNumeric(Mid(buf, 5, 1)) Then
                lTypeCode = UCase(Trim(Mid(buf, 50, 6)))
                If Len(lTypeCode) > 0 And IsArray(g_arrWarrantyLTypes) Then
                    For wi = 0 To UBound(g_arrWarrantyLTypes)
                        If Len(g_arrWarrantyLTypes(wi)) > 0 And lTypeCode = g_arrWarrantyLTypes(wi) Then
                            IsWarrantyLine = True
                            Exit Function
                        End If
                    Next
                End If
            End If
        End If
    Next
End Function

'--------------------------------------------------------------------
' Local copy of HandleWarrantyClaimsDialog (mirrors PostFinalCharges.vbs)
'--------------------------------------------------------------------
Sub HandleWarrantyClaimsDialog()
    Call LogInfo("Warranty claims dialog: detecting prompt", "HandleWarrantyClaimsDialog")

    Dim buf, row, detected, i
    detected = ""
    For i = 1 To 50
        For row = 20 To 24
            buf = ""
            On Error Resume Next
            g_bzhao.ReadScreen buf, 80, row, 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            If InStr(1, buf, "LABOR OP:", vbTextCompare) > 0 Then
                detected = "LABOROP"
                Exit For
            End If
            If InStr(1, buf, "COMMAND:", vbTextCompare) > 0 Then
                detected = "COMMAND"
                Exit For
            End If
        Next
        If Len(detected) > 0 Then Exit For
        Call WaitMs(100)
    Next

    If detected = "LABOROP" Then
        Call LogInfo("Warranty claims dialog: LABOR OP prompt (multiple L-ops) - sending blank Enter", "HandleWarrantyClaimsDialog")
        Call WaitForPrompt("LABOR OP:", "", True, 5000, "")
        Call WaitMs(2000)

        Dim causeRow, causeBuf, causeFound, ci
        For ci = 1 To 10
            causeFound = False
            For causeRow = 20 To 24
                causeBuf = ""
                On Error Resume Next
                g_bzhao.ReadScreen causeBuf, 80, causeRow, 1
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                If InStr(1, causeBuf, "CAUSE L", vbTextCompare) > 0 Then
                    causeFound = True
                    Exit For
                End If
            Next
            If Not causeFound Then Exit For
            Call LogInfo("Warranty claims dialog: CAUSE L prompt detected - sending cause text", "HandleWarrantyClaimsDialog")
            Call WaitMs(2000)
            Call FastText(g_WarrantyCauseText)
            Call WaitMs(1000)
            Call FastKey("<NumpadEnter>")
            Call WaitMs(1000)
        Next

    ElseIf detected = "COMMAND" Then
        Call LogInfo("Warranty claims dialog: COMMAND prompt (single L-op) - sending . to skip fields then E to exit", "HandleWarrantyClaimsDialog")
        Call WaitMs(2000)
        Call FastText(".")
        Call WaitMs(1000)
        Call FastKey("<NumpadEnter>")
        Call WaitMs(2000)
        Call FastText("E")
        Call WaitMs(1000)
        Call FastKey("<NumpadEnter>")
        Call WaitMs(1000)
    Else
        Call LogWarn("Warranty claims dialog: no known prompt detected within timeout", "HandleWarrantyClaimsDialog")
    End If

    Call LogInfo("Warranty claims dialog: complete", "HandleWarrantyClaimsDialog")
End Sub

'--------------------------------------------------------------------
' Assertions
'--------------------------------------------------------------------
Sub AssertEqual(label, expected, actual)
    If CStr(expected) = CStr(actual) Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (expected: """ & expected & """ got: """ & actual & """)"
    End If
End Sub

Sub AssertTrue(label, value)
    AssertEqual label, True, value
End Sub

Sub AssertFalse(label, value)
    AssertEqual label, False, value
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------
WScript.Echo "Warranty Review Flow Tests"
WScript.Echo "=========================="

g_arrWarrantyLTypes = Array("WCH", "WV", "WF")
g_WarrantyCauseText = "Device failure"

Dim fake, result

' --- Test 1: IsWarrantyLine returns True for WCH ---
Set fake = New FakeBzhaoSinglePage
fake.SetPage BuildSinglePage(BuildHeaderRow("A"), BuildLRow("WCH", "VEND TO DEALER"))
Set g_bzhao = fake
AssertTrue "IsWarrantyLine True for WCH", IsWarrantyLine("A")

' --- Test 2: IsWarrantyLine returns True for WV ---
Set fake = New FakeBzhaoSinglePage
fake.SetPage BuildSinglePage(BuildHeaderRow("A"), BuildLRow("WV", "WARRANTY VOLVO"))
Set g_bzhao = fake
AssertTrue "IsWarrantyLine True for WV", IsWarrantyLine("A")

' --- Test 3: IsWarrantyLine returns True for WF ---
Set fake = New FakeBzhaoSinglePage
fake.SetPage BuildSinglePage(BuildHeaderRow("A"), BuildLRow("WF", "WARRANTY FORD"))
Set g_bzhao = fake
AssertTrue "IsWarrantyLine True for WF", IsWarrantyLine("A")

' --- Test 4: IsWarrantyLine returns False for LTYPE=I ---
Set fake = New FakeBzhaoSinglePage
fake.SetPage BuildSinglePage(BuildHeaderRow("A"), BuildLRow("I", "INTERNAL LABOR"))
Set g_bzhao = fake
AssertFalse "IsWarrantyLine False for LTYPE=I", IsWarrantyLine("A")

' --- Test 5: IsWarrantyLine returns False for blank LTYPE ---
Set fake = New FakeBzhaoSinglePage
fake.SetPage BuildSinglePage(BuildHeaderRow("A"), BuildLRow("", "SOME LABOR"))
Set g_bzhao = fake
AssertFalse "IsWarrantyLine False for blank LTYPE", IsWarrantyLine("A")

' --- Test 6: IsWarrantyLine returns False for unrecognized code ---
Set fake = New FakeBzhaoSinglePage
fake.SetPage BuildSinglePage(BuildHeaderRow("A"), BuildLRow("ZZZ", "UNKNOWN LTYPE"))
Set g_bzhao = fake
AssertFalse "IsWarrantyLine False for unrecognized LTYPE ZZZ", IsWarrantyLine("A")

' --- Test 7: COMMAND: dialog — sends . + Enter (skip fields) then E + Enter (exit) ---
Set fake = New FakeBzhaoDialog
fake.SetPage BuildDialogPage("        COMMAND:                                                        ")
Set g_bzhao = fake
HandleWarrantyClaimsDialog()
Dim keys7 : keys7 = fake.GetKeys()
AssertEqual "COMMAND: dialog - first key is period (.)", ".", keys7(0)
AssertEqual "COMMAND: dialog - second key is NumpadEnter (skip fields)", "<NumpadEnter>", keys7(1)
AssertEqual "COMMAND: dialog - third key is E (exit)", "E", keys7(2)
AssertEqual "COMMAND: dialog - fourth key is NumpadEnter (confirm exit)", "<NumpadEnter>", keys7(3)
AssertEqual "COMMAND: dialog - exactly four keys sent", 4, fake.KeyCount

' --- Test 8: LABOR OP: dialog — sends blank Enter only (no CAUSE L follows) ---
Set fake = New FakeBzhaoDialog
fake.SetPage BuildDialogPage("      LABOR OP:                                                         ")
Set g_bzhao = fake
HandleWarrantyClaimsDialog()
Dim keys8 : keys8 = fake.GetKeys()
AssertEqual "LABOR OP: dialog - first key is NumpadEnter", "<NumpadEnter>", keys8(0)
AssertEqual "LABOR OP: dialog - only one key sent", 1, fake.KeyCount

' --- Test 9: LABOR OP: dialog followed by CAUSE L1: sub-prompt ---
' Page 0: LABOR OP: (initial detection + WaitForPrompt; Enter advances to page 1)
' Page 1: CAUSE L1: (CAUSE L poll finds it; FastText advances to page 2)
' Page 2: blank     (FastKey Enter advances to page 3; next CAUSE L poll finds nothing)
' Page 3: blank     (stable blank page for any further polling)
Dim fakeSeq : Set fakeSeq = New FakeBzhaoDialogSequenced
fakeSeq.AddPage BuildDialogPage("      LABOR OP:                                                         ")
fakeSeq.AddPage BuildDialogPage("      CAUSE L1:                                                         ")
fakeSeq.AddPage String(24 * 80, " ")
fakeSeq.AddPage String(24 * 80, " ")
Set g_bzhao = fakeSeq
HandleWarrantyClaimsDialog()
Dim keys9 : keys9 = fakeSeq.GetKeys()
AssertEqual "CAUSE L1: - first key is NumpadEnter (LABOR OP: response)", "<NumpadEnter>", keys9(0)
AssertEqual "CAUSE L1: - second key is cause text", "Device failure", keys9(1)
AssertEqual "CAUSE L1: - third key is NumpadEnter (CAUSE L1: response)", "<NumpadEnter>", keys9(2)
AssertEqual "CAUSE L1: - exactly three keys sent", 3, fakeSeq.KeyCount

'--------------------------------------------------------------------
' Summary
'--------------------------------------------------------------------
WScript.Echo ""
WScript.Echo "Results: " & g_Pass & " passed, " & g_Fail & " failed."
If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All warranty review flow tests passed."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & g_Fail & " test(s) failed."
    WScript.Quit 1
End If
