'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestWarrantyReviewFlow
' **DATE CREATED:** 2026-04-16
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Behavioral tests for IsWarrantyLine(), DetectWarrantyDialog(),
' HandleWarrantyClaimsDialog() (dispatcher), HandleFcaClaimsDialog(),
' and HandleVwWarrantyDialog() (stub).
'
' Tests:
'   1-6)  IsWarrantyLine: True for WCH/WV/WF, False for I/blank/unknown
'   7)    FCA COMMAND: dialog — dispatcher routes to FCA handler; correct keys sent
'   8)    FCA LABOR OP: dialog — dispatcher routes to FCA handler; blank Enter only
'   9)    FCA LABOR OP: + CAUSE L1: sub-prompt — correct 3-key sequence
'   10)   WV dialog — dispatcher routes to WV stub; no keys sent
'   11)   Unknown dialog type — dispatcher logs warning; no keys sent
'   12)   No dialog detected within timeout — dispatcher logs warning; no keys sent
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_Pass, g_Fail
g_Pass = 0
g_Fail = 0

Dim g_bzhao
Dim g_arrWarrantyLTypes
Dim g_WarrantyCauseText
Dim g_WarrantyDialogStepDelayMs
Dim g_WarrantyDialogSignatureTexts()
Dim g_WarrantyDialogSignatureTypes()

' Tracking variables for stub/dispatcher verification
Dim g_VwDialogHandlerCalled
Dim g_LastWarnMessage

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
' Fake BlueZone for dialog tests
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

' FCA single-op dialog: CLAIM TYPE: at row 21 (outer detection), COMMAND: at row 22 (inner branch)
Function BuildFcaCommandDialogPage()
    Dim buf : buf = String(24 * 80, " ")
    Mid(buf, 20 * 80 + 1, 11) = "CLAIM TYPE:"
    Mid(buf, 21 * 80 + 1, 8)  = "COMMAND:"
    BuildFcaCommandDialogPage = buf
End Function

' FCA multi-op dialog: LABOR OP: at row 22 (serves as both outer and inner detection)
Function BuildFcaLaborOpDialogPage()
    Dim buf : buf = String(24 * 80, " ")
    Mid(buf, 21 * 80 + 1, 9) = "LABOR OP:"
    BuildFcaLaborOpDialogPage = buf
End Function

' WV dialog: FAILURE CODE: at row 22 (outer detection → WV type)
Function BuildVwDialogPage()
    Dim buf : buf = String(24 * 80, " ")
    Mid(buf, 21 * 80 + 1, 13) = "FAILURE CODE:"
    BuildVwDialogPage = buf
End Function

' Dialog with a custom signature text at row 22
Function BuildCustomDialogPage(signatureText)
    Dim buf : buf = String(24 * 80, " ")
    Mid(buf, 21 * 80 + 1, Len(signatureText)) = Left(signatureText, 80)
    BuildCustomDialogPage = buf
End Function

'--------------------------------------------------------------------
' Stubs required by the dialog handlers
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
    g_LastWarnMessage = msg
End Sub

Function GetScreenSnapshot(numLines)
    GetScreenSnapshot = "[test-screen-snapshot]"
End Function

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
' Local copy of DetectWarrantyDialog (mirrors PostFinalCharges.vbs)
'--------------------------------------------------------------------
Function DetectWarrantyDialog()
    DetectWarrantyDialog = ""
    If Not IsArray(g_WarrantyDialogSignatureTexts) Then Exit Function
    If UBound(g_WarrantyDialogSignatureTexts) < 0 Then Exit Function

    Dim buf, row, poll, si
    For poll = 1 To 20
        For row = 1 To 24
            buf = ""
            On Error Resume Next
            g_bzhao.ReadScreen buf, 80, row, 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            For si = 0 To UBound(g_WarrantyDialogSignatureTexts)
                If InStr(1, buf, g_WarrantyDialogSignatureTexts(si), vbTextCompare) > 0 Then
                    DetectWarrantyDialog = g_WarrantyDialogSignatureTypes(si)
                    Exit Function
                End If
            Next
        Next
        Call WaitMs(500)
    Next
End Function

'--------------------------------------------------------------------
' Local copy of HandleWarrantyClaimsDialog — dispatcher
' (mirrors PostFinalCharges.vbs)
'--------------------------------------------------------------------
Sub HandleWarrantyClaimsDialog()
    Dim dialogType
    dialogType = DetectWarrantyDialog()

    If dialogType = "" Then
        Call LogWarn("Warranty claims dialog: no dialog detected within timeout", "HandleWarrantyClaimsDialog")
    ElseIf dialogType = "FCA" Then
        Call HandleFcaClaimsDialog()
    ElseIf dialogType = "WV" Then
        Call HandleVwWarrantyDialog()
    Else
        Call LogWarn("Warranty claims dialog: unhandled dialog type [" & dialogType & "] - no handler implemented", "HandleWarrantyClaimsDialog")
        Call GetScreenSnapshot(24)
    End If
End Sub

'--------------------------------------------------------------------
' Local copy of HandleFcaClaimsDialog (mirrors PostFinalCharges.vbs)
'--------------------------------------------------------------------
Sub HandleFcaClaimsDialog()
    Dim buf, row, detected, i
    detected = ""
    For i = 1 To 20
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
        Call WaitMs(500)
    Next

    If detected = "LABOROP" Then
        Call WaitForPrompt("LABOR OP:", "", True, 5000, "")
        Call WaitMs(g_WarrantyDialogStepDelayMs)

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
            Call WaitMs(g_WarrantyDialogStepDelayMs)
            Call FastText(g_WarrantyCauseText)
            Call WaitMs(g_WarrantyDialogStepDelayMs \ 2)
            Call FastKey("<NumpadEnter>")
            Call WaitMs(g_WarrantyDialogStepDelayMs \ 2)
        Next

    ElseIf detected = "COMMAND" Then
        Call WaitMs(g_WarrantyDialogStepDelayMs)
        Call FastText(".")
        Call WaitMs(g_WarrantyDialogStepDelayMs \ 2)
        Call FastKey("<NumpadEnter>")
        Call WaitMs(g_WarrantyDialogStepDelayMs)
        Call FastText("E")
        Call WaitMs(g_WarrantyDialogStepDelayMs \ 2)
        Call FastKey("<NumpadEnter>")
        Call WaitMs(g_WarrantyDialogStepDelayMs \ 2)
    Else
        Call LogWarn("FCA claims dialog: no internal prompt detected within timeout", "HandleFcaClaimsDialog")
    End If
End Sub

'--------------------------------------------------------------------
' Local copy of HandleVwWarrantyDialog (mirrors PostFinalCharges.vbs)
'--------------------------------------------------------------------
Sub HandleVwWarrantyDialog()
    g_VwDialogHandlerCalled = True

    Dim buf, row, fi, commandFound
    commandFound = False
    For fi = 1 To 15
        For row = 20 To 24
            buf = ""
            On Error Resume Next
            g_bzhao.ReadScreen buf, 80, row, 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            If InStr(1, buf, "WARRANTY COMMAND:", vbTextCompare) > 0 Then
                commandFound = True
                Exit For
            End If
        Next
        If commandFound Then Exit For
        Call WaitMs(g_WarrantyDialogStepDelayMs)
        Call FastKey("<NumpadEnter>")
        Call WaitMs(g_WarrantyDialogStepDelayMs \ 2)
    Next

    If commandFound Then
        Call WaitMs(g_WarrantyDialogStepDelayMs)
        Call FastText("E")
        Call WaitMs(g_WarrantyDialogStepDelayMs \ 2)
        Call FastKey("<NumpadEnter>")
        Call WaitMs(g_WarrantyDialogStepDelayMs \ 2)
    Else
        Call LogWarn("VW warranty dialog: WARRANTY COMMAND: prompt not found within field limit", "HandleVwWarrantyDialog")
    End If
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

Sub AssertContainsStr(label, needle, haystack)
    If InStr(1, haystack, needle, vbTextCompare) > 0 Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (expected to find """ & needle & """ in """ & haystack & """)"
    End If
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------
WScript.Echo "Warranty Review Flow Tests"
WScript.Echo "=========================="

g_arrWarrantyLTypes = Array("WCH", "WV", "WF")
g_WarrantyCauseText = "Device failure"
g_WarrantyDialogStepDelayMs = 0
g_VwDialogHandlerCalled = False
g_LastWarnMessage = ""

' Default signature arrays (mirrors config.ini defaults)
g_WarrantyDialogSignatureTexts = Array("LABOR OP:", "CLAIM TYPE:", "FAILURE CODE:", "MODIFY WARRANTY INFORMATION")
g_WarrantyDialogSignatureTypes = Array("FCA", "FCA", "WV", "WV")

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

' --- Test 7: FCA COMMAND: dialog (single L-op) ---
' Page has CLAIM TYPE: (outer detection → FCA) and COMMAND: (inner FCA branch → COMMAND)
' Expected: . + Enter + E + Enter (4 keys)
Set fake = New FakeBzhaoDialog
fake.SetPage BuildFcaCommandDialogPage()
Set g_bzhao = fake
HandleWarrantyClaimsDialog()
Dim keys7 : keys7 = fake.GetKeys()
AssertEqual "COMMAND: dialog - routed through dispatcher to FCA handler", 4, fake.KeyCount
AssertEqual "COMMAND: dialog - first key is period (.)", ".", keys7(0)
AssertEqual "COMMAND: dialog - second key is NumpadEnter (skip fields)", "<NumpadEnter>", keys7(1)
AssertEqual "COMMAND: dialog - third key is E (exit)", "E", keys7(2)
AssertEqual "COMMAND: dialog - fourth key is NumpadEnter (confirm exit)", "<NumpadEnter>", keys7(3)

' --- Test 8: FCA LABOR OP: dialog (multiple L-ops, no CAUSE L follows) ---
' LABOR OP: serves as both outer detection (FCA) and inner FCA branch (LABOROP)
' Expected: blank Enter only (1 key)
Set fake = New FakeBzhaoDialog
fake.SetPage BuildFcaLaborOpDialogPage()
Set g_bzhao = fake
HandleWarrantyClaimsDialog()
Dim keys8 : keys8 = fake.GetKeys()
AssertEqual "LABOR OP: dialog - routed through dispatcher to FCA handler", 1, fake.KeyCount
AssertEqual "LABOR OP: dialog - first key is NumpadEnter", "<NumpadEnter>", keys8(0)

' --- Test 9: FCA LABOR OP: dialog followed by CAUSE L1: sub-prompt ---
' Page 0: LABOR OP: (outer FCA detection + inner LABOROP branch; Enter advances to page 1)
' Page 1: CAUSE L1: (CAUSE L poll finds it; FastText advances to page 2)
' Page 2: blank     (FastKey Enter advances to page 3; next CAUSE L poll exits loop)
' Page 3: blank     (stable blank page)
Dim fakeSeq : Set fakeSeq = New FakeBzhaoDialogSequenced
fakeSeq.AddPage BuildFcaLaborOpDialogPage()
fakeSeq.AddPage BuildCustomDialogPage("      CAUSE L1:")
fakeSeq.AddPage String(24 * 80, " ")
fakeSeq.AddPage String(24 * 80, " ")
Set g_bzhao = fakeSeq
HandleWarrantyClaimsDialog()
Dim keys9 : keys9 = fakeSeq.GetKeys()
AssertEqual "CAUSE L1: - first key is NumpadEnter (LABOR OP: response)", "<NumpadEnter>", keys9(0)
AssertEqual "CAUSE L1: - second key is cause text", "Device failure", keys9(1)
AssertEqual "CAUSE L1: - third key is NumpadEnter (CAUSE L1: response)", "<NumpadEnter>", keys9(2)
AssertEqual "CAUSE L1: - exactly three keys sent", 3, fakeSeq.KeyCount

' --- Test 10: WV dialog — 7 blank Enters through fields, then E at WARRANTY COMMAND: ---
' Page sequence using FakeBzhaoDialogSequenced:
'   Pages 0-6: blank fields (no WARRANTY COMMAND: yet) — each Enter advances page
'   Page 7:    WARRANTY COMMAND: visible — handler sends E + Enter
'   Page 8:    blank (stable after exit)
g_VwDialogHandlerCalled = False
Dim fakeVw : Set fakeVw = New FakeBzhaoDialogSequenced
Dim vwField
For vwField = 1 To 7
    fakeVw.AddPage BuildVwDialogPage()   ' FAILURE CODE: visible but no WARRANTY COMMAND:
Next
' Page 7: WARRANTY COMMAND: prompt appears after 7th field
Dim vwCommandPage : vwCommandPage = String(24 * 80, " ")
Mid(vwCommandPage, 21 * 80 + 1, 17) = "WARRANTY COMMAND:"
fakeVw.AddPage vwCommandPage
fakeVw.AddPage String(24 * 80, " ")
Set g_bzhao = fakeVw
HandleWarrantyClaimsDialog()
AssertTrue  "WV dialog - VW handler was called", g_VwDialogHandlerCalled
' 7 Enters (one per field) + E + Enter = 9 keys total
AssertEqual "WV dialog - total keys sent", 9, fakeVw.KeyCount
Dim keysVw : keysVw = fakeVw.GetKeys()
AssertEqual "WV dialog - last-but-one key is E", "E", keysVw(7)
AssertEqual "WV dialog - last key is NumpadEnter (confirm exit)", "<NumpadEnter>", keysVw(8)

' --- Test 11: Unknown dialog type — dispatcher logs warning, no keys sent ---
' Temporarily add an unknown-type signature, use a page that matches it
Dim savedSigTexts, savedSigTypes
savedSigTexts = g_WarrantyDialogSignatureTexts
savedSigTypes  = g_WarrantyDialogSignatureTypes
g_WarrantyDialogSignatureTexts = Array("LABOR OP:", "CLAIM TYPE:", "FAILURE CODE:", "MODIFY WARRANTY INFORMATION", "FUTURE MAKER DIALOG:")
g_WarrantyDialogSignatureTypes = Array("FCA", "FCA", "WV", "WV", "FUTUREMAKER")
g_LastWarnMessage = ""
Set fake = New FakeBzhaoDialog
fake.SetPage BuildCustomDialogPage("FUTURE MAKER DIALOG:")
Set g_bzhao = fake
HandleWarrantyClaimsDialog()
AssertEqual "Unknown type - no keys sent", 0, fake.KeyCount
AssertContainsStr "Unknown type - warning contains type key", "FUTUREMAKER", g_LastWarnMessage
g_WarrantyDialogSignatureTexts = savedSigTexts
g_WarrantyDialogSignatureTypes = savedSigTypes

' --- Test 12: No dialog detected — dispatcher logs warning, no keys sent ---
g_LastWarnMessage = ""
Set fake = New FakeBzhaoDialog
fake.SetPage String(24 * 80, " ")
Set g_bzhao = fake
HandleWarrantyClaimsDialog()
AssertEqual "No dialog - no keys sent", 0, fake.KeyCount
AssertContainsStr "No dialog - warning mentions timeout", "no dialog detected", g_LastWarnMessage

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
