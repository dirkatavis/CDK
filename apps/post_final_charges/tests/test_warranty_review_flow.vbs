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
'   13)   Ford dialog — dispatcher routes to Ford handler; handler called flag set
'   14)   Ford field sequence, VLS blank — types state code chars then Enter
'   15)   Ford field sequence, VLS pre-filled — skips state chars, Enter only at VLS
'   16)   Ford CAUSE L1: loop — sends FordWarrantyCauseText + Enter for each CAUSE L prompt
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_Pass, g_Fail
g_Pass = 0
g_Fail = 0

Dim g_bzhao
Dim g_arrWarrantyLTypes
Dim g_WarrantyCauseText
Dim g_WarrantyDialogStepDelayMs
Dim g_WarrantyDialogSignatureTexts
Dim g_WarrantyDialogSignatureTypes

' Tracking variables for stub/dispatcher verification
Dim g_FordWarrantyCauseText
Dim g_FordWarrantyLicenseState
Dim g_VwDialogHandlerCalled
Dim g_FordDialogHandlerCalled
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
' Note: VBScript Mid-statement assignment (Mid(x,n,l)=v) fails in this
' environment. All builders use string concatenation only.
'--------------------------------------------------------------------
Function PadTo80(s)
    PadTo80 = Left(s & String(80, " "), 80)
End Function

' Cols: 1-3=spaces, 4-5="L1", 6=space, 7-41=desc(35), 42-49=spaces, 50-55=ltype(6), 56-80=spaces
Function BuildLRow(ltypeCode, descText)
    Dim desc35 : desc35 = Left(descText  & String(35, " "), 35)
    Dim ltype6 : ltype6 = Left(ltypeCode & String(6,  " "), 6)
    BuildLRow = String(3, " ") & "L1" & " " & desc35 & String(8, " ") & ltype6 & String(25, " ")
End Function

' Col 1 = lineLetter, rest spaces
Function BuildHeaderRow(lineLetter)
    BuildHeaderRow = Left(lineLetter, 1) & String(79, " ")
End Function

' Rows 1-8 blank, row 9 = row9, row 10 = row10, rows 11-24 blank
Function BuildSinglePage(row9, row10)
    BuildSinglePage = String(8 * 80, " ") & PadTo80(row9) & PadTo80(row10) & String(14 * 80, " ")
End Function

' FCA single-op dialog: CLAIM TYPE: at row 21 (outer detection), COMMAND: at row 22 (inner branch)
Function BuildFcaCommandDialogPage()
    BuildFcaCommandDialogPage = String(20 * 80, " ") & PadTo80("CLAIM TYPE:") & PadTo80("COMMAND:") & String(2 * 80, " ")
End Function

' FCA multi-op dialog: LABOR OP: at row 22 (serves as both outer and inner detection)
Function BuildFcaLaborOpDialogPage()
    BuildFcaLaborOpDialogPage = String(21 * 80, " ") & PadTo80("LABOR OP:") & String(2 * 80, " ")
End Function

' WV dialog: FAILURE CODE: at row 22 (outer detection → WV type)
Function BuildVwDialogPage()
    BuildVwDialogPage = String(21 * 80, " ") & PadTo80("FAILURE CODE:") & String(2 * 80, " ")
End Function

' Dialog with a custom signature text at row 22
Function BuildCustomDialogPage(signatureText)
    BuildCustomDialogPage = String(21 * 80, " ") & PadTo80(signatureText) & String(2 * 80, " ")
End Function

' Ford warranty dialog page:
'   Row 20: "MODIFY FORD REPAIR TYPE INFORMATION" (outer detection signature)
'   Row 21: "VEHICLE LICENSE STATE:" followed by vlsFieldValue (5 chars padded)
' "VEHICLE LICENSE STATE:" is 22 chars; field value read from col 23 (vlsLabelPos+22).
' vlsFieldValue = "" for blank VLS (handler types state code); non-blank = Enter-only path.
Function BuildFordDialogPage(vlsFieldValue)
    Dim vlsText : vlsText = "VEHICLE LICENSE STATE:" & Left(vlsFieldValue & String(5, " "), 5)
    BuildFordDialogPage = String(19 * 80, " ") & PadTo80("MODIFY FORD REPAIR TYPE INFORMATION") & PadTo80(vlsText) & String(3 * 80, " ")
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
    ElseIf dialogType = "FORD" Then
        Call HandleFordWarrantyDialog()
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
' Local copy of HandleFordWarrantyDialog (mirrors PostFinalCharges.vbs)
'--------------------------------------------------------------------
Sub HandleFordWarrantyDialog()
    g_FordDialogHandlerCalled = True

    ' Step 1: P & A CODE — accept auto-populated value
    Call WaitForPrompt("P & A CODE:", "", True, 5000, "")
    Call WaitMs(g_WarrantyDialogStepDelayMs)

    ' Step 2: FORD / L-M MAKE (Y/N)? — accept auto-populated Y
    Call WaitForPrompt("FORD / L-M MAKE", "", True, 5000, "")
    Call WaitMs(g_WarrantyDialogStepDelayMs)

    ' Step 3: FRANCHISE MODEL (Y/N)? — accept auto-populated Y
    Call WaitForPrompt("FRANCHISE MODEL", "", True, 5000, "")
    Call WaitMs(g_WarrantyDialogStepDelayMs)

    ' Step 4: VEHICLE LICENSE STATE — type state code if blank; Enter only if already filled
    Call WaitForPrompt("VEHICLE LICENSE STATE:", "", False, 5000, "")
    Dim vlsRow, vlsBuf, vlsLabelPos, vlsFieldText
    vlsLabelPos = 0
    vlsFieldText = ""
    For vlsRow = 1 To 24
        vlsBuf = ""
        On Error Resume Next
        g_bzhao.ReadScreen vlsBuf, 80, vlsRow, 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        vlsLabelPos = InStr(1, vlsBuf, "VEHICLE LICENSE STATE:", vbTextCompare)
        If vlsLabelPos > 0 Then
            vlsFieldText = Trim(Mid(vlsBuf, vlsLabelPos + 22, 5))
            Exit For
        End If
    Next
    If Len(vlsFieldText) = 0 Then
        Dim vlsCharIdx
        For vlsCharIdx = 1 To Len(g_FordWarrantyLicenseState)
            Call FastText(Mid(g_FordWarrantyLicenseState, vlsCharIdx, 1))
            Call WaitMs(100)
        Next
    End If
    Call FastKey("<NumpadEnter>")
    Call WaitMs(g_WarrantyDialogStepDelayMs)

    ' Step 5: REPAIR TYPE — send 1 (Warranty/ESP) + Enter
    Call WaitForPrompt("REPAIR TYPE:", "1", True, 5000, "")
    Call WaitMs(g_WarrantyDialogStepDelayMs)

    ' Step 6: COMMAND: (inside dialog) — send . + Enter to skip remaining fields
    Call WaitForPrompt("COMMAND:", ".", True, 5000, "")
    Call WaitMs(g_WarrantyDialogStepDelayMs)

    ' CAUSE L<n>: prompts after dialog closes — each receives FordWarrantyCauseText
    Dim causeRow, causeBuf, causeFound, ci, causePoll
    For ci = 1 To 10
        causeFound = False
        For causePoll = 1 To 6
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
            If causeFound Then Exit For
            Call WaitMs(500)
        Next
        If Not causeFound Then Exit For
        Call FastText(g_FordWarrantyCauseText)
        Call WaitMs(g_WarrantyDialogStepDelayMs \ 2)
        Call FastKey("<NumpadEnter>")
        Call WaitMs(g_WarrantyDialogStepDelayMs \ 2)
    Next
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
g_FordDialogHandlerCalled = False
g_FordWarrantyCauseText = "Defective Part"
g_FordWarrantyLicenseState = "GA"
g_LastWarnMessage = ""

' Default signature arrays (mirrors config.ini defaults)
g_WarrantyDialogSignatureTexts = Array("LABOR OP:", "CLAIM TYPE:", "FAILURE CODE:", "MODIFY WARRANTY INFORMATION", "MODIFY FORD REPAIR TYPE INFORMATION")
g_WarrantyDialogSignatureTypes = Array("FCA", "FCA", "WV", "WV", "FORD")

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
Dim vwCommandPage : vwCommandPage = String(21 * 80, " ") & PadTo80("WARRANTY COMMAND:") & String(2 * 80, " ")
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

' --- Test 13: Ford dialog — dispatcher routes to Ford handler ---
g_FordDialogHandlerCalled = False
Set fake = New FakeBzhaoDialog
fake.SetPage BuildFordDialogPage("")
Set g_bzhao = fake
HandleWarrantyClaimsDialog()
AssertTrue "Ford dialog - dispatcher called Ford handler", g_FordDialogHandlerCalled

' --- Test 14: Ford field sequence — VLS blank, types state code one char at a time ---
' Expected key sequence (10 keys, no CAUSE L on static page):
'   0: <NumpadEnter>   P & A CODE
'   1: <NumpadEnter>   FORD / L-M MAKE
'   2: <NumpadEnter>   FRANCHISE MODEL
'   3: G               VLS char 1  (g_FordWarrantyLicenseState = "GA")
'   4: A               VLS char 2
'   5: <NumpadEnter>   VLS confirm
'   6: 1               REPAIR TYPE responseKey
'   7: <NumpadEnter>   REPAIR TYPE Enter
'   8: .               COMMAND responseKey
'   9: <NumpadEnter>   COMMAND Enter
g_FordDialogHandlerCalled = False
Set fake = New FakeBzhaoDialog
fake.SetPage BuildFordDialogPage("")
Set g_bzhao = fake
HandleWarrantyClaimsDialog()
Dim keys14 : keys14 = fake.GetKeys()
AssertEqual "Ford VLS blank - total keys sent", 10, fake.KeyCount
AssertEqual "Ford VLS blank - key 0 is NumpadEnter (P & A CODE)", "<NumpadEnter>", keys14(0)
AssertEqual "Ford VLS blank - key 1 is NumpadEnter (FORD/L-M MAKE)", "<NumpadEnter>", keys14(1)
AssertEqual "Ford VLS blank - key 2 is NumpadEnter (FRANCHISE MODEL)", "<NumpadEnter>", keys14(2)
AssertEqual "Ford VLS blank - key 3 is G (VLS state char 1)", "G", keys14(3)
AssertEqual "Ford VLS blank - key 4 is A (VLS state char 2)", "A", keys14(4)
AssertEqual "Ford VLS blank - key 5 is NumpadEnter (VLS confirm)", "<NumpadEnter>", keys14(5)
AssertEqual "Ford VLS blank - key 6 is 1 (REPAIR TYPE)", "1", keys14(6)
AssertEqual "Ford VLS blank - key 7 is NumpadEnter (REPAIR TYPE Enter)", "<NumpadEnter>", keys14(7)
AssertEqual "Ford VLS blank - key 8 is period (COMMAND exit)", ".", keys14(8)
AssertEqual "Ford VLS blank - key 9 is NumpadEnter (COMMAND Enter)", "<NumpadEnter>", keys14(9)

' --- Test 15: Ford field sequence — VLS pre-filled, skips state chars ---
' Expected key sequence (8 keys): 3 field Enters + 1 VLS Enter (no state chars) + 1+Enter REPAIR TYPE + .+Enter COMMAND
'   0: <NumpadEnter>   P & A CODE
'   1: <NumpadEnter>   FORD / L-M MAKE
'   2: <NumpadEnter>   FRANCHISE MODEL
'   3: <NumpadEnter>   VLS confirm (field already has value — no state chars typed)
'   4: 1               REPAIR TYPE
'   5: <NumpadEnter>   REPAIR TYPE Enter
'   6: .               COMMAND
'   7: <NumpadEnter>   COMMAND Enter
g_FordDialogHandlerCalled = False
Set fake = New FakeBzhaoDialog
fake.SetPage BuildFordDialogPage("GA")
Set g_bzhao = fake
HandleWarrantyClaimsDialog()
Dim keys15 : keys15 = fake.GetKeys()
AssertEqual "Ford VLS prefilled - total keys sent", 8, fake.KeyCount
AssertEqual "Ford VLS prefilled - key 0 is NumpadEnter (P & A CODE)", "<NumpadEnter>", keys15(0)
AssertEqual "Ford VLS prefilled - key 3 is NumpadEnter (VLS confirm, no state chars)", "<NumpadEnter>", keys15(3)
AssertEqual "Ford VLS prefilled - key 4 is 1 (REPAIR TYPE, not a state char)", "1", keys15(4)
AssertEqual "Ford VLS prefilled - key 6 is period (COMMAND)", ".", keys15(6)

' --- Test 16: Ford CAUSE L1: loop — sends FordWarrantyCauseText + Enter per CAUSE L prompt ---
' FakeBzhaoDialogSequenced page sequence:
'   Pages 0-9:  Ford dialog page (detection signature + VLS blank), one per SendKey during field nav
'   Page 10:    CAUSE L1: visible in rows 20-24 (inner poll finds it)
'   Pages 11-12: blank (after FastText + Enter advance past CAUSE L; second poll exits loop)
' Total expected keys: 10 (field nav) + 1 (FastText cause text) + 1 (cause Enter) = 12
Dim fakeF16 : Set fakeF16 = New FakeBzhaoDialogSequenced
Dim fordBlankPage : fordBlankPage = BuildFordDialogPage("")
Dim fi16
For fi16 = 1 To 10
    fakeF16.AddPage fordBlankPage
Next
Dim causePage16 : causePage16 = String(20 * 80, " ") & PadTo80("CAUSE L1:") & String(3 * 80, " ")
fakeF16.AddPage causePage16
fakeF16.AddPage String(24 * 80, " ")
fakeF16.AddPage String(24 * 80, " ")
Set g_bzhao = fakeF16
HandleWarrantyClaimsDialog()
Dim keys16 : keys16 = fakeF16.GetKeys()
AssertEqual "Ford CAUSE L - total keys sent", 12, fakeF16.KeyCount
AssertEqual "Ford CAUSE L - key 0 is NumpadEnter (P & A CODE)", "<NumpadEnter>", keys16(0)
AssertEqual "Ford CAUSE L - key 10 is FordWarrantyCauseText (Defective Part)", "Defective Part", keys16(10)
AssertEqual "Ford CAUSE L - key 11 is NumpadEnter (cause confirm)", "<NumpadEnter>", keys16(11)

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
