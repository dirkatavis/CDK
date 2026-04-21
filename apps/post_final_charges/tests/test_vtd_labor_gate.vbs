'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestVtdLaborGate
' **DATE CREATED:** 2026-04-21
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Behavioral unit tests for ContainsWholeWordVtd() and EvaluateVtdLaborGate():
'   1)  Fixture screen_i91_tech_875722.txt — line C ltype I + "VTD" in L-row desc
'       -> gate fails, skipReason contains "Skipped - VTD labor line:" prefix
'   2)  Synthetic — ltype I + whole-word "VTD" in L-row desc -> gate fails
'   3)  Synthetic — "VTD" in line header desc, ltype I L-row (no VTD in L-row desc)
'       -> gate fails (header desc match triggers)
'   4)  Synthetic — ltype I, no VTD anywhere -> gate passes
'   5)  Synthetic — VTD in L-row desc, ltype != I -> gate passes
'   6)  Synthetic — "VTDEV" in L-row desc (not whole word), ltype I -> gate passes
'   7)  ContainsWholeWordVtd: "VTD" alone -> True
'   8)  ContainsWholeWordVtd: "INSPECT VTD" -> True (at end, space before)
'   9)  ContainsWholeWordVtd: "VTDEV CHECK" -> False (E follows, no word boundary after)
'  10)  ContainsWholeWordVtd: "AAAVTD" -> False (A precedes, no word boundary before)
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_Pass, g_Fail, g_fso
g_Pass = 0
g_Fail = 0
Set g_fso = CreateObject("Scripting.FileSystemObject")

Dim g_bzhao

'--------------------------------------------------------------------
' Stub — LogEvent (no-op)
'--------------------------------------------------------------------
Sub LogEvent(ByVal a, ByVal b, ByVal c, ByVal d, ByVal e, ByVal f)
End Sub

'--------------------------------------------------------------------
' FakeBzhao — single-page; SendKey / Pause are no-ops
'--------------------------------------------------------------------
Class FakeBzhao
    Private m_page

    Public Sub SetPage(ByVal pageBuf)
        m_page = pageBuf
    End Sub

    Public Sub ReadScreen(ByRef outText, ByVal length, ByVal row, ByVal col)
        Dim pos
        pos = ((row - 1) * 80) + col
        If pos < 1 Then pos = 1
        outText = Mid(m_page, pos, length)
    End Sub

    Public Sub SendKey(ByVal keyText)
    End Sub

    Public Sub Pause(ByVal ms)
    End Sub
End Class

'--------------------------------------------------------------------
' Screen builder helpers (string-concatenation only, no Mid assignment)
'--------------------------------------------------------------------
Function SetColText(ByVal rowText, ByVal colNum, ByVal textValue)
    Dim base
    base = Left(rowText & String(80, " "), 80)
    SetColText = Left(Left(base, colNum - 1) & textValue & Mid(base, colNum + Len(textValue)) & String(80, " "), 80)
End Function

Function SetRow(ByVal pageBuf, ByVal rowNum, ByVal rowText)
    Dim pos
    pos = (rowNum - 1) * 80 + 1
    SetRow = Left(pageBuf, pos - 1) & Left(rowText & String(80, " "), 80) & Mid(pageBuf, pos + 80)
End Function

' Line header: col 1=letter, col 2=space, cols 4+= description (38 chars)
Function BuildHeaderRow(ByVal lineLetter, ByVal descText)
    Dim rowText : rowText = String(80, " ")
    rowText = SetColText(rowText, 1, lineLetter)
    rowText = SetColText(rowText, 4, Left(descText, 38))
    BuildHeaderRow = rowText
End Function

' L-row: col 4-5="L1", col 7-41=description (35 chars), col 50-55=ltype
Function BuildLRow(ByVal ltypeCode, ByVal lRowDesc)
    Dim rowText : rowText = String(80, " ")
    rowText = SetColText(rowText, 4, "L1")
    rowText = SetColText(rowText, 7, Left(lRowDesc, 35))
    rowText = SetColText(rowText, 50, ltypeCode)
    BuildLRow = rowText
End Function

' "(END OF DISPLAY)" marker row (placed at row 22 to stop pagination)
Function BuildEndRow()
    Dim rowText : rowText = String(80, " ")
    rowText = SetColText(rowText, 1, "(END OF DISPLAY)")
    BuildEndRow = rowText
End Function

' Single-page screen: header at row 9, L-row at row 10, "(END OF DISPLAY)" at row 22
Function BuildPage(ByVal lineLetter, ByVal headerDesc, ByVal ltypeCode, ByVal lRowDesc)
    Dim pageBuf : pageBuf = String(24 * 80, " ")
    pageBuf = SetRow(pageBuf, 9,  BuildHeaderRow(lineLetter, headerDesc))
    pageBuf = SetRow(pageBuf, 10, BuildLRow(ltypeCode, lRowDesc))
    pageBuf = SetRow(pageBuf, 22, BuildEndRow())
    BuildPage = pageBuf
End Function

'--------------------------------------------------------------------
' ParseScreenMapToBuffer — loads "DD | <80 chars>" coordinate-map files
'--------------------------------------------------------------------
Function ParseScreenMapToBuffer(ByVal filePath)
    Dim buf : buf = String(24 * 80, " ")
    If Not g_fso.FileExists(filePath) Then
        ParseScreenMapToBuffer = buf
        Exit Function
    End If
    Dim ts : Set ts = g_fso.OpenTextFile(filePath, 1)
    Do While Not ts.AtEndOfStream
        Dim lineText : lineText = ts.ReadLine
        If Len(lineText) >= 7 And Mid(lineText, 3, 3) = " | " Then
            Dim rowNum
            On Error Resume Next
            rowNum = CInt(Left(lineText, 2))
            On Error GoTo 0
            If rowNum >= 1 And rowNum <= 24 Then
                Dim rowContent : rowContent = Mid(lineText, 6)
                rowContent = Left(rowContent & String(80, " "), 80)
                Dim bufPos : bufPos = ((rowNum - 1) * 80) + 1
                buf = Left(buf, bufPos - 1) & rowContent & Mid(buf, bufPos + 80)
            End If
        End If
    Loop
    ts.Close
    ParseScreenMapToBuffer = buf
End Function

'--------------------------------------------------------------------
' Local copy of ContainsWholeWordVtd (mirrors PostFinalCharges.vbs)
'--------------------------------------------------------------------
Function ContainsWholeWordVtd(ByVal text)
    ContainsWholeWordVtd = False
    Dim pos, charBefore, charAfter, isBefore, isAfter
    pos = 1
    Do
        pos = InStr(pos, text, "VTD", vbTextCompare)
        If pos = 0 Then Exit Do
        isBefore = (pos = 1)
        If Not isBefore Then
            charBefore = Mid(text, pos - 1, 1)
            isBefore = Not ((charBefore >= "A" And charBefore <= "Z") Or _
                            (charBefore >= "a" And charBefore <= "z") Or _
                            (charBefore >= "0" And charBefore <= "9"))
        End If
        isAfter = (pos + 3 > Len(text))
        If Not isAfter Then
            charAfter = Mid(text, pos + 3, 1)
            isAfter = Not ((charAfter >= "A" And charAfter <= "Z") Or _
                           (charAfter >= "a" And charAfter <= "z") Or _
                           (charAfter >= "0" And charAfter <= "9"))
        End If
        If isBefore And isAfter Then
            ContainsWholeWordVtd = True
            Exit Function
        End If
        pos = pos + 1
    Loop
End Function

'--------------------------------------------------------------------
' Local copy of EvaluateVtdLaborGate (mirrors PostFinalCharges.vbs)
'--------------------------------------------------------------------
Function EvaluateVtdLaborGate(ByRef skipReason)
    EvaluateVtdLaborGate = True
    skipReason = ""

    Dim row, buf
    Dim doneScanning, pagesAdvanced, p
    Dim preSig, postSig, preSig2, postSig2, preMarker, postMarker
    Dim maxPageAdvances
    Dim currentLineHeaderDesc, lRowDesc, lTypeCode
    Dim pageIndicatorVtd

    doneScanning = False
    pagesAdvanced = 0
    maxPageAdvances = 50
    currentLineHeaderDesc = ""

    Do While Not doneScanning
        For row = 9 To 23
            buf = ""
            On Error Resume Next
            g_bzhao.ReadScreen buf, 80, row, 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0

            If Len(buf) >= 5 Then
                If Mid(buf, 1, 1) >= "A" And Mid(buf, 1, 1) <= "Z" And Mid(buf, 2, 1) = " " And Mid(buf, 4, 1) <> "L" Then
                    currentLineHeaderDesc = Trim(Mid(buf, 4, 38))
                End If

                If Mid(buf, 4, 1) = "L" And IsNumeric(Mid(buf, 5, 1)) Then
                    lTypeCode = UCase(Trim(Mid(buf, 50, 6)))
                    lRowDesc  = Trim(Mid(buf, 7, 35))
                    If lTypeCode = "I" Then
                        If ContainsWholeWordVtd(lRowDesc) Or ContainsWholeWordVtd(currentLineHeaderDesc) Then
                            skipReason = "Skipped - VTD labor line: ltype=[" & lTypeCode & "] lRowDesc=[" & lRowDesc & "] header=[" & currentLineHeaderDesc & "]"
                            EvaluateVtdLaborGate = False
                            doneScanning = True
                            Exit For
                        End If
                    End If
                End If
            End If
        Next

        If Not doneScanning Then
            pageIndicatorVtd = ""
            On Error Resume Next
            g_bzhao.ReadScreen pageIndicatorVtd, 80, 22, 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0

            If InStr(1, pageIndicatorVtd, "(END OF DISPLAY)", vbTextCompare) > 0 Then
                doneScanning = True
            ElseIf InStr(1, pageIndicatorVtd, "(MORE ON NEXT SCREEN)", vbTextCompare) > 0 Then
                preMarker = pageIndicatorVtd
                preSig = "" : preSig2 = ""
                On Error Resume Next
                g_bzhao.ReadScreen preSig,  80, 9,  1
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
                g_bzhao.ReadScreen postSig,    80, 9,  1
                g_bzhao.ReadScreen postSig2,   80, 10, 1
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
Sub AssertEqual(ByVal label, ByVal expected, ByVal actual)
    If CStr(expected) = CStr(actual) Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label
        WScript.Echo "       Expected: """ & expected & """"
        WScript.Echo "       Actual  : """ & actual & """"
    End If
End Sub

Sub AssertTrue(ByVal label, ByVal value)
    AssertEqual label, True, value
End Sub

Sub AssertFalse(ByVal label, ByVal value)
    AssertEqual label, False, value
End Sub

Sub AssertStartsWith(ByVal label, ByVal prefix, ByVal actual)
    If Left(UCase(actual), Len(prefix)) = UCase(prefix) Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label
        WScript.Echo "       Expected prefix: """ & prefix & """"
        WScript.Echo "       Actual         : """ & actual & """"
    End If
End Sub

'====================================================================
' Tests
'====================================================================
WScript.Echo "VTD Labor Gate Tests"
WScript.Echo "===================="

'--------------------------------------------------------------------
' Test 1: Fixture screen_i91_tech_875722.txt
' Line C has ltype=I and "VTD VEND TO DEALER" in its L-row description.
'--------------------------------------------------------------------
Dim fixturePath
fixturePath = g_fso.BuildPath(g_fso.GetParentFolderName(WScript.ScriptFullName), "fixtures\screen_i91_tech_875722.txt")

Dim skipReason
Set g_bzhao = New FakeBzhao
g_bzhao.SetPage ParseScreenMapToBuffer(fixturePath)
AssertFalse "Fixture: gate fails for ltype=I + VTD in L-row desc", EvaluateVtdLaborGate(skipReason)
AssertStartsWith "Fixture: skipReason has expected prefix", "Skipped - VTD labor line:", skipReason

'--------------------------------------------------------------------
' Test 2: Synthetic — ltype I + whole-word "VTD" in L-row desc
'--------------------------------------------------------------------
Set g_bzhao = New FakeBzhao
g_bzhao.SetPage BuildPage("A", "VEND TO DEALER", "I", "VTD VEND TO DEALER")
skipReason = ""
AssertFalse "Synthetic: gate fails when ltype=I and L-row desc has whole-word VTD", EvaluateVtdLaborGate(skipReason)
AssertStartsWith "Synthetic: skipReason prefix", "Skipped - VTD labor line:", skipReason

'--------------------------------------------------------------------
' Test 3: Synthetic — VTD in header desc, ltype I, L-row desc has no VTD
'--------------------------------------------------------------------
Set g_bzhao = New FakeBzhao
g_bzhao.SetPage BuildPage("A", "VTD CHECK PROCEDURE", "I", "LABOR ONLY WORK")
skipReason = ""
AssertFalse "Synthetic: gate fails when header desc has VTD and ltype=I", EvaluateVtdLaborGate(skipReason)

'--------------------------------------------------------------------
' Test 4: Synthetic — ltype I, no VTD anywhere -> gate passes
'--------------------------------------------------------------------
Set g_bzhao = New FakeBzhao
g_bzhao.SetPage BuildPage("A", "ADJUST BRAKES", "I", "INSPECT AND ADJUST")
skipReason = ""
AssertTrue "Synthetic: gate passes when ltype=I but no VTD", EvaluateVtdLaborGate(skipReason)
AssertEqual "Synthetic: skipReason is empty on pass", "", skipReason

'--------------------------------------------------------------------
' Test 5: Synthetic — VTD in L-row desc but ltype != I -> gate passes
'--------------------------------------------------------------------
Set g_bzhao = New FakeBzhao
g_bzhao.SetPage BuildPage("A", "VEND TO DEALER", "R", "VTD VEND TO DEALER")
skipReason = ""
AssertTrue "Synthetic: gate passes when ltype is not I (even with VTD in desc)", EvaluateVtdLaborGate(skipReason)

'--------------------------------------------------------------------
' Test 6: Synthetic — "VTDEV" in L-row desc (not whole word), ltype I -> gate passes
'--------------------------------------------------------------------
Set g_bzhao = New FakeBzhao
g_bzhao.SetPage BuildPage("A", "VTDEV INSPECTION", "I", "VTDEV SOFTWARE UPDATE")
skipReason = ""
AssertTrue "Synthetic: gate passes when VTD is only a substring (VTDEV), not whole word", EvaluateVtdLaborGate(skipReason)

'--------------------------------------------------------------------
' Tests 7-10: ContainsWholeWordVtd unit tests
'--------------------------------------------------------------------
AssertTrue  "ContainsWholeWordVtd: 'VTD' alone -> True",          ContainsWholeWordVtd("VTD")
AssertTrue  "ContainsWholeWordVtd: 'INSPECT VTD' -> True",        ContainsWholeWordVtd("INSPECT VTD")
AssertFalse "ContainsWholeWordVtd: 'VTDEV CHECK' -> False",       ContainsWholeWordVtd("VTDEV CHECK")
AssertFalse "ContainsWholeWordVtd: 'AAAVTD' -> False",            ContainsWholeWordVtd("AAAVTD")

'--------------------------------------------------------------------
' Summary
'--------------------------------------------------------------------
WScript.Echo ""
WScript.Echo "Results: " & g_Pass & " passed, " & g_Fail & " failed."
If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All VTD labor gate tests passed."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & g_Fail & " test(s) failed."
    WScript.Quit 1
End If
