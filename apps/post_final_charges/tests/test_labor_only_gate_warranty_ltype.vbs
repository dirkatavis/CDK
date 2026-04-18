'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestLaborOnlyGateWarrantyLtype
' **DATE CREATED:** 2026-04-18
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Behavioral regression tests for EvaluateLaborOnlyGate covering:
'   1) Unsupported W* ltype skipped regardless of parts
'   2) Supported W* ltype (WCH) passes warranty check; no-parts + no exception skips
'   3) Non-W ltype, no parts, no exception -> no-parts skip
'   4) Empty g_arrWarrantyLTypes -> any W* ltype skipped
'   5) L-row with P-row following -> passes unconditionally (parts attached)
'   6) L-row without P-row, description matches exception -> passes
'   7) L-row without P-row, no exception -> skips
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_Pass, g_Fail
g_Pass = 0
g_Fail = 0

' ---- Globals required by EvaluateLaborOnlyGate ----
Dim g_bzhao
Dim g_arrWarrantyLTypes
Dim g_arrCDKDescriptionExceptions

' ---- Stub helpers ----
Sub LogEvent(ByVal a, ByVal b, ByVal c, ByVal d, ByVal e, ByVal f)
    ' no-op for unit tests
End Sub

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

' ---- Minimal FakeBzhao: single-page, no paging ----
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
        ' no-op
    End Sub

    Public Sub Pause(ByVal ms)
        ' no-op
    End Sub
End Class

' ---- Screen builder helpers ----
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

Function BuildLRow(ByVal ltypeCode, ByVal lRowDesc)
    Dim rowText
    rowText = String(80, " ")
    rowText = SetColText(rowText, 4, "L1")
    rowText = SetColText(rowText, 7, Left(lRowDesc & String(35, " "), 35))
    rowText = SetColText(rowText, 50, ltypeCode)
    BuildLRow = rowText
End Function

Function BuildPRow()
    Dim rowText
    rowText = String(80, " ")
    rowText = SetColText(rowText, 6, "P1")
    rowText = SetColText(rowText, 9, "PARTNUM")
    BuildPRow = rowText
End Function

Function BuildHeaderRow()
    Dim rowText
    rowText = String(80, " ")
    rowText = SetColText(rowText, 1, "A ")
    BuildHeaderRow = rowText
End Function

Function BuildEndRow()
    Dim rowText
    rowText = String(80, " ")
    rowText = SetColText(rowText, 1, "(END OF DISPLAY)")
    BuildEndRow = rowText
End Function

' Build a single-page screen with an L-row and no P-row following it.
Function BuildPageNoP(ByVal ltypeCode, ByVal lRowDesc)
    Dim pageBuf
    pageBuf = String(24 * 80, " ")
    pageBuf = SetRow(pageBuf, 9, BuildHeaderRow())
    pageBuf = SetRow(pageBuf, 10, BuildLRow(ltypeCode, lRowDesc))
    pageBuf = SetRow(pageBuf, 22, BuildEndRow())
    BuildPageNoP = pageBuf
End Function

' Build a single-page screen with an L-row followed immediately by a P-row.
Function BuildPageWithP(ByVal ltypeCode, ByVal lRowDesc)
    Dim pageBuf
    pageBuf = String(24 * 80, " ")
    pageBuf = SetRow(pageBuf, 9, BuildHeaderRow())
    pageBuf = SetRow(pageBuf, 10, BuildLRow(ltypeCode, lRowDesc))
    pageBuf = SetRow(pageBuf, 11, BuildPRow())
    pageBuf = SetRow(pageBuf, 22, BuildEndRow())
    BuildPageWithP = pageBuf
End Function

' ---- EvaluateLaborOnlyGate (copy-pasted from PostFinalCharges.vbs) ----
Function EvaluateLaborOnlyGate(ByRef skipReason)
    Dim row, buf
    Dim pageIndicator
    Dim doneScanning, pagesAdvanced, p
    Dim preSig, postSig, preSig2, postSig2, preMarker, postMarker
    Dim maxPageAdvances
    Dim currentLineHeaderDesc, lRowDesc, hasLRows
    Dim lineHeaderPasses, lRowPasses
    Dim lTypeCode, unsupportedWarranty, warrantyIndex
    Dim pendingLRowDesc, pendingLRowHeaderDesc

    EvaluateLaborOnlyGate = True
    skipReason = ""

    doneScanning = False
    pagesAdvanced = 0
    maxPageAdvances = 50
    currentLineHeaderDesc = ""
    hasLRows = False
    pendingLRowDesc = ""
    pendingLRowHeaderDesc = ""

    Do While Not doneScanning
        For row = 9 To 23
            buf = ""
            On Error Resume Next
            g_bzhao.ReadScreen buf, 80, row, 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0

            If Len(pendingLRowDesc) > 0 Then
                If Len(buf) >= 7 And Mid(buf, 6, 1) = "P" And IsNumeric(Mid(buf, 7, 1)) Then
                    Call LogEvent("comm", "med", "Labor-only gate L-row has parts — passes", "EvaluateLaborOnlyGate", _
                        "lRowDesc=[" & pendingLRowDesc & "]", "")
                    pendingLRowDesc = ""
                    pendingLRowHeaderDesc = ""
                Else
                    lineHeaderPasses = IsCdkLaborOnlyExceptionDesc(pendingLRowHeaderDesc)
                    lRowPasses = IsCdkLaborOnlyExceptionDesc(pendingLRowDesc)

                    Call LogEvent("comm", "med", "Labor-only gate L-row (no parts) scanned", "EvaluateLaborOnlyGate", _
                        "lineHeader=[" & pendingLRowHeaderDesc & "] lRowDesc=[" & pendingLRowDesc & "] headerPass=" & lineHeaderPasses & " lRowPass=" & lRowPasses, "")

                    If Not lineHeaderPasses And Not lRowPasses Then
                        skipReason = "Skipped - No parts charged: lrow=[" & pendingLRowDesc & "] header=[" & pendingLRowHeaderDesc & "]"
                        Call LogEvent("comm", "low", "Labor-only gate SKIP — L-row has no parts and no exception keyword", "EvaluateLaborOnlyGate", _
                            "lineHeader=[" & pendingLRowHeaderDesc & "] lRowDesc=[" & pendingLRowDesc & "]", "")
                        EvaluateLaborOnlyGate = False
                        doneScanning = True
                        Exit For
                    End If
                    pendingLRowDesc = ""
                    pendingLRowHeaderDesc = ""
                End If
            End If

            If Len(buf) >= 5 Then
                If Mid(buf, 1, 1) >= "A" And Mid(buf, 1, 1) <= "Z" And Mid(buf, 2, 1) = " " And Mid(buf, 4, 1) <> "L" Then
                    currentLineHeaderDesc = Trim(Mid(buf, 4, 38))
                End If

                If Mid(buf, 4, 1) = "L" And IsNumeric(Mid(buf, 5, 1)) Then
                    hasLRows = True
                    lRowDesc = Trim(Mid(buf, 7, 35))
                    lTypeCode = UCase(Trim(Mid(buf, 50, 6)))

                    If Left(lTypeCode, 1) = "W" Then
                        unsupportedWarranty = True
                        If IsArray(g_arrWarrantyLTypes) Then
                            For warrantyIndex = 0 To UBound(g_arrWarrantyLTypes)
                                If lTypeCode = g_arrWarrantyLTypes(warrantyIndex) Then
                                    unsupportedWarranty = False
                                    Exit For
                                End If
                            Next
                        End If
                        If unsupportedWarranty Then
                            skipReason = "Skipped - Unsupported warranty ltype: [" & lTypeCode & "] lrow=[" & lRowDesc & "]"
                            Call LogEvent("comm", "low", "Labor-only gate SKIP — unsupported warranty ltype", "EvaluateLaborOnlyGate", _
                                "ltype=[" & lTypeCode & "] not in WarrantyLTypes", "")
                            EvaluateLaborOnlyGate = False
                            doneScanning = True
                            Exit For
                        End If
                    End If

                    pendingLRowDesc = lRowDesc
                    pendingLRowHeaderDesc = currentLineHeaderDesc
                End If
            End If
        Next

        If Not doneScanning Then
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
                g_bzhao.ReadScreen postSig, 80, 9, 1
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

' ---- Assert helpers ----
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

Sub AssertContains(ByVal label, ByVal haystack, ByVal needle)
    If InStr(1, haystack, needle, vbTextCompare) > 0 Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (expected '" & needle & "' in '" & haystack & "')"
    End If
End Sub

' ============================
' Tests
' ============================
WScript.Echo "Labor-Only Gate — Behavioral Tests"
WScript.Echo "==================================="

g_arrCDKDescriptionExceptions = Array("check and adjust", "oil and filter")
g_arrWarrantyLTypes = Array("WCH", "WV")

' --- Test 1: WF ltype (unsupported W*) -> skipped ---
Dim mock1
Set mock1 = New FakeBzhao
mock1.SetPage BuildPageNoP("WF", "SOME REPAIR")
Set g_bzhao = mock1
Dim result1, reason1
result1 = EvaluateLaborOnlyGate(reason1)
AssertFalse "WF ltype (unsupported) causes gate to return False", result1
AssertContains "WF skip reason contains unsupported warranty ltype prefix", reason1, "Skipped - Unsupported warranty ltype: [WF]"
AssertContains "WF skip reason includes lrow description", reason1, "lrow=[SOME REPAIR]"

' --- Test 2: WCH ltype (supported), no parts, no exception -> no-parts skip ---
Dim mock2
Set mock2 = New FakeBzhao
mock2.SetPage BuildPageNoP("WCH", "SOME REPAIR")
Set g_bzhao = mock2
Dim result2, reason2
result2 = EvaluateLaborOnlyGate(reason2)
AssertFalse "WCH ltype (supported) does not trigger unsupported-warranty skip", _
    (InStr(1, reason2, "Unsupported warranty ltype", vbTextCompare) > 0)
AssertContains "WCH ltype with no parts and no exception skips as no-parts-charged", reason2, "Skipped - No parts charged"

' --- Test 3: Non-W ltype, no parts, no exception -> no-parts skip ---
Dim mock3
Set mock3 = New FakeBzhao
mock3.SetPage BuildPageNoP("CP", "OIL CHANGE")
Set g_bzhao = mock3
Dim result3, reason3
result3 = EvaluateLaborOnlyGate(reason3)
AssertFalse "Non-W ltype with no parts and no exception skips", result3
AssertContains "Non-W ltype skip reason is no-parts-charged", reason3, "Skipped - No parts charged"
AssertFalse "Non-W ltype skip reason does not mention unsupported warranty", _
    (InStr(1, reason3, "Unsupported warranty ltype", vbTextCompare) > 0)

' --- Test 4: Empty g_arrWarrantyLTypes -> any W* ltype skipped ---
Dim mock4
Set mock4 = New FakeBzhao
mock4.SetPage BuildPageNoP("WV", "RECALL REPAIR")
Set g_bzhao = mock4
g_arrWarrantyLTypes = Array()
Dim result4, reason4
result4 = EvaluateLaborOnlyGate(reason4)
AssertFalse "Empty WarrantyLTypes causes any W* ltype to be skipped", result4
AssertContains "Empty-list skip reason contains unsupported warranty prefix", reason4, "Skipped - Unsupported warranty ltype: [WV]"
g_arrWarrantyLTypes = Array("WCH", "WV")

' --- Test 5: L-row WITH P-row following -> passes unconditionally ---
Dim mock5
Set mock5 = New FakeBzhao
mock5.SetPage BuildPageWithP("CP", "REPLACE OIL AND FILTER")
Set g_bzhao = mock5
Dim result5, reason5
result5 = EvaluateLaborOnlyGate(reason5)
AssertTrue "L-row with P-row following passes unconditionally", result5

' --- Test 6: L-row without P-row, description matches exception -> passes ---
Dim mock6
Set mock6 = New FakeBzhao
mock6.SetPage BuildPageNoP("CP", "REPLACE OIL AND FILTER")
Set g_bzhao = mock6
Dim result6, reason6
result6 = EvaluateLaborOnlyGate(reason6)
AssertTrue "L-row without P-row but exception keyword in description passes", result6

' --- Test 7: L-row without P-row, no exception -> skips ---
Dim mock7
Set mock7 = New FakeBzhao
mock7.SetPage BuildPageNoP("CP", "REPLACE BRAKE PADS")
Set g_bzhao = mock7
Dim result7, reason7
result7 = EvaluateLaborOnlyGate(reason7)
AssertFalse "L-row without P-row and no exception keyword skips", result7
AssertContains "Skip reason is no-parts-charged with description", reason7, "Skipped - No parts charged"

WScript.Echo ""
If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All " & g_Pass & " labor-only gate tests passed."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & g_Fail & " test(s) failed."
    WScript.Quit 1
End If
