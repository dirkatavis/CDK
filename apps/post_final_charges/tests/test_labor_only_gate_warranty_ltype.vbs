'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestLaborOnlyGateWarrantyLtype
' **DATE CREATED:** 2026-04-18
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Behavioral regression tests for the unsupported-warranty-ltype skip path in
' EvaluateLaborOnlyGate. Verifies:
'   1) An L-row with a W* ltype not in g_arrWarrantyLTypes returns False and
'      sets the expected skip reason.
'   2) An L-row with a supported W* ltype (e.g. WCH) does not trigger the gate.
'   3) An L-row with a non-W ltype and no description exception triggers the
'      normal no-parts-charged skip (not the warranty skip).
'   4) An empty g_arrWarrantyLTypes causes any W* ltype to be skipped.
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
        ' no-op — single page, no navigation expected
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

Function BuildPage(ByVal ltypeCode, ByVal lRowDesc)
    ' Builds a 24-row x 80-col page buffer.
    ' Row 9:  line header row (A)
    ' Row 10: L-row with ltype at col 50, description at col 7
    ' Row 22: end-of-display marker
    Dim pageBuf, headerRow, lRow, endRow
    pageBuf = String(24 * 80, " ")

    headerRow = String(80, " ")
    headerRow = SetColText(headerRow, 1, "A ")

    lRow = String(80, " ")
    lRow = SetColText(lRow, 4, "L1")
    lRow = SetColText(lRow, 7, Left(lRowDesc & String(35, " "), 35))
    lRow = SetColText(lRow, 50, ltypeCode)

    endRow = String(80, " ")
    endRow = SetColText(endRow, 1, "(END OF DISPLAY)")

    Dim pos
    pos = (9 - 1) * 80 + 1
    pageBuf = Left(pageBuf, pos - 1) & headerRow & Mid(pageBuf, pos + 80)
    pos = (10 - 1) * 80 + 1
    pageBuf = Left(pageBuf, pos - 1) & lRow & Mid(pageBuf, pos + 80)
    pos = (22 - 1) * 80 + 1
    pageBuf = Left(pageBuf, pos - 1) & endRow & Mid(pageBuf, pos + 80)

    BuildPage = pageBuf
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

    EvaluateLaborOnlyGate = True
    skipReason = ""

    doneScanning = False
    pagesAdvanced = 0
    maxPageAdvances = 50
    currentLineHeaderDesc = ""
    hasLRows = False

    Do While Not doneScanning
        For row = 9 To 22
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

                    lineHeaderPasses = IsCdkLaborOnlyExceptionDesc(currentLineHeaderDesc)
                    lRowPasses = IsCdkLaborOnlyExceptionDesc(lRowDesc)

                    Call LogEvent("comm", "med", "Labor-only gate L-row scanned", "EvaluateLaborOnlyGate", _
                        "lineHeader=[" & currentLineHeaderDesc & "] lRowDesc=[" & lRowDesc & "] headerPass=" & lineHeaderPasses & " lRowPass=" & lRowPasses, "")

                    If Not lineHeaderPasses And Not lRowPasses Then
                        skipReason = "Skipped - No parts charged: lrow=[" & lRowDesc & "] header=[" & currentLineHeaderDesc & "]"
                        Call LogEvent("comm", "low", "Labor-only gate SKIP — L-row does not contain a listed keyword", "EvaluateLaborOnlyGate", _
                            "lineHeader=[" & currentLineHeaderDesc & "] lRowDesc=[" & lRowDesc & "]", "")
                        EvaluateLaborOnlyGate = False
                        doneScanning = True
                        Exit For
                    End If
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
WScript.Echo "Labor-Only Gate — Unsupported Warranty Ltype Tests"
WScript.Echo "==================================================="

g_arrCDKDescriptionExceptions = Array("check and adjust")

' --- Test 1: WF ltype (W* but not in list) -> skipped ---
Dim mock1
Set mock1 = New FakeBzhao
mock1.SetPage BuildPage("WF", "SOME REPAIR")
Set g_bzhao = mock1
g_arrWarrantyLTypes = Array("WCH", "WV")
Dim result1, reason1
result1 = EvaluateLaborOnlyGate(reason1)
AssertFalse "WF ltype (unsupported) causes gate to return False", result1
AssertContains "WF skip reason contains unsupported warranty ltype prefix", reason1, "Skipped - Unsupported warranty ltype: [WF]"
AssertContains "WF skip reason includes lrow description", reason1, "lrow=[SOME REPAIR]"

' --- Test 2: WCH ltype (supported) -> gate does not skip on warranty grounds ---
' WCH has no description exception either, so it falls through to the no-parts-charged check.
' With no description exception it will skip as "no parts charged", not "unsupported warranty".
Dim mock2
Set mock2 = New FakeBzhao
mock2.SetPage BuildPage("WCH", "SOME REPAIR")
Set g_bzhao = mock2
g_arrWarrantyLTypes = Array("WCH", "WV")
Dim result2, reason2
result2 = EvaluateLaborOnlyGate(reason2)
AssertFalse "WCH ltype (supported) does not trigger unsupported-warranty skip", _
    (InStr(1, reason2, "Unsupported warranty ltype", vbTextCompare) > 0)
AssertContains "WCH ltype skips as no-parts-charged (not warranty gate)", reason2, "Skipped - No parts charged"

' --- Test 3: Non-W ltype with no description exception -> normal no-parts skip ---
Dim mock3
Set mock3 = New FakeBzhao
mock3.SetPage BuildPage("CP", "OIL CHANGE")
Set g_bzhao = mock3
g_arrWarrantyLTypes = Array("WCH", "WV")
Dim result3, reason3
result3 = EvaluateLaborOnlyGate(reason3)
AssertFalse "Non-W ltype with no exception triggers no-parts skip", result3
AssertContains "Non-W ltype skip reason is no-parts-charged", reason3, "Skipped - No parts charged"
AssertFalse "Non-W ltype skip reason does not mention unsupported warranty", _
    (InStr(1, reason3, "Unsupported warranty ltype", vbTextCompare) > 0)

' --- Test 4: Empty g_arrWarrantyLTypes -> any W* ltype skipped ---
Dim mock4
Set mock4 = New FakeBzhao
mock4.SetPage BuildPage("WV", "RECALL REPAIR")
Set g_bzhao = mock4
g_arrWarrantyLTypes = Array()
Dim result4, reason4
result4 = EvaluateLaborOnlyGate(reason4)
AssertFalse "Empty WarrantyLTypes causes any W* ltype to be skipped", result4
AssertContains "Empty-list skip reason contains unsupported warranty prefix", reason4, "Skipped - Unsupported warranty ltype: [WV]"

WScript.Echo ""
If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All " & g_Pass & " labor-only gate warranty ltype tests passed."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & g_Fail & " test(s) failed."
    WScript.Quit 1
End If
