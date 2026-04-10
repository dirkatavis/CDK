'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestFcaDialogHandler
' **DATE CREATED:** 2026-04-10
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Unit tests for ExtractPartNumberForFca() using screen fixture files and
' synthetic buffers. Verifies part number extraction from rows 9-22 of the
' RO detail screen, including the case where the FCA dialog overlay covers
' the right half of the screen (cols ~37-80).
'
' Test cases:
'   1. Real fixture (RO 876518): P1 BBH6A001AA  -> "BBH6A001AA"
'   2. Blank screen (no P-lines)                -> ""
'   3. Synthetic P-line at row 12 "XYZ1234567"  -> "XYZ1234567"
'
' Fixture files are in tests\fixtures\ and are committed to source control.
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_fso
Set g_fso = CreateObject("Scripting.FileSystemObject")

' ---- Test counters ----
Dim g_Pass, g_Fail
g_Pass = 0
g_Fail = 0

' ---- Load AdvancedMock ----
Dim g_mockPath
g_mockPath = g_fso.BuildPath( _
    g_fso.GetParentFolderName( _
        g_fso.GetParentFolderName( _
            g_fso.GetParentFolderName( _
                g_fso.GetParentFolderName(WScript.ScriptFullName)))), _
    "framework\AdvancedMock.vbs")

If Not g_fso.FileExists(g_mockPath) Then
    WScript.Echo "[FAIL] Cannot find AdvancedMock.vbs at: " & g_mockPath
    WScript.Quit 1
End If
ExecuteGlobal g_fso.OpenTextFile(g_mockPath).ReadAll

' ---- Declare g_bzhao (matches PostFinalCharges.vbs global) ----
Dim g_bzhao

' ---- Inline IsWchLine (copy kept in sync with PostFinalCharges.vbs) ----
Function IsWchLine(lineLetterChar)
    IsWchLine = False
    Dim row, buf, inTargetLine, firstChar
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
                If Trim(Mid(buf, 50, 6)) = "WCH" Then
                    IsWchLine = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

' ---- Inline ExtractPartNumberForFca (copy kept in sync with PostFinalCharges.vbs) ----
Function ExtractPartNumberForFca()
    Dim row, buf, partToken, spacePos
    ExtractPartNumberForFca = ""
    For row = 9 To 22
        buf = ""
        On Error Resume Next
        g_bzhao.ReadScreen buf, 80, row, 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Len(buf) >= 20 Then
            If Mid(buf, 6, 1) = "P" And IsNumeric(Mid(buf, 7, 1)) Then
                partToken = Trim(Mid(buf, 9, 20))
                spacePos = InStr(1, partToken, " ")
                If spacePos > 1 Then partToken = Left(partToken, spacePos - 1)
                If Len(partToken) > 0 Then
                    ExtractPartNumberForFca = partToken
                    Exit Function
                End If
            End If
        End If
    Next
End Function

'-----------------------------------------------------------------------------------
' ParseScreenMapToBuffer
' Parses the coordinate map format produced by ro_screen_map.vbs:
'   "DD | <80 chars>"
' Returns a 24*80 character string suitable for AdvancedMock.SetBuffer().
'-----------------------------------------------------------------------------------
Function ParseScreenMapToBuffer(filePath)
    Dim buf: buf = String(24 * 80, " ")
    If Not g_fso.FileExists(filePath) Then
        ParseScreenMapToBuffer = buf
        Exit Function
    End If
    Dim ts: Set ts = g_fso.OpenTextFile(filePath, 1)
    Do While Not ts.AtEndOfStream
        Dim line: line = ts.ReadLine
        If Len(line) >= 7 And Mid(line, 3, 3) = " | " Then
            Dim rowNum
            On Error Resume Next
            rowNum = CInt(Left(line, 2))
            On Error GoTo 0
            If rowNum >= 1 And rowNum <= 24 Then
                Dim content: content = Mid(line, 6)
                content = Left(content & String(80, " "), 80)
                Dim pos: pos = ((rowNum - 1) * 80) + 1
                buf = Left(buf, pos - 1) & content & Mid(buf, pos + 80)
            End If
        End If
    Loop
    ts.Close
    ParseScreenMapToBuffer = buf
End Function

'-----------------------------------------------------------------------------------
' BuildSyntheticBuffer
' Builds a 24*80 buffer with a single row set to provided content.
'-----------------------------------------------------------------------------------
Function BuildSyntheticBuffer(rowNum, rowContent)
    Dim buf: buf = String(24 * 80, " ")
    Dim content: content = Left(rowContent & String(80, " "), 80)
    Dim pos: pos = ((rowNum - 1) * 80) + 1
    buf = Left(buf, pos - 1) & content & Mid(buf, pos + 80)
    BuildSyntheticBuffer = buf
End Function

'-----------------------------------------------------------------------------------
' BuildDetailScreenBuffer
' Builds a 24*80 buffer representing an RO detail screen with one line item.
' lineLetterChar: the line letter (e.g. "A")
' headerRow: screen row for the line letter header
' laborType: LTYPE value padded/placed at col 50 of the L1 row
'-----------------------------------------------------------------------------------
Function BuildDetailScreenBuffer(lineLetterChar, headerRow, laborType)
    Dim buf: buf = String(24 * 80, " ")
    ' Column header row (row 8)
    Dim hdrContent
    hdrContent = "LC DESCRIPTION                           TECH... LTYPE    ACT   SOLD    SALE AMT"
    hdrContent = Left(hdrContent & String(80, " "), 80)
    Dim hdrPos: hdrPos = ((8 - 1) * 80) + 1
    buf = Left(buf, hdrPos - 1) & hdrContent & Mid(buf, hdrPos + 80)
    ' Line letter header row
    Dim lineHdr: lineHdr = lineLetterChar & "  SOME DESCRIPTION                      C92                                    "
    lineHdr = Left(lineHdr & String(80, " "), 80)
    Dim linePos: linePos = ((headerRow - 1) * 80) + 1
    buf = Left(buf, linePos - 1) & lineHdr & Mid(buf, linePos + 80)
    ' L1 row (immediately after header)
    Dim ltype: ltype = Left(laborType & "      ", 6)  ' Pad LTYPE to 6 chars at col 50
    Dim lRow: lRow = "   L1 B SOME DESCRIPTION                  73166   " & ltype & "  0.00   0.10        8.10"
    lRow = Left(lRow & String(80, " "), 80)
    Dim lPos: lPos = ((headerRow) * 80) + 1  ' L1 is one row below header
    buf = Left(buf, lPos - 1) & lRow & Mid(buf, lPos + 80)
    BuildDetailScreenBuffer = buf
End Function

' ---- Test helpers ----
Sub AssertEqual(label, expected, actual)
    If expected = actual Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (expected: """ & expected & """, got: """ & actual & """)"
    End If
End Sub

' ---- Tests ----

WScript.Echo "IsWchLine + ExtractPartNumberForFca Unit Tests"
WScript.Echo "==============================================="

' ---- Test 1: Real screen capture (RO 876518) with P1 BBH6A001AA -> "BBH6A001AA" ----
Dim fixturePath
fixturePath = g_fso.BuildPath( _
    g_fso.GetParentFolderName(WScript.ScriptFullName), _
    "fixtures\screen_p1_charged_876518.txt")

If Not g_fso.FileExists(fixturePath) Then
    WScript.Echo "[FAIL] Fixture file not found: " & fixturePath
    WScript.Quit 1
End If

Dim mock1: Set mock1 = New AdvancedMock
mock1.Connect ""
mock1.SetBuffer ParseScreenMapToBuffer(fixturePath)
Set g_bzhao = mock1
AssertEqual "Real screen capture: P1 BBH6A001AA -> ""BBH6A001AA""", "BBH6A001AA", ExtractPartNumberForFca()

' ---- Test 2: Blank screen (no P-lines) -> "" ----
Dim mock2: Set mock2 = New AdvancedMock
mock2.Connect ""
mock2.SetBuffer String(24 * 80, " ")
Set g_bzhao = mock2
AssertEqual "Blank screen: no P-lines -> """"", "", ExtractPartNumberForFca()

' ---- Test 3: Synthetic P-line at row 12 with part "XYZ1234567" -> "XYZ1234567" ----
' Columns:   123456789012345678901234567890
'            "     P1 XYZ1234567 SOME DESC..."
Dim syntheticRow: syntheticRow = "     P1 XYZ1234567 SOME DESCRIPTION                                            "
Dim mock3: Set mock3 = New AdvancedMock
mock3.Connect ""
mock3.SetBuffer BuildSyntheticBuffer(12, syntheticRow)
Set g_bzhao = mock3
AssertEqual "Synthetic P-line at row 12 -> ""XYZ1234567""", "XYZ1234567", ExtractPartNumberForFca()

' ---- IsWchLine tests ----
WScript.Echo ""
WScript.Echo "--- IsWchLine ---"

' Test 4: Real fixture (RO 876518) - Line A has WCH -> True
Dim mock4wch: Set mock4wch = New AdvancedMock
mock4wch.Connect ""
mock4wch.SetBuffer ParseScreenMapToBuffer(fixturePath)
Set g_bzhao = mock4wch
AssertEqual "Real fixture line A has WCH -> True", True, IsWchLine("A")

' Test 5: Real fixture - Line B has labor type "I" (not WCH) -> False
Dim mock5wch: Set mock5wch = New AdvancedMock
mock5wch.Connect ""
mock5wch.SetBuffer ParseScreenMapToBuffer(fixturePath)
Set g_bzhao = mock5wch
AssertEqual "Real fixture line B labor type I -> False", False, IsWchLine("B")

' Test 6: Synthetic buffer - Line A L1 LTYPE = "WCH" -> True
Dim mock6wch: Set mock6wch = New AdvancedMock
mock6wch.Connect ""
mock6wch.SetBuffer BuildDetailScreenBuffer("A", 10, "WCH")
Set g_bzhao = mock6wch
AssertEqual "Synthetic line A LTYPE=WCH -> True", True, IsWchLine("A")

' Test 7: Synthetic buffer - Line A L1 LTYPE = "I" -> False
Dim mock7wch: Set mock7wch = New AdvancedMock
mock7wch.Connect ""
mock7wch.SetBuffer BuildDetailScreenBuffer("A", 10, "I")
Set g_bzhao = mock7wch
AssertEqual "Synthetic line A LTYPE=I -> False", False, IsWchLine("A")

' ---- Summary ----
WScript.Echo ""
If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All " & g_Pass & " FCA dialog handler tests passed."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & g_Fail & " test(s) failed."
    WScript.Quit 1
End If
