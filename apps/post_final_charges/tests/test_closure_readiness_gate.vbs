'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** test_closure_readiness_gate.vbs
' **DATE CREATED:** 2026-04-14
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Unit tests for GetFirstNonCompliantLineTech() using synthetic screen buffers
' built from coordinate-finder log captures.
'
' Column layout (1-indexed, confirmed by Coordinate_Finder.log):
'   Col 1     = line letter (A, B, C ...)
'   Cols 42+  = tech code (e.g. "C92", "I91")  <-- col 42, not 44
'
' Test cases:
'   1. All-compliant screen (C92/C93 only)        -> "" (no skip)
'   2. I91 on line C (RO 875722 from screenshot)  -> "Line C: I91" (skip)
'   3. H20 on line D (observed in coord-finder)   -> "Line D: H20" (skip)
'   4. Gate disabled (empty AllowedTechCodes)     -> "" (no skip)
'   5. Blank tech code on a line header           -> "" (treated compliant)
'-----------------------------------------------------------------------------------
Option Explicit

Dim g_fso
Set g_fso = CreateObject("Scripting.FileSystemObject")

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

' ---- Declare g_AllowedTechCodes (matches PostFinalCharges.vbs global) ----
Dim g_AllowedTechCodes

'-----------------------------------------------------------------------------------
' Inline copy of GetFirstNonCompliantLineTech (kept in sync with PostFinalCharges.vbs)
'-----------------------------------------------------------------------------------
Function GetFirstNonCompliantLineTech()
    GetFirstNonCompliantLineTech = ""
    If Not IsArray(g_AllowedTechCodes) Then Exit Function
    If UBound(g_AllowedTechCodes) < 0 Then Exit Function
    If Len(Trim(g_AllowedTechCodes(0))) = 0 Then Exit Function  ' empty config = gate disabled

    Dim row, buf, firstChar, techCode, i, isAllowed
    For row = 9 To 22
        buf = ""
        On Error Resume Next
        g_bzhao.ReadScreen buf, 80, row, 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Len(buf) >= 44 Then
            firstChar = Mid(buf, 1, 1)
            If firstChar >= "A" And firstChar <= "Z" Then
                techCode = UCase(Trim(Mid(buf, 42, 8)))
                If Len(techCode) > 0 Then
                    isAllowed = False
                    For i = 0 To UBound(g_AllowedTechCodes)
                        If techCode = g_AllowedTechCodes(i) Then
                            isAllowed = True
                            Exit For
                        End If
                    Next
                    If Not isAllowed Then
                        GetFirstNonCompliantLineTech = "Line " & firstChar & ": " & techCode
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

'-----------------------------------------------------------------------------------
' ParseScreenMapToBuffer
' Parses the coordinate map format: "DD | <80 chars>"
' Returns a 24*80 character string for AdvancedMock.SetBuffer().
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
' Builds a 24*80 buffer with one or more rows set from an array.
' rowDefs: array of "RR|<content>" pairs.
'-----------------------------------------------------------------------------------
Function BuildSyntheticBuffer(rowNum, rowContent)
    Dim buf: buf = String(24 * 80, " ")
    Dim content: content = Left(rowContent & String(80, " "), 80)
    Dim pos: pos = ((rowNum - 1) * 80) + 1
    buf = Left(buf, pos - 1) & content & Mid(buf, pos + 80)
    BuildSyntheticBuffer = buf
End Function

' Builds multi-row buffer from array of "RR|content" strings
Function BuildMultiRowBuffer(rowDefs)
    Dim buf: buf = String(24 * 80, " ")
    Dim j, parts, rowNum, content, pos
    For j = 0 To UBound(rowDefs)
        parts = Split(rowDefs(j), "|", 2)
        rowNum = CInt(Trim(parts(0)))
        content = Left(parts(1) & String(80, " "), 80)
        pos = ((rowNum - 1) * 80) + 1
        buf = Left(buf, pos - 1) & content & Mid(buf, pos + 80)
    Next
    BuildMultiRowBuffer = buf
End Function

' ---- Helpers ----
Sub AssertEqual(label, expected, actual)
    If expected = actual Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label
        WScript.Echo "       Expected: """ & expected & """"
        WScript.Echo "       Actual  : """ & actual & """"
    End If
End Sub

'===================================================================================
' TEST 1: All-compliant (C92/C93 only) - RO 876518 fixture
'===================================================================================
Sub Test_AllCompliant_FromFixture()
    Dim fixturePath
    fixturePath = g_fso.BuildPath( _
        g_fso.GetParentFolderName(WScript.ScriptFullName), _
        "fixtures\screen_p1_charged_876518.txt")

    g_AllowedTechCodes = Array("C92", "C93")
    Set g_bzhao = New AdvancedMock
    g_bzhao.Connect ""
    g_bzhao.SetBuffer ParseScreenMapToBuffer(fixturePath)

    AssertEqual "All-compliant screen (C92 only) -> no skip", "", GetFirstNonCompliantLineTech()
End Sub

'===================================================================================
' TEST 2: I91 on line C (RO 875722 from screenshot) - fixture file
'===================================================================================
Sub Test_I91_OnLineC_FromFixture()
    Dim fixturePath
    fixturePath = g_fso.BuildPath( _
        g_fso.GetParentFolderName(WScript.ScriptFullName), _
        "fixtures\screen_i91_tech_875722.txt")

    g_AllowedTechCodes = Array("C92", "C93")
    Set g_bzhao = New AdvancedMock
    g_bzhao.Connect ""
    g_bzhao.SetBuffer ParseScreenMapToBuffer(fixturePath)

    AssertEqual "I91 on line C -> skip with 'Line C: I91'", "Line C: I91", GetFirstNonCompliantLineTech()
End Sub

'===================================================================================
' TEST 3: H20 on line D (synthetic - observed in coordinate finder log)
'===================================================================================
Sub Test_H20_OnLineD_Synthetic()
    ' Build rows matching the coordinate-finder layout (row 16 = line D with H20)
    ' Col 42 is where tech code lives. Row content (1-indexed, 80 chars):
    '   "D  STATES AC BLOWS WARM AIR              H20" padded to 80
    Dim rows(1)
    rows(0) = "10|A  CHECK AND ADJUST TIRE PRESSURE        C92"
    rows(1) = "16|D  STATES AC BLOWS WARM AIR              H20"

    g_AllowedTechCodes = Array("C92", "C93")
    Set g_bzhao = New AdvancedMock
    g_bzhao.Connect ""
    g_bzhao.SetBuffer BuildMultiRowBuffer(rows)

    AssertEqual "H20 on line D -> skip with 'Line D: H20'", "Line D: H20", GetFirstNonCompliantLineTech()
End Sub

'===================================================================================
' TEST 4: Gate disabled (empty AllowedTechCodes first element)
'===================================================================================
Sub Test_GateDisabled_NoSkip()
    Dim rows(0)
    rows(0) = "10|C  VEND TO DEALER                        I91"

    g_AllowedTechCodes = Array("")   ' empty = disabled
    Set g_bzhao = New AdvancedMock
    g_bzhao.Connect ""
    g_bzhao.SetBuffer BuildMultiRowBuffer(rows)

    AssertEqual "Gate disabled (empty AllowedTechCodes) -> no skip", "", GetFirstNonCompliantLineTech()
End Sub

'===================================================================================
' TEST 5: Line header with blank tech code treated as compliant
'===================================================================================
Sub Test_BlankTechCode_Compliant()
    ' Line A has no tech code assigned yet (blank at col 42+)
    Dim rows(0)
    rows(0) = "10|A  CHECK FUELDOOR, WON'T CLOSE          "   ' trailing spaces, no tech

    g_AllowedTechCodes = Array("C92", "C93")
    Set g_bzhao = New AdvancedMock
    g_bzhao.Connect ""
    g_bzhao.SetBuffer BuildMultiRowBuffer(rows)

    AssertEqual "Blank tech code on header row -> treated compliant", "", GetFirstNonCompliantLineTech()
End Sub

'===================================================================================
' TEST 6: Col alignment sanity - verify Mid(buf,42,8) extracts code from fixture row
'===================================================================================
Sub Test_ColAlignmentSanity()
    ' Directly verify that row 10 of the I91 fixture reads "C93" at col 42
    Dim fixturePath
    fixturePath = g_fso.BuildPath( _
        g_fso.GetParentFolderName(WScript.ScriptFullName), _
        "fixtures\screen_i91_tech_875722.txt")

    Dim fullBuf: fullBuf = ParseScreenMapToBuffer(fixturePath)
    ' Row 10 starts at offset (10-1)*80+1 = 721
    Dim row10: row10 = Mid(fullBuf, 721, 80)
    Dim extracted: extracted = UCase(Trim(Mid(row10, 42, 8)))
    AssertEqual "Row 10 col 42 extracts 'C93' (line A)", "C93", extracted

    ' Row 15 = line C with I91, starts at offset (15-1)*80+1 = 1121
    Dim row15: row15 = Mid(fullBuf, 1121, 80)
    extracted = UCase(Trim(Mid(row15, 42, 8)))
    AssertEqual "Row 15 col 42 extracts 'I91' (line C)", "I91", extracted
End Sub

'===================================================================================
' RUN ALL TESTS
'===================================================================================
Test_AllCompliant_FromFixture
Test_I91_OnLineC_FromFixture
Test_H20_OnLineD_Synthetic
Test_GateDisabled_NoSkip
Test_BlankTechCode_Compliant
Test_ColAlignmentSanity

WScript.Echo ""
WScript.Echo "Results: " & g_Pass & " passed, " & g_Fail & " failed."
If g_Fail > 0 Then WScript.Quit 1
