'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestHasPartsCharged
' **DATE CREATED:** 2026-04-09
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Unit tests for HasPartsCharged() using screen fixture files captured via the
' RO Screen Mapper tool (ro_screen_map.vbs).
'
' ParseScreenMapToBuffer() parses the "DD | <80chars>" format produced by the
' mapper and builds a 24*80 character buffer suitable for AdvancedMock.SetBuffer().
'
' Test cases:
'   1. Real capture (RO 876518): P1 BBH6A001AA $295.12  -> True
'   2. P-line present but SALE AMT is 0.00               -> False
'   3. No P-lines on screen                              -> False
'   4. P-line with empty (space-padded) SALE AMT column  -> False
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

' ---- Inline HasPartsCharged (copy kept in sync with PostFinalCharges.vbs) ----
Function HasPartsCharged()
    Dim row, buf, amtRaw, amtVal
    HasPartsCharged = False
    For row = 9 To 22
        buf = ""
        On Error Resume Next
        g_bzhao.ReadScreen buf, 80, row, 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Len(buf) >= 80 Then
            If Mid(buf, 6, 1) = "P" And IsNumeric(Mid(buf, 7, 1)) Then
                amtRaw = Trim(Mid(buf, 70, 11))
                amtVal = 0
                If IsNumeric(amtRaw) Then amtVal = CDbl(amtRaw)
                If amtVal > 0 Then
                    HasPartsCharged = True
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
        ' Match lines of the form "DD | <content>" (rows 01-23)
        If Len(line) >= 7 And Mid(line, 3, 3) = " | " Then
            Dim rowNum
            On Error Resume Next
            rowNum = CInt(Left(line, 2))
            On Error GoTo 0
            If rowNum >= 1 And rowNum <= 24 Then
                ' Content starts at position 6 (after "DD | ")
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
' All other rows are blank.
'-----------------------------------------------------------------------------------
Function BuildSyntheticBuffer(rowNum, rowContent)
    Dim buf: buf = String(24 * 80, " ")
    Dim content: content = Left(rowContent & String(80, " "), 80)
    Dim pos: pos = ((rowNum - 1) * 80) + 1
    buf = Left(buf, pos - 1) & content & Mid(buf, pos + 80)
    BuildSyntheticBuffer = buf
End Function

' ---- Test helpers ----
Sub AssertTrue(label, value)
    If value Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (expected True, got False)"
    End If
End Sub

Sub AssertFalse(label, value)
    If Not value Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (expected False, got True)"
    End If
End Sub

' ---- Tests ----

WScript.Echo "HasPartsCharged Unit Tests"
WScript.Echo "=========================="

' ---- Test 1: Real screen capture with P1 $295.12 -> True ----
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
AssertTrue "Real screen capture: P1 BBH6A001AA $295.12 -> True", HasPartsCharged()

' ---- Test 2: P-line at row 12 but SALE AMT is 0.00 -> False ----
' Screen col 6 = P, col 7 = 1, SALE AMT (cols 70-80) = "      0.00 "
'                    0         1         2         3         4         5         6         7
'                    0123456789012345678901234567890123456789012345678901234567890123456789012345678
Dim zeroAmtRow: zeroAmtRow = "     P1 BBH6A001AA BATTERY-STORAGE                               1        0.00"
Dim mock2: Set mock2 = New AdvancedMock
mock2.Connect ""
mock2.SetBuffer BuildSyntheticBuffer(12, zeroAmtRow)
Set g_bzhao = mock2
AssertFalse "P-line present, SALE AMT = 0.00 -> False", HasPartsCharged()

' ---- Test 3: No P-lines on screen -> False ----
Dim mock3: Set mock3 = New AdvancedMock
mock3.Connect ""
mock3.SetBuffer String(24 * 80, " ")
Set g_bzhao = mock3
AssertFalse "No P-lines on screen -> False", HasPartsCharged()

' ---- Test 4: P-line with fully blank SALE AMT column -> False ----
Dim blankAmtRow: blankAmtRow = "     P1 BBH6A001AA BATTERY-STORAGE                                              "
Dim mock4: Set mock4 = New AdvancedMock
mock4.Connect ""
mock4.SetBuffer BuildSyntheticBuffer(12, blankAmtRow)
Set g_bzhao = mock4
AssertFalse "P-line present, SALE AMT blank -> False", HasPartsCharged()

' ---- Summary ----
WScript.Echo ""
If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All " & g_Pass & " HasPartsCharged tests passed."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & g_Fail & " test(s) failed."
    WScript.Quit 1
End If
