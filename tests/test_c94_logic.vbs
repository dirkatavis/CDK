'=====================================================================================
' Test C94 Report Analyzer Logic
' Purpose: Verify GetLineDescription, TruncateAtDoubleSpace, LoadTargetROs,
'          and AllFound using AdvancedMock — no live terminal required.
' Usage:   cscript.exe tests\test_c94_logic.vbs
'=====================================================================================

Option Explicit

' --- Bootstrap ---
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim testsDir: testsDir = fso.GetParentFolderName(WScript.ScriptFullName)
Dim repoRoot: repoRoot = fso.GetParentFolderName(testsDir)
Dim mockPath: mockPath = fso.BuildPath(repoRoot, "framework\AdvancedMock.vbs")
ExecuteGlobal fso.OpenTextFile(mockPath).ReadAll

Dim bzhao: Set bzhao = New AdvancedMock
bzhao.Connect "A"

Dim g_testsPassed: g_testsPassed = 0
Dim g_testsFailed: g_testsFailed = 0

' INPUT_CSV_PATH used by LoadTargetROs — overridden per test
Dim INPUT_CSV_PATH: INPUT_CSV_PATH = ""


' --- Helpers ---

Sub SetRow(grid, rowIdx, content)
    grid(rowIdx) = Left(content & String(80, " "), 80)
End Sub

Function BlankGrid()
    Dim g(23), i
    For i = 0 To 23
        g(i) = String(80, " ")
    Next
    BlankGrid = g
End Function

Sub Assert(testName, actual, expected)
    If actual = expected Then
        WScript.Echo "[PASS] " & testName
        g_testsPassed = g_testsPassed + 1
    Else
        WScript.Echo "[FAIL] " & testName
        WScript.Echo "       Expected: '" & expected & "'"
        WScript.Echo "       Actual:   '" & actual & "'"
        g_testsFailed = g_testsFailed + 1
    End If
End Sub

Sub AssertTrue(testName, condition)
    If condition Then
        WScript.Echo "[PASS] " & testName
        g_testsPassed = g_testsPassed + 1
    Else
        WScript.Echo "[FAIL] " & testName
        g_testsFailed = g_testsFailed + 1
    End If
End Sub


' --- Functions under test (copied from C94ReportAnalyzer.vbs) ---

Function TruncateAtDoubleSpace(text)
    Dim k
    For k = 1 To Len(text) - 1
        If Mid(text, k, 2) = "  " Then
            TruncateAtDoubleSpace = Left(text, k - 1)
            Exit Function
        End If
    Next
    TruncateAtDoubleSpace = text
End Function

Function GetLineDescription(letter)
    Dim row, buf, nextColChar, foundText
    GetLineDescription = ""
    For row = 10 To 22
        bzhao.ReadScreen buf, 1, row, 1
        If UCase(Trim(buf)) = UCase(letter) Then
            bzhao.ReadScreen nextColChar, 1, row, 2
            If Asc(nextColChar) = 32 Then
                bzhao.ReadScreen foundText, 100, row, 4
                GetLineDescription = Left(TruncateAtDoubleSpace(Trim(foundText)), 100)
                Exit Function
            End If
        End If
    Next
End Function

Function LoadTargetROs()
    Dim dict, f, line
    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set f = fso.OpenTextFile(INPUT_CSV_PATH, 1)
    If Err.Number <> 0 Then
        Set LoadTargetROs = dict
        Exit Function
    End If
    On Error GoTo 0
    Do Until f.AtEndOfStream
        line = Trim(f.ReadLine())
        If line <> "" And IsNumeric(line) Then
            dict(line) = False
        End If
    Loop
    f.Close
    Set LoadTargetROs = dict
End Function

Function AllFound(dict)
    Dim key
    AllFound = True
    For Each key In dict.Keys
        If Not dict(key) Then
            AllFound = False
            Exit Function
        End If
    Next
End Function


' --- Tests ---

Sub TestDoubleSpaceTruncation()
    WScript.Echo ""
    WScript.Echo "== TruncateAtDoubleSpace =="
    Assert "T0a: No double-space — returns full string", _
        TruncateAtDoubleSpace("REPLACE BRAKE PADS"), "REPLACE BRAKE PADS"
    Assert "T0b: Double-space present — stops before it", _
        TruncateAtDoubleSpace("REPLACE BRAKE PADS  TECH1234"), "REPLACE BRAKE PADS"
    Assert "T0c: Double-space at start — returns empty", _
        TruncateAtDoubleSpace("  TECH1234"), ""
End Sub

Sub TestGetLineDescription()
    WScript.Echo ""
    WScript.Echo "== GetLineDescription =="

    ' T1: Line A with double-space stopping before TECH column
    Dim g1: g1 = BlankGrid()
    SetRow g1, 11, "A  REPLACE BRAKE PADS  TECH1234"
    bzhao.SetBuffer Join(g1, "")
    Assert "T1: Line A stops before double-space", _
        GetLineDescription("A"), "REPLACE BRAKE PADS"

    ' T2: Line A description longer than 100 chars (no double-space)
    Dim g2: g2 = BlankGrid()
    Dim longDesc: longDesc = String(110, "X")
    SetRow g2, 11, "A  " & longDesc
    bzhao.SetBuffer Join(g2, "")
    Dim result2: result2 = GetLineDescription("A")
    AssertTrue "T2: Line A truncated to 100 chars", Len(result2) = 100

    ' T3: Line C absent — returns empty string
    Dim g3: g3 = BlankGrid()
    SetRow g3, 11, "A  OIL CHANGE"
    SetRow g3, 13, "B  TIRE ROTATION"
    bzhao.SetBuffer Join(g3, "")
    Assert "T3: Line C absent returns empty", GetLineDescription("C"), ""

    ' T4: Line B present
    bzhao.SetBuffer Join(g3, "")
    Assert "T4: Line B returns correct text", GetLineDescription("B"), "TIRE ROTATION"
End Sub

Sub TestLoadTargetROs()
    WScript.Echo ""
    WScript.Echo "== LoadTargetROs =="

    ' Write a temp input file
    Dim tmpPath: tmpPath = fso.BuildPath(fso.GetSpecialFolder(2), "test_c94_input.csv")
    Dim f: Set f = fso.CreateTextFile(tmpPath, True)
    f.WriteLine ""          ' blank line — should be skipped
    f.WriteLine "866409"
    f.WriteLine "871814"
    f.WriteLine "875001"
    f.Close

    INPUT_CSV_PATH = tmpPath
    Dim dict: Set dict = LoadTargetROs()

    AssertTrue "T5: LoadTargetROs loads 3 ROs", dict.Count = 3
    AssertTrue "T6: All values initialised to False", _
        (Not dict("866409")) And (Not dict("871814")) And (Not dict("875001"))

    ' Clean up temp file
    fso.DeleteFile tmpPath
End Sub

Sub TestAllFound()
    WScript.Echo ""
    WScript.Echo "== AllFound =="

    Dim dict: Set dict = CreateObject("Scripting.Dictionary")
    dict("866409") = False
    dict("871814") = False
    dict("875001") = False

    AssertTrue "T7: AllFound returns False when none found", Not AllFound(dict)

    dict("866409") = True
    dict("871814") = True
    AssertTrue "T8: AllFound returns False when partially found", Not AllFound(dict)

    dict("875001") = True
    AssertTrue "T9: AllFound returns True when all found", AllFound(dict)
End Sub


' --- Run all tests ---
WScript.Echo "C94 Report Analyzer — Logic Tests"
WScript.Echo "==================================="

TestDoubleSpaceTruncation
TestGetLineDescription
TestLoadTargetROs
TestAllFound

WScript.Echo ""
WScript.Echo "==================================="
WScript.Echo "Total:  " & (g_testsPassed + g_testsFailed)
WScript.Echo "Passed: " & g_testsPassed
WScript.Echo "Failed: " & g_testsFailed

If g_testsFailed > 0 Then
    WScript.Echo "RESULT: FAIL"
    WScript.Quit 1
Else
    WScript.Echo "RESULT: PASS"
End If
