'==============================================================================
' test_close_single_ro.vbs
' Tests the review-phase logic added to tools\close_single_ro.vbs:
'   - DiscoverLineLetters           : dynamic line-letter detection
'   - IsCommandPromptVisible        : COMMAND: visibility check
'   - WaitForReviewCommandCompletion: single R <letter> dispatch + verify
'   - ExecuteReviewSequence         : full R A, R B, R C sequence
'
' Run: cscript.exe //nologo test_close_single_ro.vbs
'==============================================================================

Option Explicit

' ---------------------------------------------------------------------------
' Test harness globals
' ---------------------------------------------------------------------------
Dim g_pass, g_fail
g_pass = 0
g_fail = 0

' ---------------------------------------------------------------------------
' Load AdvancedMock to drive bzhao without BlueZone
' ---------------------------------------------------------------------------
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_repoRoot: g_repoRoot = g_fso.GetAbsolutePathName( _
    g_fso.BuildPath(g_fso.GetParentFolderName(WScript.ScriptFullName), "..\.."))

Dim mockPath: mockPath = g_fso.BuildPath(g_repoRoot, "framework\AdvancedMock.vbs")
If Not g_fso.FileExists(mockPath) Then
    WScript.Echo "FAIL: AdvancedMock.vbs not found at " & mockPath
    WScript.Quit 1
End If
ExecuteGlobal g_fso.OpenTextFile(mockPath).ReadAll

' bzhao is the global consumed by all functions under test (same name as production).
Dim bzhao

' ---------------------------------------------------------------------------
' Functions under test  (identical implementations to close_single_ro.vbs)
'
' They are defined inline here because close_single_ro.vbs is not an include-
' able library — it executes Call Main() at module level, making ExecuteGlobal
' unsafe in a test context.
' ---------------------------------------------------------------------------

Function IsCommandPromptVisible()
    Dim buf23, buf24
    bzhao.ReadScreen buf23, 80, 23, 1
    bzhao.ReadScreen buf24, 80, 24, 1
    IsCommandPromptVisible = (InStr(UCase(buf23 & " " & buf24), "COMMAND:") > 0)
End Function

Function DiscoverLineLetters()
    Dim i, capturedLetter, screenContentBuffer, nextColChar
    Dim tempLetters(25), foundCount, emptyRowCount, startReadRow
    foundCount = 0
    emptyRowCount = 0

    For startReadRow = 10 To 22
        bzhao.ReadScreen screenContentBuffer, 1, startReadRow, 1
        capturedLetter = Trim(screenContentBuffer)

        If Len(capturedLetter) = 1 Then
            If Asc(UCase(capturedLetter)) >= Asc("A") And Asc(UCase(capturedLetter)) <= Asc("Z") Then
                bzhao.ReadScreen nextColChar, 1, startReadRow, 2
                If Len(nextColChar) > 0 And Asc(nextColChar) = 32 Then
                    tempLetters(foundCount) = UCase(capturedLetter)
                    foundCount = foundCount + 1
                    emptyRowCount = 0
                Else
                    emptyRowCount = emptyRowCount + 1
                End If
            Else
                emptyRowCount = emptyRowCount + 1
            End If
        Else
            emptyRowCount = emptyRowCount + 1
        End If

        If emptyRowCount >= 3 Then Exit For
    Next

    If foundCount = 0 Then
        DiscoverLineLetters = Array()
        Exit Function
    End If

    Dim foundLetters()
    ReDim foundLetters(foundCount - 1)
    For i = 0 To foundCount - 1
        foundLetters(i) = tempLetters(i)
    Next
    DiscoverLineLetters = foundLetters
End Function

Function WaitForReviewCommandCompletion(reviewCommand)
    Dim timeoutMs: timeoutMs = 500  ' Fast timeout for tests
    Dim found: found = False
    Dim waitStart, elapsed
    
    WaitForReviewCommandCompletion = False
    waitStart = Timer

    ' Send the review command
    bzhao.SendKey reviewCommand
    bzhao.Pause 100
    bzhao.SendKey "<NumpadEnter>"
    bzhao.Pause 500

    ' Poll for COMMAND: prompt to return
    Do
        If IsCommandPromptVisible() Then
            found = True
            Exit Do
        End If
        
        bzhao.Pause 500
        elapsed = (Timer - waitStart) * 1000
        If elapsed < 0 Then elapsed = elapsed + 86400000
        
        If elapsed > timeoutMs Then
            Exit Do
        End If
    Loop

    WaitForReviewCommandCompletion = found
End Function

Function ExecuteReviewSequence()
    Dim letters, i, letter, reviewCommand
    ExecuteReviewSequence = False

    ' Discover line letters on the current RO detail screen
    letters = DiscoverLineLetters()
    If UBound(letters) = -1 Then
        ' No letters found; use fallback sequence
        letters = Array("A", "B", "C")
    End If

    ' Execute R <letter> for each discovered line
    For i = 0 To UBound(letters)
        letter = letters(i)
        reviewCommand = "R " & letter
        
        ' Send command and wait for COMMAND: to return
        If Not WaitForReviewCommandCompletion(reviewCommand) Then
            Exit Function
        End If
    Next

    ExecuteReviewSequence = True
End Function

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

Sub Assert(testName, condition)
    If condition Then
        WScript.Echo "PASS: " & testName
        g_pass = g_pass + 1
    Else
        WScript.Echo "FAIL: " & testName
        g_fail = g_fail + 1
    End If
End Sub

' Return a 24*80 blank screen with `text` placed at (row, col).
Function MakeScreenBuffer(text, row, col)
    Dim buf: buf = String(24 * 80, " ")
    Dim pos: pos = ((row - 1) * 80) + (col - 1) + 1
    buf = Left(buf, pos - 1) & text & Mid(buf, pos + Len(text))
    MakeScreenBuffer = buf
End Function

' Place line letter in an existing buffer string: letter at (row,1), space at (row,2).
Function PlaceLineLetter(buf, letter, row)
    Dim pos: pos = ((row - 1) * 80) + 1
    PlaceLineLetter = Left(buf, pos - 1) & letter & " " & Mid(buf, pos + 2)
End Function

' ---------------------------------------------------------------------------
' Tests
' ---------------------------------------------------------------------------

Sub RunAllTests()
    WScript.Echo "Running close_single_ro review-phase tests..."
    WScript.Echo String(60, "-")

    Test_IsCommandPromptVisible_True()
    Test_IsCommandPromptVisible_False()
    Test_DiscoverLineLetters_FindsABC()
    Test_DiscoverLineLetters_EmptyScreen()
    Test_DiscoverLineLetters_StopsAfterGap()
    Test_WaitForReviewCommandCompletion_Success()
    Test_WaitForReviewCommandCompletion_Timeout()
    Test_ExecuteReviewSequence_AllPass()
    Test_ExecuteReviewSequence_FallbackToABC()
    Test_ExecuteReviewSequence_StopsOnFirstFailure()

    WScript.Echo String(60, "-")
    WScript.Echo "Results: " & g_pass & " passed, " & g_fail & " failed"
    If g_fail > 0 Then WScript.Quit 1
End Sub

' --- IsCommandPromptVisible ---

Sub Test_IsCommandPromptVisible_True()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetBuffer MakeScreenBuffer("COMMAND:", 23, 1)
    Assert "IsCommandPromptVisible returns True when COMMAND: on row 23", IsCommandPromptVisible()
End Sub

Sub Test_IsCommandPromptVisible_False()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetBuffer String(24 * 80, " ")
    Assert "IsCommandPromptVisible returns False on blank screen", Not IsCommandPromptVisible()
End Sub

' --- DiscoverLineLetters ---

Sub Test_DiscoverLineLetters_FindsABC()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    Dim buf: buf = String(24 * 80, " ")
    buf = PlaceLineLetter(buf, "A", 10)
    buf = PlaceLineLetter(buf, "B", 11)
    buf = PlaceLineLetter(buf, "C", 12)
    bzhao.SetBuffer buf

    Dim letters: letters = DiscoverLineLetters()
    Assert "DiscoverLineLetters finds A, B, C", _
        UBound(letters) = 2 And letters(0) = "A" And letters(1) = "B" And letters(2) = "C"
End Sub

Sub Test_DiscoverLineLetters_EmptyScreen()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetBuffer String(24 * 80, " ")
    Dim letters: letters = DiscoverLineLetters()
    Assert "DiscoverLineLetters returns empty array on blank screen", UBound(letters) = -1
End Sub

Sub Test_DiscoverLineLetters_StopsAfterGap()
    ' A at row 10, then 3 blank rows (11-13), then D at row 14 -> D should NOT be found.
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    Dim buf: buf = String(24 * 80, " ")
    buf = PlaceLineLetter(buf, "A", 10)
    buf = PlaceLineLetter(buf, "D", 14)
    bzhao.SetBuffer buf

    Dim letters: letters = DiscoverLineLetters()
    Assert "DiscoverLineLetters stops after 3 consecutive blank rows", _
        UBound(letters) = 0 And letters(0) = "A"
End Sub

' --- WaitForReviewCommandCompletion ---

Sub Test_WaitForReviewCommandCompletion_Success()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    ' After the first Enter, SetPromptSequence advances to show COMMAND: on row 23.
    bzhao.SetPromptSequence Array("COMMAND:")

    Dim result: result = WaitForReviewCommandCompletion("R A")
    Dim keys: keys = bzhao.GetSentKeys()
    Assert "WaitForReviewCommandCompletion sends 'R A' followed by Enter", _
        InStr(keys, "R A") > 0 And InStr(keys, "<NumpadEnter>") > 0
    Assert "WaitForReviewCommandCompletion returns True when COMMAND: appears", result
End Sub

Sub Test_WaitForReviewCommandCompletion_Timeout()
    ' Screen never shows COMMAND: -> hits timeout (500 ms in test mode).
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetBuffer String(24 * 80, " ")

    Dim result: result = WaitForReviewCommandCompletion("R B")
    Assert "WaitForReviewCommandCompletion returns False on timeout", Not result
End Sub

' --- ExecuteReviewSequence ---

Sub Test_ExecuteReviewSequence_AllPass()
    ' Screen has A, B on rows 10-11; COMMAND: always on row 23.
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    Dim buf: buf = MakeScreenBuffer("COMMAND:", 23, 1)
    buf = PlaceLineLetter(buf, "A", 10)
    buf = PlaceLineLetter(buf, "B", 11)
    bzhao.SetBuffer buf

    Assert "ExecuteReviewSequence returns True when all R commands succeed", ExecuteReviewSequence()
End Sub

Sub Test_ExecuteReviewSequence_FallbackToABC()
    ' No line letters discovered -> falls back to A, B, C.
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetBuffer MakeScreenBuffer("COMMAND:", 23, 1)

    Dim result: result = ExecuteReviewSequence()
    Dim keys: keys = bzhao.GetSentKeys()
    Assert "ExecuteReviewSequence uses fallback A,B,C when no letters discovered", _
        InStr(keys, "R A") > 0 And InStr(keys, "R B") > 0 And InStr(keys, "R C") > 0
    Assert "ExecuteReviewSequence returns True with fallback A,B,C", result
End Sub

Sub Test_ExecuteReviewSequence_StopsOnFirstFailure()
    ' Screen has A, B, C but COMMAND: never appears -> A fails -> B never attempted.
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    Dim buf: buf = String(24 * 80, " ")
    buf = PlaceLineLetter(buf, "A", 10)
    buf = PlaceLineLetter(buf, "B", 11)
    buf = PlaceLineLetter(buf, "C", 12)
    bzhao.SetBuffer buf

    Dim result: result = ExecuteReviewSequence()
    Dim keys: keys = bzhao.GetSentKeys()
    Assert "ExecuteReviewSequence returns False when first review fails", Not result
    Assert "ExecuteReviewSequence stops after first failure (R B not sent)", InStr(keys, "R B") = 0
End Sub

RunAllTests()
