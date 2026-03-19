'==============================================================================
' test_close_single_ro.vbs
' Tests the review-phase logic added to tools\close_single_ro.vbs:
'   - DiscoverLineLetters  : dynamic line-letter detection
'   - IsAtCommandPrompt    : COMMAND: visibility check
'   - SendReviewCommand    : single R <letter> dispatch + verify
'   - ReviewLineItems      : full review sequence with fallback
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

' g_ReviewTimeoutMs mirrors the production default; tests override to 500 ms to avoid
' 10-second busy-waits during failure/timeout test cases.
Dim g_ReviewTimeoutMs: g_ReviewTimeoutMs = 500

' ---------------------------------------------------------------------------
' Functions under test  (identical implementations to close_single_ro.vbs)
'
' They are defined inline here because close_single_ro.vbs is not an include-
' able library — it executes Call Main() at module level, making ExecuteGlobal
' unsafe in a test context.
' ---------------------------------------------------------------------------

Function IsAtCommandPrompt()
    Dim buf23, buf24
    bzhao.ReadScreen buf23, 80, 23, 1
    bzhao.ReadScreen buf24, 80, 24, 1
    IsAtCommandPrompt = (InStr(UCase(buf23 & " " & buf24), "COMMAND:") > 0)
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

Function SendReviewCommand(letter)
    Dim elapsed
    SendReviewCommand = False
    elapsed = 0

    bzhao.SendKey "R " & letter
    bzhao.Pause 100
    bzhao.SendKey "<NumpadEnter>"

    Do
        bzhao.Pause 500
        elapsed = elapsed + 500
        If IsAtCommandPrompt() Then
            SendReviewCommand = True
            Exit Do
        End If
        If elapsed >= g_ReviewTimeoutMs Then Exit Do
    Loop
End Function

Function ReviewLineItems()
    Dim letters, i, letter
    ReviewLineItems = False

    letters = DiscoverLineLetters()
    If UBound(letters) = -1 Then
        letters = Array("A", "B", "C")
    End If

    For i = 0 To UBound(letters)
        letter = letters(i)
        If Not SendReviewCommand(letter) Then
            ' In tests: do not call MsgBox; just return False.
            Exit Function
        End If
    Next

    ReviewLineItems = True
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

    Test_IsAtCommandPrompt_True()
    Test_IsAtCommandPrompt_False()
    Test_DiscoverLineLetters_FindsABC()
    Test_DiscoverLineLetters_EmptyScreen()
    Test_DiscoverLineLetters_StopsAfterGap()
    Test_SendReviewCommand_Success()
    Test_SendReviewCommand_Timeout()
    Test_ReviewLineItems_AllPass()
    Test_ReviewLineItems_FallbackToABC()
    Test_ReviewLineItems_StopsOnFirstFailure()

    WScript.Echo String(60, "-")
    WScript.Echo "Results: " & g_pass & " passed, " & g_fail & " failed"
    If g_fail > 0 Then WScript.Quit 1
End Sub

' --- IsAtCommandPrompt ---

Sub Test_IsAtCommandPrompt_True()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetBuffer MakeScreenBuffer("COMMAND:", 23, 1)
    Assert "IsAtCommandPrompt returns True when COMMAND: on row 23", IsAtCommandPrompt()
End Sub

Sub Test_IsAtCommandPrompt_False()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetBuffer String(24 * 80, " ")
    Assert "IsAtCommandPrompt returns False on blank screen", Not IsAtCommandPrompt()
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

' --- SendReviewCommand ---

Sub Test_SendReviewCommand_Success()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    ' After the first Enter, SetPromptSequence advances to show COMMAND: on row 23.
    bzhao.SetPromptSequence Array("COMMAND:")

    Dim result: result = SendReviewCommand("A")
    Dim keys: keys = bzhao.GetSentKeys()
    Assert "SendReviewCommand sends 'R A' followed by Enter", _
        InStr(keys, "R A") > 0 And InStr(keys, "<NumpadEnter>") > 0
    Assert "SendReviewCommand returns True when COMMAND: appears", result
End Sub

Sub Test_SendReviewCommand_Timeout()
    ' Screen never shows COMMAND: -> hits g_ReviewTimeoutMs (500 ms in test mode).
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetBuffer String(24 * 80, " ")

    Dim result: result = SendReviewCommand("B")
    Assert "SendReviewCommand returns False on timeout", Not result
End Sub

' --- ReviewLineItems ---

Sub Test_ReviewLineItems_AllPass()
    ' Screen has A, B on rows 10-11; COMMAND: always on row 23.
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    Dim buf: buf = MakeScreenBuffer("COMMAND:", 23, 1)
    buf = PlaceLineLetter(buf, "A", 10)
    buf = PlaceLineLetter(buf, "B", 11)
    bzhao.SetBuffer buf

    Assert "ReviewLineItems returns True when all R commands succeed", ReviewLineItems()
End Sub

Sub Test_ReviewLineItems_FallbackToABC()
    ' No line letters discovered -> falls back to A, B, C.
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetBuffer MakeScreenBuffer("COMMAND:", 23, 1)

    Dim result: result = ReviewLineItems()
    Dim keys: keys = bzhao.GetSentKeys()
    Assert "ReviewLineItems uses fallback A,B,C when no letters discovered", _
        InStr(keys, "R A") > 0 And InStr(keys, "R B") > 0 And InStr(keys, "R C") > 0
    Assert "ReviewLineItems returns True with fallback A,B,C", result
End Sub

Sub Test_ReviewLineItems_StopsOnFirstFailure()
    ' Screen has A, B, C but COMMAND: never appears -> A fails -> B never attempted.
    ' g_ReviewTimeoutMs = 500 keeps this test fast.
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    Dim buf: buf = String(24 * 80, " ")
    buf = PlaceLineLetter(buf, "A", 10)
    buf = PlaceLineLetter(buf, "B", 11)
    buf = PlaceLineLetter(buf, "C", 12)
    bzhao.SetBuffer buf

    Dim result: result = ReviewLineItems()
    Dim keys: keys = bzhao.GetSentKeys()
    Assert "ReviewLineItems returns False when first review fails", Not result
    Assert "ReviewLineItems stops after first failure (R B not sent)", InStr(keys, "R B") = 0
End Sub

RunAllTests()
