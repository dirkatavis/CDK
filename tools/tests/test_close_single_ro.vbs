'==============================================================================
' test_close_single_ro.vbs
' Tests the prompt-driven review phase in tools\close_single_ro.vbs.
'==============================================================================

Option Explicit

Dim g_pass, g_fail
g_pass = 0
g_fail = 0

Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_repoRoot: g_repoRoot = g_fso.GetAbsolutePathName(g_fso.BuildPath(g_fso.GetParentFolderName(WScript.ScriptFullName), "..\.."))

Dim mockPath: mockPath = g_fso.BuildPath(g_repoRoot, "framework\AdvancedMock.vbs")
If Not g_fso.FileExists(mockPath) Then
    WScript.Echo "FAIL: AdvancedMock.vbs not found at " & mockPath
    WScript.Quit 1
End If
ExecuteGlobal g_fso.OpenTextFile(mockPath).ReadAll

Dim bzhao

Class Prompt
    Public TriggerText
    Public ResponseText
    Public KeyPress
    Public IsSuccess
    Public AcceptDefault
    Public IsRegex
End Class

Function InferRegexPattern(pattern)
    InferRegexPattern = False
    If Left(pattern, 1) = "^" Or InStr(pattern, "(") > 0 Or InStr(pattern, "[") > 0 Or InStr(pattern, ".*") > 0 Or InStr(pattern, "\") > 0 Then
        InferRegexPattern = True
    End If
End Function

Sub AddPromptToDict(dict, trigger, response, key, isSuccess)
    Dim prompt
    Set prompt = New Prompt
    prompt.TriggerText = trigger
    prompt.ResponseText = response
    prompt.KeyPress = key
    prompt.IsSuccess = isSuccess
    prompt.AcceptDefault = False
    prompt.IsRegex = InferRegexPattern(trigger)
    dict.Add trigger, prompt
End Sub

Sub AddPromptToDictEx(dict, trigger, response, key, isSuccess, acceptDefault)
    Dim prompt
    Set prompt = New Prompt
    prompt.TriggerText = trigger
    prompt.ResponseText = response
    prompt.KeyPress = key
    prompt.IsSuccess = isSuccess
    prompt.AcceptDefault = acceptDefault
    prompt.IsRegex = InferRegexPattern(trigger)
    dict.Add trigger, prompt
End Sub

Function CreateReviewPromptDictionary()
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")
    Call AddPromptToDict(dict, "COMMAND:", "", "", True)
    Call AddPromptToDictEx(dict, "OPERATION CODE FOR LINE.*(\(.*\))?\?", "I", "<NumpadEnter>", False, True)
    Call AddPromptToDict(dict, "OPERATION CODE FOR LINE[^\(]*\?", "I", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "LABOR TYPE FOR LINE", "", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "DESC:", "", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "Enter a technician number", "", "<F3>", False)
    Call AddPromptToDictEx(dict, "TECHNICIAN \([A-Za-z0-9]+\)\?", "99", "<NumpadEnter>", False, True)
    Call AddPromptToDictEx(dict, "TECHNICIAN\?", "99", "<NumpadEnter>", False, True)
    Call AddPromptToDictEx(dict, "TECHNICIAN\s*\?", "99", "<NumpadEnter>", False, True)
    Call AddPromptToDict(dict, "TECHNICIAN FINISHING WORK", "99", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "IS ASSIGNED TO LINE", "Y", "<NumpadEnter>", False)
    Call AddPromptToDictEx(dict, "TECHNICIAN \(Y/N\)", "Y", "<NumpadEnter>", True, False)
    Call AddPromptToDictEx(dict, "ACTUAL HOURS \(\d+\)", "0", "<NumpadEnter>", False, True)
    Call AddPromptToDictEx(dict, "SOLD HOURS( \(\d+\))?\?", "0", "<NumpadEnter>", False, True)
    Call AddPromptToDict(dict, "ADD A LABOR OPERATION( \(N\)\?)?", "N", "<NumpadEnter>", True)
    Call AddPromptToDict(dict, "PRESS RETURN TO CONTINUE", "", "<Enter>", False)
    Call AddPromptToDict(dict, "Press F3 to exit.", "", "<F3>", False)
    Set CreateReviewPromptDictionary = dict
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

Function ProcessPromptSequence(prompts, timeoutMs)
    Dim startTime, elapsed, promptKey, lineToCheck, lineText
    Dim bestMatchKey, bestMatchLength, promptDetails, mainPromptText, bestMatchLineText

    If timeoutMs <= 0 Then timeoutMs = 10000

    ProcessPromptSequence = False
    startTime = Timer

    Do
        mainPromptText = GetScreenLine(23)
        If Len(mainPromptText) > 0 Then
            If Not IsPromptInConfig(mainPromptText, prompts) Then
                Exit Function
            End If
        End If

        bestMatchKey = ""
        bestMatchLength = 0
        bestMatchLineText = ""

        For Each lineToCheck In Array(23, 22, 24, 21, 20)
            lineText = GetScreenLine(lineToCheck)
            If Len(lineText) > 0 Then
                For Each promptKey In prompts.Keys
                    Set promptDetails = prompts.Item(promptKey)
                    If IsPromptMatch(lineText, promptKey, promptDetails.IsRegex) Then
                        If Len(promptKey) > bestMatchLength Then
                            bestMatchKey = promptKey
                            bestMatchLength = Len(promptKey)
                            bestMatchLineText = lineText
                        End If
                    End If
                Next
            End If
        Next

        If bestMatchLength > 0 Then
            Set promptDetails = prompts.Item(bestMatchKey)

            If promptDetails.ResponseText <> "" And Not (HasDefaultValueInPrompt(bestMatchLineText) And Not IsYesNoPrompt(bestMatchLineText)) Then
                bzhao.SendKey promptDetails.ResponseText
                bzhao.Pause 100
            End If

            If promptDetails.KeyPress <> "" Then
                bzhao.SendKey promptDetails.KeyPress
                bzhao.Pause 500
            End If

            If promptDetails.IsSuccess Then
                ProcessPromptSequence = True
                Exit Function
            End If
        Else
            If InStr(1, mainPromptText, "COMMAND:", vbTextCompare) = 1 Then
                ProcessPromptSequence = True
                Exit Function
            End If
            bzhao.Pause 250
        End If

        elapsed = (Timer - startTime) * 1000
        If elapsed < 0 Then elapsed = elapsed + 86400000
        If elapsed > timeoutMs Then Exit Do
    Loop
End Function

Function ExecuteReviewSequence()
    Dim letters, i, letter, reviewCommand, prompts
    ExecuteReviewSequence = False

    letters = DiscoverLineLetters()
    If UBound(letters) = -1 Then
        letters = Array("A", "B", "C")
    End If

    For i = 0 To UBound(letters)
        letter = letters(i)
        reviewCommand = "R " & letter
        Call EnterTextAndWait(reviewCommand)
        Set prompts = CreateReviewPromptDictionary()
        If Not ProcessPromptSequence(prompts, 1000) Then
            Exit Function
        End If
    Next

    ExecuteReviewSequence = True
End Function

Function IsPromptMatch(screenLine, triggerText, isRegex)
    IsPromptMatch = False

    If InStr(1, screenLine, "COMMAND:", vbTextCompare) = 1 And InStr(1, triggerText, "COMMAND:", vbTextCompare) = 1 Then
        IsPromptMatch = True
        Exit Function
    End If

    If isRegex Then
        Dim re
        On Error Resume Next
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = triggerText
        re.IgnoreCase = True
        re.Global = False
        If Err.Number = 0 Then
            IsPromptMatch = re.Test(screenLine)
        Else
            Err.Clear
        End If
        On Error GoTo 0
    Else
        IsPromptMatch = (InStr(1, screenLine, triggerText, vbTextCompare) > 0)
    End If
End Function

Function IsPromptInConfig(promptText, promptsDict)
    Dim key

    If InStr(1, promptText, "COMMAND:", vbTextCompare) = 1 Then
        For Each key In promptsDict.Keys
            If InStr(1, key, "COMMAND:", vbTextCompare) = 1 Then
                IsPromptInConfig = True
                Exit Function
            End If
        Next
    End If

    For Each key In promptsDict.Keys
        If IsPromptMatch(promptText, key, InferRegexPattern(key)) Then
            IsPromptInConfig = True
            Exit Function
        End If
    Next

    IsPromptInConfig = False
End Function

Function GetScreenLine(rowNumber)
    Dim buffer
    buffer = ""
    bzhao.ReadScreen buffer, 80, rowNumber, 1
    GetScreenLine = Trim(buffer)
End Function

Function HasDefaultValueInPrompt(promptText)
    Dim re
    HasDefaultValueInPrompt = False

    On Error Resume Next
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\([^\)]*\)\?"
    re.IgnoreCase = True
    re.Global = False
    If Err.Number = 0 Then
        HasDefaultValueInPrompt = re.Test(promptText)
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Function

Function IsYesNoPrompt(promptText)
    IsYesNoPrompt = (InStr(1, promptText, "(Y/N)", vbTextCompare) > 0)
End Function

Sub EnterTextAndWait(text)
    If text <> "" Then bzhao.SendKey text
    bzhao.Pause 100
    bzhao.SendKey "<NumpadEnter>"
    bzhao.Pause 500
End Sub

Sub Assert(testName, condition)
    If condition Then
        WScript.Echo "PASS: " & testName
        g_pass = g_pass + 1
    Else
        WScript.Echo "FAIL: " & testName
        g_fail = g_fail + 1
    End If
End Sub

Function MakeScreenBuffer(text, row, col)
    Dim buf: buf = String(24 * 80, " ")
    Dim pos: pos = ((row - 1) * 80) + (col - 1) + 1
    buf = Left(buf, pos - 1) & text & Mid(buf, pos + Len(text))
    MakeScreenBuffer = buf
End Function

Function PlaceLineLetter(buf, letter, row)
    Dim pos: pos = ((row - 1) * 80) + 1
    PlaceLineLetter = Left(buf, pos - 1) & letter & " " & Mid(buf, pos + 2)
End Function

Sub RunAllTests()
    WScript.Echo "Running close_single_ro review-phase tests..."
    WScript.Echo String(60, "-")

    Test_DiscoverLineLetters_FindsABC()
    Test_ProcessPromptSequence_SucceedsOnCommandPrompt()
    Test_ProcessPromptSequence_HandlesOperationCodePrompt()
    Test_ProcessPromptSequence_HandlesLaborTypePrompt()
    Test_ProcessPromptSequence_AcceptsTechnicianDefault()
    Test_ProcessPromptSequence_Sends99WhenTechnicianHasNoDefault()
    Test_ProcessPromptSequence_OverridesYesNoPrompt()
    Test_ProcessPromptSequence_HandlesReturnPrompt()
    Test_ProcessPromptSequence_FailsOnUnknownPrompt()
    Test_ExecuteReviewSequence_AllPass()
    Test_ExecuteReviewSequence_FallbackToABC()
    Test_ExecuteReviewSequence_StopsOnUnknownPrompt()

    WScript.Echo String(60, "-")
    WScript.Echo "Results: " & g_pass & " passed, " & g_fail & " failed"
    If g_fail > 0 Then WScript.Quit 1
End Sub

Sub Test_DiscoverLineLetters_FindsABC()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    Dim buf: buf = String(24 * 80, " ")
    buf = PlaceLineLetter(buf, "A", 10)
    buf = PlaceLineLetter(buf, "B", 11)
    buf = PlaceLineLetter(buf, "C", 12)
    bzhao.SetBuffer buf

    Dim letters: letters = DiscoverLineLetters()
    Assert "DiscoverLineLetters finds A, B, C", UBound(letters) = 2 And letters(0) = "A" And letters(1) = "B" And letters(2) = "C"
End Sub

Sub Test_ProcessPromptSequence_SucceedsOnCommandPrompt()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetBuffer MakeScreenBuffer("COMMAND: R A", 23, 1)
    Assert "ProcessPromptSequence accepts COMMAND-prefixed success prompt", ProcessPromptSequence(CreateReviewPromptDictionary(), 1000)
End Sub

Sub Test_ProcessPromptSequence_HandlesOperationCodePrompt()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetPromptSequence Array("OPERATION CODE FOR LINE A, L1 (PMSCRT)?", "COMMAND:")
    Dim result: result = ProcessPromptSequence(CreateReviewPromptDictionary(), 1000)
    Dim keys: keys = bzhao.GetSentKeys()
    Assert "ProcessPromptSequence handles OPERATION CODE prompt", result
    Assert "ProcessPromptSequence accepts OPERATION CODE default", InStr(keys, "I") = 0
End Sub

Sub Test_ProcessPromptSequence_HandlesLaborTypePrompt()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetPromptSequence Array("LABOR TYPE FOR LINE A, L1 (I)?", "COMMAND:")
    Dim result: result = ProcessPromptSequence(CreateReviewPromptDictionary(), 1000)
    Dim keys: keys = bzhao.GetSentKeys()
    Assert "ProcessPromptSequence handles LABOR TYPE FOR LINE prompt", result
    Assert "ProcessPromptSequence sends Enter for LABOR TYPE prompt", InStr(keys, "<NumpadEnter>") > 0
End Sub

Sub Test_ProcessPromptSequence_AcceptsTechnicianDefault()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetPromptSequence Array("TECHNICIAN (72925)?", "COMMAND:")
    Dim result: result = ProcessPromptSequence(CreateReviewPromptDictionary(), 1000)
    Dim keys: keys = bzhao.GetSentKeys()
    Assert "ProcessPromptSequence accepts technician default", result
    Assert "ProcessPromptSequence does not send override when default exists", InStr(keys, "99") = 0
End Sub

Sub Test_ProcessPromptSequence_Sends99WhenTechnicianHasNoDefault()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetPromptSequence Array("TECHNICIAN ?", "COMMAND:")
    Dim result: result = ProcessPromptSequence(CreateReviewPromptDictionary(), 1000)
    Dim keys: keys = bzhao.GetSentKeys()
    Assert "ProcessPromptSequence handles TECHNICIAN no-default prompt", result
    Assert "ProcessPromptSequence sends 99 when TECHNICIAN has no default", InStr(keys, "99") > 0
End Sub

Sub Test_ProcessPromptSequence_OverridesYesNoPrompt()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetPromptSequence Array("TECHNICIAN (Y/N)", "COMMAND:")
    Dim result: result = ProcessPromptSequence(CreateReviewPromptDictionary(), 1000)
    Dim keys: keys = bzhao.GetSentKeys()
    Assert "ProcessPromptSequence handles (Y/N) prompt", result
    Assert "ProcessPromptSequence sends Y for (Y/N) prompt", InStr(keys, "Y") > 0
End Sub

Sub Test_ProcessPromptSequence_HandlesReturnPrompt()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetPromptSequence Array("PRESS RETURN TO CONTINUE", "COMMAND:")
    Assert "ProcessPromptSequence handles PRESS RETURN TO CONTINUE then COMMAND", ProcessPromptSequence(CreateReviewPromptDictionary(), 1000)
End Sub

Sub Test_ProcessPromptSequence_FailsOnUnknownPrompt()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetBuffer MakeScreenBuffer("UNEXPECTED REVIEW PROMPT", 23, 1)
    Assert "ProcessPromptSequence fails immediately on unknown prompt", Not ProcessPromptSequence(CreateReviewPromptDictionary(), 1000)
End Sub

Sub Test_ExecuteReviewSequence_AllPass()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    Dim buf: buf = MakeScreenBuffer("COMMAND:", 23, 1)
    buf = PlaceLineLetter(buf, "A", 10)
    buf = PlaceLineLetter(buf, "B", 11)
    bzhao.SetBuffer buf
    Assert "ExecuteReviewSequence returns True when COMMAND is already visible", ExecuteReviewSequence()
End Sub

Sub Test_ExecuteReviewSequence_FallbackToABC()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    bzhao.SetBuffer MakeScreenBuffer("COMMAND:", 23, 1)
    Dim result: result = ExecuteReviewSequence()
    Dim keys: keys = bzhao.GetSentKeys()
    Assert "ExecuteReviewSequence uses fallback A,B,C when no letters discovered", InStr(keys, "R A") > 0 And InStr(keys, "R B") > 0 And InStr(keys, "R C") > 0
    Assert "ExecuteReviewSequence returns True with fallback A,B,C", result
End Sub

Sub Test_ExecuteReviewSequence_StopsOnUnknownPrompt()
    Set bzhao = New AdvancedMock
    bzhao.Connect ""
    Dim buf: buf = MakeScreenBuffer("UNEXPECTED REVIEW PROMPT", 23, 1)
    buf = PlaceLineLetter(buf, "A", 10)
    buf = PlaceLineLetter(buf, "B", 11)
    bzhao.SetBuffer buf
    Dim result: result = ExecuteReviewSequence()
    Dim keys: keys = bzhao.GetSentKeys()
    Assert "ExecuteReviewSequence returns False on unknown prompt", Not result
    Assert "ExecuteReviewSequence stops after first failed review", InStr(keys, "R B") = 0
End Sub

RunAllTests()