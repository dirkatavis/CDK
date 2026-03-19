Option Explicit

Dim bzhao

Call Main()

Sub Main()
    Set bzhao = CreateObject("BZWhll.WhllObj")
    bzhao.Connect ""

    ' 1. Ensure at COMMAND: prompt on the RO detail screen
    WaitForTextAtBottom "COMMAND:"

    ' 2. Review phase: execute R A, R B, R C sequence before issuing Final Charge
    If Not ExecuteReviewSequence() Then
        Exit Sub
    End If

    ' 3. FC command (Final Charge) — only after all reviews pass
    WaitForTextAtBottom "COMMAND:"
    EnterTextAndWait "FC"
    bzhao.Pause 1000

    ' 4. ALL LABOR POSTED prompt
    WaitForTextAtBottom "ALL LABOR POSTED"
    EnterTextAndWait "Y"
    bzhao.Pause 1000

    ' 5. MILEAGE OUT prompt (accept default)
    WaitForTextAtBottom "MILEAGE OUT"
    EnterTextAndWait ""
    bzhao.Pause 1000

    ' 6. MILEAGE IN prompt
    WaitForTextAtBottom "MILEAGE IN"
    EnterTextAndWait ""
    bzhao.Pause 1000

    ' 7. OK TO CLOSE prompt
    WaitForTextAtBottom "O.K. TO CLOSE RO"
    EnterTextAndWait "Y"
    bzhao.Pause 1000

    ' 8. INVOICE PRINTER prompt
    WaitForTextAtBottom "INVOICE PRINTER"
    EnterTextAndWait "2"
    bzhao.Pause 1000
End Sub

'-----------------------------------------------------------
' Prompt object and helpers adapted from PostFinalCharges.vbs
' so review handling follows the same prompt-driven pattern.
'-----------------------------------------------------------
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

'-----------------------------------------------------------
' Executes the review sequence: R A, R B, R C in order.
' Uses DiscoverLineLetters for dynamic discovery; falls back
' to A, B, C if no letters found on screen.
' Returns True if all reviews succeed, False on any failure.
'-----------------------------------------------------------
Function ExecuteReviewSequence()
    Dim letters, i, letter, reviewCommand, prompts
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

        Call EnterTextAndWait(reviewCommand)

        Set prompts = CreateReviewPromptDictionary()
        If Not ProcessPromptSequence(prompts, 10000) Then
            Call LogErrorMessage("Review command '" & reviewCommand & "' failed. COMMAND: prompt did not return.")
            Exit Function
        End If
    Next

    ExecuteReviewSequence = True
End Function

'-----------------------------------------------------------
' Processes prompts using the same dictionary-driven model
' used in PostFinalCharges.vbs.
' Returns True when a success prompt is reached.
'-----------------------------------------------------------
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
                Call LogErrorMessage("Unknown prompt on line 23: '" & mainPromptText & "'" & vbCrLf & BuildPromptAreaSnapshot())
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

    Call LogErrorMessage("Prompt sequence timed out after " & timeoutMs & " ms." & vbCrLf & BuildPromptAreaSnapshot())
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

Function BuildPromptAreaSnapshot()
    BuildPromptAreaSnapshot = "[22] " & GetScreenLine(22) & vbCrLf & _
                              "[23] " & GetScreenLine(23) & vbCrLf & _
                              "[24] " & GetScreenLine(24)
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

Sub LogErrorMessage(message)
    On Error Resume Next
    bzhao.MsgBox "ERROR: " & message, 16
    On Error GoTo 0
End Sub

'-----------------------------------------------------------
' Discovers which line letters (A-Z) are present on the
' current RO detail screen (rows 10-22, column 1).
' Returns an array of letter strings, or an empty array.
'-----------------------------------------------------------
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

'-----------------------------------------------------------
' Waits for text to appear at bottom of screen (rows 23-24)
'-----------------------------------------------------------
Sub WaitForTextAtBottom(targetText)
    Dim elapsed, screenContentBuffer, screenLength, found, col
    elapsed = 0
    col = 1
    screenLength = 80
    
    Dim targets: targets = Split(targetText, "|")
    Dim i
    
    Do
        bzhao.Pause 500
        elapsed = elapsed + 500
        
        Dim buffer23, buffer24
        bzhao.ReadScreen buffer23, screenLength, 23, col
        bzhao.ReadScreen buffer24, screenLength, 24, col
        screenContentBuffer = UCase(buffer23 & " " & buffer24)

        found = False
        For i = 0 To UBound(targets)
            If InStr(screenContentBuffer, UCase(targets(i))) > 0 Then
                found = True
                Exit For
            End If
        Next
        
        If found Then Exit Do
        
        If elapsed >= 10000 Then
            bzhao.msgBox "ERROR: Timeout waiting for '" & targetText & "'" & vbCrLf & "Last 2 lines: " & vbCrLf & buffer23 & vbCrLf & buffer24, 16
            Exit Sub
        End If
    Loop
End Sub

'-----------------------------------------------------------
' Sends text and presses Enter
'-----------------------------------------------------------
Sub EnterTextAndWait(text)
    If text <> "" Then bzhao.SendKey text
    bzhao.Pause 100
    bzhao.SendKey "<NumpadEnter>"
    bzhao.Pause 500
End Sub
