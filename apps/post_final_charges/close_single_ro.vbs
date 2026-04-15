Option Explicit

Dim bzhao
Const REVIEW_PAUSE = 500

Call Main()

Sub Main()
    Set bzhao = CreateObject("BZWhll.WhllObj")
    bzhao.Connect ""

    ' 1. Ensure at COMMAND: prompt on the RO detail screen
    If Not WaitForTextAtBottom("COMMAND:") Then Exit Sub

    ' 2. Review phase: execute R A, R B, R C sequence before issuing Final Charge
    If Not ExecuteReviewSequence() Then
        Exit Sub
    End If

    ' 3. FC command (Final Charge) — only after all reviews pass
    If Not WaitForTextAtBottom("COMMAND:") Then Exit Sub
    EnterTextAndWait "FC"
    bzhao.Pause 1000

    ' 4. ALL LABOR POSTED prompt
    If Not WaitForTextAtBottom("ALL LABOR POSTED") Then Exit Sub
    EnterTextAndWait "Y"
    bzhao.Pause 1000

    ' 5. MILEAGE OUT prompt (accept default)
    If Not WaitForTextAtBottom("MILEAGE OUT") Then Exit Sub
    EnterTextAndWait ""
    bzhao.Pause 1000

    ' 6. MILEAGE IN prompt
    If Not WaitForTextAtBottom("MILEAGE IN") Then Exit Sub
    EnterTextAndWait ""
    bzhao.Pause 1000

    ' 7. OK TO CLOSE prompt
    If Not WaitForTextAtBottom("O.K. TO CLOSE RO") Then Exit Sub
    EnterTextAndWait "Y"
    bzhao.Pause 1000

    ' 8. INVOICE PRINTER prompt
    If Not WaitForTextAtBottom("INVOICE PRINTER") Then Exit Sub
    EnterTextAndWait "2"
    bzhao.Pause 1000
End Sub

Function DiscoverLineLetters()
    Dim i, capturedLetter, screenContentBuffer, readLength
    Dim foundLetters, foundCount
    Dim startReadRow, startReadColumn, emptyRowCount, nextColChar
    Dim startRow, endRow
    Dim tempLetters(25)

    foundCount = 0
    emptyRowCount = 0
    startRow = 10
    endRow = 22

    For startReadRow = startRow To endRow
        startReadColumn = 1
        readLength = 1

        On Error Resume Next
        bzhao.ReadScreen screenContentBuffer, readLength, startReadRow, startReadColumn
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        On Error GoTo 0

        capturedLetter = Trim(screenContentBuffer)
        If Len(capturedLetter) = 1 Then
            If Asc(UCase(capturedLetter)) >= Asc("A") And Asc(UCase(capturedLetter)) <= Asc("Z") Then
                nextColChar = ""
                On Error Resume Next
                bzhao.ReadScreen nextColChar, 1, startReadRow, startReadColumn + 1
                If Err.Number <> 0 Then
                    Err.Clear
                    nextColChar = ""
                End If
                On Error GoTo 0

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

    ReDim foundLetters(foundCount - 1)
    For i = 0 To foundCount - 1
        foundLetters(i) = tempLetters(i)
    Next

    DiscoverLineLetters = foundLetters
End Function

Function ProcessRoReview()
    Dim lineLetters, i, lineLetter, missingLineCount, reviewedLineCount
    Dim highestDiscoveredLetter, discoveredCount
    Dim screenContent

    ProcessRoReview = True
    lineLetters = DiscoverLineLetters()

    If UBound(lineLetters) = -1 Then Exit Function

    missingLineCount = 0
    reviewedLineCount = 0
    discoveredCount = UBound(lineLetters) + 1
    highestDiscoveredLetter = lineLetters(UBound(lineLetters))

    For i = 65 To 90
        lineLetter = Chr(i)
        If Not WaitForTextAtBottom("COMMAND:") Then
            ProcessRoReview = False
            Exit Function
        End If
        bzhao.SendKey "R " & lineLetter
        bzhao.Pause 100
        bzhao.SendKey "<NumpadEnter>"
        bzhao.Pause REVIEW_PAUSE

        screenContent = ""
        bzhao.ReadScreen screenContent, 1920, 1, 1
        If InStr(1, screenContent, "LINE CODE " & lineLetter & " IS NOT ON FILE", vbTextCompare) > 0 Then
            bzhao.SendKey "<Enter>"
            bzhao.Pause REVIEW_PAUSE
            missingLineCount = missingLineCount + 1

            If reviewedLineCount > 0 Then
                If discoveredCount >= 2 Then
                    If Asc(lineLetter) > Asc(highestDiscoveredLetter) Then Exit For
                ElseIf missingLineCount >= 3 Then
                    Exit For
                End If
            End If
        Else
            missingLineCount = 0
            reviewedLineCount = reviewedLineCount + 1

            If Not HandleReviewPrompts(lineLetter) Then
                ProcessRoReview = False
                Exit Function
            End If
        End If
    Next
End Function

Function HandleReviewPrompts(lineLetter)
    Dim screenContent, startTime, elapsed, regEx

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = False

    startTime = Timer

    Do
        bzhao.Pause REVIEW_PAUSE
        bzhao.ReadScreen screenContent, 1920, 1, 1

        If TestPrompt(regEx, screenContent, "LINE\s+" & lineLetter & "\s+IS\s+NOT\s+FINISHED.*REVIEW IT ANYWAY\s*\(Y/N\)") Then
            bzhao.SendKey "<Enter>"
            bzhao.Pause REVIEW_PAUSE

            If Not FinishLineForReview(lineLetter) Then
                HandleReviewPrompts = False
                Exit Function
            End If

            If Not WaitForTextAtBottom("COMMAND:") Then
                HandleReviewPrompts = False
                Exit Function
            End If
            bzhao.SendKey "R " & lineLetter
            bzhao.Pause 100
            bzhao.SendKey "<NumpadEnter>"
            bzhao.Pause REVIEW_PAUSE
        ElseIf IsIdleCommandPrompt() Then
            HandleReviewPrompts = True
            Exit Function
        End If

        If TestPrompt(regEx, screenContent, "PRESS RETURN TO CONTINUE") Then
            bzhao.SendKey "<Enter>"
            bzhao.Pause REVIEW_PAUSE
        ElseIf TestPrompt(regEx, screenContent, "Press F3 to exit\.?$") Then
            bzhao.SendKey "<F3>"
            bzhao.Pause REVIEW_PAUSE
        ElseIf TestPrompt(regEx, screenContent, "LABOR TYPE FOR LINE|LABOR TYPE") Then
            EnterReviewPrompt ""
        ElseIf TestPrompt(regEx, screenContent, "OP CODE.*\([A-Za-z0-9]+\)\?|OPERATION CODE.*\([A-Za-z0-9]+\)\?") Then
            EnterReviewPrompt ""
        ElseIf TestPrompt(regEx, screenContent, "OP CODE.*\?|OPERATION CODE.*\?") Then
            EnterReviewPrompt "I"
        ElseIf TestPrompt(regEx, screenContent, "DESC:") Then
            EnterReviewPrompt ""
        ElseIf TestPrompt(regEx, screenContent, "Enter a technician number") Then
            bzhao.SendKey "<F3>"
            bzhao.Pause REVIEW_PAUSE
        ElseIf TestPrompt(regEx, screenContent, "TECHNICIAN\s*\(Y/N\)") Then
            EnterReviewPrompt "Y"
        ElseIf TestPrompt(regEx, screenContent, "TECHNICIAN FINISHING WORK") Then
            EnterReviewPrompt "99"
        ElseIf TestPrompt(regEx, screenContent, "IS ASSIGNED TO LINE") Then
            EnterReviewPrompt "Y"
        ElseIf TestPrompt(regEx, screenContent, "TECHNICIAN.*\([A-Za-z0-9]+\)\?") Then
            EnterReviewPrompt ""
        ElseIf TestPrompt(regEx, screenContent, "TECHNICIAN\s*\?") Then
            EnterReviewPrompt "99"
        ElseIf TestPrompt(regEx, screenContent, "ACTUAL HOURS") Then
            EnterReviewPrompt ""
        ElseIf TestPrompt(regEx, screenContent, "SOLD HOURS") Then
            EnterReviewPrompt ""
        ElseIf TestPrompt(regEx, screenContent, "ADD A LABOR OPERATION") Then
            EnterReviewPrompt "N"
        End If

        elapsed = Timer - startTime
        If elapsed < 0 Then elapsed = elapsed + 86400

        If elapsed > 45 Then
            HandleReviewPrompts = False
            Exit Function
        End If
    Loop
End Function

Function FinishLineForReview(lineLetter)
    Dim screenContent, startTime, elapsed, regEx

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = False

    FinishLineForReview = False

    If Not WaitForTextAtBottom("COMMAND:") Then Exit Function
    bzhao.SendKey "FNL " & lineLetter
    bzhao.Pause 100
    bzhao.SendKey "<NumpadEnter>"
    bzhao.Pause REVIEW_PAUSE

    startTime = Timer

    Do
        bzhao.Pause REVIEW_PAUSE
        screenContent = ""
        bzhao.ReadScreen screenContent, 1920, 1, 1

        If IsIdleCommandPrompt() Then
            FinishLineForReview = True
            Exit Function
        ElseIf InStr(1, screenContent, "LINE " & lineLetter & " IS ALREADY FINISHED", vbTextCompare) > 0 Then
            bzhao.SendKey "<Enter>"
            bzhao.Pause REVIEW_PAUSE
            If Not WaitForTextAtBottom("COMMAND:") Then Exit Function
            FinishLineForReview = True
            Exit Function
        ElseIf InStr(1, screenContent, "LINE CODE " & lineLetter & " IS NOT ON FILE", vbTextCompare) > 0 Then
            bzhao.SendKey "<Enter>"
            bzhao.Pause REVIEW_PAUSE
            Exit Function
        ElseIf TestPrompt(regEx, screenContent, "TECHNICIAN FINISHING WORK\s*\?") Then
            EnterReviewPrompt "99"
        ElseIf TestPrompt(regEx, screenContent, "TECHNICIAN\s*\([A-Za-z0-9]+\)\?") Then
            EnterReviewPrompt ""
        ElseIf TestPrompt(regEx, screenContent, "TECHNICIAN\s*\?") Then
            EnterReviewPrompt "99"
        ElseIf TestPrompt(regEx, screenContent, "IS ASSIGNED TO LINE") Then
            EnterReviewPrompt "Y"
        ElseIf TestPrompt(regEx, screenContent, "IS LINE '[A-Z]' COMPLETED\s*\(Y/N\)") Then
            EnterReviewPrompt "Y"
        ElseIf TestPrompt(regEx, screenContent, "ACTUAL HOURS") Then
            EnterReviewPrompt ""
        ElseIf TestPrompt(regEx, screenContent, "SOLD HOURS") Then
            EnterReviewPrompt ""
        End If

        elapsed = Timer - startTime
        If elapsed < 0 Then elapsed = elapsed + 86400

        If elapsed > 45 Then Exit Function
    Loop
End Function

Function IsIdleCommandPrompt()
    Dim line23, line24

    line23 = ""
    line24 = ""
    bzhao.ReadScreen line23, 80, 23, 1
    bzhao.ReadScreen line24, 80, 24, 1

    line23 = Trim(line23)
    line24 = Trim(line24)

    If InStr(1, line23, "COMMAND:", vbTextCompare) = 1 Then
        IsIdleCommandPrompt = (line24 = "")
    ElseIf InStr(1, line24, "COMMAND:", vbTextCompare) = 1 Then
        IsIdleCommandPrompt = (line23 = "")
    Else
        IsIdleCommandPrompt = False
    End If
End Function

Sub EnterReviewPrompt(text)
    If text <> "" Then bzhao.SendKey CStr(text)
    bzhao.Pause 50
    bzhao.SendKey "<NumpadEnter>"
    bzhao.Pause REVIEW_PAUSE
End Sub

Function TestPrompt(regEx, text, pattern)
    regEx.Pattern = pattern
    TestPrompt = regEx.Test(text)
End Function

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
    Call AddPromptToDict(dict, "ADD A LABOR OPERATION( \(N\)\?)?", "N", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "ADD A LABOR OPERATION", "", "<Enter>", False)
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

            ' Give each prompt transition its own timeout window.
            startTime = Timer
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

    Call LogErrorMessage("Prompt sequence timed out after " & timeoutMs & " ms." & vbCrLf & "Elapsed: " & Int(elapsed) & " ms." & vbCrLf & BuildPromptAreaSnapshot())
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
' Waits for text to appear at bottom of screen (rows 23-24)
'-----------------------------------------------------------
Function WaitForTextAtBottom(targetText)
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
        
        If found Then
            WaitForTextAtBottom = True
            Exit Function
        End If
        
        If elapsed >= 30000 Then
            bzhao.msgBox "ERROR: Timeout waiting for '" & targetText & "'" & vbCrLf & "Last 2 lines: " & vbCrLf & buffer23 & vbCrLf & buffer24, 16
            WaitForTextAtBottom = False
            Exit Function
        End If
    Loop
End Function

'-----------------------------------------------------------
' Sends text and presses Enter
'-----------------------------------------------------------
Sub EnterTextAndWait(text)
    If text <> "" Then bzhao.SendKey text
    bzhao.Pause 100
    bzhao.SendKey "<NumpadEnter>"
    bzhao.Pause 500
End Sub
