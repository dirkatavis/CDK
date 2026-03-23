Option Explicit

Dim bzhao
Const REVIEW_PAUSE = 500

Call Main()

Sub Main()
    Set bzhao = CreateObject("BZWhll.WhllObj")
    bzhao.Connect ""

    ' Review all discovered lines before the existing final charge flow.
    If Not ProcessRoReview() Then
        bzhao.msgBox "ERROR: Review phase failed before FC. RO was not closed.", 16
        Exit Sub
    End If

    ' 1. FC command (Final Charge)
    WaitForTextAtBottom "COMMAND:"
    EnterTextAndWait "FC"
    bzhao.Pause 1000

    ' 2. ALL LABOR POSTED prompt
    WaitForTextAtBottom "ALL LABOR POSTED"
    EnterTextAndWait "Y"
    bzhao.Pause 1000

    ' 3. MILEAGE OUT prompt (accept default)
    WaitForTextAtBottom "MILEAGE OUT"
    EnterTextAndWait ""
    bzhao.Pause 1000

    ' 4. MILEAGE IN prompt
    WaitForTextAtBottom "MILEAGE IN"
    EnterTextAndWait ""
    bzhao.Pause 1000

    ' 5. OK TO CLOSE prompt
    WaitForTextAtBottom "O.K. TO CLOSE RO"
    EnterTextAndWait "Y"
    bzhao.Pause 1000

    ' 6. INVOICE PRINTER prompt
    WaitForTextAtBottom "INVOICE PRINTER"
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
        WaitForTextAtBottom "COMMAND:"
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

            WaitForTextAtBottom "COMMAND:"
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

    WaitForTextAtBottom "COMMAND:"
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
            WaitForTextAtBottom "COMMAND:"
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
