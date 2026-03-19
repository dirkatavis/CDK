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
        bzhao.MsgBox "Review sequence failed. RO not closed.", 16
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
' Executes the review sequence: R A, R B, R C in order.
' Uses DiscoverLineLetters for dynamic discovery; falls back
' to A, B, C if no letters found on screen.
' Returns True if all reviews succeed, False on any failure.
'-----------------------------------------------------------
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
            bzhao.MsgBox "ERROR: Review command '" & reviewCommand & "' failed." & vbCrLf & _
                         "COMMAND: prompt did not return within timeout.", 16
            Exit Function
        End If
    Next

    ExecuteReviewSequence = True
End Function

'-----------------------------------------------------------
' Sends a review command (e.g., "R A") and waits for
' COMMAND: prompt to return, verifying normal return.
' Returns True on success, False on timeout/failure.
'-----------------------------------------------------------
Function WaitForReviewCommandCompletion(reviewCommand)
    Dim timeoutMs: timeoutMs = 10000
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
        If elapsed < 0 Then elapsed = elapsed + 86400000 ' Handle midnight rollover
        
        If elapsed > timeoutMs Then
            Exit Do
        End If
    Loop

    WaitForReviewCommandCompletion = found
End Function

'-----------------------------------------------------------
' Returns True if "COMMAND:" is visible on rows 23-24.
'-----------------------------------------------------------
Function IsCommandPromptVisible()
    Dim buf23, buf24
    bzhao.ReadScreen buf23, 80, 23, 1
    bzhao.ReadScreen buf24, 80, 24, 1
    IsCommandPromptVisible = (InStr(UCase(buf23 & " " & buf24), "COMMAND:") > 0)
End Function

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
