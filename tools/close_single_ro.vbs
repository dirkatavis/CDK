Option Explicit

Dim bzhao
Dim g_ReviewTimeoutMs: g_ReviewTimeoutMs = 10000

Call Main()

Sub Main()
    Set bzhao = CreateObject("BZWhll.WhllObj")
    bzhao.Connect ""

    ' 1. Ensure at COMMAND: prompt on the RO detail screen
    WaitForTextAtBottom "COMMAND:"

    ' 2. Review phase: review each line item before issuing Final Charge
    If Not ReviewLineItems() Then
        Exit Sub
    End If

    ' 3. FC command (Final Charge)
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
' Reviews each line item (R A, R B, ...) before Final Charge.
' Uses DiscoverLineLetters for dynamic discovery; falls back
' to A, B, C if no letters are found on screen.
' Returns True if all reviews succeed; False on any failure.
'-----------------------------------------------------------
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
            bzhao.MsgBox "ERROR: 'R " & letter & "' did not return COMMAND: prompt." & vbCrLf & "Script cancelled.", 16
            Exit Function
        End If
    Next

    ReviewLineItems = True
End Function

'-----------------------------------------------------------
' Sends "R <letter>" and waits up to 10 s for COMMAND: to return.
' Returns True on success, False on timeout/failure.
'-----------------------------------------------------------
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

'-----------------------------------------------------------
' Returns True if "COMMAND:" is visible on rows 23-24.
'-----------------------------------------------------------
Function IsAtCommandPrompt()
    Dim buf23, buf24
    bzhao.ReadScreen buf23, 80, 23, 1
    bzhao.ReadScreen buf24, 80, 24, 1
    IsAtCommandPrompt = (InStr(UCase(buf23 & " " & buf24), "COMMAND:") > 0)
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
