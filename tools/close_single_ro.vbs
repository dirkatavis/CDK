Option Explicit

Dim bzhao

Call Main()

Sub Main()
    Set bzhao = CreateObject("BZWhll.WhllObj")
    bzhao.Connect ""

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
