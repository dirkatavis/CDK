
Dim POLL_INTERVAL: POLL_INTERVAL = 1 ' 1 second polling interval for development
Dim CSV_FILE_PATH
Dim LOG_FILE_PATH
Dim DebugLevel ' 0=None, 1=Error, 2=Info, 3=Debug
DebugLevel = 2 ' Set default debug level (change as needed)
Dim fso, ts, strLine, number
Dim bzhao: Set bzhao = CreateObject("BZWhll.WhllObj")

'-----------------------------------------------------------
' Define file paths and connect to BlueZone
'-----------------------------------------------------------
CSV_FILE_PATH = "C:\Temp\Code\Scripts\VBScript\CDK\Close_ROs\Close_ROs_Pt1.csv"
LOG_FILE_PATH = "C:\Temp\Code\Scripts\VBScript\CDK\Close_ROs\Close_ROs_Pt2.log"


'-----------------------------------------------------------
' Main script execution loop
'-----------------------------------------------------------
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(CSV_FILE_PATH) Then
    bzhao.Connect ""
    Set ts = fso.OpenTextFile(CSV_FILE_PATH, 1)
    ts.ReadLine   ' Skip header row if present

    Do While Not ts.AtEndOfStream
        strLine = ts.ReadLine
        number = Trim(strLine)
        If Len(number) > 0 And IsNumeric(number) Then
            Call Main(number)
        End If
    Loop

    ts.Close
    Set ts = Nothing
Else
    MsgBox "Error: The file '" & CSV_FILE_PATH & "' was not found.", vbCritical, "File Not Found"
End If

Set fso = Nothing

'-----------------------------------------------------------
' Main subroutine to check and process each RO number
'-----------------------------------------------------------
Sub Main(number)
    ' Enter the number and send Enter
    Call EnterTextAndWait(number, 2000)

    ' Check for NOT ON FILE error in line 2
    If CheckForTextInLine2("NOT ON FILE") Then
    LogResult "INFO", "RO NOT ON FILE - Skipping to next. RO: " & number
        Exit Sub
    End If

    ' Check for "is closed" response in line 2
    If CheckForTextInLine2("is closed") Then
    LogResult "INFO", "RO IS CLOSED - Skipping to next. RO: " & number
        Exit Sub
    End If

    Call Closeout_Ro()
End Sub

'-----------------------------------------------------------
' Checks if line 2 contains the specified text and returns True if so
'-----------------------------------------------------------
Function CheckForTextInLine2(targetText)
    Dim screenContentBuffer, screenLength
    screenLength = 80
    bzhao.Pause 2000 ' Give screen time to update
    bzhao.ReadScreen screenContentBuffer, screenLength, 2, 1
    If InStr(screenContentBuffer, targetText) > 0 Then
        CheckForTextInLine2 = True
    Else
        CheckForTextInLine2 = False
    End If
End Function


'-----------------------------------------------------------
' Closeout_Ro script subroutines
' (replace EnterText(...) calls with EnterTextAndWait(..., 1))
'-----------------------------------------------------------
Sub Closeout_Ro()
    ' Use the new subroutine to add the 'B' story.
        WaitForTextAtBottom "COMMAND:"
        AddStory bzhao, "B"
    'If HandleCloseoutErrors() Then Exit Sub
    
    ' Use the new subroutine to add the 'C' story.
    WaitForTextAtBottom "COMMAND:"
    AddStory bzhao, "C"
    'If HandleCloseoutErrors() Then Exit Sub

    
    '*******************************************************
    ' Final Closeout Steps
    '*******************************************************
    WaitForTextAtBottom "COMMAND:"
    EnterTextAndWait "FC", 1000
    If HandleCloseoutErrors() Then Exit Sub
    
    ' Have all hours been entered
    WaitForTextAtBottom "ALL LABOR POSTED"
    EnterTextAndWait "Y", 1000
    If HandleCloseoutErrors() Then Exit Sub

    
    ' OUT MILEAGE
    WaitForTextAtBottom "MILEAGE OUT"
    EnterTextAndWait "", 1000
    If HandleCloseoutErrors() Then Exit Sub
    
    ' IN MILEAGE
    WaitForTextAtBottom "MILEAGE IN"
    EnterTextAndWait "", 1000
    If HandleCloseoutErrors() Then Exit Sub
    
    ' OK TO CLOSE THE RO?
    WaitForTextAtBottom "O.K. TO CLOSE RO"
    EnterTextAndWait "Y", 1000
    If HandleCloseoutErrors() Then Exit Sub
    
    ' SEND TO PRINTER 2
    bzhao.Pause 2000
    WaitForTextAtBottom "INVOICE PRINTER"
    EnterTextAndWait "2", 1000
    lastRoResult = "Successfully closed"
End Sub

'-----------------------------------------------------------
' Waits for a specific text to appear at the bottom line of the screen within a timeout
' targetText: the string to wait for
' pollInterval: ms between checks
' timeout: ms before giving up
'-----------------------------------------------------------
Sub WaitForTextAtBottom(targetText)
    Dim elapsed, screenContentBuffer, screenLength, found, row, col
    elapsed = 0
    row = 23 ' correct line for debug
    col = 1
    screenLength = 80 ' one line
    LogResult "DEBUG", "Waiting for text at bottom: '" & targetText & "'"
    Do
        'LogResult "DEBUG", "Waiting for text at bottom: '" & targetText & "'. Elapsed time: " & elapsed & " ms"
        bzhao.Pause 500
        elapsed = elapsed + 500
        bzhao.ReadScreen screenContentBuffer, screenLength, row, col
        Dim debugLine
        debugLine = Left(screenContentBuffer, 40)
    LogResult "DEBUG", "Last line (row " & row & "): '" & debugLine & "' | Expected: '" & targetText & "' | Match: " & (InStr(debugLine, targetText) > 0)
        found = (InStr(screenContentBuffer, targetText) > 0)
        If found Then
            Exit Do
        End If
        'LogResult "DEBUG", "Text not found yet. Continuing to wait."
        If elapsed >= 5000 Then
            MsgBox "ERROR: Timeout waiting for text '" & targetText & "' to appear at bottom of screen. Script will exit.", vbCritical
            bzhao.StopScript
        End If
    Loop
End Sub

'-----------------------------------------------------------
' Helper functions and subroutines
'-----------------------------------------------------------
Function IsTextPresent(textToFind)
    Dim screenContentBuffer, screenLength
    screenLength = 24 * 80
    bzhao.ReadScreen screenContentBuffer, screenLength, 1, 1
    IsTextPresent = (InStr(1, screenContentBuffer, textToFind, vbTextCompare) > 0)
End Function

'-----------------------------------------------------------
' EnterTextAndWait subroutine
' If the entry requires an enter key press, use this subroutine.
'-----------------------------------------------------------
Sub EnterTextAndWait(text, wait)
    bzhao.SendKey text
    bzhao.Pause 100 ' Small delay to allow text to register
    Call PressKey ("<NumpadEnter>")
    bzhao.Pause 500
End Sub

'-----------------------------------------------------------
' PressKey subroutine
' If the entry requires no enter key press, use this subroutine.
'-----------------------------------------------------------
Sub PressKey(key)
   bzhao.SendKey key
   bzhao.Pause 100 ' Small delay to allow key press to register  
End Sub

Sub LogResult(logType, message)
    ' logType: "ERROR"=1, "INFO"=2, "DEBUG"=3
    Dim logFSO, logFile, typeLevel
    Select Case UCase(logType)
        Case "ERROR": typeLevel = 1
        Case "INFO": typeLevel = 2
        Case "DEBUG": typeLevel = 3
        Case Else: typeLevel = 2 ' Default to INFO
    End Select
    If typeLevel <= DebugLevel Then
        Set logFSO = CreateObject("Scripting.FileSystemObject")
        Set logFile = logFSO.OpenTextFile(LOG_FILE_PATH, 8, True)
        logFile.WriteLine Now & "  [" & logType & "] " & message
        logFile.Close
        Set logFile = Nothing
        Set logFSO = Nothing
    End If
End Sub
Function HandleCloseoutErrors()
    Dim screenContentBuffer, screenLength
    screenLength = 5 * 80 ' Read last 5 rows for error messages
    bzhao.ReadScreen screenContentBuffer, screenLength, 20, 1

    If InStr(1, screenContentBuffer, "ERROR", vbTextCompare) > 0 Then
    LogResult "ERROR", "Closeout failed due to error on screen."
        ' Send 'E' to exit back to main screen
        bzhao.SendKey "E"
        bzhao.Pause 100
        bzhao.SendKey "<NumpadEnter>"
        bzhao.Pause 1000
        HandleCloseoutErrors = True
    Else
        HandleCloseoutErrors = False
    End If
End Function

Sub AddStory(bzhao, storyCode)
    ' Use the storyCode variable (e.g., "B" or "C") to make the code dynamic.
    EnterText bzhao, "R " & storyCode 

    If storyCode = "B" Then

        ' Wait for the expected prompt at the bottom before sending the story command
        WaitForTextAtBottom "LABOR TYPE FOR LINE" ' Wait up to 15s for command prompt (adjust text as needed)
        EnterText bzhao, ""

        'Entering Operations Code. Defaulting to 'CATP'
        WaitForTextAtBottom "OPERATION CODE FOR "
        EnterText bzhao, ""
                
        'Entering story description. Accepting default
        WaitForTextAtBottom "DESC: CHECK AND ADJ"
        EnterText bzhao, ""

        'Entering technician id
        WaitForTextAtBottom "TECHNICIAN"
        EnterText bzhao, "99"    

        'Entering Actual hours. Defaulting to 0
        WaitForTextAtBottom "ACTUAL HOURS"
        EnterText bzhao, ""

        'Entering sold hours. defaulting to 10
        WaitForTextAtBottom "SOLD HOURS"
        EnterText bzhao, ""        
    
        'Add a labor operation? Defaulting to No
        WaitForTextAtBottom "ADD A LABOR OPERATION"
        EnterText bzhao, ""
    End If  

    If storyCode = "C" Then
        ' Wait for the expected prompt at the bottom before sending the story command
        WaitForTextAtBottom "LABOR TYPE FOR LINE" ' Wait up to 15s for command prompt (adjust text as needed)
        EnterText bzhao, ""
        
        'Entering Operations Code. Defaulting to 'CATP'
        WaitForTextAtBottom "OPERATION CODE FOR "
        EnterText bzhao, ""

        WaitForTextAtBottom "DESC: MEASURE AND"
        EnterText bzhao, ""
        

        'Entering technician id
        WaitForTextAtBottom "TECHNICIAN"
        EnterText bzhao, "99"  

        
        'Entering Actual hours. Defaulting to 0
        WaitForTextAtBottom "ACTUAL HOURS"
        EnterText bzhao, ""

        'Entering sold hours. defaulting to 10
        WaitForTextAtBottom "SOLD HOURS"
        EnterText bzhao, ""     
         
        'Add a labor operation? Defaulting to No
        WaitForTextAtBottom "ADD A LABOR OPERATION"
        EnterText bzhao, ""

    End If


End Sub

Sub EnterText(bzhao, textToEnter)
    bzhao.SendKey textToEnter
    bzhao.Pause 100 ' Small delay to allow text to register
    bzhao.SendKey "<Enter>"
End Sub

bzhao.Disconnect
