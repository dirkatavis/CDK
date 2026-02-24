Option Explicit

' --- Load PathHelper for centralized path management ---
' We use a minimal bootstrap here to load the shared framework
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim sh: Set sh = CreateObject("WScript.Shell")
Dim basePath: basePath = sh.Environment("USER")("CDK_BASE")
If basePath = "" Or Not fso.FolderExists(basePath) Then
    MsgBox "Error: CDK_BASE environment variable not set or path does not exist.", vbCritical
    WScript.Quit
End If

Dim helperPath: helperPath = fso.BuildPath(basePath, "framework\PathHelper.vbs")
ExecuteGlobal fso.OpenTextFile(helperPath).ReadAll

Dim POLL_INTERVAL: POLL_INTERVAL = 1 ' 1 second polling interval for development
Dim CSV_FILE_PATH: CSV_FILE_PATH = GetConfigPath("Finalize_Close", "CSV")
Dim LOG_FILE_PATH: LOG_FILE_PATH = GetConfigPath("Finalize_Close", "Log")
Dim DebugLevel ' 0=None, 1=Error, 2=Info, 3=Debug
DebugLevel = 2 ' Set default debug level (change as needed)

Dim ts, strLine, roNumber, lastRoResult
Dim bzhao: Set bzhao = CreateObject("BZWhll.WhllObj")

'-----------------------------------------------------------
' Main script execution loop
'-----------------------------------------------------------
If fso.FileExists(CSV_FILE_PATH) Then
    Dim connResult: connResult = bzhao.Connect("")
    If connResult <> 0 Then
        MsgBox "Error: Could not connect to BlueZone session. Ensure BlueZone is open and active.", vbCritical, "Connection Failed"
        WScript.Quit 1
    End If
    
    Set ts = fso.OpenTextFile(CSV_FILE_PATH, 1)
    ts.ReadLine   ' Skip header row if present

    Do While Not ts.AtEndOfStream
        strLine = ts.ReadLine
        roNumber = Trim(strLine)
        If Len(roNumber) > 0 And IsNumeric(roNumber) Then
            Call Main(roNumber)
        End If
    Loop

    ts.Close
    Set ts = Nothing
    bzhao.Disconnect
Else
    MsgBox "Error: The file '" & CSV_FILE_PATH & "' was not found.", vbCritical, "File Not Found"
End If

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
' DiscoverLineLetters: Detects which line letters (A, B, C, etc.) are present
' on the current RO Detail screen by reading the LC column.
' Returns: Array of line letters found (e.g., Array("A", "C") if B is missing)
' Note: This function is duplicated in Close_ROs_Pt1.vbs for independence.
'       Consider extracting to shared include file if more scripts need this.
'-----------------------------------------------------------
Function DiscoverLineLetters()
    Dim maxLinesToCheck, i, capturedLetter, screenContentBuffer, readLength, nextColChar
    Dim foundLetters, foundCount
    Dim startReadRow, startReadColumn, emptyRowCount
    
    ' Array to store discovered line letters
    Dim tempLetters(25) ' Max 26 letters A-Z
    foundCount = 0
    emptyRowCount = 0
    
    ' The prompt area starts at row 23, so we must stop at row 22 to avoid 
    ' misidentifying prompt characters (like 'C' in 'COMMAND:') as line letters.
    Dim startRow, endRow
    startRow = 10 ' Anchor at first actual data row (skip header rows)
    endRow = 22   ' Last possible data row before prompt area
    
    For startReadRow = startRow To endRow
        startReadColumn = 1
        readLength = 1 ' Read just 1 character (the line letter)
        
        On Error Resume Next
        bzhao.ReadScreen screenContentBuffer, readLength, startReadRow, startReadColumn
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        On Error GoTo 0
        
        ' Trim and check if it's a valid letter (A-Z)
        capturedLetter = Trim(screenContentBuffer)
        If Len(capturedLetter) = 1 Then
            If Asc(UCase(capturedLetter)) >= Asc("A") And Asc(UCase(capturedLetter)) <= Asc("Z") Then
                ' Peek column 2 to ensure this is a line letter (typical form: "A  DESCRIPTION")
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
                    emptyRowCount = 0 ' Reset when a letter is found
                Else
                    emptyRowCount = emptyRowCount + 1
                End If
            Else
                emptyRowCount = emptyRowCount + 1
            End If
        Else
            emptyRowCount = emptyRowCount + 1
        End If

        ' If we hit 3 consecutive rows without a letter, we've likely finished the list
        If emptyRowCount >= 3 Then Exit For
    Next
    
    ' If no line letters found, log error and return empty array to skip this RO
    If foundCount = 0 Then
        LogResult "ERROR", "No line letters discovered - skipping RO"
        DiscoverLineLetters = Array()
        Exit Function
    End If
    
    ' Create properly sized array with found letters
    ReDim foundLetters(foundCount - 1)
    For i = 0 To foundCount - 1
        foundLetters(i) = tempLetters(i)
    Next
    
    ' Log discovered line letters for debugging
    Dim lettersList
    lettersList = Join(foundLetters, ", ")
    LogResult "INFO", "Discovered line letters: " & lettersList
    
    DiscoverLineLetters = foundLetters
End Function

'-----------------------------------------------------------
' Closeout_Ro script subroutines
' (replace EnterText(...) calls with EnterTextAndWait(..., 1))
'-----------------------------------------------------------
Sub Closeout_Ro()
    ' Discover which line letters are present on the screen
    Dim lineLetters, i, screenContent
    lineLetters = DiscoverLineLetters()
    
    ' If no line letters discovered, log error and exit
    If IsEmpty(lineLetters) Or UBound(lineLetters) = -1 Then
        LogResult "ERROR", "No line letters discovered - Skipping closeout"
        Exit Sub
    End If
    
    ' Add stories for each discovered line letter
    For i = 0 To UBound(lineLetters)
        ' Wait for either the main COMMAND prompt or any story review prompt
        ' This handles cases where CDK might auto-advance to the next line
        LogResult "INFO", "Syncing for line " & lineLetters(i)
        Call WaitForTextAtBottom("COMMAND:|LABOR TYPE|OPERATION CODE|DESC:|TECHNICIAN")
        
        ' Read the screen to see if we need to send the 'R' command
        bzhao.ReadScreen screenContent, 160, 23, 1
        If InStr(UCase(screenContent), "COMMAND:") > 0 Then
            EnterText bzhao, "R " & lineLetters(i) 
        End If

        ' Process the prompts for this story
        AddStory bzhao, lineLetters(i)
    Next

    
    '*******************************************************
    ' Filing Steps
    '*******************************************************
    WaitForTextAtBottom "COMMAND:"
    EnterTextAndWait "FC", 1000
    If HandleCloseoutErrors() Then Exit Sub
    
    ' Have all hours been entered
    WaitForTextAtBottom "ALL LABOR POSTED"
    EnterTextAndWait "Y", 1000
    If HandleCloseoutErrors() Then Exit Sub

    
    ' ' OUT MILEAGE
    WaitForTextAtBottom "MILEAGE OUT"
    EnterTextAndWait "", 1000
    If HandleCloseoutErrors() Then Exit Sub
    
    ' ' IN MILEAGE
    WaitForTextAtBottom "MILEAGE IN"
    EnterTextAndWait "", 1000
    If HandleCloseoutErrors() Then Exit Sub
    
    ' ' OK TO CLOSE THE RO?
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
    Dim elapsed, screenContentBuffer, screenLength, found, col
    elapsed = 0
    col = 1
    screenLength = 80 ' one line
    LogResult "DEBUG", "Waiting for text at bottom: '" & targetText & "'"
    
    ' Split targetText by | in case multiple options are provided
    Dim targets: targets = Split(targetText, "|")
    Dim i
    
    Do
        bzhao.Pause 500
        elapsed = elapsed + 500
        
        ' Read both rows 23 and 24
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
            Exit Do
        End If
        
        If elapsed >= 10000 Then
            MsgBox "ERROR: Timeout waiting for text '" & targetText & "' to appear at bottom of screen. Script will exit." & vbCrLf & "Last 2 lines: " & vbCrLf & buffer23 & vbCrLf & buffer24, vbCritical
            WScript.Quit
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
Sub EnterTextAndWait(textVal, waitMs)
    bzhao.SendKey textVal
    bzhao.Pause 100 ' Small delay to allow text to register
    Call PressKey("<NumpadEnter>")
    bzhao.Pause waitMs
End Sub

'-----------------------------------------------------------
' PressKey subroutine
' If the entry requires no enter key press, use this subroutine.
'-----------------------------------------------------------
Sub PressKey(keyName)
   bzhao.SendKey keyName
   bzhao.Pause 100 ' Small delay to allow key press to register  
End Sub

Sub LogResult(logType, message)
    ' logType: "ERROR"=1, "INFO"=2, "DEBUG"=3
    Dim logFSO, logFile, typeLevel
    Select Case UCase(logType)
        Case "ERROR"
            typeLevel = 1
        Case "INFO"
            typeLevel = 2
        Case "DEBUG"
            typeLevel = 3
        Case Else
            typeLevel = 2 ' Default to INFO
    End Select
    If typeLevel <= DebugLevel Then
        Set logFSO = CreateObject("Scripting.FileSystemObject")
        
        ' Ensure parent folder exists
        Dim logDir: logDir = logFSO.GetParentFolderName(LOG_FILE_PATH)
        If Not logFSO.FolderExists(logDir) Then
            logFSO.CreateFolder(logDir)
        End If

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
    ' This subroutine now uses a state-aware loop to handle prompts in any order.
    ' It exits when it detects the "COMMAND:" prompt or "ADD A LABOR OPERATION" is completed.
    
    Dim screenContent, startTime, elapsed, lastPrompt, sameCount
    startTime = Timer
    lastPrompt = ""
    sameCount = 0
    
    Do
        bzhao.Pause 500
        ' Read the prompt area (bottom 2 lines)
        bzhao.ReadScreen screenContent, 160, 23, 1
        screenContent = UCase(screenContent)
        
        ' 1. Check for exit condition (back to main screen)
        If InStr(screenContent, "COMMAND:") > 0 Then
            LogResult "INFO", "Story " & storyCode & " review complete (COMMAND: detected)."
            Exit Sub
        End If
        
        ' 2. Detect and respond to specific prompts
        Dim currentMatched: currentMatched = ""
        
        If InStr(screenContent, "LABOR TYPE") > 0 Then
            currentMatched = "LABOR TYPE"
            EnterText bzhao, ""
        ElseIf InStr(screenContent, "OPERATION CODE") > 0 Then
            currentMatched = "OPERATION CODE"
            EnterText bzhao, ""
        ElseIf InStr(screenContent, "DESC:") > 0 Then
            currentMatched = "DESC:"
            EnterText bzhao, ""
        ElseIf InStr(screenContent, "TECHNICIAN") > 0 Or InStr(screenContent, "TECH...") > 0 Then
            currentMatched = "TECHNICIAN"
            EnterText bzhao, "99"
        ElseIf InStr(screenContent, "ACTUAL HOURS") > 0 Then
            currentMatched = "ACTUAL HOURS"
            EnterText bzhao, ""
        ElseIf InStr(screenContent, "SOLD HOURS") > 0 Then
            currentMatched = "SOLD HOURS"
            EnterText bzhao, ""
        ElseIf InStr(screenContent, "ADD A LABOR OPERATION") > 0 Then
            currentMatched = "ADD A LABOR OPERATION"
            EnterText bzhao, ""
            ' After "Add a labor operation", we expect to exit soon
        ElseIf InStr(screenContent, "(END OF DISPLAY)") > 0 Then
            currentMatched = "END OF DISPLAY"
            EnterText bzhao, ""
        End If
        
        ' 3. Stuck detection
        If currentMatched <> "" Then
            If currentMatched = lastPrompt Then
                sameCount = sameCount + 1
            Else
                sameCount = 0
                lastPrompt = currentMatched
            End If
            
            ' If we've sent the same response 3 times, try a generic Enter to force a move
            If sameCount >= 3 Then
                LogResult "DEBUG", "Stuck at prompt '" & currentMatched & "'. Sending extra Enter."
                EnterText bzhao, ""
                sameCount = 0
            End If
        End If
        
        ' 4. Timeout (20 seconds per story)
        elapsed = Timer - startTime
        If elapsed > 20 Then
            MsgBox "ERROR: Timeout in AddStory for " & storyCode & ". Script will exit to prevent data corruption." & vbCrLf & "Current screen prompt area: " & vbCrLf & screenContent, vbCritical
            WScript.Quit
        End If
    Loop
End Sub

Sub EnterText(bzhao, textToEnter)
    bzhao.SendKey textToEnter
    bzhao.Pause 150 ' Small delay to allow text to register
    bzhao.SendKey "<NumpadEnter>"
End Sub
