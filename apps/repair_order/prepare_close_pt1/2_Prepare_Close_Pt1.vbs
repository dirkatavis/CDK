Option Explicit

'====================================================================
' Script: Close_RO_Automation.vbs
' Description: Simplified script to read 6-digit numbers from a CSV
'              and perform BlueZone CCC operations (A, B, C) with pauses.
'              All screen checks and error handling have been removed.
'====================================================================

' --- Load PathHelper for centralized path management ---
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Const BASE_ENV_VAR_LOCAL = "CDK_BASE"

' Find repo root by searching for .cdkroot marker
Function FindRepoRootForBootstrap()
    Dim sh: Set sh = CreateObject("WScript.Shell")
    Dim basePath: basePath = sh.Environment("USER")(BASE_ENV_VAR_LOCAL)

    If basePath = "" Or Not g_fso.FolderExists(basePath) Then
        Err.Raise 53, "Bootstrap", "Invalid or missing CDK_BASE. Value: " & basePath
    End If

    If Not g_fso.FileExists(g_fso.BuildPath(basePath, ".cdkroot")) Then
        Err.Raise 53, "Bootstrap", "Cannot find .cdkroot in base path:" & vbCrLf & basePath
    End If

    FindRepoRootForBootstrap = basePath
End Function

Dim helperPath: helperPath = g_fso.BuildPath(FindRepoRootForBootstrap(), "framework\PathHelper.vbs")
ExecuteGlobal g_fso.OpenTextFile(helperPath).ReadAll

' --- Configuration ---
Dim CSV_FILE: CSV_FILE = GetConfigPath("Prepare_Close_Pt1", "CSV")
Const NUM_COLUMN = 0 ' This constant is now largely redundant but kept for clarity

' --- VBScript Objects ---
Dim fso, tsInput
Dim bzhao ' BlueZone Host Access Object

' --- Variables ---
Dim strLine ' Holds the entire line from the file, which is the RO number
Dim RoNumber

' --- Main Execution ---
' Initialize File System Object (FSO)
Set fso = CreateObject("Scripting.FileSystemObject")

' Connect to BlueZone
Set bzhao = CreateObject("BZWhll.WhllObj")
bzhao.Connect ""

' Check and Open Input File
If fso.FileExists(CSV_FILE) Then
    Set tsInput = fso.OpenTextFile(CSV_FILE, 1) ' 1=ForReading
Else
    ' Leaving this basic check for file existence.
    MsgBox "Error: Input file not found at " & CSV_FILE, vbCritical
        bzhao.StopScript
End If

' 4. Process Records - THIS IS THE LOOP
Do While Not tsInput.AtEndOfStream
    strLine = tsInput.ReadLine
    
    ' Since there is only 1 element per line (the RO Number), we no longer need to split the line.
    RoNumber = Trim(strLine)
    
    ' Removed the UBound(arrValues) check since there is no array now.
    
    ' Simple check for a valid 6-digit number before processing
    If Len(RoNumber) = 6 And IsNumeric(RoNumber) Then
        'bzhao.msgBox "Processing RO Number: " & RoNumber
        Call ProcessRo(RoNumber)
    End If
Loop

' 5. Cleanup
tsInput.Close
bzhao.Disconnect

Set tsInput = Nothing
Set fso = Nothing
Set bzhao = Nothing


' --- Subroutines ---

'-----------------------------------------------------------
' DiscoverLineLetters: Detects which line letters (A, B, C, etc.) are present
' on the current RO Detail screen by reading the LC column.
' Returns: Array of line letters found (e.g., Array("A", "C") if B is missing)
' Note: This function is duplicated in Close_ROs_Pt2.vbs for independence.
'       Consider extracting to shared include file if more scripts need this.
'-----------------------------------------------------------
Function DiscoverLineLetters()
    Dim maxLinesToCheck, i, capturedLetter, screenContentBuffer, readLength
    Dim foundLetters, foundCount
    Dim startReadRow, startReadColumn
    Dim missingLetters
    
    ' Array to store discovered line letters
    Dim tempLetters(25) ' Max 26 letters A-Z (sized for theoretical maximum)
    foundCount = 0
    maxLinesToCheck = 10 ' Practical limit: Check up to 10 line letters (business logic constraint)
    missingLetters = 0
    
    ' The LC column header is typically on row 6, and line letters start on row 10
    ' Column 1 contains the line letter (under the "L" in "LC")
    Dim startRow
    startRow = 10 ' First data row (line letters always start at row 10)
    
    ' Read the screen area where line letters appear (column 1, multiple rows)
    For i = 0 To maxLinesToCheck - 1
        startReadRow = startRow + i
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
                tempLetters(foundCount) = UCase(capturedLetter)
                foundCount = foundCount + 1
                missingLetters = 0 ' Reset counter when we find a letter
            Else
                missingLetters = missingLetters + 1
            End If
        Else
            missingLetters = missingLetters + 1
        End If
        
        ' Stop if we encounter 2 consecutive non-letter rows (end of line items)
        If missingLetters >= 2 Then
            Exit For
        End If
    Next
    
    ' If no line letters found, log error and return empty array to skip this RO
    ' Using "DiscoverLineLetters" as the identifier in the log instead of RO number since we don't have one here
    If foundCount = 0 Then
        LogResult "DiscoverLineLetters", "ERROR: No line letters discovered - skipping RO"
        DiscoverLineLetters = Array()
        Exit Function
    End If
    
    ' Create properly sized array with found letters
    ReDim foundLetters(foundCount - 1)
    For i = 0 To foundCount - 1
        foundLetters(i) = tempLetters(i)
    Next
    
    ' Log discovered line letters for debugging
    ' Using "DiscoverLineLetters" as the identifier in the log instead of RO number since we don't have one here
    Dim lettersList
    lettersList = Join(foundLetters, ", ")
    LogResult "DiscoverLineLetters", "Discovered line letters: " & lettersList
    
    DiscoverLineLetters = foundLetters
End Function

'Subroutine to perform BlueZone automation steps for a single RO
Sub ProcessRo(RoNumber)
    Dim commands, i
    
    ' 1. Send RO Number and Enter
    bzhao.SendKey RoNumber
    bzhao.SendKey "<NumpadEnter>"


    
    ' Check for NOT ON FILE error in line 1
    Dim foundError
    foundError = CheckForROError()
    'bzhao.msgBox "Debug: CheckForROError returned " & foundError
    If foundError = "NOT ON FILE" Then
        LogResult RoNumber, "RO NOT ON FILE - Skipping to next."
        'bzhao.msgBox "RO " & RoNumber & " NOT ON FILE - Skipping to next."
        Exit Sub
    ElseIf InStr(foundError, "closed") > 0 Then
        LogResult RoNumber, "RO IS CLOSED - Skipping to next."
        'bzhao.msgBox "RO " & RoNumber & " IS CLOSED - Skipping to next."
        Exit Sub
    End If

    ' 2. Discover which line letters are present and execute CCC commands
    commands = DiscoverLineLetters()
    
    ' If no line letters discovered, log error and skip this RO
    If IsEmpty(commands) Or UBound(commands) = -1 Then
        LogResult RoNumber, "ERROR: No line letters discovered - Skipping to next."
        Exit Sub
    End If
    
    For i = 0 To UBound(commands)
        bzhao.Pause 1000 ' Pause 1 second before each command
        bzhao.SendKey "CCC " & commands(i)
        bzhao.SendKey "<Enter>"
        bzhao.Pause 2000 ' Give screen time to update
        
        ' Wait for story to close by monitoring screen text
        Call WaitForStoryClosure(commands(i))
    Next
    
    ' 3. Command E Execution (To exit/move to next record screen)
    bzhao.SendKey "E"
    bzhao.SendKey "<Enter>"
    bzhao.Pause 1000 ' Final pause to allow the screen to fully reset for the next RO
End Sub

' Waits for the story closure text to disappear from screen

' Waits for the story closure text to disappear from screen

'-----------------------------------------------------------
' Checks if line 1 contains 'NOT ON FILE' and returns True if so
'-----------------------------------------------------------
Function CheckForROError()
    Dim screenContentBuffer, screenLength
    screenLength = 80

    'Give screen a moment to update
    bzhao.Pause 2000

    ' Check line 2
    bzhao.ReadScreen screenContentBuffer, screenLength, 2, 1
    If InStr(screenContentBuffer, "NOT ON FILE") > 0 Then
        CheckForROError = "NOT ON FILE"
    ElseIf InStr(screenContentBuffer, "closed") > 0 Then
        CheckForROError = screenContentBuffer
    Else
        CheckForROError = False
    End If
End Function


'-----------------------------------------------------------
' Waits for the story closure text to disappear from screen
' storyType: "A", "B", or "C"
'-----------------------------------------------------------
Sub WaitForStoryClosure(storyType)
    Dim pollInterval, storyText
    pollInterval = 1000
    storyText = "Please close the STY, before exiting this screen"

    ' Wait for the story closure text to APPEAR (before waiting for it to disappear)
    Call WaitForTextState(storyText, pollInterval, 20000, True, "Timeout waiting for story closure message to APPEAR (Story " & storyType & ")", storyType)

    ' Now wait for the story closure text to disappear
    Call WaitForTextState(storyText, pollInterval, 60000, False, "Timeout waiting for story closure message to disappear (Story " & storyType & ")", storyType)
End Sub


' Waits for a specific text to appear or disappear on the screen within a timeout
' wantPresent: True to wait for text to appear, False to wait for it to disappear
Sub WaitForTextState(targetText, pollInterval, timeout, wantPresent, errorMsg, storyType)
    Dim elapsed, screenContentBuffer, screenLength, found
    elapsed = 0
    Do
        bzhao.Pause pollInterval
        elapsed = elapsed + pollInterval
        screenLength = 8 * 80
        bzhao.ReadScreen screenContentBuffer, screenLength, 8, 1
        found = (InStr(screenContentBuffer, targetText) > 0)
        If (wantPresent And found) Or (Not wantPresent And Not found) Then
            Exit Do
        End If
        If elapsed >= timeout Then
            MsgBox "ERROR: " & errorMsg & ". Script will exit.", vbCritical
            bzhao.StopScript
        End If
    Loop
End Sub

'-----------------------------------------------------------
' Reads the entire BlueZone screen and displays it in a MsgBox
'-----------------------------------------------------------
Sub ReadAndShowFullScreen()
    Dim screenContentBuffer, screenLength
    screenLength = 24 * 80 ' 24 rows, 80 columns
    bzhao.ReadScreen screenContentBuffer, screenLength, 1, 1
    bzhao.Pause 2000 ' Give screen time to update
    MsgBox screenContentBuffer, vbOKOnly, "Full Screen Content: " & screenContentBuffer
End Sub

'-----------------------------------------------------------
' LogResult subroutine for logging results/errors
'-----------------------------------------------------------
Sub LogResult(ro, result)
    Dim fsoLog, logFile, logPath
    logPath = GetConfigPath("Prepare_Close_Pt1", "Log")
    Set fsoLog = CreateObject("Scripting.FileSystemObject")
    Set logFile = fsoLog.OpenTextFile(logPath, 8, True)
    logFile.WriteLine Now & "  " & ro & " - Result: " & result
    logFile.Close
    Set logFile = Nothing
    Set fsoLog = Nothing
End Sub