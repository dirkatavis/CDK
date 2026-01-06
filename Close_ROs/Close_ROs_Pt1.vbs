Option Explicit

'====================================================================
' Script: Close_RO_Automation.vbs
' Description: Simplified script to read 6-digit numbers from a CSV
'              and perform BlueZone CCC operations (A, B, C) with pauses.
'              All screen checks and error handling have been removed.
'====================================================================


' --- Configuration ---
Const CSV_FILE = "C:\Temp\Code\Scripts\VBScript\CDK\Close_ROs\Close_ROs_Pt1.csv" ' Update path if needed
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

    ' 2. Execute CCC commands A, B, and C in a loop
    commands = Array("A", "B", "C")
    For i = 0 To UBound(commands)
        bzhao.Pause 300 ' Brief pause before each command (optimized from 1000ms)
        bzhao.SendKey "CCC " & commands(i)
        bzhao.SendKey "<Enter>"
        bzhao.Pause 500 ' Give screen time to update (optimized from 2000ms)
        
        ' Wait for story to close by monitoring screen text
        Call WaitForStoryClosure(commands(i))
    Next
    
    ' 3. Command E Execution (To exit/move to next record screen)
    bzhao.SendKey "E"
    bzhao.SendKey "<Enter>"
    bzhao.Pause 500 ' Final pause to allow the screen to fully reset (optimized from 1000ms)
End Sub

' Waits for the story closure text to disappear from screen

' Waits for the story closure text to disappear from screen

'-----------------------------------------------------------
' Checks if line 1 contains 'NOT ON FILE' and returns True if so
'-----------------------------------------------------------
Function CheckForROError()
    Dim screenContentBuffer, screenLength
    screenLength = 80

    'Give screen a moment to update (optimized from 2000ms)
    bzhao.Pause 500

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
    bzhao.Pause 200 ' Brief screen update delay (optimized from 2000ms)
    MsgBox screenContentBuffer, vbOKOnly, "Full Screen Content: " & screenContentBuffer
End Sub

'-----------------------------------------------------------
' LogResult subroutine for logging results/errors
'-----------------------------------------------------------
Sub LogResult(ro, result)
    Dim fsoLog, logFile, logPath
    logPath = "C:\Temp\Code\Scripts\VBScript\CDK\Close_ROs\Close_ROs_Pt1.log"
    Set fsoLog = CreateObject("Scripting.FileSystemObject")
    Set logFile = fsoLog.OpenTextFile(logPath, 8, True)
    logFile.WriteLine Now & "  " & ro & " - Result: " & result
    logFile.Close
    Set logFile = Nothing
    Set fsoLog = Nothing
End Sub