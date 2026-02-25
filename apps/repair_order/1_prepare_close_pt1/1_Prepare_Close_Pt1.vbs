Option Explicit

'====================================================================
' Script: 1_Prepare_Close_Pt1.vbs
' Description: Localized copy of 2_Prepare_Close_Pt1.vbs. Reads RO numbers
'              from a CSV and performs BlueZone CCC operations (A, B, C)
'              with minimal checks. Intended for localized use in
'              apps/repair_order/1_prepare_close_pt1.
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
Const NUM_COLUMN = 0 ' Kept for clarity

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
Dim connResult: connResult = bzhao.Connect("")
If connResult <> 0 Then
    MsgBox "Error: Could not connect to BlueZone session. Ensure BlueZone is open and active.", vbCritical, "Connection Failed"
    WScript.Quit 1
End If

' Check and Open Input File
If fso.FileExists(CSV_FILE) Then
    Set tsInput = fso.OpenTextFile(CSV_FILE, 1) ' 1=ForReading
Else
    MsgBox "Error: Input file not found at " & CSV_FILE, vbCritical
        bzhao.StopScript
End If

' Process Records
Do While Not tsInput.AtEndOfStream
    strLine = tsInput.ReadLine
    RoNumber = Trim(strLine)
    If Len(RoNumber) = 6 And IsNumeric(RoNumber) Then
        Call ProcessRo(RoNumber)
    End If
Loop

' Cleanup
tsInput.Close
bzhao.Disconnect

Set tsInput = Nothing
Set fso = Nothing
Set bzhao = Nothing

' --- Subroutines ---

' DiscoverLineLetters: Detects which line letters (A, B, C, etc.) are present
Function DiscoverLineLetters()
    Dim maxLinesToCheck, i, capturedLetter, screenContentBuffer, readLength
    Dim foundLetters, foundCount
    Dim startReadRow, startReadColumn
    Dim missingLetters
    Dim tempLetters(25)
    foundCount = 0
    maxLinesToCheck = 10
    missingLetters = 0
    Dim startRow
    startRow = 7
    For i = 0 To maxLinesToCheck - 1
        startReadRow = startRow + i
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
                tempLetters(foundCount) = UCase(capturedLetter)
                foundCount = foundCount + 1
                missingLetters = 0
            Else
                missingLetters = missingLetters + 1
            End If
        Else
            missingLetters = missingLetters + 1
        End If
        If missingLetters >= 2 Then
            Exit For
        End If
    Next
    If foundCount = 0 Then
        LogResult "DiscoverLineLetters", "ERROR: No line letters discovered - skipping RO"
        DiscoverLineLetters = Array()
        Exit Function
    End If
    ReDim foundLetters(foundCount - 1)
    For i = 0 To foundCount - 1
        foundLetters(i) = tempLetters(i)
    Next
    Dim lettersList
    lettersList = Join(foundLetters, ", ")
    LogResult "DiscoverLineLetters", "Discovered line letters: " & lettersList
    DiscoverLineLetters = foundLetters
End Function

' Subroutine to perform BlueZone automation steps for a single RO
Sub ProcessRo(RoNumber)
    Dim commands, i
    bzhao.SendKey RoNumber
    bzhao.SendKey "<NumpadEnter>"
    Dim foundError
    foundError = CheckForROError()
    If foundError = "NOT ON FILE" Then
        LogResult RoNumber, "RO NOT ON FILE - Skipping to next."
        Exit Sub
    ElseIf InStr(foundError, "closed") > 0 Then
        LogResult RoNumber, "RO IS CLOSED - Skipping to next."
        Exit Sub
    End If
    commands = DiscoverLineLetters()
    If IsEmpty(commands) Or UBound(commands) = -1 Then
        LogResult RoNumber, "ERROR: No line letters discovered - Skipping to next."
        Exit Sub
    End If
    For i = 0 To UBound(commands)
        bzhao.Pause 1000
        bzhao.SendKey "CCC " & commands(i)
        bzhao.SendKey "<Enter>"
        bzhao.Pause 2000
        Call WaitForStoryClosure(commands(i))
    Next
    bzhao.SendKey "E"
    bzhao.SendKey "<Enter>"
    bzhao.Pause 1000
End Sub

' Checks if line 2 contains 'NOT ON FILE' or 'closed'
Function CheckForROError()
    Dim screenContentBuffer, screenLength
    screenLength = 80
    bzhao.Pause 2000
    bzhao.ReadScreen screenContentBuffer, screenLength, 2, 1
    If InStr(screenContentBuffer, "NOT ON FILE") > 0 Then
        CheckForROError = "NOT ON FILE"
    ElseIf InStr(screenContentBuffer, "closed") > 0 Then
        CheckForROError = screenContentBuffer
    Else
        CheckForROError = False
    End If
End Function

' Waits for the story closure text to disappear from screen
Sub WaitForStoryClosure(storyType)
    Dim pollInterval, storyText
    pollInterval = 1000
    storyText = "Please close the STY, before exiting this screen"
    Call WaitForTextState(storyText, pollInterval, 20000, True, "Timeout waiting for story closure message to APPEAR (Story " & storyType & ")", storyType)
    Call WaitForTextState(storyText, pollInterval, 60000, False, "Timeout waiting for story closure message to disappear (Story " & storyType & ")", storyType)
End Sub

' Waits for a specific text to appear or disappear on the screen within a timeout
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

' Reads the entire BlueZone screen and displays it in a MsgBox
Sub ReadAndShowFullScreen()
    Dim screenContentBuffer, screenLength
    screenLength = 24 * 80
    bzhao.ReadScreen screenContentBuffer, screenLength, 1, 1
    bzhao.Pause 2000
    MsgBox screenContentBuffer, vbOKOnly, "Full Screen Content: " & screenContentBuffer
End Sub

' LogResult subroutine for logging results/errors
Sub LogResult(ro, result)
    Dim fsoLog, logFile, logPath
    logPath = GetConfigPath("Prepare_Close_Pt1", "Log")
    Set fsoLog = CreateObject("Scripting.FileSystemObject")
    Dim logDir: logDir = fsoLog.GetParentFolderName(logPath)
    If Not fsoLog.FolderExists(logDir) Then
        fsoLog.CreateFolder(logDir)
    End If
    Set logFile = fsoLog.OpenTextFile(logPath, 8, True)
    logFile.WriteLine Now & "  " & ro & " - Result: " & result
    logFile.Close
    Set logFile = Nothing
    Set fsoLog = Nothing
End Sub
