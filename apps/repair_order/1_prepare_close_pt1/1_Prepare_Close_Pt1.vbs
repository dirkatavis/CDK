Option Explicit

'====================================================================
' Script: 1_Prepare_Close_Pt1.vbs
' Description: Localized copy of 2_Prepare_Close_Pt1.vbs. Reads RO numbers
'              from a CSV and performs BlueZone CCC operations (A, B, C)
'              with minimal checks. Intended for localized use in
'              apps/repair_order/1_prepare_close_pt1.
'====================================================================

' --- Bootstrap ---
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_sh: Set g_sh = CreateObject("WScript.Shell")
Dim g_root: g_root = g_sh.Environment("USER")("CDK_BASE")
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\PathHelper.vbs")).ReadAll

' --- Configuration ---
Dim CSV_FILE: CSV_FILE = GetConfigPath("Prepare_Close_Pt1", "CSV")
Dim RO_SCREEN_LOAD_TIMEOUT_MS: RO_SCREEN_LOAD_TIMEOUT_MS = GetConfigInt("Prepare_Close_Pt1", "RoScreenLoadTimeoutMs", 25000)
Dim RO_SCREEN_POLL_MS: RO_SCREEN_POLL_MS = GetConfigInt("Prepare_Close_Pt1", "RoScreenPollMs", 1000)
Const NUM_COLUMN = 0 ' Kept for clarity

' --- VBScript Objects ---
Dim tsInput
Dim g_bzhao ' BlueZone Host Access Object

' --- Variables ---
Dim strLine ' Holds the entire line from the file, which is the RO number
Dim RoNumber

' --- Main Execution ---
' Connect to BlueZone
Set g_bzhao = CreateObject("BZWhll.WhllObj")
Dim connResult: connResult = g_g_bzhao.Connect("")
If connResult <> 0 Then
    MsgBox "Error: Could not connect to BlueZone session. Ensure BlueZone is open and active.", vbCritical, "Connection Failed"
    WScript.Quit 1
End If

' Check and Open Input File
If g_fso.FileExists(CSV_FILE) Then
    Set tsInput = g_fso.OpenTextFile(CSV_FILE, 1) ' 1=ForReading
Else
    MsgBox "Error: Input file not found at " & CSV_FILE, vbCritical
        g_bzhao.StopScript
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
g_bzhao.Disconnect

Set tsInput = Nothing
Set g_fso = Nothing
Set g_bzhao = Nothing

' --- Subroutines ---

Function GetConfigInt(sectionName, keyName, defaultValue)
    Dim valueText
    On Error Resume Next
    valueText = GetConfigPath(sectionName, keyName)
    If Err.Number <> 0 Then
        Err.Clear
        GetConfigInt = defaultValue
        Exit Function
    End If
    On Error GoTo 0

    If IsNumeric(valueText) Then
        GetConfigInt = CLng(valueText)
    Else
        GetConfigInt = defaultValue
    End If
End Function

Function HasAnyLineLettersOnScreen()
    Dim i, row, capturedLetter, nextColChar, screenContentBuffer
    HasAnyLineLettersOnScreen = False

    For i = 0 To 12
        row = 10 + i

        On Error Resume Next
        g_bzhao.ReadScreen screenContentBuffer, 1, row, 1
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
                g_bzhao.ReadScreen nextColChar, 1, row, 2
                If Err.Number <> 0 Then
                    Err.Clear
                    nextColChar = ""
                End If
                On Error GoTo 0

                If Len(nextColChar) > 0 And Asc(nextColChar) = 32 Then
                    HasAnyLineLettersOnScreen = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Function WaitForRoDetailScreen(ro)
    Dim elapsedMs, foundError
    elapsedMs = 0

    Do While elapsedMs < RO_SCREEN_LOAD_TIMEOUT_MS
        foundError = CheckForROError()
        If foundError = "NOT ON FILE" Or InStr(foundError, "closed") > 0 Then
            WaitForRoDetailScreen = foundError
            Exit Function
        End If

        If HasAnyLineLettersOnScreen() Then
            WaitForRoDetailScreen = "READY"
            Exit Function
        End If

        g_bzhao.Pause RO_SCREEN_POLL_MS
        elapsedMs = elapsedMs + RO_SCREEN_POLL_MS
    Loop

    WaitForRoDetailScreen = "TIMEOUT"
End Function

' DiscoverLineLetters: Detects which line letters (A, B, C, etc.) are present
Function DiscoverLineLetters()
    Dim maxLinesToCheck, i, capturedLetter, screenContentBuffer, readLength
    Dim foundLetters, foundCount
    Dim startReadRow, startReadColumn
    Dim missingLetters, nextColChar
    Dim tempLetters(25)
    foundCount = 0
    maxLinesToCheck = 10
    missingLetters = 0
    Dim startRow
    startRow = 10 ' Anchor at first actual data row (skip header rows)
    For i = 0 To maxLinesToCheck - 1
        startReadRow = startRow + i
        startReadColumn = 1
        readLength = 1
        On Error Resume Next
        g_bzhao.ReadScreen screenContentBuffer, readLength, startReadRow, startReadColumn
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        On Error GoTo 0
        capturedLetter = Trim(screenContentBuffer)
        If Len(capturedLetter) = 1 Then
            If Asc(UCase(capturedLetter)) >= Asc("A") And Asc(UCase(capturedLetter)) <= Asc("Z") Then
                ' Peek column 2 to ensure this is a line letter (typical form: "A  DESCRIPTION")
                nextColChar = ""
                On Error Resume Next
                g_bzhao.ReadScreen nextColChar, 1, startReadRow, startReadColumn + 1
                If Err.Number <> 0 Then
                    Err.Clear
                    nextColChar = ""
                End If
                On Error GoTo 0

                If Len(nextColChar) > 0 And Asc(nextColChar) = 32 Then
                    tempLetters(foundCount) = UCase(capturedLetter)
                    foundCount = foundCount + 1
                    missingLetters = 0
                Else
                    missingLetters = missingLetters + 1
                End If
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
    Dim commands, i, screenState
    g_bzhao.SendKey RoNumber
    g_bzhao.SendKey "<NumpadEnter>"

    screenState = WaitForRoDetailScreen(RoNumber)

    If screenState = "NOT ON FILE" Then
        LogResult RoNumber, "RO NOT ON FILE - Skipping to next."
        Exit Sub
    ElseIf InStr(screenState, "closed") > 0 Then
        LogResult RoNumber, "RO IS CLOSED - Skipping to next."
        Exit Sub
    ElseIf screenState = "TIMEOUT" Then
        LogResult RoNumber, "TIMEOUT waiting for RO detail screen to load - Skipping to next."
        Exit Sub
    End If

    commands = DiscoverLineLetters()
    If IsEmpty(commands) Or UBound(commands) = -1 Then
        LogResult RoNumber, "ERROR: No line letters discovered - Skipping to next."
        Exit Sub
    End If
    For i = 0 To UBound(commands)
        g_bzhao.Pause 1000
        g_bzhao.SendKey "CCC " & commands(i)
        g_bzhao.SendKey "<Enter>"
        g_bzhao.Pause 2000
        Call WaitForStoryClosure(commands(i))
    Next
    g_bzhao.SendKey "E"
    g_bzhao.SendKey "<Enter>"
    g_bzhao.Pause 1000
End Sub

' Checks if line 2 contains 'NOT ON FILE' or 'closed'
Function CheckForROError()
    Dim screenContentBuffer, screenLength
    screenLength = 80
    g_bzhao.Pause 2000
    g_bzhao.ReadScreen screenContentBuffer, screenLength, 2, 1
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
        g_bzhao.Pause pollInterval
        elapsed = elapsed + pollInterval
        screenLength = 8 * 80
        g_bzhao.ReadScreen screenContentBuffer, screenLength, 8, 1
        found = (InStr(screenContentBuffer, targetText) > 0)
        If (wantPresent And found) Or (Not wantPresent And Not found) Then
            Exit Do
        End If
        If elapsed >= timeout Then
            MsgBox "ERROR: " & errorMsg & ". Script will exit.", vbCritical
            g_bzhao.StopScript
        End If
    Loop
End Sub

' Reads the entire BlueZone screen and displays it in a MsgBox
Sub ReadAndShowFullScreen()
    Dim screenContentBuffer, screenLength
    screenLength = 24 * 80
    g_bzhao.ReadScreen screenContentBuffer, screenLength, 1, 1
    g_bzhao.Pause 2000
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
