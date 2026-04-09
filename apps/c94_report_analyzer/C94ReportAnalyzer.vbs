'=====================================================================================
' C94 Report Analyzer
' Part of the CDK DMS Automation Suite
'
' Reads a list of RO numbers from input.csv, walks terminal sequences,
' and extracts Lines A, B, C into a 4-column CSV (RO, Line A, Line B, Line C).
'=====================================================================================

Option Explicit

' --- Bootstrap ---
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_sh: Set g_sh = CreateObject("WScript.Shell")
Dim g_root: g_root = g_sh.Environment("USER")("CDK_BASE")
If g_root = "" Or Not g_fso.FolderExists(g_root) Then
    Err.Raise 53, "Bootstrap", "Invalid or missing CDK_BASE. Value: " & g_root
End If
If Not g_fso.FileExists(g_fso.BuildPath(g_root, ".cdkroot")) Then
    Err.Raise 53, "Bootstrap", "Cannot find .cdkroot in base path:" & vbCrLf & g_root
End If
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\PathHelper.vbs")).ReadAll

' --- CDK Terminal Object (must be declared before loading BZHelper) ---
Dim g_bzhao: Set g_bzhao = CreateObject("BZWhll.WhllObj")
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\BZHelper.vbs")).ReadAll

' --- Configuration ---
Dim INPUT_CSV_PATH:        INPUT_CSV_PATH        = GetConfigPath("C94ReportAnalyzer", "InputCSV")
Dim OUTPUT_CSV_PATH:       OUTPUT_CSV_PATH       = GetConfigPath("C94ReportAnalyzer", "OutputCSV")
Dim LOG_FILE_PATH:         LOG_FILE_PATH         = GetConfigPath("C94ReportAnalyzer", "Log")
Dim SCREEN_WAIT_DELAY:     SCREEN_WAIT_DELAY     = CInt(GetIniSetting("C94ReportAnalyzer", "ScreenWaitDelay", "1000"))
Dim EMPLOYEE_NUMBER:       EMPLOYEE_NUMBER       = GetIniSetting("C94ReportAnalyzer", "EmployeeNumber", "")
Dim EMPLOYEE_NAME_CONFIRM: EMPLOYEE_NAME_CONFIRM = GetIniSetting("C94ReportAnalyzer", "EmployeeNameConfirm", "")


' --- Main Script ---
Sub RunScrapper()
    Dim targetROs, csvFile, totalWritten

    ' Load target RO list
    Set targetROs = LoadTargetROs()
    If targetROs.Count = 0 Then
        LogResult "ERROR", "No target ROs loaded from " & INPUT_CSV_PATH & ". Aborting."
        MsgBox "No RO numbers found in input.csv. Aborting.", vbCritical
        Exit Sub
    End If
    LogResult "INFO", "Loaded " & targetROs.Count & " target RO(s) from " & INPUT_CSV_PATH

    ' Connect to terminal
    On Error Resume Next
    g_bzhao.Connect ""
    If Err.Number <> 0 Then
        LogResult "ERROR", "Failed to connect to BlueZone: " & Err.Description
        MsgBox "Failed to connect to BlueZone terminal session.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Initialise output CSV (overwrite)
    Set csvFile = g_fso.CreateTextFile(OUTPUT_CSV_PATH, True)
    csvFile.WriteLine "RO,Line A,Line B,Line C"

    totalWritten = 0

    ' Loop directly over the RO list — type each RO number at the prompt
    Dim roNum
    For Each roNum In targetROs.Keys

        LogResult "INFO", "Navigating to RO: " & roNum

        ' Confirm we are at the R.O. NUMBER prompt before sending
        Dim atPrompt
        atPrompt = WaitForPrompt("R.O. NUMBER", "", False, 5000, "R.O. NUMBER prompt before RO " & roNum)

        If Not atPrompt Then
            LogResult "ERROR", "R.O. NUMBER prompt not found before RO " & roNum & ". Skipping."
        Else
            ' Type the RO number at the prompt
            g_bzhao.SendKey roNum & "<NumpadEnter>"
            g_bzhao.Pause SCREEN_WAIT_DELAY

            ' Wait for the RO screen to load
            Dim screenText, startTime, screenFound
            startTime = Timer
            screenFound = False
            Do
                g_bzhao.ReadScreen screenText, 1920, 1, 1

                If InStr(1, screenText, "RO:", vbTextCompare) > 0 Or InStr(1, screenText, "RO STATUS:", vbTextCompare) > 0 Then
                    screenFound = True
                    Exit Do
                End If

                If IsKnownErrorPresent(screenText) Then
                    If DetectAndRecover() Then
                        LogResult "INFO", "Recovery successful for RO " & roNum & "."
                    Else
                        LogResult "ERROR", "Recovery failed for RO " & roNum & ". Skipping."
                    End If
                    Exit Do
                End If

                If Timer - startTime > 10 Then
                    LogResult "ERROR", "Timeout waiting for RO screen for RO " & roNum
                    Exit Do
                End If
                g_bzhao.Pause 500
            Loop

            If screenFound Then
                Dim rowData
                rowData = ScrapeCurrentRO()
                If rowData <> "" Then
                    csvFile.WriteLine rowData
                    totalWritten = totalWritten + 1
                    LogResult "INFO", "Wrote RO " & roNum & " (" & totalWritten & " of " & targetROs.Count & ")"
                End If

                ' Exit RO detail screen and confirm return to R.O. NUMBER prompt
                g_bzhao.SendKey "E<NumpadEnter>"
                If Not WaitForPrompt("R.O. NUMBER", "", False, 5000, "R.O. NUMBER prompt after RO " & roNum) Then
                    LogResult "ERROR", "R.O. NUMBER prompt not found after exiting RO " & roNum & ". Stopping to avoid corrupt input."
                    Exit For
                End If
            Else
                LogResult "ERROR", "RO " & roNum & " skipped — screen did not load."
            End If
        End If

    Next

    csvFile.Close
    LogResult "INFO", "Finished. Total rows written: " & totalWritten
    MsgBox "C94 Report Analyzer Finished." & vbCrLf & "Rows written: " & totalWritten, vbInformation
End Sub


' --- Target RO Management ---

Function LoadTargetROs()
    Dim dict, f, line
    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set f = g_fso.OpenTextFile(INPUT_CSV_PATH, 1)
    If Err.Number <> 0 Then
        LogResult "ERROR", "Cannot open input CSV: " & INPUT_CSV_PATH
        Set LoadTargetROs = dict
        Exit Function
    End If
    On Error GoTo 0
    Do Until f.AtEndOfStream
        line = Trim(f.ReadLine())
        If line <> "" And IsNumeric(line) Then
            dict(line) = False
        End If
    Loop
    f.Close
    Set LoadTargetROs = dict
End Function

Function AllFound(dict)
    Dim key
    AllFound = True
    For Each key In dict.Keys
        If Not dict(key) Then
            AllFound = False
            Exit Function
        End If
    Next
End Function


' --- Scraping Functions ---

Function ScrapeCurrentRO()
    Dim roNum, lineA, lineB, lineC

    roNum = GetROFromScreen()
    lineA = GetLineDescription("A")
    lineB = GetLineDescription("B")
    lineC = GetLineDescription("C")

    ' Sanitise commas for CSV safety
    roNum = Replace(roNum, ",", " ")
    lineA = Replace(lineA, ",", " ")
    lineB = Replace(lineB, ",", " ")
    lineC = Replace(lineC, ",", " ")

    ScrapeCurrentRO = roNum & "," & lineA & "," & lineB & "," & lineC
End Function

Function GetLineDescription(letter)
    Dim row, buf, nextColChar, foundText
    GetLineDescription = ""

    For row = 10 To 22
        g_bzhao.ReadScreen buf, 1, row, 1
        If UCase(Trim(buf)) = UCase(letter) Then
            g_bzhao.ReadScreen nextColChar, 1, row, 2
            If Asc(nextColChar) = 32 Then
                g_bzhao.ReadScreen foundText, 100, row, 4
                GetLineDescription = Left(TruncateAtDoubleSpace(Trim(foundText)), 100)
                Exit Function
            End If
        End If
    Next
End Function

Function TruncateAtDoubleSpace(text)
    Dim k
    For k = 1 To Len(text) - 1
        If Mid(text, k, 2) = "  " Then
            TruncateAtDoubleSpace = Left(text, k - 1)
            Exit Function
        End If
    Next
    TruncateAtDoubleSpace = text
End Function


' --- Shared Helpers (verbatim from PFC_Scrapper) ---

Function GetROFromScreen()
    Dim buf, re, matches
    g_bzhao.ReadScreen buf, 240, 1, 1
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "RO:?\s*(\d{4,})"
    re.IgnoreCase = True
    If re.Test(buf) Then
        Set matches = re.Execute(buf)
        GetROFromScreen = Trim(matches(0).SubMatches(0))
    Else
        re.Pattern = "(^|\s)(\d{6})(\s|$)"
        If re.Test(buf) Then
            Set matches = re.Execute(buf)
            GetROFromScreen = Trim(matches(0).SubMatches(1))
        Else
            GetROFromScreen = "UNKNOWN"
        End If
    End If
End Function

Function IsKnownErrorPresent(screenContent)
    IsKnownErrorPresent = (InStr(1, screenContent, "PRESS RETURN TO CONTINUE", vbTextCompare) > 0 Or _
                           InStr(1, screenContent, "Process is locked by", vbTextCompare) > 0 Or _
                           InStr(1, screenContent, "NOT ON FILE", vbTextCompare) > 0 Or _
                           InStr(1, screenContent, "is closed", vbTextCompare) > 0 Or _
                           InStr(1, screenContent, "ALREADY CLOSED", vbTextCompare) > 0)
End Function

Function DetectAndRecover()
    Dim screenContent
    g_bzhao.ReadScreen screenContent, 1920, 1, 1

    If InStr(1, screenContent, "PRESS RETURN TO CONTINUE", vbTextCompare) > 0 Then
        LogResult "INFO", "Error detected: VEHID not on file."
        DetectAndRecover = RecoverFromVehidError()
    ElseIf InStr(1, screenContent, "Process is locked by", vbTextCompare) > 0 Then
        LogResult "INFO", "Error detected: Process locked."
        DetectAndRecover = RecoverFromLockedProcess()
    ElseIf InStr(1, screenContent, "NOT ON FILE", vbTextCompare) > 0 Then
        LogResult "INFO", "Error detected: RO not on file. Skipping."
        DetectAndRecover = RecoverFromSkippableError()
    ElseIf InStr(1, screenContent, "is closed", vbTextCompare) > 0 Or _
           InStr(1, screenContent, "ALREADY CLOSED", vbTextCompare) > 0 Then
        LogResult "INFO", "Error detected: RO already closed. Skipping."
        DetectAndRecover = RecoverFromSkippableError()
    Else
        LogResult "ERROR", "Unrecognised screen state — no recovery handler matched."
        DetectAndRecover = False
    End If
End Function

Function RecoverFromSkippableError()
    ' The screen already shows the R.O. NUMBER prompt alongside the message.
    ' No keys needed — just wait for the screen to settle before the next RO.
    g_bzhao.Pause SCREEN_WAIT_DELAY
    RecoverFromSkippableError = True
End Function

Function RecoverFromLockedProcess()
    RecoverFromLockedProcess = False
    LogResult "INFO", "Recovery: dismissing locked process, waiting for sequence prompt."
    g_bzhao.SendKey "<Enter>"
    If Not WaitForPrompt("R.O. NUMBER", "", False, 10000, "") Then
        LogResult "ERROR", "Recovery failed: sequence prompt not found after locked process dismiss."
        Exit Function
    End If
    LogResult "INFO", "Recovery complete. Back at sequence prompt."
    RecoverFromLockedProcess = True
End Function

Function RecoverFromVehidError()
    RecoverFromVehidError = False
    If Not BZH_RecoverFromVehidError(EMPLOYEE_NUMBER, EMPLOYEE_NAME_CONFIRM, "2") Then
        LogResult "ERROR", "Recovery failed: BZH_RecoverFromVehidError returned False."
        Exit Function
    End If
    If Not WaitForPrompt("R.O. NUMBER", "", False, 10000, "sequence prompt after VEHID recovery") Then
        LogResult "ERROR", "Recovery failed: sequence prompt not found after BZH_RecoverFromVehidError."
        Exit Function
    End If
    LogResult "INFO", "Recovery complete. Back at sequence prompt."
    RecoverFromVehidError = True
End Function

Sub LogResult(ByVal level, ByVal message)
    Dim logFile
    On Error Resume Next
    Set logFile = g_fso.OpenTextFile(LOG_FILE_PATH, 8, True)
    logFile.WriteLine Now & " [" & level & "] " & message
    logFile.Close
    On Error GoTo 0
End Sub

Function GetIniSetting(section, key, defaultValue)
    Dim configPath, val
    GetIniSetting = defaultValue
    On Error Resume Next
    configPath = g_fso.BuildPath(GetRepoRoot(), "config\config.ini")
    val = ReadIniValue(configPath, section, key)
    If val <> "" Then GetIniSetting = val
    On Error GoTo 0
End Function


' --- Start Execution ---
RunScrapper
