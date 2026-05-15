'=====================================================================================
' RO Status Summary
' Part of the CDK DMS Automation Suite - tools\ro_status_summary.vbs
'
' Purpose: Loop through PFC sequences starting at 1, scrape the RO status of each,
'          aggregate counts by status into a Scripting.Dictionary, and display a
'          formatted two-column MsgBox. Runs until the end-of-sequence sentinel is
'          detected — no configured upper bound.
'
' Speed tuning: lower StepDelayMs in config\config.ini once stable (e.g. 300ms).
'=====================================================================================

Option Explicit

' --- Bootstrap ---
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_sh:  Set g_sh  = CreateObject("WScript.Shell")
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
Dim LOG_FILE_PATH:         LOG_FILE_PATH         = GetConfigPath("RoStatusSummary", "Log")
Dim OUTPUT_CSV_PATH:       OUTPUT_CSV_PATH       = GetConfigPath("RoStatusSummary", "OutputCSV")
Dim STEP_DELAY_MS:         STEP_DELAY_MS         = CInt(GetIniSetting("RoStatusSummary", "StepDelayMs", "1000"))
Dim EMPLOYEE_NUMBER:       EMPLOYEE_NUMBER       = GetIniSetting("RoStatusSummary", "EmployeeNumber", "")
Dim EMPLOYEE_NAME_CONFIRM: EMPLOYEE_NAME_CONFIRM = GetIniSetting("RoStatusSummary", "EmployeeNameConfirm", "")

' --- Entry Point ---
Call ProcessRONumbers()


'==============================================================================
' ProcessRONumbers
' Entry point. Connects to BlueZone, gathers status counts, displays summary.
'==============================================================================
Sub ProcessRONumbers()
    Dim statusDict
    Dim csvFile
    Set statusDict = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set csvFile = g_fso.CreateTextFile(OUTPUT_CSV_PATH, True)
    If Err.Number <> 0 Then
        LogResult "ERROR", "Failed to open output CSV: " & OUTPUT_CSV_PATH & " | " & Err.Description
        MsgBox "Failed to open output CSV file." & vbCrLf & OUTPUT_CSV_PATH, vbCritical, "RO Status Summary"
        Set statusDict = Nothing
        Exit Sub
    End If
    On Error GoTo 0
    csvFile.WriteLine "MVA,Status"

    On Error Resume Next
    g_bzhao.Connect ""
    If Err.Number <> 0 Then
        LogResult "ERROR", "Failed to connect to BlueZone: " & Err.Description
        MsgBox "Failed to connect to BlueZone terminal session.", vbCritical, "RO Status Summary"
        csvFile.Close
        Set csvFile = Nothing
        Set statusDict = Nothing
        Exit Sub
    End If
    On Error GoTo 0

    Call GatherROStatuses(statusDict, csvFile)
    csvFile.Close
    Set csvFile = Nothing
    Call DisplaySummary(statusDict)

    Set statusDict = Nothing
End Sub


'==============================================================================
' GatherROStatuses
' Loops from sequence 1 until the end-of-sequence sentinel. For each sequence:
'   1. Error check  — clear any terminal interruptions before reading anything
'   2. Sentinel check — exact pattern from PostFinalCharges.vbs ProcessRONumbers
'   3. Status read  — GetRepairOrderStatus() → tally into statusDict
'   4. E+Enter      — return to COMMAND prompt
'==============================================================================
Sub GatherROStatuses(statusDict, csvFile)
    Dim roNumber: roNumber = 1

    Do
        ' Wait for COMMAND: prompt before sending next sequence
        If Not WaitForPrompt("COMMAND:", "", False, 5000, "COMMAND prompt before sequence " & roNumber) Then
            LogResult "ERROR", "Timed out waiting for COMMAND prompt at sequence " & roNumber & ". Stopping."
            Exit Sub
        End If

        g_bzhao.SendKey roNumber & "<NumpadEnter>"
        g_bzhao.Pause STEP_DELAY_MS

        ' --- 1. Error check — uses IsTextPresent (row-by-row scan, same as BZHelper/PFC) ---
        ' (1) "for this RO is not on file" — VEHID body text, checked before NOT ON FILE
        '     because the VEHID screen also contains that string.
        ' (2) "PRESS RETURN TO CONTINUE" — all-caps VEHID variant.
        ' (3) "NOT ON FILE" — separate, no VEHID involvement.
        ' (4) "is closed" / "ALREADY CLOSED"
        ' (5) "Process is locked by"
        Dim skipToNext: skipToNext = False

        If IsTextPresent("for this RO is not on file") Then
            LogResult "INFO", "Seq " & roNumber & ": VEHID not on file — recovering via full PFC re-auth."
            Call TallyStatus(statusDict, "DEFECTIVE RO")
            Call WriteStatusCsvRow(csvFile, "UNKNOWN", "DEFECTIVE RO")
            Call RecoverFromVehidError()
            skipToNext = True

        ElseIf IsTextPresent("PRESS RETURN TO CONTINUE") Then
            LogResult "INFO", "Seq " & roNumber & ": VEHID error (all-caps) — recovering via full PFC re-auth."
            Call TallyStatus(statusDict, "DEFECTIVE RO")
            Call WriteStatusCsvRow(csvFile, "UNKNOWN", "DEFECTIVE RO")
            Call RecoverFromVehidError()
            skipToNext = True

        ElseIf IsTextPresent("NOT ON FILE") Then
            Call TallyStatus(statusDict, "NOT ON FILE")
            Call WriteStatusCsvRow(csvFile, "UNKNOWN", "NOT ON FILE")
            Call RecoverFromSkippableError()
            skipToNext = True

        ElseIf IsTextPresent("is closed") Or IsTextPresent("ALREADY CLOSED") Then
            Call TallyStatus(statusDict, "ALREADY CLOSED")
            Call WriteStatusCsvRow(csvFile, "UNKNOWN", "ALREADY CLOSED")
            Call RecoverFromSkippableError()
            skipToNext = True

        ElseIf IsTextPresent("Process is locked by") Then
            If Not RecoverFromLockedProcess() Then
                LogResult "ERROR", "Recovery failed at sequence " & roNumber & ". Stopping."
                Exit Sub
            End If
        End If

        If Not skipToNext Then
            ' --- 2. Sentinel check (exact pattern from PostFinalCharges.vbs) ---
            If IsTextPresent("SEQUENCE NUMBER " & roNumber & " DOES NOT EXIST") Then
                Call LogEvent("maj", "low", "End of sequence detected", "ProcessRONumbers", "SEQUENCE NUMBER " & roNumber & " DOES NOT EXIST", "Stopping script")
                Exit Sub
            End If

            ' --- 3. Status read and tally ---
            Dim mva: mva = GetMVAFromScreen()
            Dim roStatus: roStatus = GetRepairOrderStatus()
            Call TallyStatus(statusDict, roStatus)
            Call WriteStatusCsvRow(csvFile, mva, roStatus)
            LogResult "INFO", "Seq " & roNumber & " MVA " & mva & ": " & roStatus

            ' --- 4. Return to COMMAND prompt ---
            g_bzhao.SendKey "E<NumpadEnter>"
            g_bzhao.Pause STEP_DELAY_MS
        End If

        roNumber = roNumber + 1
    Loop
End Sub

Sub WriteStatusCsvRow(csvFile, mva, status)
    Dim safeMva: safeMva = CsvSafe(mva)
    Dim safeStatus: safeStatus = CsvSafe(status)
    csvFile.WriteLine safeMva & "," & safeStatus
End Sub

Function CsvSafe(value)
    Dim text: text = CStr(value)
    If InStr(text, Chr(34)) > 0 Then text = Replace(text, Chr(34), Chr(34) & Chr(34))
    If InStr(text, ",") > 0 Or InStr(text, Chr(34)) > 0 Then
        CsvSafe = Chr(34) & text & Chr(34)
    Else
        CsvSafe = text
    End If
End Function


'==============================================================================
' DisplaySummary
' Formats statusDict as a padded two-column table and shows it via MsgBox.
' State names are padded to COL_WIDTH chars using Space() for alignment.
'==============================================================================
Sub DisplaySummary(statusDict)
    Const COL_WIDTH = 22
    Dim k, count, total, pad, lines

    total = 0
    lines = "RO Status Summary" & vbCrLf & _
            String(COL_WIDTH + 8, "-") & vbCrLf

    For Each k In statusDict.Keys
        count = statusDict(k)
        total = total + count
        pad = COL_WIDTH - Len(CStr(k))
        If pad < 1 Then pad = 1
        lines = lines & k & Space(pad) & count & vbCrLf
    Next

    lines = lines & String(COL_WIDTH + 8, "-") & vbCrLf
    pad = COL_WIDTH - Len("Total")
    lines = lines & "Total" & Space(pad) & total

    MsgBox lines, vbInformation, "RO Status Summary"
End Sub


'==============================================================================
' TallyStatus
' Increments the count for statusKey in dict, adding the key if new.
' Normalises empty/blank keys to "UNKNOWN".
'==============================================================================
Sub TallyStatus(dict, statusKey)
    If Len(Trim(statusKey)) = 0 Then statusKey = "UNKNOWN"
    If dict.Exists(statusKey) Then
        dict(statusKey) = dict(statusKey) + 1
    Else
        dict.Add statusKey, 1
    End If
End Sub


'==============================================================================
' LogEvent stub
' Matches the 6-arg signature used in PostFinalCharges.vbs.
' Routes through LogResult so recovery events appear in the log file.
'==============================================================================
Sub LogEvent(criticality, verbosity, headline, stage, reason, technical)
    Dim msg: msg = headline
    If Len(Trim(reason)) > 0 Then msg = msg & " | " & reason
    LogResult UCase(criticality), msg
End Sub


' --- Status Scraping ---
' Sourced from apps\post_final_charges\Pfc_Summary.vbs

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

Function GetMVAFromScreen()
    Dim buf, re, matches
    g_bzhao.ReadScreen buf, 240, 1, 1

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "MVA:?\s*([A-Z0-9\-]{4,})"
    re.IgnoreCase = True
    If re.Test(buf) Then
        Set matches = re.Execute(buf)
        GetMVAFromScreen = Trim(matches(0).SubMatches(0))
        Exit Function
    End If

    ' Fallback: when MVA label is unavailable, use the same header identifier used by current logging.
    GetMVAFromScreen = GetROFromScreen()
End Function

Function GetRepairOrderStatus()
    ' "RO STATUS: " is always at row 5, col 1 (11 chars), so the value is always at col 12.
    Dim buf
    g_bzhao.ReadScreen buf, 15, 5, 12
    GetRepairOrderStatus = Trim(buf)
End Function


' --- Error Detection and Recovery ---
' Sourced from apps\c94_report_analyzer\C94ReportAnalyzer.vbs.
' Recovery functions wait for COMMAND:(SEQ# (sequence context, not R.O. NUMBER).

Function IsKnownErrorPresent(screenContent)
    IsKnownErrorPresent = (InStr(1, screenContent, "PRESS RETURN TO CONTINUE", vbTextCompare) > 0 Or _
                           InStr(1, screenContent, "Process is locked by",     vbTextCompare) > 0 Or _
                           InStr(1, screenContent, "NOT ON FILE",              vbTextCompare) > 0 Or _
                           InStr(1, screenContent, "is closed",                vbTextCompare) > 0 Or _
                           InStr(1, screenContent, "ALREADY CLOSED",           vbTextCompare) > 0)
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
    g_bzhao.Pause STEP_DELAY_MS
    RecoverFromSkippableError = True
End Function

Function RecoverFromLockedProcess()
    RecoverFromLockedProcess = False
    LogResult "INFO", "Recovery: dismissing locked process, waiting for COMMAND prompt."
    g_bzhao.SendKey "<Enter>"
    If Not WaitForPrompt("COMMAND:(SEQ#", "", False, 10000, "") Then
        LogResult "ERROR", "Recovery failed: COMMAND prompt not found after locked process dismiss."
        Exit Function
    End If
    LogResult "INFO", "Recovery complete. Back at COMMAND prompt."
    RecoverFromLockedProcess = True
End Function

Function RecoverFromVehidError()
    RecoverFromVehidError = False
    If Not BZH_RecoverFromVehidError(EMPLOYEE_NUMBER, EMPLOYEE_NAME_CONFIRM, "2") Then
        LogResult "ERROR", "Recovery failed: BZH_RecoverFromVehidError returned False."
        Exit Function
    End If
    If Not WaitForPrompt("COMMAND:(SEQ#", "", False, 10000, "COMMAND prompt after VEHID recovery") Then
        LogResult "ERROR", "Recovery failed: COMMAND prompt not found after VEHID recovery."
        Exit Function
    End If
    LogResult "INFO", "Recovery complete. Back at COMMAND prompt."
    RecoverFromVehidError = True
End Function


' --- Logging ---

Sub LogResult(ByVal level, ByVal message)
    Dim logFile
    On Error Resume Next
    Set logFile = g_fso.OpenTextFile(LOG_FILE_PATH, 8, True)
    logFile.WriteLine Now & " [" & level & "] " & message
    logFile.Close
    Set logFile = Nothing
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
