'=====================================================================================
' Pfc_Summary.vbs
' Scrape RO number + RO status from PFC sequence range and write CSV summary.
' Pattern intentionally mirrors apps\pfc_scrapper\PFC_Scrapper.vbs.
'=====================================================================================

Option Explicit

' --- Bootstrap ---
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_sh: Set g_sh = CreateObject("WScript.Shell")
Dim g_root: g_root = g_sh.Environment("USER")("CDK_BASE")
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\PathHelper.vbs")).ReadAll

' --- CDK Terminal Object (must be declared before loading BZHelper) ---
Dim g_bzhao: Set g_bzhao = CreateObject("BZWhll.WhllObj")
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\BZHelper.vbs")).ReadAll

' --- Configuration ---
Dim LOG_FILE_PATH: LOG_FILE_PATH = GetConfigPath("Pfc_Summary", "Log")
Dim OUTPUT_CSV_PATH: OUTPUT_CSV_PATH = GetConfigPath("Pfc_Summary", "OutputCSV")
Dim START_SEQUENCE: START_SEQUENCE = CInt(GetIniSetting("Pfc_Summary", "StartSequenceNumber", "1"))
Dim END_SEQUENCE: END_SEQUENCE = CInt(GetIniSetting("Pfc_Summary", "EndSequenceNumber", "100"))
Dim STEP_DELAY_MS: STEP_DELAY_MS = CInt(GetIniSetting("Pfc_Summary", "StepDelayMs", "1000"))


Sub RunSummary()
    Dim i, totalScraped, csvFile
    totalScraped = 0

    LogResult "INFO", "Starting Pfc_Summary. Output: " & OUTPUT_CSV_PATH
    LogResult "INFO", "Sequence range: " & START_SEQUENCE & " to " & END_SEQUENCE

    On Error Resume Next
    g_bzhao.Connect ""
    If Err.Number <> 0 Then
        LogResult "ERROR", "Failed to connect to BlueZone: " & Err.Description
        Exit Sub
    End If
    On Error GoTo 0

    ' Give the terminal a moment to settle before starting the loop
    g_bzhao.Pause STEP_DELAY_MS

    Set csvFile = g_fso.CreateTextFile(OUTPUT_CSV_PATH, True)
    csvFile.WriteLine "RO_Number,Status"

    For i = START_SEQUENCE To END_SEQUENCE
        LogResult "INFO", "Processing sequence: " & i

        If Not WaitForPrompt("COMMAND:", "", False, 5000, "") Then
            LogResult "ERROR", "Timed out waiting for COMMAND prompt at sequence " & i
            Exit For
        End If


        g_bzhao.Pause STEP_DELAY_MS

        g_bzhao.SendKey i & "<NumpadEnter>"
        g_bzhao.Pause STEP_DELAY_MS

        Dim screenText, startTime, screenFound
        startTime = Timer
        screenFound = False

        Do
            g_bzhao.ReadScreen screenText, 1920, 1, 1

            If InStr(1, screenText, "DOES NOT EXIST", vbTextCompare) > 0 Then
                LogResult "INFO", "Reached end of sequence at " & i & ". Termination signal detected."
                csvFile.Close
                LogResult "INFO", "Pfc_Summary finished. Total ROs scraped: " & totalScraped
                Exit Sub
            End If

            If InStr(1, screenText, "RO:", vbTextCompare) > 0 Or InStr(1, screenText, "RO STATUS:", vbTextCompare) > 0 Then
                screenFound = True
                Exit Do
            End If

            If Timer - startTime > 10 Then
                LogResult "ERROR", "Timeout waiting for RO screen at sequence " & i
                Exit Do
            End If

            g_bzhao.Pause 500
        Loop

        If screenFound Then
            Dim roNumber, roStatus
            roNumber = GetROFromScreen()
            roStatus = GetRepairOrderStatus()

            roNumber = Replace(roNumber, ",", " ")
            roStatus = Replace(roStatus, ",", " ")

            csvFile.WriteLine roNumber & "," & roStatus
            totalScraped = totalScraped + 1
            LogResult "INFO", "Wrote row: " & roNumber & "," & roStatus

            g_bzhao.Pause STEP_DELAY_MS
            g_bzhao.SendKey "E<NumpadEnter>"
            g_bzhao.Pause STEP_DELAY_MS
        Else
            LogResult "ERROR", "Sequence " & i & " skipped due to screen transition timeout."
        End If
    Next

    csvFile.Close
    LogResult "INFO", "Pfc_Summary finished. Total ROs scraped: " & totalScraped
End Sub

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

Function GetRepairOrderStatus()
    Dim buf, re, matches

    g_bzhao.ReadScreen buf, 80, 5, 1

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "RO STATUS:\s*([A-Z\s]{1,20})"
    re.IgnoreCase = True

    If re.Test(buf) Then
        Set matches = re.Execute(buf)
        GetRepairOrderStatus = Trim(matches(0).SubMatches(0))
    Else
        g_bzhao.ReadScreen buf, 15, 5, 12
        If Trim(buf) = "" Then
            GetRepairOrderStatus = "UNKNOWN"
        Else
            GetRepairOrderStatus = Trim(buf)
        End If
    End If
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

RunSummary
