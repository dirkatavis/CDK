'=====================================================================================
' Pfc_Summary.vbs
' Scrape RO number + RO status from PFC sequence range and write CSV summary.
' Pattern intentionally mirrors apps\pfc_scrapper\PFC_Scrapper.vbs.
'=====================================================================================

Option Explicit

' --- Load PathHelper for centralized path management ---
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Const BASE_ENV_VAR_LOCAL = "CDK_BASE"

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
Dim LOG_FILE_PATH: LOG_FILE_PATH = GetConfigPath("Pfc_Summary", "Log")
Dim OUTPUT_CSV_PATH: OUTPUT_CSV_PATH = GetConfigPath("Pfc_Summary", "OutputCSV")
Dim START_SEQUENCE: START_SEQUENCE = CInt(GetIniSetting("Pfc_Summary", "StartSequenceNumber", "1"))
Dim END_SEQUENCE: END_SEQUENCE = CInt(GetIniSetting("Pfc_Summary", "EndSequenceNumber", "100"))
Dim STEP_DELAY_MS: STEP_DELAY_MS = CInt(GetIniSetting("Pfc_Summary", "StepDelayMs", "1000"))

' --- CDK Objects ---
Dim bzhao: Set bzhao = CreateObject("BZWhll.WhllObj")

Sub RunSummary()
    Dim i, totalScraped, csvFile
    totalScraped = 0

    LogResult "INFO", "Starting Pfc_Summary. Output: " & OUTPUT_CSV_PATH
    LogResult "INFO", "Sequence range: " & START_SEQUENCE & " to " & END_SEQUENCE

    On Error Resume Next
    bzhao.Connect ""
    If Err.Number <> 0 Then
        LogResult "ERROR", "Failed to connect to BlueZone: " & Err.Description
        Exit Sub
    End If
    On Error GoTo 0

    ' Give the terminal a moment to settle before starting the loop
    bzhao.Pause STEP_DELAY_MS

    Set csvFile = g_fso.CreateTextFile(OUTPUT_CSV_PATH, True)
    csvFile.WriteLine "RO_Number,Status"

    For i = START_SEQUENCE To END_SEQUENCE
        LogResult "INFO", "Processing sequence: " & i

        If Not WaitForPrompt("COMMAND:", 5) Then
            LogResult "ERROR", "Timed out waiting for COMMAND prompt at sequence " & i
            Exit For
        End If


        bzhao.Pause STEP_DELAY_MS

        bzhao.SendKey i & "<NumpadEnter>"
        bzhao.Pause STEP_DELAY_MS

        Dim screenText, startTime, screenFound
        startTime = Timer
        screenFound = False

        Do
            bzhao.ReadScreen screenText, 1920, 1, 1

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

            bzhao.Pause 500
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

            bzhao.Pause STEP_DELAY_MS
            bzhao.SendKey "E<NumpadEnter>"
            bzhao.Pause STEP_DELAY_MS
        Else
            LogResult "ERROR", "Sequence " & i & " skipped due to screen transition timeout."
        End If
    Next

    csvFile.Close
    LogResult "INFO", "Pfc_Summary finished. Total ROs scraped: " & totalScraped
End Sub

Function GetROFromScreen()
    Dim buf, re, matches

    bzhao.ReadScreen buf, 240, 1, 1

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

    bzhao.ReadScreen buf, 80, 5, 1

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "RO STATUS:\s*([A-Z\s]{1,20})"
    re.IgnoreCase = True

    If re.Test(buf) Then
        Set matches = re.Execute(buf)
        GetRepairOrderStatus = Trim(matches(0).SubMatches(0))
    Else
        bzhao.ReadScreen buf, 15, 5, 12
        If Trim(buf) = "" Then
            GetRepairOrderStatus = "UNKNOWN"
        Else
            GetRepairOrderStatus = Trim(buf)
        End If
    End If
End Function

Function WaitForPrompt(text, timeoutSec)
    Dim startTime, elapsed, screenContent
    startTime = Timer

    Do
        bzhao.ReadScreen screenContent, 1920, 1, 1
        If InStr(1, screenContent, text, vbTextCompare) > 0 Then
            WaitForPrompt = True
            Exit Function
        End If

        bzhao.Pause 500
        elapsed = Timer - startTime
    Loop While elapsed < timeoutSec

    WaitForPrompt = False
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
