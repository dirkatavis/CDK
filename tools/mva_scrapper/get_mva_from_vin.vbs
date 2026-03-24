Option Explicit

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

Dim LOG_FILE_PATH: LOG_FILE_PATH = GetConfigPath("GetMvaFromVin", "Log")
Dim DIAG_LOG_FILE_PATH: DIAG_LOG_FILE_PATH = GetConfigPath("GetMvaFromVin", "DiagnosticLog")
Dim INPUT_CSV_PATH: INPUT_CSV_PATH = GetConfigPath("GetMvaFromVin", "InputCSV")
Dim OUTPUT_CSV_PATH: OUTPUT_CSV_PATH = GetConfigPath("GetMvaFromVin", "OutputCSV")

Dim VIN_COLUMN_NAME: VIN_COLUMN_NAME = GetIniSetting("GetMvaFromVin", "VinColumn", "VIN")
Dim VIN_PROMPT_TEXT: VIN_PROMPT_TEXT = GetIniSetting("GetMvaFromVin", "VinPromptText", "VIN")
Dim RESULTS_READY_TEXT: RESULTS_READY_TEXT = GetIniSetting("GetMvaFromVin", "ResultsReadyText", "MVA")
Dim NO_RESULT_TEXT: NO_RESULT_TEXT = GetIniSetting("GetMvaFromVin", "NoResultText", "NO VEHICLE FOUND")
Dim MVA_LABEL_TEXT: MVA_LABEL_TEXT = GetIniSetting("GetMvaFromVin", "MvaLabelText", "MVA")
Dim MVA_REGEX_PATTERN: MVA_REGEX_PATTERN = GetIniSetting("GetMvaFromVin", "MvaRegex", "MVA[: ]*([A-Z0-9\-]+)")
Dim SEARCH_SUBMIT_KEY: SEARCH_SUBMIT_KEY = GetIniSetting("GetMvaFromVin", "SearchSubmitKey", "<NumpadEnter>")
Dim RETURN_TO_SEARCH_KEY: RETURN_TO_SEARCH_KEY = GetIniSetting("GetMvaFromVin", "ReturnToSearchKey", "")
Dim COMMAND_WAIT_SEC: COMMAND_WAIT_SEC = CLng(GetIniSetting("GetMvaFromVin", "CommandWaitSec", "10"))
Dim RESULTS_WAIT_SEC: RESULTS_WAIT_SEC = CLng(GetIniSetting("GetMvaFromVin", "ResultsWaitSec", "15"))
Dim POLL_MS: POLL_MS = CLng(GetIniSetting("GetMvaFromVin", "PollMs", "300"))
Dim CAPTURE_SCREEN_ON_ERROR: CAPTURE_SCREEN_ON_ERROR = ParseBoolean(GetIniSetting("GetMvaFromVin", "CaptureScreenOnError", "true"), True)
Dim CONTINUE_ON_ERROR: CONTINUE_ON_ERROR = ParseBoolean(GetIniSetting("GetMvaFromVin", "ContinueOnError", "true"), True)

Dim bzhao: Set bzhao = CreateObject("BZWhll.WhllObj")

Sub Main()
    Dim inFile, outFile, headerLine, vinColumnIndex
    Dim lineText, vinValue, mvaValue, statusValue, errorValue
    Dim totalRows, successRows, failedRows

    Call EnsureFolderForFile(LOG_FILE_PATH)
    Call EnsureFolderForFile(DIAG_LOG_FILE_PATH)
    Call EnsureFolderForFile(OUTPUT_CSV_PATH)

    Call LogResult("INFO", "Starting get_mva_from_vin")
    Call LogResult("INFO", "InputCSV=" & INPUT_CSV_PATH)
    Call LogResult("INFO", "OutputCSV=" & OUTPUT_CSV_PATH)

    If Not g_fso.FileExists(INPUT_CSV_PATH) Then
        Call LogResult("ERROR", "Input CSV not found: " & INPUT_CSV_PATH)
        WScript.Quit 1
    End If

    On Error Resume Next
    bzhao.Connect ""
    If Err.Number <> 0 Then
        Call LogResult("ERROR", "Failed to connect to BlueZone/Compass session: " & Err.Description)
        Err.Clear
        WScript.Quit 1
    End If
    On Error GoTo 0

    Set inFile = g_fso.OpenTextFile(INPUT_CSV_PATH, 1, False)
    Set outFile = g_fso.CreateTextFile(OUTPUT_CSV_PATH, True)
    outFile.WriteLine "VIN,MVA,Status,Error"

    If inFile.AtEndOfStream Then
        Call LogResult("ERROR", "Input CSV is empty")
        inFile.Close
        outFile.Close
        WScript.Quit 1
    End If

    headerLine = inFile.ReadLine
    vinColumnIndex = GetColumnIndex(headerLine, VIN_COLUMN_NAME)
    If vinColumnIndex < 0 Then
        Call LogResult("ERROR", "VIN column '" & VIN_COLUMN_NAME & "' not found in header")
        inFile.Close
        outFile.Close
        WScript.Quit 1
    End If

    totalRows = 0
    successRows = 0
    failedRows = 0

    Do While Not inFile.AtEndOfStream
        lineText = Trim(CStr(inFile.ReadLine))
        If Len(lineText) = 0 Then
            ' skip
        Else
            vinValue = Trim(GetField(lineText, vinColumnIndex))
            If Len(vinValue) = 0 Then
                Call WriteOutputRow(outFile, "", "", "SKIPPED", "Empty VIN")
                failedRows = failedRows + 1
                totalRows = totalRows + 1
            Else
                totalRows = totalRows + 1
                mvaValue = ""
                statusValue = ""
                errorValue = ""

                Call ProcessVin(vinValue, mvaValue, statusValue, errorValue)
                Call WriteOutputRow(outFile, vinValue, mvaValue, statusValue, errorValue)

                If UCase(statusValue) = "OK" Then
                    successRows = successRows + 1
                Else
                    failedRows = failedRows + 1
                    If Not CONTINUE_ON_ERROR Then
                        Call LogResult("ERROR", "Stopping run due to failure and ContinueOnError=false")
                        Exit Do
                    End If
                End If
            End If
        End If
    Loop

    inFile.Close
    outFile.Close

    Call LogResult("INFO", "Completed get_mva_from_vin. Total=" & totalRows & ", Success=" & successRows & ", Failed=" & failedRows)
End Sub

Sub ProcessVin(vinValue, ByRef mvaValue, ByRef statusValue, ByRef errorValue)
    Dim resultScreen, mvaExtracted

    Call LogResult("INFO", "Processing VIN " & vinValue)

    If Not WaitForText(VIN_PROMPT_TEXT, COMMAND_WAIT_SEC) Then
        statusValue = "ERROR"
        errorValue = "VIN prompt not found before input"
        Call HandleFailure(vinValue, errorValue)
        Exit Sub
    End If

    bzhao.SendKey vinValue & SEARCH_SUBMIT_KEY
    Call WaitMs(POLL_MS)

    If Not WaitForResultScreen(resultScreen, RESULTS_WAIT_SEC) Then
        statusValue = "TIMEOUT"
        errorValue = "Timed out waiting for Compass result"
        Call HandleFailure(vinValue, errorValue)
        Exit Sub
    End If

    If InStr(1, resultScreen, NO_RESULT_TEXT, vbTextCompare) > 0 Then
        statusValue = "NOT_FOUND"
        errorValue = "No vehicle result returned"
        mvaValue = ""
        Call LogResult("WARN", "VIN " & vinValue & " -> NOT_FOUND")
        Call MaybeReturnToSearch()
        Exit Sub
    End If

    mvaExtracted = ExtractMva(resultScreen)
    If Len(mvaExtracted) = 0 Then
        statusValue = "ERROR"
        errorValue = "MVA not found on results screen"
        Call HandleFailure(vinValue, errorValue)
        Call MaybeReturnToSearch()
        Exit Sub
    End If

    mvaValue = mvaExtracted
    statusValue = "OK"
    errorValue = ""
    Call LogResult("INFO", "VIN " & vinValue & " -> MVA " & mvaValue)

    Call MaybeReturnToSearch()
End Sub

Sub HandleFailure(vinValue, message)
    Call LogResult("ERROR", "VIN " & vinValue & " failed: " & message)
    If CAPTURE_SCREEN_ON_ERROR Then
        Call LogDiagnostic("VIN " & vinValue & " failure screen", ReadScreenAll())
    End If
End Sub

Function WaitForResultScreen(ByRef resultScreen, timeoutSec)
    Dim startedAt, elapsedSec, currentScreen
    WaitForResultScreen = False
    resultScreen = ""

    startedAt = Timer
    Do
        currentScreen = ReadScreenAll()

        If InStr(1, currentScreen, NO_RESULT_TEXT, vbTextCompare) > 0 Then
            resultScreen = currentScreen
            WaitForResultScreen = True
            Exit Function
        End If

        If InStr(1, currentScreen, RESULTS_READY_TEXT, vbTextCompare) > 0 Or InStr(1, currentScreen, MVA_LABEL_TEXT, vbTextCompare) > 0 Then
            resultScreen = currentScreen
            WaitForResultScreen = True
            Exit Function
        End If

        Call WaitMs(POLL_MS)
        elapsedSec = SecondsSince(startedAt)
    Loop While elapsedSec < timeoutSec
End Function

Function ExtractMva(screenText)
    Dim re, matches
    ExtractMva = ""

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = MVA_REGEX_PATTERN
    re.IgnoreCase = True
    re.Global = False

    If re.Test(screenText) Then
        Set matches = re.Execute(screenText)
        If matches.Count > 0 And matches(0).SubMatches.Count > 0 Then
            ExtractMva = Trim(CStr(matches(0).SubMatches(0)))
        End If
    End If
End Function

Sub MaybeReturnToSearch()
    If Len(Trim(CStr(RETURN_TO_SEARCH_KEY))) = 0 Then Exit Sub

    If WaitForText(VIN_PROMPT_TEXT, 1) Then Exit Sub

    bzhao.SendKey RETURN_TO_SEARCH_KEY
    Call WaitMs(POLL_MS)
End Sub

Function WaitForText(textToFind, timeoutSec)
    Dim startedAt, elapsedSec, screenText
    WaitForText = False

    startedAt = Timer
    Do
        screenText = ReadScreenAll()
        If InStr(1, screenText, textToFind, vbTextCompare) > 0 Then
            WaitForText = True
            Exit Function
        End If

        Call WaitMs(POLL_MS)
        elapsedSec = SecondsSince(startedAt)
    Loop While elapsedSec < timeoutSec
End Function

Function ReadScreenAll()
    Dim content
    content = ""
    On Error Resume Next
    bzhao.ReadScreen content, 1920, 1, 1
    If Err.Number <> 0 Then
        content = ""
        Err.Clear
    End If
    On Error GoTo 0
    ReadScreenAll = content
End Function

Sub WaitMs(ms)
    If ms < 0 Then ms = 0
    bzhao.Pause CLng(ms)
End Sub

Function SecondsSince(startTimer)
    Dim elapsed
    elapsed = Timer - startTimer
    If elapsed < 0 Then elapsed = elapsed + 86400
    SecondsSince = elapsed
End Function

Sub WriteOutputRow(outFile, vinValue, mvaValue, statusValue, errorValue)
    outFile.WriteLine CsvEscape(vinValue) & "," & CsvEscape(mvaValue) & "," & CsvEscape(statusValue) & "," & CsvEscape(errorValue)
End Sub

Function CsvEscape(value)
    Dim text, quoteChar
    quoteChar = Chr(34)
    text = CStr(value)
    text = Replace(text, quoteChar, quoteChar & quoteChar)
    If InStr(text, ",") > 0 Or InStr(text, quoteChar) > 0 Or InStr(text, vbCr) > 0 Or InStr(text, vbLf) > 0 Then
        CsvEscape = quoteChar & text & quoteChar
    Else
        CsvEscape = text
    End If
End Function

Function GetColumnIndex(headerLine, targetColumn)
    Dim parts, i
    GetColumnIndex = -1
    parts = Split(headerLine, ",")

    For i = 0 To UBound(parts)
        If UCase(Trim(parts(i))) = UCase(Trim(targetColumn)) Then
            GetColumnIndex = i
            Exit Function
        End If
    Next
End Function

Function GetField(csvLine, idx)
    Dim parts
    parts = Split(csvLine, ",")
    If idx >= 0 And idx <= UBound(parts) Then
        GetField = Trim(parts(idx))
    Else
        GetField = ""
    End If
End Function

Sub LogResult(level, message)
    Dim logFile
    On Error Resume Next
    Set logFile = g_fso.OpenTextFile(LOG_FILE_PATH, 8, True)
    If Err.Number = 0 Then
        logFile.WriteLine Now & " [" & level & "] " & message
        logFile.Close
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Sub LogDiagnostic(title, screenText)
    Dim diagFile
    On Error Resume Next
    Set diagFile = g_fso.OpenTextFile(DIAG_LOG_FILE_PATH, 8, True)
    If Err.Number = 0 Then
        diagFile.WriteLine "==== " & Now & " :: " & title & " ===="
        diagFile.WriteLine Replace(Replace(CStr(screenText), vbCrLf, " "), vbCr, " ")
        diagFile.WriteLine ""
        diagFile.Close
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Function ParseBoolean(value, defaultValue)
    Dim text
    text = LCase(Trim(CStr(value)))

    Select Case text
        Case "1", "true", "yes", "y", "on"
            ParseBoolean = True
        Case "0", "false", "no", "n", "off"
            ParseBoolean = False
        Case Else
            ParseBoolean = defaultValue
    End Select
End Function

Sub EnsureFolderForFile(filePath)
    Dim folderPath
    folderPath = g_fso.GetParentFolderName(filePath)
    If Len(folderPath) = 0 Then Exit Sub
    Call EnsureFolderExists(folderPath)
End Sub

Sub EnsureFolderExists(folderPath)
    Dim parentFolder

    If Len(folderPath) = 0 Then Exit Sub
    If g_fso.FolderExists(folderPath) Then Exit Sub

    parentFolder = g_fso.GetParentFolderName(folderPath)
    If Len(parentFolder) > 0 And Not g_fso.FolderExists(parentFolder) Then
        Call EnsureFolderExists(parentFolder)
    End If

    On Error Resume Next
    g_fso.CreateFolder folderPath
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Function GetIniSetting(section, key, defaultValue)
    Dim configPath, value
    GetIniSetting = defaultValue

    On Error Resume Next
    configPath = g_fso.BuildPath(GetRepoRoot(), "config\config.ini")
    value = ReadIniValue(configPath, section, key)
    If value <> "" Then
        GetIniSetting = value
    End If
    On Error GoTo 0
End Function

Main
