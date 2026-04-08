'=====================================================================================
' PFC Scrapper
' Part of the CDK DMS Automation Suite
'
' Strategic Context: Legacy system scheduled for retirement in 3-6 months.
' Purpose: Scrape RO details from sequences until "DOES NOT EXIST" is encountered.
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
Dim LOG_FILE_PATH: LOG_FILE_PATH = GetConfigPath("PFC_Scrapper", "Log")
Dim OUTPUT_CSV_PATH: OUTPUT_CSV_PATH = GetConfigPath("PFC_Scrapper", "OutputCSV")
Dim SCREEN_WAIT_DELAY: SCREEN_WAIT_DELAY = CInt(GetIniSetting("PFC_Scrapper", "ScreenWaitDelay", "1000"))
Dim START_SEQUENCE: START_SEQUENCE = CInt(GetIniSetting("PFC_Scrapper", "StartSequence", "1"))
Dim SKIP_SEQUENCES: SKIP_SEQUENCES = GetIniSetting("PFC_Scrapper", "SkipSequences", "")
Dim EMPLOYEE_NUMBER: EMPLOYEE_NUMBER = GetIniSetting("PFC_Scrapper", "EmployeeNumber", "")
Dim EMPLOYEE_NAME_CONFIRM: EMPLOYEE_NAME_CONFIRM = GetIniSetting("PFC_Scrapper", "EmployeeNameConfirm", "")


' --- Main Script ---
Sub RunScrapper()
    Dim i, totalScraped, csvFile
    totalScraped = 0
    i = START_SEQUENCE

    LogResult "INFO", "Starting PFC Scrapper. Output: " & OUTPUT_CSV_PATH

    ' Connect to terminal
    On Error Resume Next
    g_bzhao.Connect ""
    If Err.Number <> 0 Then
        LogResult "ERROR", "Failed to connect to BlueZone: " & Err.Description
        MsgBox "Failed to connect to BlueZone terminal session.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Initialize CSV (Overwrite)
    Set csvFile = g_fso.CreateTextFile(OUTPUT_CSV_PATH, True)
    csvFile.WriteLine "RO number, Tech ID, RO status, Line A, Line B, Line C, Open Date"

    Dim abortAll: abortAll = False

    Do
        If abortAll Then Exit Do

        LogResult "INFO", "Processing sequence: " & i

        ' Skip logic
        If ShouldSkipSequence(i) Then
            LogResult "INFO", "Skipping sequence " & i & " as per config."
            i = i + 1
        Else
            ' Ensure we are at COMMAND prompt — recover if security menu is showing
            If Not WaitForPrompt("COMMAND:", "", False, 5000, "") Then
                LogResult "ERROR", "Timed out waiting for COMMAND prompt at sequence " & i
                If DetectAndRecover() Then
                    LogResult "INFO", "Recovery successful. Retrying sequence " & i & "."
                Else
                    LogResult "ERROR", "Recovery failed. Exiting."
                    abortAll = True
                End If
            End If

            If abortAll Then Exit Do

            ' Enter sequence number
            g_bzhao.SendKey i & "<NumpadEnter>"
            g_bzhao.Pause SCREEN_WAIT_DELAY

            ' Wait for state change - either RO screen, security menu, or error
            Dim screenText, startTime, screenFound
            startTime = Timer
            screenFound = False
            Do
                g_bzhao.ReadScreen screenText, 1920, 1, 1
                If InStr(1, screenText, "DOES NOT EXIST", vbTextCompare) > 0 Then
                    LogResult "INFO", "Reached end of sequence at " & i & ". Termination signal detected."
                    csvFile.Close
                    LogResult "INFO", "Scrapper finished. Total ROs scraped: " & totalScraped
                    MsgBox "PFC Scraper Finished." & vbCrLf & "Total Scraped: " & totalScraped, vbInformation
                    Exit Sub
                End If

                ' Look for RO header or status line as confirmation we are in an RO
                If InStr(1, screenText, "RO:", vbTextCompare) > 0 Or InStr(1, screenText, "RO STATUS:", vbTextCompare) > 0 Then
                    screenFound = True
                    Exit Do ' Proceed to scrape
                End If

                ' Detect known error conditions — recover and skip this sequence
                If IsKnownErrorPresent(screenText) Then
                    If DetectAndRecover() Then
                        LogResult "INFO", "Recovery successful. Sequence " & i & " will be skipped."
                    Else
                        LogResult "ERROR", "Recovery failed. Exiting."
                        abortAll = True
                    End If
                    Exit Do
                End If

                If Timer - startTime > 10 Then
                    LogResult "ERROR", "Timeout waiting for RO screen at sequence " & i
                    Exit Do
                End If
                g_bzhao.Pause 500
            Loop

            If abortAll Then Exit Do

            If screenFound Then
                ' Scrape Data
                Dim roData
                roData = ScrapeCurrentRO()

                If roData <> "" Then
                    csvFile.WriteLine roData
                    totalScraped = totalScraped + 1
                End If

                ' Return to command prompt
                g_bzhao.SendKey "E<NumpadEnter>"
                g_bzhao.Pause SCREEN_WAIT_DELAY
            Else
                LogResult "ERROR", "Sequence " & i & " skipped due to screen transition timeout."
            End If

            i = i + 1
        End If
    Loop

    csvFile.Close
    LogResult "INFO", "Scrapper finished. Total ROs scraped: " & totalScraped
    MsgBox "PFC Scraper Finished." & vbCrLf & "Total Scraped: " & totalScraped, vbInformation
End Sub

' --- Scraping functions ---

Function ScrapeCurrentRO()
    Dim roNum, roStatus, lineA, lineB, lineC, openDate, techId
    
    ' Scrape Header (RO and Date)
    roNum = GetROFromScreen()
    openDate = GetOpenDateFromScreen()
    
    ' Scrape Status
    roStatus = GetRepairOrderStatus()
    
    ' Scrape Lines
    lineA = GetLineDescription("A")
    lineB = GetLineDescription("B")
    lineC = GetLineDescription("C")
    
    ' Scrape Tech ID (specifically for Line A)
    techId = GetTechId()
    If techId = "" Then techId = "<empty>"

    ' Clean commas for CSV safety
    roNum = Replace(roNum, ",", " ")
    roStatus = Replace(roStatus, ",", " ")
    lineA = Replace(lineA, ",", " ")
    lineB = Replace(lineB, ",", " ")
    lineC = Replace(lineC, ",", " ")
    openDate = Replace(openDate, ",", " ")
    techId = Replace(techId, ",", " ")

    ScrapeCurrentRO = roNum & "," & techId & "," & roStatus & "," & lineA & "," & lineB & "," & lineC & "," & openDate
End Function

Function GetTechId()
    Dim row, buf, foundText, i, re, matches, wholeLine
    GetTechId = ""
    
    ' Setup Regex for tech ID (2-5 digits OR literal "MULTI")
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\b(\d{2,5}|MULTI)\b"
    re.IgnoreCase = True
    re.Global = False

    ' Find Line A header first to anchor our search
    For row = 10 To 22 
        g_bzhao.ReadScreen buf, 1, row, 1
        ' Look for 'A' in the line code column (Column 1)
        If UCase(Trim(buf)) = "A" Then
            ' Once 'A' is found, scan rows below for the L1 labor line
            For i = 0 To 3
                If row + i <= 24 Then
                    g_bzhao.ReadScreen wholeLine, 80, row + i, 1
                    ' Check for L1 marker (indicators say it starts around Col 4)
                    If InStr(1, wholeLine, "L1", vbTextCompare) > 0 Then
                        ' Based on debug: Tech ID is visible if reading from Col 40
                        ' We read a larger block covering the tech field and ltype
                        g_bzhao.ReadScreen foundText, 15, row + i, 40 
                        
                        If re.Test(foundText) Then
                            Set matches = re.Execute(foundText)
                            GetTechId = UCase(matches(0).Value)
                            Exit Function
                        End If
                    End If
                End If
            Next
            Exit Function
        End If
    Next
End Function

Function GetROFromScreen()
    Dim buf, re, matches
    ' Based on Header Map: RO number is on Row 3
    g_bzhao.ReadScreen buf, 240, 1, 1 ' Read top 3 rows
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "RO:?\s*(\d{4,})"
    re.IgnoreCase = True
    If re.Test(buf) Then
        Set matches = re.Execute(buf)
        GetROFromScreen = Trim(matches(0).SubMatches(0))
    Else
        ' Fallback: look for 6-digit number in the header block
        re.Pattern = "(^|\s)(\d{6})(\s|$)"
        If re.Test(buf) Then
            Set matches = re.Execute(buf)
            GetROFromScreen = Trim(matches(0).SubMatches(1))
        Else
            GetROFromScreen = "UNKNOWN"
        End If
    End If
End Function

Function GetOpenDateFromScreen()
    Dim buf, re, matches
    ' Based on Header Map: Row 4 contains "OPENED DATE: 05NOV25"
    g_bzhao.ReadScreen buf, 80, 4, 1 
    
    Set re = CreateObject("VBScript.RegExp")
    ' Match "OPENED DATE: " followed by alphanumeric date (e.g. 05NOV25)
    re.Pattern = "OPENED DATE:\s*([A-Z0-9]{6,10})"
    re.IgnoreCase = True
    
    If re.Test(buf) Then
        Set matches = re.Execute(buf)
        GetOpenDateFromScreen = Trim(matches(0).SubMatches(0))
    Else
        ' Fallback to scanning rows 1-3 if Row 4 format differs
        g_bzhao.ReadScreen buf, 240, 1, 1
        re.Pattern = "(?:DATE|OPN|OPEN):?\s*([A-Z0-9/]{6,10})"
        If re.Test(buf) Then
            Set matches = re.Execute(buf)
            GetOpenDateFromScreen = Trim(matches(0).SubMatches(0))
        Else
            GetOpenDateFromScreen = "UNKNOWN"
        End If
    End If
End Function

Function GetRepairOrderStatus()
    Dim buf, re, matches
    ' Based on Header Map: Row 5 contains "RO STATUS: WORKING"
    g_bzhao.ReadScreen buf, 80, 5, 1
    
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "RO STATUS:\s*([A-Z\s]{1,15})"
    re.IgnoreCase = True
    
    If re.Test(buf) Then
        Set matches = re.Execute(buf)
        GetRepairOrderStatus = Trim(matches(0).SubMatches(0))
    Else
        ' Fallback: check Row 5, Col 12 specifically
        g_bzhao.ReadScreen buf, 15, 5, 12
        GetRepairOrderStatus = Trim(buf)
    End If
End Function

Function GetLineDescription(letter)
    Dim row, buf, foundText, nextColChar
    GetLineDescription = ""
    ' Header ends at Row 6 (REMARKS). Lines A, B, C start at Row 7.
    ' We scan from Row 10 to skip potential multi-line headers (e.g. REPAIR, REMARKS)
    For row = 10 To 22
        g_bzhao.ReadScreen buf, 1, row, 1
        ' Look for the letter specifically in column 1
        If UCase(Trim(buf)) = UCase(letter) Then
            ' Peek column 2 to ensure this is a line letter (typical form: "A  DESCRIPTION")
            g_bzhao.ReadScreen nextColChar, 1, row, 2
            If Asc(nextColChar) = 32 Then
                ' Found the line letter anchor in Col 1
                ' Based on previous working state, description starts around Col 4
                g_bzhao.ReadScreen foundText, 50, row, 4
                GetLineDescription = Left(Trim(foundText), 25)
                Exit Function
            End If
        End If
    Next
End Function

' --- Shared Helpers ---

' --- Error Detection and Recovery ---
' To add a new error: add a trigger to IsKnownErrorPresent(), and an ElseIf block in DetectAndRecover().

Function IsKnownErrorPresent(screenContent)
    IsKnownErrorPresent = (InStr(1, screenContent, "PRESS RETURN TO CONTINUE", vbTextCompare) > 0 Or _
                           InStr(1, screenContent, "Process is locked by", vbTextCompare) > 0)
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
    Else
        LogResult "ERROR", "Unrecognised screen state — no recovery handler matched."
        DetectAndRecover = False
    End If
End Function

Function RecoverFromLockedProcess()
    RecoverFromLockedProcess = False
    LogResult "INFO", "Recovery: dismissing locked process, waiting for sequence prompt."
    g_bzhao.SendKey "<Enter>"
    If Not WaitForPrompt("COMMAND:(SEQ#", "", False, 10000, "") Then
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
    If Not WaitForPrompt("COMMAND:(SEQ#", "", False, 10000, "sequence prompt after VEHID recovery") Then
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

Function ShouldSkipSequence(seqNumber)
    ShouldSkipSequence = False
    If SKIP_SEQUENCES = "" Then Exit Function
    
    Dim skipList, j
    skipList = Split(SKIP_SEQUENCES, ",")
    For j = 0 To UBound(skipList)
        If Trim(skipList(j)) = CStr(seqNumber) Then
            ShouldSkipSequence = True
            Exit Function
        End If
    Next
End Function

' Start execution
RunScrapper
