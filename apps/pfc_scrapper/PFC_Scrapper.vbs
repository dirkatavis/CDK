'=====================================================================================
' PFC Scrapper
' Part of the CDK DMS Automation Suite
'
' Strategic Context: Legacy system scheduled for retirement in 3-6 months.
' Purpose: Scrape RO details from sequences until "DOES NOT EXIST" is encountered.
'=====================================================================================

Option Explicit

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
Dim LOG_FILE_PATH: LOG_FILE_PATH = GetConfigPath("PFC_Scrapper", "Log")
Dim OUTPUT_CSV_PATH: OUTPUT_CSV_PATH = GetConfigPath("PFC_Scrapper", "OutputCSV")
Dim SCREEN_WAIT_DELAY: SCREEN_WAIT_DELAY = CInt(GetIniSetting("PFC_Scrapper", "ScreenWaitDelay", "1000"))
Dim SKIP_SEQUENCES: SKIP_SEQUENCES = GetIniSetting("PFC_Scrapper", "SkipSequences", "")

' --- CDK Objects ---
Dim bzhao: Set bzhao = CreateObject("BZWhll.WhllObj")

' --- Main Script ---
Sub RunScrapper()
    Dim i, totalScraped, csvFile
    totalScraped = 0
    i = 1

    LogResult "INFO", "Starting PFC Scrapper. Output: " & OUTPUT_CSV_PATH

    ' Connect to terminal
    On Error Resume Next
    bzhao.Connect ""
    If Err.Number <> 0 Then
        LogResult "ERROR", "Failed to connect to BlueZone: " & Err.Description
        MsgBox "Failed to connect to BlueZone terminal session.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Initialize CSV (Overwrite)
    Set csvFile = g_fso.CreateTextFile(OUTPUT_CSV_PATH, True)
    csvFile.WriteLine "RO number, RO status, Line A, Line B, Line C, Open Date"

    Do
        LogResult "INFO", "Processing sequence: " & i
        
        ' Skip logic
        If ShouldSkipSequence(i) Then
            LogResult "INFO", "Skipping sequence " & i & " as per config."
            i = i + 1
        Else
            ' Ensure we are at COMMAND prompt
            If Not WaitForPrompt("COMMAND:", 5) Then
                LogResult "ERROR", "Timed out waiting for COMMAND prompt at sequence " & i
                Exit Do
            End If

            ' Enter sequence number
            bzhao.SendKey i & "<NumpadEnter>"
            bzhao.Pause SCREEN_WAIT_DELAY

            ' Wait for state change - either RO screen or error
            Dim screenText, startTime, screenFound
            startTime = Timer
            screenFound = False
            Do
                bzhao.ReadScreen screenText, 1920, 1, 1
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
                
                If Timer - startTime > 10 Then
                    LogResult "ERROR", "Timeout waiting for RO screen at sequence " & i
                    Exit Do
                End If
                bzhao.Pause 500
            Loop

            If screenFound Then
                ' Scrape Data
                Dim roData
                roData = ScrapeCurrentRO()
                
                If roData <> "" Then
                    csvFile.WriteLine roData
                    totalScraped = totalScraped + 1
                End If

                ' Return to command prompt
                bzhao.SendKey "E<NumpadEnter>"
                bzhao.Pause SCREEN_WAIT_DELAY
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
    Dim roNum, roStatus, lineA, lineB, lineC, openDate
    
    ' Scrape Header (RO and Date)
    roNum = GetROFromScreen()
    openDate = GetOpenDateFromScreen()
    
    ' Scrape Status
    roStatus = GetRepairOrderStatus()
    
    ' Scrape Lines
    lineA = GetLineDescription("A")
    lineB = GetLineDescription("B")
    lineC = GetLineDescription("C")
    
    ' Clean commas for CSV safety
    roNum = Replace(roNum, ",", " ")
    roStatus = Replace(roStatus, ",", " ")
    lineA = Replace(lineA, ",", " ")
    lineB = Replace(lineB, ",", " ")
    lineC = Replace(lineC, ",", " ")
    openDate = Replace(openDate, ",", " ")

    ScrapeCurrentRO = roNum & "," & roStatus & "," & lineA & "," & lineB & "," & lineC & "," & openDate
End Function

Function GetROFromScreen()
    Dim buf, re, matches
    ' Based on Header Map: RO number is on Row 3
    bzhao.ReadScreen buf, 240, 1, 1 ' Read top 3 rows
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
    bzhao.ReadScreen buf, 80, 4, 1 
    
    Set re = CreateObject("VBScript.RegExp")
    ' Match "OPENED DATE: " followed by alphanumeric date (e.g. 05NOV25)
    re.Pattern = "OPENED DATE:\s*([A-Z0-9]{6,10})"
    re.IgnoreCase = True
    
    If re.Test(buf) Then
        Set matches = re.Execute(buf)
        GetOpenDateFromScreen = Trim(matches(0).SubMatches(0))
    Else
        ' Fallback to scanning rows 1-3 if Row 4 format differs
        bzhao.ReadScreen buf, 240, 1, 1
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
    bzhao.ReadScreen buf, 80, 5, 1
    
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "RO STATUS:\s*([A-Z\s]{1,15})"
    re.IgnoreCase = True
    
    If re.Test(buf) Then
        Set matches = re.Execute(buf)
        GetRepairOrderStatus = Trim(matches(0).SubMatches(0))
    Else
        ' Fallback: check Row 5, Col 12 specifically
        bzhao.ReadScreen buf, 15, 5, 12
        GetRepairOrderStatus = Trim(buf)
    End If
End Function

Function GetLineDescription(letter)
    Dim row, buf, foundText, nextColChar
    GetLineDescription = ""
    ' Header ends at Row 6 (REMARKS). Lines A, B, C start at Row 7.
    ' We scan from Row 10 to skip potential multi-line headers (e.g. REPAIR, REMARKS)
    For row = 10 To 22
        bzhao.ReadScreen buf, 1, row, 1
        ' Look for the letter specifically in column 1
        If UCase(Trim(buf)) = UCase(letter) Then
            ' Peek column 2 to ensure this is a line letter (typical form: "A  DESCRIPTION")
            bzhao.ReadScreen nextColChar, 1, row, 2
            If Asc(nextColChar) = 32 Then
                ' Found the line letter anchor in Col 1
                ' Based on previous working state, description starts around Col 7
                bzhao.ReadScreen foundText, 50, row, 7
                GetLineDescription = Left(Trim(foundText), 25)
                Exit Function
            End If
        End If
    Next
End Function

' --- Shared Helpers ---

Function WaitForPrompt(text, timeoutSec)
    Dim start, elapsed, screenContent
    start = Timer
    Do
        bzhao.ReadScreen screenContent, 1920, 1, 1
        If InStr(1, screenContent, text, vbTextCompare) > 0 Then
            WaitForPrompt = True
            Exit Function
        End If
        bzhao.Pause 500
        elapsed = Timer - start
    Loop While elapsed < timeoutSec
    WaitForPrompt = False
End Function

Function IsTextPresent(text)
    Dim screenContent
    bzhao.ReadScreen screenContent, 1920, 1, 1
    If InStr(1, screenContent, text, vbTextCompare) > 0 Then
        IsTextPresent = True
    Else
        IsTextPresent = False
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
