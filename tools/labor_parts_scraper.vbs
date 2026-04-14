'=====================================================================================
' Labor Parts Scraper
' Part of the CDK DMS Automation Suite — tools\labor_parts_scraper.vbs
'
' Purpose: Scrape L-line (labor) and P-line (parts) data from PFC sequences into
'          a CSV for downstream analysis by apps\labor_parts_analyzer\analyze.py.
'
' Strategic Context: Legacy system nearing retirement. Optimize for simplicity.
'=====================================================================================

Option Explicit

' --- Bootstrap ---
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_sh:  Set g_sh  = CreateObject("WScript.Shell")
Dim g_root: g_root = g_sh.Environment("USER")("CDK_BASE")
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\PathHelper.vbs")).ReadAll

' --- CDK Terminal Object (must be declared before loading BZHelper) ---
Dim g_bzhao: Set g_bzhao = CreateObject("BZWhll.WhllObj")
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\BZHelper.vbs")).ReadAll

' --- Configuration ---
Dim OUTPUT_CSV:       OUTPUT_CSV       = GetConfigPath("LaborPartsScraper", "OutputCSV")
Dim START_SEQUENCE:   START_SEQUENCE   = CInt(GetIniSetting("LaborPartsScraper", "StartSequence",   "100"))
Dim END_SEQUENCE:     END_SEQUENCE     = CInt(GetIniSetting("LaborPartsScraper", "EndSequence",     "200"))
Dim SCREEN_WAIT_MS:   SCREEN_WAIT_MS   = CInt(GetIniSetting("LaborPartsScraper", "ScreenWaitDelay",  "1000"))


' ==============================================================================
' ENTRY POINT
' ==============================================================================
Sub RunScraper()
    Dim seqNum, totalScraped, resumeSeq

    On Error Resume Next
    g_bzhao.Connect ""
    If Err.Number <> 0 Then
        MsgBox "Failed to connect to BlueZone: " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    EnsureCsvHeader

    resumeSeq = GetResumeSeq()
    If resumeSeq > START_SEQUENCE Then
        MsgBox "Resuming from sequence " & resumeSeq & " (previous run aborted at " & (resumeSeq - 1) & ").", vbInformation
    End If

    totalScraped = 0

    For seqNum = resumeSeq To END_SEQUENCE

        If Not WaitForCommandPrompt(seqNum) Then
            MsgBox "Lost COMMAND: prompt at sequence " & seqNum & ". Scraper aborted." & vbCrLf & _
                   "Re-run to resume from this sequence.", vbCritical
            Exit Sub
        End If

        g_bzhao.SendKey CStr(seqNum) & "<NumpadEnter>"
        g_bzhao.Pause SCREEN_WAIT_MS

        Dim screenBuf
        g_bzhao.ReadScreen screenBuf, 1920, 1, 1

        If InStr(1, screenBuf, "DOES NOT EXIST", vbTextCompare) > 0 Then
            g_bzhao.SendKey "E<NumpadEnter>"
            g_bzhao.Pause SCREEN_WAIT_MS
            MsgBox "Sequence " & seqNum & " does not exist. End of range reached." & vbCrLf & _
                   "Total scraped: " & totalScraped, vbInformation
            Exit Sub
        End If

        Dim roNum, roOpenDate
        roNum     = GetROFromScreen()
        roOpenDate = GetOpenDateFromScreen()

        Dim scraped
        scraped = ScrapeLines(roNum, roOpenDate, seqNum)
        totalScraped = totalScraped + scraped

        ' Return to COMMAND prompt
        g_bzhao.SendKey "E<NumpadEnter>"
        g_bzhao.Pause SCREEN_WAIT_MS

    Next

    MsgBox "Labor Parts Scraper finished." & vbCrLf & "Total rows written: " & totalScraped, vbInformation
End Sub


' ==============================================================================
' WaitForCommandPrompt
' Retries 5 times at 500 ms intervals before giving up.
' ==============================================================================
Function WaitForCommandPrompt(seqNum)
    Dim attempt, buf
    WaitForCommandPrompt = False

    For attempt = 1 To 5
        g_bzhao.ReadScreen buf, 80, 24, 1
        If InStr(1, buf, "COMMAND:", vbTextCompare) > 0 Then
            WaitForCommandPrompt = True
            Exit Function
        End If
        g_bzhao.Pause 500
    Next
End Function


' ==============================================================================
' GetResumeSeq
' Reads the max Sequence value from existing CSV rows; returns that + 1.
' Returns START_SEQUENCE when the file is missing, empty, or has only a header.
' ==============================================================================
Function GetResumeSeq()
    GetResumeSeq = START_SEQUENCE

    If Not g_fso.FileExists(OUTPUT_CSV) Then Exit Function

    Dim ts: Set ts = g_fso.OpenTextFile(OUTPUT_CSV, 1, False)
    If ts.AtEndOfStream Then
        ts.Close
        Exit Function
    End If

    ' Skip header
    Dim headerLine: headerLine = ts.ReadLine
    Dim maxSeq: maxSeq = 0
    Dim parts, seqVal

    Do Until ts.AtEndOfStream
        Dim line: line = ts.ReadLine
        If Len(Trim(line)) = 0 Then GoTo NextLine
        parts = Split(line, ",")
        ' Sequence is column index 5 (0-based)
        If UBound(parts) >= 5 Then
            On Error Resume Next
            seqVal = CInt(Trim(parts(5)))
            If Err.Number = 0 Then
                If seqVal > maxSeq Then maxSeq = seqVal
            End If
            On Error GoTo 0
        End If
        NextLine:
    Loop

    ts.Close

    If maxSeq >= START_SEQUENCE Then
        GetResumeSeq = maxSeq + 1
    End If
End Function


' ==============================================================================
' EnsureCsvHeader
' Creates the CSV with header if absent OR if the file exists but is empty.
' ==============================================================================
Sub EnsureCsvHeader()
    If Not g_fso.FileExists(OUTPUT_CSV) Then
        Dim newFile: Set newFile = g_fso.CreateTextFile(OUTPUT_CSV, True)
        newFile.WriteLine "Timestamp,RO_Number,Labor_ID,Description,Parts_Found,Sequence,RO_Open_Date"
        newFile.Close
        Exit Sub
    End If

    ' File exists — check if empty
    Dim f: Set f = g_fso.GetFile(OUTPUT_CSV)
    If f.Size = 0 Then
        Dim emptyFile: Set emptyFile = g_fso.OpenTextFile(OUTPUT_CSV, 2, True)
        emptyFile.WriteLine "Timestamp,RO_Number,Labor_ID,Description,Parts_Found,Sequence,RO_Open_Date"
        emptyFile.Close
    End If
End Sub


' ==============================================================================
' ScrapeLines
' Scans rows 9–22 for L-lines and P-lines.
' For each L-line found, checks whether a corresponding P-line exists.
' Writes one CSV row per L-line.
' Returns count of rows written.
' ==============================================================================
Function ScrapeLines(roNum, roOpenDate, seqNum)
    ScrapeLines = 0

    Dim row, colBuf, lLines()
    Dim lCount: lCount = 0
    ReDim lLines(14)  ' 14 rows max (9-22)

    ' Pass 1: collect all L-line row numbers
    For row = 9 To 22
        g_bzhao.ReadScreen colBuf, 3, row, 4  ' cols 4-6
        If UCase(Left(Trim(colBuf), 1)) = "L" And IsNumeric(Mid(colBuf, 2, 1)) Then
            lLines(lCount) = row
            lCount = lCount + 1
        End If
    Next

    If lCount = 0 Then Exit Function

    ' Pass 2: for each L-line, check if a P-line follows before the next L-line
    Dim i
    For i = 0 To lCount - 1
        Dim lRow: lRow = lLines(i)
        Dim nextLRow
        If i + 1 < lCount Then
            nextLRow = lLines(i + 1)
        Else
            nextLRow = 23
        End If

        ' Read labor line details
        Dim laborId, descBuf
        g_bzhao.ReadScreen laborId,  4, lRow, 4   ' e.g. "L  1" → capture the ID
        g_bzhao.ReadScreen descBuf, 50, lRow, 10  ' description starts ~col 10

        laborId = Trim(laborId)
        Dim desc: desc = Trim(descBuf)

        ' Check for P-line between this L-line and the next
        Dim partsFound: partsFound = "False"
        Dim chkRow
        For chkRow = lRow + 1 To nextLRow - 1
            Dim pBuf
            g_bzhao.ReadScreen pBuf, 3, chkRow, 6  ' cols 6-8
            If UCase(Left(Trim(pBuf), 1)) = "P" And IsNumeric(Mid(Trim(pBuf), 2, 1)) Then
                partsFound = "True"
                Exit For
            End If
        Next

        ' Sanitize for CSV
        Dim safeDesc:    safeDesc    = Replace(desc,      ",", " ")
        Dim safeRoNum:   safeRoNum   = Replace(roNum,     ",", " ")
        Dim safeLaborId: safeLaborId = Replace(laborId,   ",", " ")
        Dim safeOpenDate: safeOpenDate = Replace(roOpenDate, ",", " ")

        Dim ts: Set ts = g_fso.OpenTextFile(OUTPUT_CSV, 8, True)
        ts.WriteLine Now() & "," & safeRoNum & "," & safeLaborId & "," & safeDesc & "," & partsFound & "," & seqNum & "," & safeOpenDate
        ts.Close

        ScrapeLines = ScrapeLines + 1
    Next
End Function


' ==============================================================================
' GetROFromScreen — reads RO number from rows 1-3
' ==============================================================================
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


' ==============================================================================
' GetOpenDateFromScreen — reads OPENED DATE from row 4
' ==============================================================================
Function GetOpenDateFromScreen()
    Dim buf, re, matches
    g_bzhao.ReadScreen buf, 80, 4, 1

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "OPENED DATE:\s*([A-Z0-9]{6,10})"
    re.IgnoreCase = True

    If re.Test(buf) Then
        Set matches = re.Execute(buf)
        GetOpenDateFromScreen = Trim(matches(0).SubMatches(0))
        Exit Function
    End If

    ' Fallback: scan rows 1-3
    g_bzhao.ReadScreen buf, 240, 1, 1
    re.Pattern = "(?:DATE|OPN|OPEN):?\s*([A-Z0-9/]{6,10})"
    If re.Test(buf) Then
        Set matches = re.Execute(buf)
        GetOpenDateFromScreen = Trim(matches(0).SubMatches(0))
    Else
        GetOpenDateFromScreen = "UNKNOWN"
    End If
End Function


' ==============================================================================
' GetIniSetting — 3-param wrapper for ReadIniValue with default
' ==============================================================================
Function GetIniSetting(section, key, defaultValue)
    GetIniSetting = defaultValue
    On Error Resume Next
    Dim configPath: configPath = g_fso.BuildPath(GetRepoRoot(), "config\config.ini")
    Dim val: val = ReadIniValue(configPath, section, key)
    If val <> "" Then GetIniSetting = val
    On Error GoTo 0
End Function


' --- Run ---
RunScraper
Option Explicit

'==============================================================================
' TOOL:    labor_parts_scraper.vbs
' PURPOSE: Scrapes L-line (labor) and P-line (parts) data from the CDK PFC
'          detail screen into a Master CSV for offline analysis.
'          Used to determine which labor descriptions require parts.
'
' OUTPUT:  Appends to the configured OutputCSV (Master_Labor_Log.csv).
'          Columns: Timestamp, RO_Number, Labor_ID, Description, Parts_Found
'
' USAGE:   Run from BlueZone Script Host while PFC COMMAND:(SEQ#/E/N/B/?)
'          prompt is active.
'==============================================================================

' --- Bootstrap ---
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_sh:  Set g_sh  = CreateObject("WScript.Shell")
Dim g_root: g_root = g_sh.Environment("USER")("CDK_BASE")
If g_root = "" Then
    MsgBox "CDK_BASE environment variable is not set. Please run setup_cdk_base.vbs.", vbCritical, "LaborPartsScraper"
    WScript.Quit
End If
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\PathHelper.vbs")).ReadAll

' --- CDK Terminal Object ---
Dim g_bzhao: Set g_bzhao = CreateObject("BZWhll.WhllObj")

' --- Configuration ---
Dim OUTPUT_CSV: OUTPUT_CSV  = GetConfigPath("LaborPartsScraper", "OutputCSV")
Dim START_SEQ:  START_SEQ   = CInt(GetIniSetting("LaborPartsScraper", "StartSequence",  "1"))
Dim END_SEQ:    END_SEQ     = CInt(GetIniSetting("LaborPartsScraper", "EndSequence",    "999"))
Dim SCRN_WAIT:  SCRN_WAIT   = CInt(GetIniSetting("LaborPartsScraper", "ScreenWaitDelay","1000"))

' --- Entry Point ---
Call RunScraper()

'==============================================================================
Sub RunScraper()
    On Error Resume Next
    g_bzhao.Connect ""
    If Err.Number <> 0 Then
        MsgBox "Failed to connect to BlueZone terminal session." & vbCrLf & Err.Description, vbCritical, "LaborPartsScraper"
        Exit Sub
    End If
    On Error GoTo 0

    ' Verify we are at the PFC COMMAND prompt
    Dim screenText
    g_bzhao.ReadScreen screenText, 1920, 1, 1
    If InStr(1, screenText, "COMMAND:", vbTextCompare) = 0 Then
        MsgBox "Session is not at COMMAND: prompt. Navigate to PFC first.", vbExclamation, "LaborPartsScraper"
        Exit Sub
    End If

    Call EnsureCsvHeader()

    Dim resumeFrom: resumeFrom = GetResumeSeq()
    Dim effectiveStart
    If resumeFrom > START_SEQ Then
        effectiveStart = resumeFrom
        MsgBox "Resuming from sequence " & effectiveStart & " (last run reached " & (resumeFrom - 1) & ").", vbInformation, "LaborPartsScraper"
    Else
        effectiveStart = START_SEQ
    End If

    Dim scrapedROs:   scrapedROs   = 0
    Dim scrapedLines: scrapedLines = 0
    Dim skippedSeqs:  skippedSeqs  = 0
    Dim seqNum

    For seqNum = effectiveStart To END_SEQ
        ' Confirm still at COMMAND prompt before each sequence (retry up to 5x to handle slow CDK response)
        Dim cmdRetry: cmdRetry = 0
        Do
            g_bzhao.ReadScreen screenText, 1920, 1, 1
            If InStr(1, screenText, "COMMAND:", vbTextCompare) > 0 Then Exit Do
            cmdRetry = cmdRetry + 1
            g_bzhao.Pause 500
        Loop While cmdRetry < 5
        If InStr(1, screenText, "COMMAND:", vbTextCompare) = 0 Then
            MsgBox "Lost COMMAND: prompt at sequence " & seqNum & " - aborting.", vbExclamation, "LaborPartsScraper"
            Exit For
        End If

        ' Navigate to sequence
        g_bzhao.SendKey CStr(seqNum) & "<NumpadEnter>"
        g_bzhao.Pause SCRN_WAIT

        ' Read result
        g_bzhao.ReadScreen screenText, 1920, 1, 1

        If InStr(1, screenText, "DOES NOT EXIST", vbTextCompare) > 0 Then
            ' End of sequence range - clear message and stop
            g_bzhao.SendKey "<NumpadEnter>"
            g_bzhao.Pause 300
            Exit For
        End If

        If InStr(1, screenText, "RO:", vbTextCompare) > 0 Then
            Dim roNumber: roNumber = GetRoFromScreen()
            Dim written:  written  = ScrapeAndAppend(roNumber, seqNum)
            scrapedLines = scrapedLines + written
            If written > 0 Then scrapedROs = scrapedROs + 1
            ' Return to COMMAND prompt
            g_bzhao.SendKey "E<NumpadEnter>"
            g_bzhao.Pause SCRN_WAIT
        Else
            skippedSeqs = skippedSeqs + 1
            ' Clear any lingering message
            g_bzhao.SendKey "<NumpadEnter>"
            g_bzhao.Pause 300
        End If
    Next

    MsgBox "LaborPartsScraper complete." & vbCrLf & vbCrLf & _
           "ROs scraped  : " & scrapedROs   & vbCrLf & _
           "Labor lines  : " & scrapedLines & vbCrLf & _
           "Seqs skipped : " & skippedSeqs  & vbCrLf & vbCrLf & _
           "Output: " & OUTPUT_CSV, vbInformation, "LaborPartsScraper"
End Sub

'==============================================================================
' Stateful scan of rows 9-22. Returns count of L-lines written.
'==============================================================================
Function ScrapeAndAppend(roNumber, seqNum)
    ScrapeAndAppend = 0
    Dim ids(49), descs(49), hasParts(49)
    Dim count:      count      = 0
    Dim currentIdx: currentIdx = -1
    Dim openDate:   openDate   = GetOpenDateFromScreen()
    Dim row, buf

    For row = 9 To 22
        buf = ""
        On Error Resume Next
        g_bzhao.ReadScreen buf, 80, row, 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0

        If Len(buf) >= 7 Then
            ' Line letter header (col 1 = A-Z): reset L-row context
            If Mid(buf, 1, 1) >= "A" And Mid(buf, 1, 1) <= "Z" Then
                currentIdx = -1
            ' L-row: col 4 = "L", col 5 = digit
            ElseIf Mid(buf, 4, 1) = "L" And IsNumeric(Mid(buf, 5, 1)) Then
                ids(count)      = Trim(Mid(buf, 4, 2))
                descs(count)    = Trim(Mid(buf, 7, 37))
                hasParts(count) = False
                currentIdx = count
                count = count + 1
            ' P-row: col 6 = "P", col 7 = digit - mark the owning L-row
            ElseIf Mid(buf, 6, 1) = "P" And IsNumeric(Mid(buf, 7, 1)) Then
                If currentIdx >= 0 Then hasParts(currentIdx) = True
            End If
        End If
    Next

    If count = 0 Then Exit Function

    Dim ts: ts = BuildTimestamp()
    Dim f:  Set f = g_fso.OpenTextFile(OUTPUT_CSV, 8, True)
    Dim i, cleanDesc
    For i = 0 To count - 1
        cleanDesc = Replace(descs(i), """", "")
        f.WriteLine ts & "," & roNumber & "," & ids(i) & "," & _
                    Chr(34) & cleanDesc & Chr(34) & "," & CStr(hasParts(i)) & "," & seqNum & "," & openDate
    Next
    f.Close
    ScrapeAndAppend = count
End Function

'==============================================================================
Function GetRoFromScreen()
    Dim buf: buf = ""
    On Error Resume Next
    g_bzhao.ReadScreen buf, 80, 3, 1
    On Error GoTo 0
    Dim roPos: roPos = InStr(buf, "RO:")
    If roPos = 0 Then GetRoFromScreen = "UNKNOWN" : Exit Function
    GetRoFromScreen = Trim(Mid(buf, roPos + 3, 10))
End Function

'==============================================================================
Sub EnsureCsvHeader()
    Dim parentDir: parentDir = g_fso.GetParentFolderName(OUTPUT_CSV)
    If Not g_fso.FolderExists(parentDir) Then g_fso.CreateFolder parentDir
    ' Write header if the file does not exist OR is empty (e.g., user cleared contents)
    Dim needsHeader: needsHeader = False
    If Not g_fso.FileExists(OUTPUT_CSV) Then
        needsHeader = True
    ElseIf g_fso.GetFile(OUTPUT_CSV).Size = 0 Then
        needsHeader = True
    End If
    If needsHeader Then
        Dim f: Set f = g_fso.OpenTextFile(OUTPUT_CSV, 2, True)
        f.WriteLine "Timestamp,RO_Number,Labor_ID,Description,Parts_Found,Sequence,RO_Open_Date"
        f.Close
    End If
End Sub

'==============================================================================
' Reads the CDK open date from Row 4 of the PFC detail screen.
' Returns a date string like "05NOV25" or "UNKNOWN" if not found.
'==============================================================================
Function GetOpenDateFromScreen()
    Dim buf, re, matches
    g_bzhao.ReadScreen buf, 80, 4, 1
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "OPENED DATE:\s*([A-Z0-9]{6,10})"
    re.IgnoreCase = True
    If re.Test(buf) Then
        Set matches = re.Execute(buf)
        GetOpenDateFromScreen = Trim(matches(0).SubMatches(0))
    Else
        ' Fallback: scan rows 1-5 for any date-like pattern
        g_bzhao.ReadScreen buf, 400, 1, 1
        re.Pattern = "(?:DATE|OPN|OPEN):\s*([A-Z0-9/]{6,10})"
        If re.Test(buf) Then
            Set matches = re.Execute(buf)
            GetOpenDateFromScreen = Trim(matches(0).SubMatches(0))
        Else
            GetOpenDateFromScreen = "UNKNOWN"
        End If
    End If
End Function

'==============================================================================
' Reads the CSV and returns the sequence number to resume from (max + 1).
' Returns 0 if no valid sequence data found (caller falls back to START_SEQ).
'==============================================================================
Function GetResumeSeq()
    GetResumeSeq = 0
    If Not g_fso.FileExists(OUTPUT_CSV) Then Exit Function
    Dim f: Set f = g_fso.OpenTextFile(OUTPUT_CSV, 1, False)
    Dim maxSeq: maxSeq = 0
    Dim firstLine: firstLine = True
    Do While Not f.AtEndOfStream
        Dim line: line = f.ReadLine()
        If firstLine Then
            firstLine = False
        Else
            Dim parts: parts = Split(line, ",")
            Dim lastField: lastField = Trim(parts(UBound(parts)))
            If IsNumeric(lastField) Then
                If CInt(lastField) > maxSeq Then maxSeq = CInt(lastField)
            End If
        End If
    Loop
    f.Close
    If maxSeq > 0 Then GetResumeSeq = maxSeq + 1
End Function

'==============================================================================
Function BuildTimestamp()
    Dim d: d = Now()
    BuildTimestamp = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & _
                    Right("0" & Day(d), 2) & " " & Right("0" & Hour(d), 2) & ":" & _
                    Right("0" & Minute(d), 2) & ":" & Right("0" & Second(d), 2)
End Function

'==============================================================================
Function GetIniSetting(section, key, defaultValue)
    Dim configPath, val
    GetIniSetting = defaultValue
    On Error Resume Next
    configPath = g_fso.BuildPath(GetRepoRoot(), "config\config.ini")
    val = ReadIniValue(configPath, section, key)
    If val <> "" Then GetIniSetting = val
    On Error GoTo 0
End Function
