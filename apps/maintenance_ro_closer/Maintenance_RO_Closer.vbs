'=====================================================================================
' Maintenance RO Auto-Closer
' Part of the CDK DMS Automation Suite
'
' Strategic Context: Legacy system scheduled for retirement in 3-6 months.
' Purpose: Automate closing of specific Maintenance ROs with exact footprint match.
'=====================================================================================

Option Explicit

' --- Bootstrap ---
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_sh: Set g_sh = CreateObject("WScript.Shell")
Dim g_root: g_root = g_sh.Environment("USER")("CDK_BASE")
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\PathHelper.vbs")).ReadAll
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\HostCompat.vbs")).ReadAll
Dim g_bzhao
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\BZHelper.vbs")).ReadAll

' --- Execution Parameters ---
Dim MAIN_PROMPT: MAIN_PROMPT = "R.O. NUMBER"
Dim LOG_FILE_PATH: LOG_FILE_PATH = GetConfigPath("Maintenance_RO_Closer", "Log")
Dim DEBUG_LEVEL: DEBUG_LEVEL = 2 ' 1=Error, 2=Info
Dim RO_LIST_PATH: RO_LIST_PATH = GetConfigPath("Maintenance_RO_Closer", "ROList")
Dim SKIP_RO_LIST_PATH: SKIP_RO_LIST_PATH = GetConfigPath("Maintenance_RO_Closer", "SkipRoList")

' --- Configurable Pauses ---
Function GetConfigSetting(section, key, defaultValue)
    Dim configFile: configFile = g_fso.BuildPath(g_root, "config\config.ini")
    Dim val: val = ReadIniValue(configFile, section, key)
    If val = "" Then
        GetConfigSetting = defaultValue
    Else
        If IsNumeric(val) Then
            GetConfigSetting = CInt(val)
        Else
            GetConfigSetting = val
        End If
    End If
End Function

Dim STABILITY_PAUSE: STABILITY_PAUSE = GetConfigSetting("Maintenance_RO_Closer", "StabilityPause", 2000)
Dim LOOP_PAUSE: LOOP_PAUSE = GetConfigSetting("Maintenance_RO_Closer", "LoopPause", 1000)
Dim REVIEW_PAUSE: REVIEW_PAUSE = GetConfigSetting("Maintenance_RO_Closer", "ReviewPause", 500)
Dim BLACKLIST_TERMS: BLACKLIST_TERMS = GetConfigSetting("Maintenance_RO_Closer", "blacklist_terms", "")
Dim OLD_RO_DAYS_THRESHOLD: OLD_RO_DAYS_THRESHOLD = GetConfigSetting("Maintenance_RO_Closer", "AssumeClosedAfterDays", 120)
Dim EMPLOYEE_NUMBER: EMPLOYEE_NUMBER = GetConfigSetting("Maintenance_RO_Closer", "EmployeeNumber", "")
Dim EMPLOYEE_NAME_CONFIRM: EMPLOYEE_NAME_CONFIRM = GetConfigSetting("Maintenance_RO_Closer", "EmployeeNameConfirm", "")
Dim WARRANTY_LTYPES_RAW: WARRANTY_LTYPES_RAW = GetConfigSetting("Maintenance_RO_Closer", "WarrantyLTypes", "WCH,WF")
Dim WARRANTY_DIALOG_STEP_DELAY_MS: WARRANTY_DIALOG_STEP_DELAY_MS = GetConfigSetting("Maintenance_RO_Closer", "WarrantyDialogStepDelayMs", 2000)
Dim FORD_WARRANTY_CAUSE_TEXT: FORD_WARRANTY_CAUSE_TEXT = GetConfigSetting("Maintenance_RO_Closer", "FordWarrantyCauseText", "Defective Part")
Dim FORD_WARRANTY_LICENSE_STATE: FORD_WARRANTY_LICENSE_STATE = GetConfigSetting("Maintenance_RO_Closer", "FordWarrantyLicenseState", "GA")
Dim SCRIPT_BUILD_TAG: SCRIPT_BUILD_TAG = "2026-05-14-fix-asc-empty-string"

Dim g_SkipRoLookup
Dim g_SupportedWarrantyLTypes
Dim g_ReviewPhaseResult

' --- CDK Objects ---
Set g_bzhao = CreateObject("BZWhll.WhllObj")

InitializeSupportedWarrantyLTypes

' --- Main Loop ---
Sub RunAutomation()
    Dim currentRo, successfulCount, fso, scriptDir, csvPath, ts, strLine, roFromCsv
    successfulCount = 0

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check file existence *before* terminal connection to avoid orphaned objects
    If Not fso.FileExists(RO_LIST_PATH) Then
        LogResult "ERROR", "Mandatory RO List file missing: " & RO_LIST_PATH
        MsgBox "Error: RO List file not found at: " & RO_LIST_PATH, vbCritical, "File Not Found"
        Exit Sub
    End If

    LogResult "INFO", "Starting Maintenance RO Auto-Closer using list: " & RO_LIST_PATH
    LogResult "INFO", "Script build: " & SCRIPT_BUILD_TAG
    
    ' Load SkipRoList (required configuration for deterministic skip behavior)
    LoadSkipRoLookup SKIP_RO_LIST_PATH
    LogResult "INFO", "SkipRoList loaded from: " & SKIP_RO_LIST_PATH & " (entries=" & g_SkipRoLookup.Count & ")"
    
    ' Connect to terminal only after configuration and file existence are verified
    g_bzhao.Connect ""
    
    ' Start processing with unified error handling
    ProcessRoList fso, successfulCount
    
    ' Final graceful disconnect
    If Not g_bzhao Is Nothing Then g_bzhao.Disconnect

    LogResult "INFO", "Automation complete. Total successful closures: " & successfulCount
    MsgBox "Maintenance RO Auto-Closer Finished." & vbCrLf & "Successful Closures: " & successfulCount, vbInformation
End Sub

' --- Helper Subroutines & Functions ---

Sub ProcessRoList(fso, ByRef successfulCount)
    Dim ts, strLine, roFromCsv, currentRo, reviewOk
    
    Set ts = fso.OpenTextFile(RO_LIST_PATH, 1) ' 1 = ForReading

    Do While Not ts.AtEndOfStream

        strLine = Trim(ts.ReadLine)
        If strLine <> "" Then
            ' Handle potential CSV splitting (take first column)
            roFromCsv = Split(strLine, ",")(0)
            currentRo = Trim(roFromCsv)
            
            ' Validate 6-digit RO
            If Len(currentRo) = 6 And IsNumeric(currentRo) Then
                If ShouldSkipRo(currentRo) Then
                    LogResult "INFO", "RO " & currentRo & " found in SkipRoList. Skipping before entry."
                Else
                LogResult "INFO", String(50, "=")
                LogResult "INFO", "Processing RO: " & currentRo
                
                ' Ensure we are at the main prompt
                WaitForText MAIN_PROMPT
                
                ' Enter RO Number
                EnterTextWithStability currentRo
                
                ' Check for errors or closed status
                If IsRoProcessable(currentRo) Then
                    If ShouldProcessRoByBusinessRules(currentRo) Then
                        reviewOk = ProcessRoReview()
                        If Err.Number <> 0 Then
                            LogResult "ERROR", "TRACE: ProcessRoReview runtime error for RO " & currentRo & " | Err=" & Err.Number & " | Desc=" & Err.Description
                            Err.Clear
                            reviewOk = False
                            g_ReviewPhaseResult = "FAILED"
                        End If

                        If reviewOk Then
                            If CloseRoFinal() Then
                                LogResult "INFO", "SUCCESS: RO " & currentRo & " finalized (FILE command sent)."
                                successfulCount = successfulCount + 1
                            Else
                                LogResult "ERROR", "Failed to finalize RO: " & currentRo & " during Phase II."
                            End If
                        Else
                            If g_ReviewPhaseResult = "SKIPPED" Then
                                LogResult "INFO", "RO " & currentRo & " skipped during review phase."
                            Else
                                LogResult "ERROR", "Failed to complete review for RO: " & currentRo & " during Phase I."
                            End If
                        End If
                    Else
                        LogResult "INFO", "RO " & currentRo & " skipped by business gates. Sending 'E'."
                        EnterTextWithStability "E"
                    End If
                End If

                ' Always return to main prompt for safety
                ReturnToMainPrompt()
                End If
            ElseIf Len(currentRo) > 0 Then
                LogResult "INFO", "Skipping invalid format row: '" & currentRo & "'"
                ReturnToMainPrompt()
            End If
        End If
    Loop
    
    If Not ts Is Nothing Then
        ts.Close
        Set ts = Nothing
    End If
End Sub

Sub LoadSkipRoLookup(skipListPath)
    Dim ts, lineText, normalizedRo, fso
    Set g_SkipRoLookup = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(skipListPath) Then
        TerminateScript "Configured SkipRoList file not found: " & skipListPath
        Exit Sub
    End If

    Set ts = fso.OpenTextFile(skipListPath, 1)
    Do While Not ts.AtEndOfStream
        lineText = Trim(ts.ReadLine)
        If lineText <> "" Then
            If Left(lineText, 1) <> "#" And Left(lineText, 1) <> ";" Then
                normalizedRo = Trim(Split(lineText, ",")(0))
                If normalizedRo <> "" Then
                    If Not g_SkipRoLookup.Exists(normalizedRo) Then
                        g_SkipRoLookup.Add normalizedRo, True
                    End If
                End If
            End If
        End If
    Loop

    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub

Function ShouldSkipRo(roNumber)
    ShouldSkipRo = False

    If IsObject(g_SkipRoLookup) Then
        If g_SkipRoLookup.Exists(Trim(CStr(roNumber))) Then
            ShouldSkipRo = True
        End If
    End If
End Function

' --- Helper Subroutines & Functions ---

Function IsRoProcessable(roNumber)
    Dim screenContent
    IsRoProcessable = True
    g_bzhao.Pause STABILITY_PAUSE
    ' Read screen starting from Row 2 down to Row 6 to catch status (Row 5) and RO info
    ' We also read more to catch system errors (Pick/BASIC errors)
    g_bzhao.ReadScreen screenContent, 1920, 1, 1 
    
    If InStr(1, screenContent, "PRESS RETURN TO CONTINUE", vbTextCompare) > 0 Then
        LogResult "INFO", "RO " & roNumber & " VEHID not on file. Attempting recovery."
        If Not BZH_RecoverFromVehidError(EMPLOYEE_NUMBER, EMPLOYEE_NAME_CONFIRM, "1") Then
            LogResult "ERROR", "RO " & roNumber & " VEHID recovery failed. Terminal state unknown — stopping to avoid incorrect keystrokes."
            TerminateScript "VEHID recovery failed for RO " & roNumber & ". Manual intervention required."
        End If
        IsRoProcessable = False
        Exit Function
    ElseIf InStr(1, screenContent, "NOT ON FILE", vbTextCompare) > 0 Then
        LogResult "INFO", "RO " & roNumber & " NOT ON FILE. Skipping."
        IsRoProcessable = False
        Exit Function
    ElseIf InStr(1, screenContent, "is closed", vbTextCompare) > 0 Or InStr(1, screenContent, "ALREADY CLOSED", vbTextCompare) > 0 Then
        LogResult "INFO", "RO " & roNumber & " ALREADY CLOSED. Skipping."
        IsRoProcessable = False
        Exit Function
    ElseIf InStr(1, screenContent, "VARIABLE HAS NOT BEEN ASSIGNED", vbTextCompare) > 0 Then
        LogResult "ERROR", "DMS System Error detected for RO " & roNumber & ". Skipping."
        IsRoProcessable = False
        Exit Function
    ElseIf InStr(1, screenContent, "ENTER SEQUENCE NUMBER", vbTextCompare) > 0 Then
        ' This is actually a valid prompt now, but we skip it here to let the main loop handle it
        LogResult "INFO", "RO " & roNumber & " prompted for Sequence Number. Treating as valid prompt."
        IsRoProcessable = False
        Exit Function
    End If
    
    IsRoProcessable = True
End Function

Function ShouldProcessRoByBusinessRules(roNumber)
    ' === Business Rules: Close/Skip Decision Table ===
    '
    ' RO Status                    | Condition                        | Action
    ' -----------------------------+----------------------------------+--------
    ' Any                          | Blacklisted (any page)           | SKIP
    ' Any (non-blacklisted)        | Age >= AssumeClosedAfterDays     | PROCESS (overrides status)
    ' READY TO POST                | (none)                           | PROCESS
    ' Any other                    | (none)                           | SKIP
    '
    ' Rules evaluated top to bottom. First match wins.
    ' Age exception overrides status check but not the blacklist.
    ' =================================================
    Dim ageDays, openedDateToken, isOldEnough
    Dim screenContent, isReadyToPost, matchedBlacklistTerm, currentStatus

    matchedBlacklistTerm = BZH_GetMatchedBlacklistTerm(BLACKLIST_TERMS, STABILITY_PAUSE)
    screenContent = GetCurrentScreenContent()
    isReadyToPost = (InStr(1, screenContent, "READY TO POST", vbTextCompare) > 0)
    isOldEnough = IsRoOldEnoughForOverride(ageDays, openedDateToken)
    currentStatus = ExtractStatusText(screenContent)

    LogResult "INFO", "RO " & roNumber & " | Status: " & IIf(isReadyToPost, "READY TO POST", currentStatus) & " | Age: " & IIf(ageDays >= 0, ageDays & " days", "unknown")

    ' Gate 1: Blacklist
    If matchedBlacklistTerm <> "" Then
        LogResult "INFO", "RO " & roNumber & " | Blacklisted ('" & matchedBlacklistTerm & "'). Skipping."
        ShouldProcessRoByBusinessRules = False
        Exit Function
    End If

    ' Gate 2: Age exception — overrides status check but not blacklist
    If isOldEnough Then
        LogResult "INFO", "RO " & roNumber & " | Age exception: " & ageDays & " days old (threshold: " & OLD_RO_DAYS_THRESHOLD & "). Proceeding regardless of status."
        ShouldProcessRoByBusinessRules = True
        Exit Function
    End If

    ' Gate 3: READY TO POST
    If isReadyToPost Then
        LogResult "INFO", "RO " & roNumber & " | READY TO POST. Proceeding."
        ShouldProcessRoByBusinessRules = True
        Exit Function
    End If

    ' Gate 4: Skip anything else
    LogResult "INFO", "RO " & roNumber & " | Not READY TO POST and age threshold not met. Skipping."
    ShouldProcessRoByBusinessRules = False
End Function

Sub InitializeSupportedWarrantyLTypes()
    Dim parts, i, token

    Set g_SupportedWarrantyLTypes = CreateObject("Scripting.Dictionary")
    parts = Split(UCase(CStr(WARRANTY_LTYPES_RAW)), ",")

    For i = 0 To UBound(parts)
        token = Trim(parts(i))
        If token <> "" Then
            If Not g_SupportedWarrantyLTypes.Exists(token) Then
                g_SupportedWarrantyLTypes.Add token, True
            End If
        End If
    Next

    LogResult "INFO", "Configured supported warranty labor types: " & WARRANTY_LTYPES_RAW
End Sub

Function IsSupportedWarrantyLType(lTypeCode)
    IsSupportedWarrantyLType = False
    If Not IsObject(g_SupportedWarrantyLTypes) Then Exit Function
    If g_SupportedWarrantyLTypes.Exists(UCase(Trim(CStr(lTypeCode)))) Then
        IsSupportedWarrantyLType = True
    End If
End Function

Function GetFirstUnsupportedWarrantyLaborType()
    Dim row, buf, lTypeCode
    GetFirstUnsupportedWarrantyLaborType = ""

    For row = 9 To 23
        buf = ""
        g_bzhao.ReadScreen buf, 80, row, 1

        If Len(buf) >= 55 Then
            If Mid(buf, 4, 1) = "L" And IsNumeric(Mid(buf, 5, 1)) Then
                lTypeCode = UCase(Trim(Mid(buf, 50, 6)))
                If Left(lTypeCode, 1) = "W" Then
                    If Not IsSupportedWarrantyLType(lTypeCode) Then
                        GetFirstUnsupportedWarrantyLaborType = lTypeCode
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

Function IIf(condition, trueVal, falseVal)
    If condition Then
        IIf = trueVal
    Else
        IIf = falseVal
    End If
End Function

Function GetCurrentScreenContent()
    Dim screenContent
    g_bzhao.ReadScreen screenContent, 1920, 1, 1
    GetCurrentScreenContent = screenContent
End Function

Function IsRoOldEnoughForOverride(ByRef ageDays, ByRef openedDateToken)
    Dim screenContent, parsedDate
    ageDays = -1
    openedDateToken = ""

    screenContent = GetCurrentScreenContent()
    openedDateToken = ExtractOpenedDateToken(screenContent)

    If openedDateToken = "" Then
        IsRoOldEnoughForOverride = False
        Exit Function
    End If

    If Not TryParseCdkDate(openedDateToken, parsedDate) Then
        LogResult "WARN", "Unable to parse OPENED DATE token '" & openedDateToken & "' for age override evaluation."
        IsRoOldEnoughForOverride = False
        Exit Function
    End If

    If CInt(OLD_RO_DAYS_THRESHOLD) <= 0 Then
        IsRoOldEnoughForOverride = False
        Exit Function
    End If

    ageDays = DateDiff("d", parsedDate, Date)
    IsRoOldEnoughForOverride = (ageDays >= CInt(OLD_RO_DAYS_THRESHOLD))
End Function

Function ExtractOpenedDateToken(screenContent)
    Dim regEx, matches
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = False
    ' Match CDK date format DDMMMYY/DDMMMYYYY (e.g. 05NOV25) or slash format (e.g. 01/20/26)
    ' Anchoring digit+letter+digit prevents grabbing trailing tokens like 'RO' from the next field
    regEx.Pattern = "OPENED DATE:\s*(\d{1,2}[A-Z]{3}\d{2,4}|\d{1,2}/\d{1,2}/\d{2,4})"

    If regEx.Test(screenContent) Then
        Set matches = regEx.Execute(screenContent)
        ExtractOpenedDateToken = Trim(matches(0).SubMatches(0))
    Else
        ExtractOpenedDateToken = ""
    End If
End Function

Function TryParseCdkDate(rawDateValue, ByRef parsedDate)
    Dim token, regEx, matches
    Dim dayPart, monthPart, yearPart
    Dim dayNumber, monthNumber, yearNumber

    token = UCase(Trim(rawDateValue))
    If token = "" Then
        TryParseCdkDate = False
        Exit Function
    End If

    If InStr(token, "/") > 0 Then
        If IsDate(token) Then
            parsedDate = CDate(token)
            TryParseCdkDate = True
            Exit Function
        End If
    End If

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = False
    regEx.Pattern = "^(\d{1,2})([A-Z]{3})(\d{2,4})$"

    If Not regEx.Test(token) Then
        TryParseCdkDate = False
        Exit Function
    End If

    Set matches = regEx.Execute(token)
    dayPart = matches(0).SubMatches(0)
    monthPart = matches(0).SubMatches(1)
    yearPart = matches(0).SubMatches(2)

    dayNumber = CInt(dayPart)
    monthNumber = MonthNumberFromAbbrev(monthPart)
    If monthNumber = 0 Then
        TryParseCdkDate = False
        Exit Function
    End If

    If Len(yearPart) = 2 Then
        yearNumber = CInt(yearPart)
        If yearNumber >= 70 Then
            yearNumber = 1900 + yearNumber
        Else
            yearNumber = 2000 + yearNumber
        End If
    Else
        yearNumber = CInt(yearPart)
    End If

    parsedDate = DateSerial(yearNumber, monthNumber, dayNumber)
    TryParseCdkDate = True
End Function

Function MonthNumberFromAbbrev(monthAbbrev)
    Select Case UCase(Trim(monthAbbrev))
        Case "JAN": MonthNumberFromAbbrev = 1
        Case "FEB": MonthNumberFromAbbrev = 2
        Case "MAR": MonthNumberFromAbbrev = 3
        Case "APR": MonthNumberFromAbbrev = 4
        Case "MAY": MonthNumberFromAbbrev = 5
        Case "JUN": MonthNumberFromAbbrev = 6
        Case "JUL": MonthNumberFromAbbrev = 7
        Case "AUG": MonthNumberFromAbbrev = 8
        Case "SEP": MonthNumberFromAbbrev = 9
        Case "OCT": MonthNumberFromAbbrev = 10
        Case "NOV": MonthNumberFromAbbrev = 11
        Case "DEC": MonthNumberFromAbbrev = 12
        Case Else: MonthNumberFromAbbrev = 0
    End Select
End Function

Function GetStatusSnip(screenContent)
    ' Helper to grab a small snip of where the status usually is for logging
    Dim pos: pos = InStr(1, screenContent, "STATUS:", vbTextCompare)
    If pos > 0 Then
        GetStatusSnip = "'" & Trim(Mid(screenContent, pos, 30)) & "'"
    Else
        GetStatusSnip = "(Status line not found in read buffer)"
    End If
End Function

Function ExtractStatusText(screenContent)
    Dim pos, snip
    pos = InStr(1, screenContent, "STATUS:", vbTextCompare)
    If pos = 0 Then
        ExtractStatusText = ""
        Exit Function
    End If
    snip = Trim(Mid(screenContent, pos + 7, 25))
    ' Trim at next field boundary (double space or end)
    Dim spPos: spPos = InStr(snip, "  ")
    If spPos > 0 Then snip = Left(snip, spPos - 1)
    ExtractStatusText = Trim(snip)
End Function

Function DiscoverLineLetters()
    Dim i, recordKey, record, letters, letterIndex
    
    ' Use extensible line-record scanner
    ScanVisibleLineHeaders()
    
    ' Extract just the line letters from scanned records
    If g_CurrentPageLineRecords.Count = 0 Then
        DiscoverLineLetters = Array()
        Exit Function
    End If
    
    ReDim letters(g_CurrentPageLineRecords.Count - 1)
    letterIndex = 0
    
    ' Extract letters in order by iterating the dictionary
    For Each recordKey In g_CurrentPageLineRecords.Keys
        Set record = g_CurrentPageLineRecords(recordKey)
        letters(letterIndex) = record("lineLetter")
        letterIndex = letterIndex + 1
    Next
    
    DiscoverLineLetters = letters
End Function


' --- Line Status Detection: Extensible Line-Record Architecture ---

' Global cache for scanned line records from current page
Dim g_CurrentPageLineRecords
Set g_CurrentPageLineRecords = CreateObject("Scripting.Dictionary")

Function CreateLineRecord(lineLetter, statusCode, rowNumber, description)
    Dim record
    Set record = CreateObject("Scripting.Dictionary")
    record("lineLetter") = UCase(Trim(CStr(lineLetter)))
    record("statusCode") = UCase(Trim(CStr(statusCode)))
    record("row") = rowNumber
    record("description") = Trim(CStr(description))
    Set CreateLineRecord = record
End Function

Function ScanVisibleLineHeaders()
    Dim i, row, buf, lineLetter, statusCode, description, recordKey
    Dim records, recordCount, lineLettersSummary, key
    
    ' Clear cache before scan
    Set g_CurrentPageLineRecords = CreateObject("Scripting.Dictionary")
    
    recordCount = 0
    
    ' Scan rows 10-22 (line-letter header rows before prompt area)
    For row = 10 To 22
        buf = ""
        g_bzhao.ReadScreen buf, 80, row, 1
        
        If Len(buf) >= 1 Then
            lineLetter = ExtractLineLetterFromHeaderRow(buf)

            ' Validate it's a line letter (A-Z)
            If Len(lineLetter) = 1 And Len(Trim(lineLetter)) > 0 Then
                If Asc(UCase(Trim(lineLetter))) >= Asc("A") And Asc(UCase(Trim(lineLetter))) <= Asc("Z") Then
                ' Extract status code with fixed-column first, then pattern fallback
                statusCode = ""
                If Len(buf) >= 49 Then
                    statusCode = Trim(Mid(buf, 42, 8))
                End If
                If Not IsLikelyLineStatusCode(statusCode) Then
                    statusCode = ExtractStatusCodeFromHeaderRow(buf)
                End If

                ' Extract description from columns 4-41
                description = ""
                If Len(buf) >= 41 Then
                    description = Trim(Mid(buf, 4, 38))
                End If

                ' Store in cache with line letter as key (prevent duplicate row captures)
                recordKey = UCase(lineLetter)
                If Not g_CurrentPageLineRecords.Exists(recordKey) Then
                    Dim lineRecord
                    Set lineRecord = CreateLineRecord(lineLetter, statusCode, row, description)
                    g_CurrentPageLineRecords.Add recordKey, lineRecord

                    recordCount = recordCount + 1
                    LogResult "INFO", "ScanVisibleLineHeaders: Line " & lineLetter & " status=" & statusCode & " desc=" & description
                End If
                End If
            End If
        End If
    Next

    If recordCount > 0 Then
        lineLettersSummary = ""
        For Each key In g_CurrentPageLineRecords.Keys
            If lineLettersSummary <> "" Then
                lineLettersSummary = lineLettersSummary & ","
            End If
            lineLettersSummary = lineLettersSummary & CStr(key)
        Next
        LogResult "INFO", "ScanVisibleLineHeaders: Visible lines=" & lineLettersSummary
    End If
    
    ScanVisibleLineHeaders = recordCount
End Function

Function ExtractLineLetterFromHeaderRow(rowText)
    Dim colIdx, ch
    ExtractLineLetterFromHeaderRow = ""

    ' Primary: first three columns where RO line letters appear most often.
    For colIdx = 1 To 3
        If Len(rowText) >= colIdx Then
            ch = UCase(Mid(rowText, colIdx, 1))
            If ch >= "A" And ch <= "Z" Then
                ExtractLineLetterFromHeaderRow = ch
                Exit Function
            End If
        End If
    Next
End Function

Function IsLikelyLineStatusCode(statusCode)
    Dim normalized
    normalized = UCase(Trim(CStr(statusCode)))

    If Len(normalized) < 3 Then
        IsLikelyLineStatusCode = False
        Exit Function
    End If

    If Left(normalized, 1) <> "C" And Left(normalized, 1) <> "I" And Left(normalized, 1) <> "H" Then
        IsLikelyLineStatusCode = False
        Exit Function
    End If

    IsLikelyLineStatusCode = IsNumeric(Mid(normalized, 2, Len(normalized) - 1))
End Function

Function ExtractStatusCodeFromHeaderRow(rowText)
    Dim re, matches
    ExtractStatusCodeFromHeaderRow = ""

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "([CIH][0-9]{2,3})"

    If re.Test(rowText) Then
        Set matches = re.Execute(rowText)
        ExtractStatusCodeFromHeaderRow = UCase(matches(0).SubMatches(0))
    End If
End Function

Function GetLineStatus(lineLetter)
    Dim key, record
    key = UCase(Trim(CStr(lineLetter)))
    
    If g_CurrentPageLineRecords.Exists(key) Then
        Set record = g_CurrentPageLineRecords(key)
        GetLineStatus = record("statusCode")
    Else
        GetLineStatus = ""
    End If
End Function

Function CreateStatusActionMap()
    Dim statusMap
    Set statusMap = CreateObject("Scripting.Dictionary")
    
    ' Map status patterns to actions
    ' Action: "REVIEW" (send R <line>)
    statusMap("C92") = "REVIEW"
    
    ' Action: "SKIP_REVIEWED" (line already reviewed, skip)
    statusMap("C93") = "SKIP_REVIEWED"
    
    ' Note: Ixx and Hxx and unknown are handled separately as patterns, not exact keys
    
    Set CreateStatusActionMap = statusMap
End Function

Function GetLineActionFromStatus(statusCode)
    Dim statusMap
    Set statusMap = CreateStatusActionMap()
    
    Dim status, normalizedStatus
    normalizedStatus = UCase(Trim(CStr(statusCode)))
    
    ' Check exact match first
    If statusMap.Exists(normalizedStatus) Then
        GetLineActionFromStatus = statusMap(normalizedStatus)
        Exit Function
    End If
    
    ' Check patterns
    If Left(normalizedStatus, 1) = "C" And Len(normalizedStatus) >= 3 Then
        If statusMap.Exists(Left(normalizedStatus, 3)) Then
            GetLineActionFromStatus = statusMap(Left(normalizedStatus, 3))
            Exit Function
        End If
    End If
    
    ' Pattern checks for I* (in progress), H* (hold), unknown
    If Left(normalizedStatus, 1) = "I" Then
        GetLineActionFromStatus = "FINISH_AND_REROUTE"
    ElseIf Left(normalizedStatus, 1) = "H" Then
        GetLineActionFromStatus = "SKIP_RO_ON_HOLD"
    ElseIf normalizedStatus <> "" Then
        GetLineActionFromStatus = "SKIP_UNKNOWN"
    Else
        GetLineActionFromStatus = "SKIP_UNKNOWN"
    End If
End Function

Function CheckRoLineStatuses()
    Dim i, recordKey, record, status, action
    Dim foundHold, allReviewed, hasLines
    
    foundHold = False
    allReviewed = True
    hasLines = (g_CurrentPageLineRecords.Count > 0)
    
    ' Iterate all scanned line records
    For Each recordKey In g_CurrentPageLineRecords.Keys
        Set record = g_CurrentPageLineRecords(recordKey)
        status = record("statusCode")
        action = GetLineActionFromStatus(status)
        
        ' Check for hold
        If action = "SKIP_RO_ON_HOLD" Then
            LogResult "INFO", "CheckRoLineStatuses: Line " & record("lineLetter") & " is on hold (" & status & "). Skipping entire RO."
            foundHold = True
            Exit For
        End If
        
        ' Check if all are reviewed (C93)
        If action <> "SKIP_REVIEWED" Then
            allReviewed = False
        End If
    Next
    
    ' Return early-skip reason if found
    If foundHold Then
        CheckRoLineStatuses = "HOLD_DETECTED"
    ElseIf allReviewed And hasLines Then
        LogResult "INFO", "CheckRoLineStatuses: All lines are C93 (already reviewed). Skipping entire RO."
        CheckRoLineStatuses = "ALL_REVIEWED"
    Else
        CheckRoLineStatuses = ""
    End If
End Function

Function CheckWarrantyPageGate()
    Dim unsupportedType
    unsupportedType = GetFirstUnsupportedWarrantyLaborType()
    If unsupportedType <> "" Then
        LogResult "INFO", "ProcessRoReview: Unsupported warranty labor type (" & unsupportedType & ") detected. Skipping review."
        CheckWarrantyPageGate = "UNSUPPORTED_WARRANTY"
    Else
        CheckWarrantyPageGate = ""
    End If
End Function

Function CheckHoldPageGate(ByVal gateResult)
    If gateResult = "HOLD_DETECTED" Then
        LogResult "INFO", "ProcessRoReview: RO has line(s) on hold. Skipping review."
        CheckHoldPageGate = "HOLD_DETECTED"
    Else
        CheckHoldPageGate = ""
    End If
End Function

Function CheckAllReviewedPageGate(ByVal gateResult)
    Dim screenContent
    LogResult "INFO", "ProcessRoReview: gateResult=" & gateResult

    If gateResult = "ALL_REVIEWED" Then
        ' Deterministic single-page fast path: if this page is the end, skip now.
        g_bzhao.ReadScreen screenContent, 1920, 1, 1
        If InStr(1, screenContent, "(END OF DISPLAY)", vbTextCompare) > 0 Then
            LogResult "INFO", "ProcessRoReview: All lines are C93 and end-of-display reached. Skipping review phase."
            CheckAllReviewedPageGate = "ALL_REVIEWED"
        Else
            CheckAllReviewedPageGate = ""
        End If
    Else
        CheckAllReviewedPageGate = ""
    End If
End Function

Function ProcessRoReview()
    Dim screenContent, pageCount, processedLetters, lineActionByLetter, lineLetterQueue, lineLetterQueueCount, i, recordKey, record
    Dim lineLetter, status, action, lineCount
    Dim hasScannedLines, hasActionableLine, pageWarrantyGate, pageHoldGate

    pageCount = 0
    Set processedLetters = CreateObject("Scripting.Dictionary")
    Set lineActionByLetter = CreateObject("Scripting.Dictionary")
    ReDim lineLetterQueue(0)
    lineLetterQueueCount = 0
    g_ReviewPhaseResult = "PROCEED"
    ProcessRoReview = True
    hasScannedLines = False
    hasActionableLine = False
    pageWarrantyGate = False
    pageHoldGate = False

    ' === PAGE SCAN PHASE: Accumulate all line statuses across all pages ===
    Do
        pageCount = pageCount + 1
        LogResult "INFO", "ProcessRoReview: Scanning page " & pageCount

        lineCount = ScanVisibleLineHeaders()
        If lineCount = 0 Then
            LogResult "INFO", "ProcessRoReview: No lines found on page " & pageCount & ". Review complete."
            Exit Do
        End If
        hasScannedLines = True

        ' Warranty gate (fail-fast, but only if found on any page)
        If Not pageWarrantyGate Then
            If CheckWarrantyPageGate() <> "" Then
                pageWarrantyGate = True
            End If
        End If

        ' Hold gate (fail-fast, but only if found on any page)
        Dim gateResult
        gateResult = CheckRoLineStatuses()
        If Not pageHoldGate Then
            If gateResult = "HOLD_DETECTED" Then
                pageHoldGate = True
            End If
        End If

        ' Accumulate only non-C93 lines across all pages.
        ' C93 (already reviewed) lines are intentionally excluded from the queue.
        For Each recordKey In g_CurrentPageLineRecords.Keys
            Set record = g_CurrentPageLineRecords(recordKey)
            lineLetter = record("lineLetter")
            status = record("statusCode")
            action = GetLineActionFromStatus(status)
            If action <> "SKIP_REVIEWED" And Not processedLetters.Exists(lineLetter) Then
                lineActionByLetter.Add lineLetter, action
                If lineLetterQueueCount = 0 Then
                    ReDim lineLetterQueue(0)
                Else
                    ReDim Preserve lineLetterQueue(lineLetterQueueCount)
                End If
                lineLetterQueue(lineLetterQueueCount) = lineLetter
                lineLetterQueueCount = lineLetterQueueCount + 1
                processedLetters.Add lineLetter, True
            End If
        Next

        ' Pagination check
        g_bzhao.ReadScreen screenContent, 1920, 1, 1
        If InStr(1, screenContent, "(END OF DISPLAY)", vbTextCompare) > 0 Then
            LogResult "INFO", "ProcessRoReview: Reached end of display. Page scan complete. PageCount=" & pageCount
            Exit Do
        Else
            LogResult "INFO", "ProcessRoReview: More lines exist. Advancing page."
            WaitForText "COMMAND:"
            EnterTextWithStability "N"
            g_bzhao.Pause STABILITY_PAUSE
        End If
    Loop

    ' === GATE EVALUATION PHASE: Decide what to do based on all accumulated lines ===
    If pageWarrantyGate Then
        LogResult "INFO", "ProcessRoReview: Unsupported warranty labor type detected on at least one page. Skipping review."
        EnterTextWithStability "E"
        g_ReviewPhaseResult = "SKIPPED"
        ProcessRoReview = False
        Exit Function
    End If

    If pageHoldGate Then
        LogResult "INFO", "ProcessRoReview: RO has line(s) on hold on at least one page. Skipping review."
        EnterTextWithStability "E"
        g_ReviewPhaseResult = "SKIPPED"
        ProcessRoReview = False
        Exit Function
    End If

    ' If no non-C93 lines were captured across all pages, skip with E.
    If hasScannedLines And lineLetterQueueCount = 0 Then
        LogResult "INFO", "ProcessRoReview: All scanned lines are C93 across all pages. Skipping review phase."
        EnterTextWithStability "E"
        g_ReviewPhaseResult = "SKIPPED"
        ProcessRoReview = False
        Exit Function
    End If

    ' === ACTION PHASE: Route all actionable lines ===
    If lineLetterQueueCount > 0 Then
        Dim reviewQueueText
        reviewQueueText = ""
        For i = 0 To lineLetterQueueCount - 1
            If reviewQueueText <> "" Then reviewQueueText = reviewQueueText & ","
            reviewQueueText = reviewQueueText & CStr(lineLetterQueue(i))
        Next
        LogResult "INFO", "Lines(" & reviewQueueText & ") need to be reviewed"
    End If

    For i = 0 To lineLetterQueueCount - 1
        lineLetter = CStr(lineLetterQueue(i))
        action = lineActionByLetter(lineLetter)
        status = ""
        ' Find the latest status for this lineLetter (from last page it appeared)
        If g_CurrentPageLineRecords.Exists(lineLetter) Then
            status = g_CurrentPageLineRecords(lineLetter)("statusCode")
        End If
        LogResult "INFO", "ProcessRoReview: Line " & lineLetter & " action=" & action
        Select Case action
            Case "REVIEW"
                WaitForText "COMMAND:"
                EnterTextWithStability "R " & lineLetter
                g_bzhao.Pause STABILITY_PAUSE
                If Not HandleReviewPrompts(lineLetter) Then
                    LogResult "ERROR", "ProcessRoReview: Review failed for line " & lineLetter
                    g_ReviewPhaseResult = "FAILED"
                    ProcessRoReview = False
                    Exit Function
                End If
            Case "SKIP_REVIEWED"
                LogResult "INFO", "ProcessRoReview: Line " & lineLetter & " already reviewed (C93). Skipping review."
            Case "FINISH_AND_REROUTE"
                LogResult "INFO", "ProcessRoReview: Line " & lineLetter & " not finished (I*). Sending FNL command."
                WaitForText "COMMAND:"
                EnterTextWithStability "FNL " & lineLetter
                g_bzhao.Pause STABILITY_PAUSE
                ' Re-scan this specific line to get updated status
                Dim updatedStatus
                updatedStatus = GetLineStatus(lineLetter)
                If updatedStatus = "" Then
                    ScanVisibleLineHeaders()
                    updatedStatus = GetLineStatus(lineLetter)
                End If
                Dim updatedAction
                updatedAction = GetLineActionFromStatus(updatedStatus)
                LogResult "INFO", "ProcessRoReview: After FNL, Line " & lineLetter & " status=" & updatedStatus & " new-action=" & updatedAction
                If updatedAction = "REVIEW" Then
                    WaitForText "COMMAND:"
                    EnterTextWithStability "R " & lineLetter
                    g_bzhao.Pause STABILITY_PAUSE
                    If Not HandleReviewPrompts(lineLetter) Then
                        LogResult "ERROR", "ProcessRoReview: Review failed for line " & lineLetter & " after FNL"
                        g_ReviewPhaseResult = "FAILED"
                        ProcessRoReview = False
                        Exit Function
                    End If
                ElseIf updatedAction = "SKIP_REVIEWED" Then
                    LogResult "INFO", "ProcessRoReview: Line " & lineLetter & " transitioned to C93 after FNL. Skipping review."
                End If
            Case "SKIP_RO_ON_HOLD"
                LogResult "ERROR", "ProcessRoReview: Line " & lineLetter & " is on hold. RO-level gate should have caught this."
                g_ReviewPhaseResult = "FAILED"
                ProcessRoReview = False
                Exit Function
            Case "SKIP_UNKNOWN"
                LogResult "INFO", "ProcessRoReview: Line " & lineLetter & " has unknown status (" & status & "). Skipping."
        End Select
    Next

    LogResult "INFO", "ProcessRoReview: hasScannedLines=" & hasScannedLines & " -> proceed"
    ProcessRoReview = True
End Function

Function HandleReviewPrompts(lineLetter)
    Dim screenContent, startTime, elapsed, regEx, warrantyDialogType, handledWarrantyDialog
    Dim promptResponse, promptLogLevel, promptLogMessage
    Dim unhandledPromptText, unhandledPromptSignature, lastUnhandledPromptSignature, lastUnhandledPromptLogTime
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = False
    
    startTime = Timer
    lastUnhandledPromptSignature = ""
    lastUnhandledPromptLogTime = -1
    
    Do
        g_bzhao.Pause REVIEW_PAUSE ' Use faster review pause
        ' Read the entire screen to handle prompts that might appear mid-screen
        g_bzhao.ReadScreen screenContent, 1920, 1, 1
        
        ' Exit condition: Back to COMMAND prompt
        If InStr(1, screenContent, "COMMAND:", vbTextCompare) > 0 Then
            HandleReviewPrompts = True
            Exit Function
        End If

        ' Warranty manufacturer dialogs can appear after review prompts (WF/WCH/W).
        ' Handle them in-line, then continue polling until COMMAND: returns.
        handledWarrantyDialog = False
        warrantyDialogType = DetectMaintenanceWarrantyDialog()
        If warrantyDialogType <> "" Then
            LogResult "INFO", "HandleReviewPrompts: detected warranty dialog type " & warrantyDialogType & " for Line " & lineLetter
            HandleMaintenanceWarrantyClaimsDialog warrantyDialogType
            startTime = Timer
            handledWarrantyDialog = True
        End If
        
        If Not handledWarrantyDialog Then
            promptResponse = ""
            promptLogLevel = ""
            promptLogMessage = ""
            If GetReviewPromptAction(regEx, screenContent, lineLetter, promptResponse, promptLogLevel, promptLogMessage) Then
                If promptLogMessage <> "" Then
                    LogResult promptLogLevel, promptLogMessage
                End If
                EnterReviewPrompt promptResponse
                lastUnhandledPromptSignature = ""
                lastUnhandledPromptLogTime = -1
            Else
                unhandledPromptText = GetPromptAreaText()
                unhandledPromptSignature = UCase(Trim(Replace(unhandledPromptText, vbCrLf, "|")))

                ' Throttle repeated logs while still surfacing new prompt text quickly.
                If unhandledPromptSignature <> "" Then
                    If (unhandledPromptSignature <> lastUnhandledPromptSignature) Or _
                       (lastUnhandledPromptLogTime < 0) Or _
                       ((Timer - lastUnhandledPromptLogTime) > 3) Then
                        LogResult "WARN", "Unhandled review prompt for Line " & lineLetter & ": " & unhandledPromptText
                        lastUnhandledPromptSignature = unhandledPromptSignature
                        lastUnhandledPromptLogTime = Timer
                    End If
                End If
            End If
        End If
        
        elapsed = Timer - startTime
        If elapsed > 45 Then ' Increased timeout for slow terminal moves
            LogResult "ERROR", "Timeout in HandleReviewPrompts for Line " & lineLetter
            HandleReviewPrompts = False
            Exit Function
        End If
    Loop
End Function

Function GetPromptAreaText()
    Dim row, buf, lines, lineText
    lines = ""

    ' Prompt/input area is typically near the bottom of the 24x80 screen.
    For row = 20 To 24
        buf = ""
        g_bzhao.ReadScreen buf, 80, row, 1

        lineText = Trim(CStr(buf))
        If lineText <> "" Then
            If lines <> "" Then lines = lines & " | "
            lines = lines & "R" & CStr(row) & ":" & lineText
        End If
    Next

    GetPromptAreaText = lines
End Function

Function GetReviewPromptAction(regEx, screenContent, lineLetter, ByRef responseText, ByRef logLevel, ByRef logMessage)
    responseText = ""
    logLevel = ""
    logMessage = ""
    GetReviewPromptAction = True

    ' Match prompts in priority order to avoid broad patterns swallowing specific ones.
    If TestPrompt(regEx, screenContent, "LABOR TYPE") Then
        responseText = ""
    ' Prefer pattern that explicitly contains a parenthesized default (accept default)
    ElseIf TestPrompt(regEx, screenContent, "OP CODE.*\([A-Za-z0-9]+\)\?|OPERATION CODE.*\([A-Za-z0-9]+\)\?") Then
        responseText = ""
    ' Fallback: no-parenthesis variant (e.g. "OPERATION CODE FOR LINE A, L1?")
    ElseIf TestPrompt(regEx, screenContent, "OP CODE.*\?|OPERATION CODE.*\?") Then
        responseText = "I"
    ElseIf TestPrompt(regEx, screenContent, "DESC:") Then
        responseText = ""
    ' TECHNICIAN: accept valid default when present; otherwise force fallback 99
    ElseIf TestPrompt(regEx, screenContent, "TECHNICIAN.*\([A-Za-z0-9]+\)\?") Then
        responseText = ""
    ElseIf TestPrompt(regEx, screenContent, "TECHNICIAN.*\?") Then
        responseText = "99"
        logLevel = "WARN"
        logMessage = "TECHNICIAN prompt has no default for Line " & lineLetter & " — sending 99"
    ' ACTUAL HOURS: accept default if present, otherwise send 0
    ElseIf TestPrompt(regEx, screenContent, "ACTUAL HOURS.*\([A-Za-z0-9\.]+\)") Then
        responseText = ""
    ElseIf TestPrompt(regEx, screenContent, "ACTUAL HOURS.*\?") Then
        responseText = "0"
    ' SOLD HOURS: accept default if present, otherwise send 0
    ElseIf TestPrompt(regEx, screenContent, "SOLD HOURS.*\([A-Za-z0-9\.]+\)") Then
        responseText = ""
    ElseIf TestPrompt(regEx, screenContent, "SOLD HOURS.*\?") Then
        responseText = "0"
    ElseIf TestPrompt(regEx, screenContent, "IS THIS A COMEBACK.*\(Y/N\)") Then
        responseText = "Y"
        logLevel = "INFO"
        logMessage = "Comeback prompt detected for Line " & lineLetter & " - sending Y"
    ElseIf TestPrompt(regEx, screenContent, "ADD A LABOR OPERATION") Then
        responseText = "" ' Defaults to "N"
    Else
        GetReviewPromptAction = False
    End If
End Function

Sub EnterReviewPrompt(text)
    ' Fast entry for review fields that don't trigger large screen transitions
    If text <> "" Then g_bzhao.SendKey CStr(text)
    g_bzhao.Pause 50
    g_bzhao.SendKey "<NumpadEnter>"
    g_bzhao.Pause REVIEW_PAUSE ' Use faster review pause instead of stability pause
End Sub

Function TestPrompt(regEx, text, pattern)
    regEx.Pattern = pattern
    TestPrompt = regEx.Test(text)
End Function

Function DetectMaintenanceWarrantyDialog()
    Dim row, buf, hasLaborOp, hasClaimType
    DetectMaintenanceWarrantyDialog = ""
    hasLaborOp = False
    hasClaimType = False

    For row = 1 To 24
        buf = ""
        g_bzhao.ReadScreen buf, 80, row, 1

        If InStr(1, buf, "MODIFY FORD REPAIR TYPE INFORMATION", vbTextCompare) > 0 Then
            DetectMaintenanceWarrantyDialog = "FORD"
            Exit Function
        End If
        If InStr(1, buf, "LABOR OP:", vbTextCompare) > 0 Then hasLaborOp = True
        If InStr(1, buf, "CLAIM TYPE:", vbTextCompare) > 0 Then hasClaimType = True
        If InStr(1, buf, "FAILURE CODE:", vbTextCompare) > 0 Or InStr(1, buf, "MODIFY WARRANTY INFORMATION", vbTextCompare) > 0 Then
            DetectMaintenanceWarrantyDialog = "W"
            Exit Function
        End If
    Next

    ' Require both markers to reduce false positives on non-dialog screens.
    If hasLaborOp And hasClaimType Then
        DetectMaintenanceWarrantyDialog = "FCA"
    End If
End Function

Sub HandleMaintenanceWarrantyClaimsDialog(dialogType)
    If dialogType = "FCA" Then
        HandleMaintenanceFcaWarrantyDialog
    ElseIf dialogType = "W" Then
        HandleMaintenanceWWarrantyDialog
    ElseIf dialogType = "FORD" Then
        HandleMaintenanceFordWarrantyDialog
    Else
        LogResult "WARN", "HandleMaintenanceWarrantyClaimsDialog: unknown dialog type " & dialogType
    End If
End Sub

Sub HandleMaintenanceFcaWarrantyDialog()
    Dim row, buf, detected, i
    detected = ""

    For i = 1 To 20
        For row = 20 To 24
            buf = ""
            g_bzhao.ReadScreen buf, 80, row, 1

            If InStr(1, buf, "LABOR OP:", vbTextCompare) > 0 Then
                detected = "LABOROP"
                Exit For
            End If
            If InStr(1, buf, "COMMAND:", vbTextCompare) > 0 Then
                detected = "COMMAND"
                Exit For
            End If
        Next
        If detected <> "" Then Exit For
        g_bzhao.Pause 500
    Next

    If detected = "LABOROP" Then
        g_bzhao.SendKey "<NumpadEnter>"
        g_bzhao.Pause WARRANTY_DIALOG_STEP_DELAY_MS
        HandleMaintenanceCausePromptLoop FORD_WARRANTY_CAUSE_TEXT
    ElseIf detected = "COMMAND" Then
        g_bzhao.SendKey "."
        g_bzhao.Pause REVIEW_PAUSE
        g_bzhao.SendKey "<NumpadEnter>"
        g_bzhao.Pause WARRANTY_DIALOG_STEP_DELAY_MS
        g_bzhao.SendKey "E"
        g_bzhao.Pause REVIEW_PAUSE
        g_bzhao.SendKey "<NumpadEnter>"
        g_bzhao.Pause REVIEW_PAUSE
    Else
        LogResult "WARN", "HandleMaintenanceFcaWarrantyDialog: expected FCA markers not found before timeout."
    End If
End Sub

Sub HandleMaintenanceWWarrantyDialog()
    Dim row, buf, i, commandFound
    commandFound = False

    For i = 1 To 15
        For row = 20 To 24
            buf = ""
            g_bzhao.ReadScreen buf, 80, row, 1
            If InStr(1, buf, "WARRANTY COMMAND:", vbTextCompare) > 0 Then
                commandFound = True
                Exit For
            End If
        Next

        If commandFound Then Exit For
        g_bzhao.SendKey "<NumpadEnter>"
        g_bzhao.Pause WARRANTY_DIALOG_STEP_DELAY_MS
    Next

    If commandFound Then
        g_bzhao.SendKey "E"
        g_bzhao.Pause REVIEW_PAUSE
        g_bzhao.SendKey "<NumpadEnter>"
        g_bzhao.Pause REVIEW_PAUSE
    End If
End Sub

Sub HandleMaintenanceFordWarrantyDialog()
    Dim vlsRow, vlsBuf, vlsLabelPos, vlsFieldText, vlsCharIdx, startupScreen

    startupScreen = ""
    g_bzhao.ReadScreen startupScreen, 1920, 1, 1

    If InStr(1, startupScreen, "MODIFY FORD REPAIR TYPE INFORMATION", vbTextCompare) = 0 Then
        LogResult "WARN", "HandleMaintenanceFordWarrantyDialog: FORD dialog header not found. Skipping Ford dialog handler."
        Exit Sub
    End If

    ' Step 1-3: P&A, Ford/L-M Make, Franchise Model (accept defaults)
    g_bzhao.SendKey "<NumpadEnter>"
    g_bzhao.Pause WARRANTY_DIALOG_STEP_DELAY_MS
    g_bzhao.SendKey "<NumpadEnter>"
    g_bzhao.Pause WARRANTY_DIALOG_STEP_DELAY_MS
    g_bzhao.SendKey "<NumpadEnter>"
    g_bzhao.Pause WARRANTY_DIALOG_STEP_DELAY_MS

    ' Step 4: Vehicle License State (type only if blank)
    vlsLabelPos = 0
    vlsFieldText = ""
    For vlsRow = 1 To 24
        vlsBuf = ""
        g_bzhao.ReadScreen vlsBuf, 80, vlsRow, 1
        vlsLabelPos = InStr(1, vlsBuf, "VEHICLE LICENSE STATE:", vbTextCompare)
        If vlsLabelPos > 0 Then
            vlsFieldText = Trim(Mid(vlsBuf, vlsLabelPos + 22, 5))
            Exit For
        End If
    Next

    If Len(vlsFieldText) = 0 Then
        For vlsCharIdx = 1 To Len(FORD_WARRANTY_LICENSE_STATE)
            g_bzhao.SendKey Mid(FORD_WARRANTY_LICENSE_STATE, vlsCharIdx, 1)
            g_bzhao.Pause 100
        Next
    End If
    g_bzhao.SendKey "<NumpadEnter>"
    g_bzhao.Pause WARRANTY_DIALOG_STEP_DELAY_MS

    ' Step 5-6: Repair Type=1, then Command='.'
    g_bzhao.SendKey "1"
    g_bzhao.Pause REVIEW_PAUSE
    g_bzhao.SendKey "<NumpadEnter>"
    g_bzhao.Pause WARRANTY_DIALOG_STEP_DELAY_MS

    g_bzhao.SendKey "."
    g_bzhao.Pause REVIEW_PAUSE
    g_bzhao.SendKey "<NumpadEnter>"
    g_bzhao.Pause WARRANTY_DIALOG_STEP_DELAY_MS

    HandleMaintenanceCausePromptLoop FORD_WARRANTY_CAUSE_TEXT
End Sub

Sub HandleMaintenanceCausePromptLoop(causeText)
    Dim causeRow, causeBuf, causeFound, ci, causePoll

    For ci = 1 To 10
        causeFound = False
        For causePoll = 1 To 6
            For causeRow = 20 To 24
                causeBuf = ""
                g_bzhao.ReadScreen causeBuf, 80, causeRow, 1
                If InStr(1, causeBuf, "CAUSE L", vbTextCompare) > 0 Then
                    causeFound = True
                    Exit For
                End If
            Next
            If causeFound Then Exit For
            g_bzhao.Pause 500
        Next

        If Not causeFound Then Exit For

        g_bzhao.SendKey causeText
        g_bzhao.Pause REVIEW_PAUSE
        g_bzhao.SendKey "<NumpadEnter>"
        g_bzhao.Pause REVIEW_PAUSE
    Next
End Sub

Function CloseRoFinal()
    ' Phase II: The Closing
    WaitForText "COMMAND:"
    EnterTextWithStability "F"

    ' ALL LABOR POSTED (Y/N)?
    WaitForText "ALL LABOR POSTED"
    EnterTextWithStability "Y"

    ' Verify CDK returned to a known-good state — without this, false positives
    ' occur when a follow-up prompt appears after the Y response.
    Dim screenContent, elapsed
    elapsed = 0
    Do
        g_bzhao.Pause LOOP_PAUSE
        elapsed = elapsed + (LOOP_PAUSE / 1000)
        g_bzhao.ReadScreen screenContent, 1920, 1, 1
        If InStr(1, screenContent, "COMMAND:", vbTextCompare) > 0 Or _
           InStr(1, screenContent, MAIN_PROMPT, vbTextCompare) > 0 Then
            LogResult "INFO", "SUCCESS: RO finalized (FILE command sent)."
            CloseRoFinal = True
            Exit Function
        End If
        If elapsed >= 30 Then
            LogResult "ERROR", "CloseRoFinal: Timeout waiting for post-close state"
            CloseRoFinal = False
            Exit Function
        End If
    Loop
End Function

Sub ReturnToMainPrompt()
    Dim screenContent, i, targets, j, isFound, waitStep
    targets = Split(MAIN_PROMPT, "|")
    
    ' Phase 1: Patience. Wait for the terminal to land on the prompt naturally.
    ' This prevents sending "E" during slow transitions (the "E" bug).
    For waitStep = 1 To 10 ' Wait up to 5 seconds total (10 * 500ms)
        g_bzhao.Pause LOOP_PAUSE
        g_bzhao.ReadScreen screenContent, 1920, 1, 1
        
        For j = 0 To UBound(targets)
            If InStr(1, screenContent, targets(j), vbTextCompare) > 0 Then
                LogResult "INFO", "Confirmed at main prompt: " & targets(j)
                Exit Sub
            End If
        Next
    Next
    
    ' Phase 2: Recovery. If still lost, try to exit/clear using "E".
    For i = 1 To 3
        LogResult "INFO", "ReturnToMainPrompt: Still not at target. Attempting recovery 'E' (" & i & "/3)..."
        g_bzhao.SendKey "E"
        g_bzhao.SendKey "<NumpadEnter>"
        
        ' Wait for response after sending E
        For waitStep = 1 To 4 ' Wait up to 2 seconds
            g_bzhao.Pause LOOP_PAUSE
            g_bzhao.ReadScreen screenContent, 1920, 1, 1
            
            For j = 0 To UBound(targets)
                If InStr(1, screenContent, targets(j), vbTextCompare) > 0 Then
                    LogResult "INFO", "Recovered to main prompt: " & targets(j)
                    Exit Sub
                End If
            Next
        Next
    Next
    
    LogResult "ERROR", "ReturnToMainPrompt failed to find target: " & MAIN_PROMPT
End Sub

Sub WaitForText(targetText)
    Dim elapsed, screenContent, targets, found, i, isMainPrompt
    targets = Split(targetText, "|")
    elapsed = 0
    isMainPrompt = (InStr(1, targetText, MAIN_PROMPT, vbTextCompare) > 0)
    
    Do
        g_bzhao.Pause LOOP_PAUSE
        elapsed = elapsed + LOOP_PAUSE
        
        g_bzhao.ReadScreen screenContent, 1920, 1, 1
        
        found = False
        For i = 0 To UBound(targets)
            If InStr(1, screenContent, targets(i), vbTextCompare) > 0 Then
                found = True
                Exit For
            End If
        Next
        
        If found Then Exit Sub
        
        ' Simple recovery if lost while seeking main prompt
        If isMainPrompt And elapsed >= 5000 Then
            If elapsed Mod 5000 = 0 Then 
                LogResult "INFO", "Seeking main prompt. Sending 'E' to clear screen."
                g_bzhao.SendKey "E"
                g_bzhao.SendKey "<NumpadEnter>"
                g_bzhao.Pause LOOP_PAUSE
            End If
        End If

        If elapsed >= 60000 Then 
            TerminateScript "Critical timeout waiting for: " & targetText
            Exit Do
        End If
    Loop
End Sub

Sub EnterTextWithStability(text)
    LogResult "INFO", "Input State: Sending text '" & text & "' to terminal."
    g_bzhao.SendKey CStr(text)
    g_bzhao.Pause 150
    g_bzhao.SendKey "<NumpadEnter>"
    g_bzhao.Pause STABILITY_PAUSE ' Configurable stability pause
End Sub

Sub LogResult(logType, message)
    Dim fso, logFile, typeLevel
    Select Case UCase(logType)
        Case "ERROR": typeLevel = 1
        Case "INFO": typeLevel = 2
        Case Else: typeLevel = 2
    End Select
    
    If typeLevel <= DEBUG_LEVEL Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set logFile = fso.OpenTextFile(LOG_FILE_PATH, 8, True)
        logFile.WriteLine Now & " [" & logType & "] " & message
        logFile.Close
        Set logFile = Nothing
        Set fso = Nothing
    End If
End Sub

Sub TerminateScript(reason)
    LogResult "ERROR", "TERMINATING SCRIPT: " & reason
    If Not g_bzhao Is Nothing Then
        g_bzhao.Disconnect
        g_bzhao.StopScript
    End If
    Host_Quit
End Sub

' Execute
RunAutomation
