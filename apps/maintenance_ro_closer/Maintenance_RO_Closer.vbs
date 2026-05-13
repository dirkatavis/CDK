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

Dim g_SkipRoLookup
Dim g_SupportedWarrantyLTypes

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
    
    ' Load SkipRoList (required configuration for deterministic skip behavior)
    LoadSkipRoLookup SKIP_RO_LIST_PATH
    LogResult "INFO", "SkipRoList loaded from: " & SKIP_RO_LIST_PATH & " (entries=" & g_SkipRoLookup.Count & ")"
    
    ' Connect to terminal only after configuration and file existence are verified
    On Error Resume Next
    g_bzhao.Connect ""
    If Err.Number <> 0 Then
        LogResult "ERROR", "Failed to connect to BlueZone: " & Err.Description
        MsgBox "Failed to connect to BlueZone terminal session.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Start processing with unified error handling
    ProcessRoList fso, successfulCount
    
    ' Final graceful disconnect
    On Error Resume Next
    If Not g_bzhao Is Nothing Then g_bzhao.Disconnect
    On Error GoTo 0

    LogResult "INFO", "Automation complete. Total successful closures: " & successfulCount
    MsgBox "Maintenance RO Auto-Closer Finished." & vbCrLf & "Successful Closures: " & successfulCount, vbInformation
End Sub

' --- Helper Subroutines & Functions ---

Sub ProcessRoList(fso, ByRef successfulCount)
    Dim ts, strLine, roFromCsv, currentRo
    
    On Error Resume Next
    Set ts = fso.OpenTextFile(RO_LIST_PATH, 1) ' 1 = ForReading
    
    If Err.Number <> 0 Then
        LogResult "ERROR", "CRITICAL: Failed to open RO List file: " & Err.Description
        MsgBox "Failed to open RO List: " & RO_LIST_PATH, vbCritical
        Exit Sub
    End If

    Do While Not ts.AtEndOfStream
        If Err.Number <> 0 Then 
            LogResult "ERROR", "Unexpected runtime error: " & Err.Description
            Err.Clear
        End If

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
                    LogResult "INFO", "RO " & currentRo & " is processable."
                    If ShouldProcessRoByBusinessRules(currentRo) Then
                        If ProcessRoReview() Then
                            If CloseRoFinal() Then
                                LogResult "INFO", "SUCCESS: RO " & currentRo & " finalized and closed."
                                successfulCount = successfulCount + 1
                            Else
                                LogResult "ERROR", "Failed to close RO: " & currentRo & " during Phase II."
                            End If
                        Else
                            LogResult "ERROR", "Failed to complete review for RO: " & currentRo & " during Phase I."
                        End If
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
    On Error GoTo 0
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
    ' Any (non-blacklisted)        | Age >= AssumeClosedAfterDays     | CLOSE  (overrides status)
    ' READY TO POST                | (none)                           | CLOSE
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

    ' Gate 0: Unsupported warranty labor types (only configured types are allowed)
    Dim unsupportedWarrantyLType
    unsupportedWarrantyLType = GetFirstUnsupportedWarrantyLaborType()
    If unsupportedWarrantyLType <> "" Then
        LogResult "INFO", "RO " & roNumber & " | Unsupported warranty labor type (" & unsupportedWarrantyLType & ") detected. Skipping."
        ShouldProcessRoByBusinessRules = False
        Exit Function
    End If

    ' Gate 1: Blacklist
    If matchedBlacklistTerm <> "" Then
        LogResult "INFO", "RO " & roNumber & " | Blacklisted ('" & matchedBlacklistTerm & "'). Skipping."
        ShouldProcessRoByBusinessRules = False
        Exit Function
    End If

    ' Gate 2: Age exception — overrides status check but not blacklist
    If isOldEnough Then
        LogResult "INFO", "RO " & roNumber & " | Age exception: " & ageDays & " days old (threshold: " & OLD_RO_DAYS_THRESHOLD & "). Closing regardless of status."
        ShouldProcessRoByBusinessRules = True
        Exit Function
    End If

    ' Gate 3: READY TO POST
    If isReadyToPost Then
        LogResult "INFO", "RO " & roNumber & " | READY TO POST. Closing."
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
        On Error Resume Next
        g_bzhao.ReadScreen buf, 80, row, 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0

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

    On Error Resume Next
    parsedDate = DateSerial(yearNumber, monthNumber, dayNumber)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        TryParseCdkDate = False
        Exit Function
    End If
    On Error GoTo 0

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
    Dim i, capturedLetter, screenContentBuffer, readLength
    Dim foundLetters, foundCount
    Dim startReadRow, startReadColumn, emptyRowCount, nextColChar
    Dim startRow, endRow
    
    ' Array to store discovered line letters
    Dim tempLetters(25) ' Max 26 letters A-Z
    foundCount = 0
    emptyRowCount = 0
    
    ' The prompt area starts at row 23, so we must stop at row 22
    ' Anchor scanning at row 10 to skip header rows (e.g., REPAIR, REMARKS)
    startRow = 10 ' First data row in CDK
    endRow = 22  ' Last possible data row before prompt area
    
    For startReadRow = startRow To endRow
        startReadColumn = 1
        readLength = 1 ' Read just 1 character (the line letter)
        
        On Error Resume Next
        g_bzhao.ReadScreen screenContentBuffer, readLength, startReadRow, startReadColumn
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        On Error GoTo 0
        
        ' Trim and check if it's a valid letter (A-Z)
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
                    emptyRowCount = 0 ' Reset when a letter is found
                Else
                    emptyRowCount = emptyRowCount + 1
                End If
            Else
                emptyRowCount = emptyRowCount + 1
            End If
        Else
            emptyRowCount = emptyRowCount + 1
        End If

        ' If we hit 3 consecutive rows without a letter, we've likely finished the list
        If emptyRowCount >= 3 Then Exit For
    Next
    
    If foundCount = 0 Then
        DiscoverLineLetters = Array()
        Exit Function
    End If
    
    ' Create properly sized array with found letters
    ReDim foundLetters(foundCount - 1)
    For i = 0 To foundCount - 1
        foundLetters(i) = tempLetters(i)
    Next
    
    DiscoverLineLetters = foundLetters
End Function

Function ProcessRoReview()
    Dim letter, screenContent

    ' Step through lines A, B, C... until CDK signals the line does not exist.
    ' Sending "R <letter>" from COMMAND: navigates directly to that line's review
    ' prompts regardless of screen pagination — CDK handles the navigation.
    ' When a letter has no corresponding line, CDK returns immediately to COMMAND:.
    letter = "A"
    Do While Asc(letter) <= Asc("Z")

        WaitForText "COMMAND:"
        EnterTextWithStability "R " & letter

        ' Allow the screen to settle before checking state
        g_bzhao.Pause STABILITY_PAUSE
        g_bzhao.ReadScreen screenContent, 1920, 1, 1

        ' If CDK returned to COMMAND: immediately, this line does not exist — done
        If InStr(1, screenContent, "COMMAND:", vbTextCompare) > 0 Then
            LogResult "INFO", "ProcessRoReview: Line '" & letter & "' not found. Review complete."
            ProcessRoReview = True
            Exit Function
        End If

        ' Line exists — handle its review prompts
        LogResult "INFO", "ProcessRoReview: Reviewing Line '" & letter & "'."
        If Not HandleReviewPrompts(letter) Then
            LogResult "ERROR", "ProcessRoReview: Review timed out for Line '" & letter & "'."
            ProcessRoReview = False
            Exit Function
        End If

        letter = Chr(Asc(letter) + 1)
    Loop

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
        On Error Resume Next
        g_bzhao.ReadScreen buf, 80, row, 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0

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
        On Error Resume Next
        g_bzhao.ReadScreen buf, 80, row, 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0

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
            On Error Resume Next
            g_bzhao.ReadScreen buf, 80, row, 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0

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
            On Error Resume Next
            g_bzhao.ReadScreen buf, 80, row, 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
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
    On Error Resume Next
    g_bzhao.ReadScreen startupScreen, 1920, 1, 1
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

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
        On Error Resume Next
        g_bzhao.ReadScreen vlsBuf, 80, vlsRow, 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
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
                On Error Resume Next
                g_bzhao.ReadScreen causeBuf, 80, causeRow, 1
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
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
        On Error Resume Next
        Set logFile = fso.OpenTextFile(LOG_FILE_PATH, 8, True)
        If Err.Number = 0 Then
            logFile.WriteLine Now & " [" & logType & "] " & message
            logFile.Close
        End If
        On Error GoTo 0
        Set logFile = Nothing
        Set fso = Nothing
    End If
End Sub

Sub TerminateScript(reason)
    LogResult "ERROR", "TERMINATING SCRIPT: " & reason
    On Error Resume Next
    If Not g_bzhao Is Nothing Then
        g_bzhao.Disconnect
        g_bzhao.StopScript
    End If
    On Error GoTo 0
    Host_Quit
End Sub

' Execute
RunAutomation
