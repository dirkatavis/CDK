'====================================================================
' Script: Efficient VehicleDataAutomation.vbs
' Focus: Maximum speed with smart prompt detection - FULLY STANDARDIZED
'====================================================================

Option Explicit

' --- Bootstrap ---
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_sh: Set g_sh = CreateObject("WScript.Shell")
Dim g_root: g_root = g_sh.Environment("USER")("CDK_BASE")
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\PathHelper.vbs")).ReadAll

' --- CDK Terminal Object (must be declared before loading BZHelper) ---
Dim g_bzhao: Set g_bzhao = CreateObject("BZWhll.WhllObj")
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\BZHelper.vbs")).ReadAll

' --- Configuration Constants ---
Const POLL_INTERVAL = 1000   ' Check every 1000ms (1 time per second)
Const POST_ENTRY_WAIT = 200  ' Minimal wait after entry
Const PRE_KEY_WAIT = 1000    ' Pause before sending special keys
Const POST_KEY_WAIT = 1000   ' Pause after sending special keys
Const PROMPT_TIMEOUT_MS = 5000 ' Default prompt timeout
Const LOG_LEVEL_LOW = 1
Const LOG_LEVEL_MED = 2
Const LOG_LEVEL_HIGH = 3

Dim CSV_FILE_PATH: CSV_FILE_PATH = GetConfigPath("Open_RO", "CSV")
Dim OUTPUT_CSV_PATH: OUTPUT_CSV_PATH = GetConfigPath("Open_RO", "OutputCSV")
Dim SCRIPT_FOLDER: SCRIPT_FOLDER = "scripts\archive"
Dim SLOW_MARKER_PATH: SLOW_MARKER_PATH = GetConfigPath("Open_RO", "DebugMarker")
Const AppendMode = 8
Dim LOG_FILE_PATH: LOG_FILE_PATH = GetConfigPath("Open_RO", "Log")
Dim g_LogVerbosity: g_LogVerbosity = ResolveLogVerbosity()

Dim ts, strLine, arrValues, i, MVA, Mileage

' Initialize log file
Dim logInit
Set logInit = g_fso.OpenTextFile(LOG_FILE_PATH, AppendMode, True)
logInit.WriteLine "===================================================="
logInit.WriteLine "SESSION START: " & Now
logInit.WriteLine "===================================================="
logInit.Close
Set logInit = Nothing

' Initialize output CSV with headers
Dim csvOut
Set csvOut = g_fso.CreateTextFile(OUTPUT_CSV_PATH, True)
csvOut.WriteLine "RO_Number"
csvOut.Close
Set csvOut = Nothing

' Test logging immediately to verify log file creation
LOG "Script started - Log file path: " & LOG_FILE_PATH, "med"
LOG "Initialized output CSV (overwritten): " & OUTPUT_CSV_PATH, "med"


If g_fso.FileExists(CSV_FILE_PATH) Then
    LOG "CSV file found: " & CSV_FILE_PATH, "low"
    g_bzhao.Connect ""
    LOG "Connected to BlueZone", "low"
    Set ts = g_fso.OpenTextFile(CSV_FILE_PATH, 1)
    ts.ReadLine   ' Skip header row
    LOG "Processing CSV data...", "low"

    Do While Not ts.AtEndOfStream
        strLine = ts.ReadLine
        arrValues = Split(strLine, ",")
        
        If UBound(arrValues) >= 1 Then
            MVA = Trim(arrValues(0))
            Mileage = Trim(arrValues(1))
            Call Main(MVA, Mileage)
        End If
    Loop

    ts.Close
    Set ts = Nothing
End If

Set g_fso = Nothing
g_bzhao.Disconnect

'--------------------------------------------------------------------
' Subroutine: Main - Fully Standardized with WaitForPrompt
'--------------------------------------------------------------------
Sub Main(mva, mileage)
    
    '==== INPUT POINT 1: BEFORE ENTERING MVA ====
    ' NEED TO IDENTIFY: What prompt appears when CDK is ready for Vehicle ID?
    ' CURRENT: Using "Vehid....." - NEEDS VERIFICATION
    WaitForPromptSlow "Vehid.....", mva, True, PROMPT_TIMEOUT_MS, ""
    ' g_bzhao.Pause 1000


    ' Skip if no matching vehicle - check but don't enter anything
    If IsTextPresent("No matching") Then Exit Sub

    '==== INPUT POINT 1B: SEQUENCE NUMBER SELECTION ====
    ' Handles "CHOOSE ONE" or "SEQUENCE NUMBER" prompt - select option #1
    WaitForPromptSlow "CHOOSE ONE|SEQUENCE NUMBER", "1", True, PROMPT_TIMEOUT_MS, ""

    '==== INPUT POINT 2: BEFORE ENTERING COMMAND SELECTION ====
    ' NEED TO IDENTIFY: What menu/prompt shows before selecting command?
    ' CURRENT: Looking for "Command?" - NEEDS VERIFICATION
    WaitForPromptSlow "Command?", "<NumpadEnter>", False, PROMPT_TIMEOUT_MS, ""
    


    '==== INPUT POINT 4: BEFORE ENTERING MILEAGE ====
    ' NEED TO IDENTIFY: What prompt shows when mileage field is ready?
    ' CURRENT: Using "Miles In...:" - NEEDS VERIFICATION
    WaitForPromptSlow "Miles In", mileage, True, PROMPT_TIMEOUT_MS, ""
    

    '==== INPUT POINT 6: BEFORE ENTERING TAG ====
    ' NEED TO IDENTIFY: What field label appears for tag entry?
    ' CURRENT: Using "Tag......" - NEEDS VERIFICATION
    WaitForPromptSlow "Tag......", mva, True, PROMPT_TIMEOUT_MS, ""
    ' g_bzhao.Pause 1000

    '==== INPUT POINT 7: BEFORE ENTERING VENDOR ====
    ' NEED TO IDENTIFY: What prompt shows for vendor field?
    ' CURRENT: Using "PMVEND" - NEEDS VERIFICATION
    WaitForPromptSlow "Quick Codes", "PMVEND", True, PROMPT_TIMEOUT_MS, ""
    ' g_bzhao.Pause 1000

    '==== INPUT POINT 8: BEFORE F3 KEY ====
    ' NEED TO IDENTIFY: What screen/text indicates ready for F3?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    WaitForPromptSlow "Quick Code Description", "<F3>", False, PROMPT_TIMEOUT_MS, ""
    ' g_bzhao.Pause 1000

    '==== INPUT POINT 9: BEFORE F8 KEY ====
    ' NEED to IDENTIFY: What screen/text indicates ready for F8?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    WaitForPromptSlow "Quick Codes", "<F8>", False, PROMPT_TIMEOUT_MS, ""
    ' g_bzhao.Pause 1000
    
    '==== INPUT POINT 10: BEFORE ENTERING "99" ====
    ' NEED TO IDENTIFY: What prompt shows for "99" entry?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    WaitForPromptSlow "Tech", "99", False, PROMPT_TIMEOUT_MS, ""
    ' g_bzhao.Pause 1000
    
    '==== INPUT POINT 11: BEFORE SECOND F3 ====
    ' NEED TO IDENTIFY: What indicates ready for second F3?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    WaitForPromptSlow "Tech", "<F3>", False, PROMPT_TIMEOUT_MS, ""
    ' g_bzhao.Pause 1000
    
    '==== INPUT POINT 12: BEFORE THIRD F3 ====
    ' NEED TO IDENTIFY: What indicates ready for third F3?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    WaitForPromptSlow "Quick Codes", "<F3>", False, PROMPT_TIMEOUT_MS, ""
    ' g_bzhao.Pause 1000    
    
    '==== INPUT POINT 13: BEFORE FIRST ENTER KEY ====
    ' NEED TO IDENTIFY: What text shows system is ready for Enter?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    WaitForPromptSlow "Choose an option", "<NumpadEnter>", False, PROMPT_TIMEOUT_MS, ""
    ' g_bzhao.Pause 1000
    
    '==== INPUT POINT 14: BEFORE SECOND ENTER KEY ====
    ' NEED TO IDENTIFY: What prompt appears before second Enter?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    WaitForPromptSlow "MILEAGE OUT", "<NumpadEnter>", False, 30000, ""
    ' g_bzhao.Pause 1000
    
    '==== INPUT POINT 15: BEFORE THIRD ENTER KEY ====
    ' NEED TO IDENTIFY: What prompt appears before third Enter?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION

    WaitForPromptSlow "MILEAGE IN", "<NumpadEnter>", False, 10000, ""
    ' g_bzhao.Pause 1000
    
    '==== INPUT POINT 16: BEFORE ENTERING FINAL "N" ====
    ' NEED TO IDENTIFY: What question/prompt is asking for N response?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    WaitForPromptSlow "O.K. TO CLOSE RO", "N", True, 30000, ""
    ' g_bzhao.Pause 1000

    '==== INPUT POINT 17: WAIT FOR CONFIRMATION SCREEN, SCRAPE RO, THEN DISMISS ====
    ' Wait for the "Created repair order" confirmation screen before scraping.
    ' Scrape must happen AFTER this wait - the screen has not transitioned yet
    ' when called immediately after INPUT POINT 16.
    Dim confirmFound
    confirmFound = WaitForPromptSlow("Created repair order|R.O. NUMBER", "", False, 10000, "")

    ' Scrape and log
    Dim roNumber
    roNumber = GetRepairOrderEnhanced()
    Call LogEntryWithRO(mva, roNumber)

    ' Dismiss confirmation screen only if prompt was detected;
    ' avoid navigating away from an unexpected screen on timeout.
    If confirmFound Then
        g_bzhao.SendKey "<F3>"
    Else
        LOG "WARNING: Confirmation screen not detected - F3 not sent for MVA: " & mva, "med"
    End If
End Sub

'--------------------------------------------------------------------
' Function: WaitForPromptSlow
' Enforces per-action pacing: waits for the prompt, then applies
' PRE_KEY_WAIT before sending and POST_KEY_WAIT after sending.
'--------------------------------------------------------------------
Function WaitForPromptSlow(promptText, inputValue, sendEnter, timeoutMs, description)
    ' Phase 1: wait for the prompt to appear (no send)
    Dim found
    found = WaitForPrompt(promptText, "", False, timeoutMs, description)

    ' Phase 2: if found, apply delays around the send
    If found Then
        If Len(inputValue) > 0 Or sendEnter Then
            WaitMs PRE_KEY_WAIT
            If Len(inputValue) > 0 Then g_bzhao.SendKey inputValue
            If sendEnter Then g_bzhao.SendKey "<NumpadEnter>"
            WaitMs POST_KEY_WAIT
        End If
    End If

    WaitForPromptSlow = found
End Function





'--------------------------------------------------------------------
' Subroutine: FastText - Minimal delay text entry
'--------------------------------------------------------------------
Sub FastText(text)
    LOG "Sending text: " & text, "high"
    g_bzhao.SendKey text
    Call WaitMs(POST_KEY_WAIT)
End Sub

'--------------------------------------------------------------------
' Subroutine: FastKey - Minimal delay key press
'--------------------------------------------------------------------
Sub FastKey(key)
    LOG "Sending key command: " & key, "high"
    ' Pause briefly before sending a special key to avoid injecting escape sequences into active fields
    Call WaitMs(PRE_KEY_WAIT)
    g_bzhao.SendKey key
    ' Allow the host some time to process the special key and transition screens
    Call WaitMs(POST_KEY_WAIT)
End Sub

'--------------------------------------------------------------------
' Function: IsSlowModeEnabled
' Returns True if the debug marker file (Create_RO.debug) is present next to the script
'--------------------------------------------------------------------
Function IsSlowModeEnabled()
    On Error Resume Next
    Dim f
    Set f = CreateObject("Scripting.FileSystemObject")
    IsSlowModeEnabled = f.FileExists(SLOW_MARKER_PATH)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

'--------------------------------------------------------------------
' Function: GetRepairOrderEnhanced - Quick repair order extraction
'--------------------------------------------------------------------
Function GetRepairOrderEnhanced()
    Dim screenContent, screenLength, pos, startPos, ch, roNumber
    screenLength = 24 * 80
    
    g_bzhao.ReadScreen screenContent, screenLength, 1, 1
    pos = InStr(1, screenContent, "Created repair order ", vbTextCompare)
    
    If pos > 0 Then
        startPos = pos + 21 ' Length of "Created repair order "
        roNumber = ""
        
        Do While startPos <= Len(screenContent)
            ch = Mid(screenContent, startPos, 1)
            If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "Z") Then
                roNumber = roNumber & ch
                startPos = startPos + 1
            Else
                Exit Do
            End If
        Loop
        
        GetRepairOrderEnhanced = roNumber
    Else
        GetRepairOrderEnhanced = ""
    End If
End Function


'--------------------------------------------------------------------
' Function: ResolveLogVerbosity
' Reads [Logging] Verbosity from config.ini with default Med
'--------------------------------------------------------------------
Function ResolveLogVerbosity()
    Dim configPath, rawValue, normalized
    configPath = g_fso.BuildPath(GetRepoRoot(), "config\config.ini")
    rawValue = ReadIniValue(configPath, "Open_RO", "Verbosity")

    If Len(Trim(rawValue)) = 0 Then rawValue = "Med"
    normalized = LCase(Trim(rawValue))

    Select Case normalized
        Case "low"
            ResolveLogVerbosity = LOG_LEVEL_LOW
        Case "med"
            ResolveLogVerbosity = LOG_LEVEL_MED
        Case "high"
            ResolveLogVerbosity = LOG_LEVEL_HIGH
        Case Else
            Err.Raise 5, "ResolveLogVerbosity", "Invalid [Logging] Verbosity value: " & rawValue & ". Expected Low, Med, or High."
    End Select
End Function

'--------------------------------------------------------------------
' Function: NormalizeLogLevel
' Validates and normalizes a requested log level
'--------------------------------------------------------------------
Function NormalizeLogLevel(levelValue)
    Dim normalized
    normalized = LCase(Trim(CStr(levelValue)))
    If normalized = "" Then normalized = "high"

    Select Case normalized
        Case "low", "med", "high"
            NormalizeLogLevel = normalized
        Case Else
            Err.Raise 5, "NormalizeLogLevel", "Invalid log level requested: " & CStr(levelValue)
    End Select
End Function

'--------------------------------------------------------------------
' Function: ShouldLog
' Returns True when requested log level should be written
'--------------------------------------------------------------------
Function ShouldLog(levelValue)
    Dim requestedLevel, normalized
    normalized = NormalizeLogLevel(levelValue)

    Select Case normalized
        Case "low"
            requestedLevel = LOG_LEVEL_LOW
        Case "med"
            requestedLevel = LOG_LEVEL_MED
        Case "high"
            requestedLevel = LOG_LEVEL_HIGH
    End Select

    ShouldLog = (requestedLevel <= g_LogVerbosity)
End Function

'--------------------------------------------------------------------
' Subroutine: LogEntryWithRO - Write to both log and CSV output
'--------------------------------------------------------------------
Sub LogEntryWithRO(mva, roNumber)
    If Trim(mva) = "" Then Exit Sub

    Dim logFSO, logFile, csvOut
    Set logFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Write to transaction log
    Set logFile = logFSO.OpenTextFile(LOG_FILE_PATH, 8, True)
    If roNumber = "" Then
        logFile.WriteLine Now & " - MVA: " & mva
    Else
        logFile.WriteLine Now & " - MVA: " & mva & " - RO: " & roNumber
    End If
    logFile.Close
    Set logFile = Nothing
    
    ' Write to CSV output (only if RO number was successfully scraped)
    If roNumber <> "" Then
        Set csvOut = logFSO.OpenTextFile(OUTPUT_CSV_PATH, 8, True)
        csvOut.WriteLine roNumber
        csvOut.Close
        Set csvOut = Nothing
    End If
    
    Set logFSO = Nothing
End Sub


'--------------------------------------------------------------------
' Subroutine: LOG - lightweight logger used by this archived script
'--------------------------------------------------------------------
Sub LOG(msg, level)
    Dim lfile, errorNum, errorDesc

    If Not ShouldLog(level) Then Exit Sub

    On Error Resume Next

    ' Try to create log entry
    Set lfile = g_fso.OpenTextFile(LOG_FILE_PATH, 8, True)
    errorNum = Err.Number
    errorDesc = Err.Description

    If errorNum = 0 Then
        lfile.WriteLine Now & " - " & CStr(msg)
        lfile.Close
    Else
        ' If main log fails, try creating a fallback log with error info
        Dim fallbackPath
        fallbackPath = GetConfigPath("Open_RO", "FallbackLog")
        Set lfile = g_fso.OpenTextFile(fallbackPath, 8, True)
        If Err.Number = 0 Then
            lfile.WriteLine Now & " - LOG ERROR: " & errorNum & " - " & errorDesc
            lfile.WriteLine Now & " - Failed LOG_FILE_PATH: " & LOG_FILE_PATH
            lfile.WriteLine Now & " - Original message: " & CStr(msg)
            lfile.Close
        End If
    End If

    Set lfile = Nothing
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

