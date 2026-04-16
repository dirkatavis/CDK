Option Explicit


' --- Bootstrap ---
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_sh: Set g_sh = CreateObject("WScript.Shell")
Dim g_root: g_root = g_sh.Environment("USER")("CDK_BASE")
If g_root = "" Or Not g_fso.FolderExists(g_root) Then
    Err.Raise 53, "Bootstrap", "Invalid or missing CDK_BASE. Value: " & g_root
End If
If Not g_fso.FileExists(g_fso.BuildPath(g_root, ".cdkroot")) Then
    Err.Raise 53, "Bootstrap", "Cannot find .cdkroot in base path:" & vbCrLf & g_root
End If
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\PathHelper.vbs")).ReadAll
Dim g_bzhao  ' Declared here for Option Explicit; assigned in InitializeObjects (supports MockBzhao test mode)
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\BZHelper.vbs")).ReadAll

' --- Load ValidateSetup for dependency checking ---
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\ValidateSetup.vbs")).ReadAll

' Global script variables
Dim CSV_FILE_PATH, LOG_FILE_PATH
Dim roNumber
Dim lastRoResult
Dim currentRODisplay
Dim commonLibLoaded
Dim POST_PROMPT_WAIT_MS
Dim DEBUG_LOGGING
Dim g_IsTestMode
Dim g_DefaultWait, g_LongWait, g_SendRetryCount, g_TimeoutMs, g_PromptWait, g_DelayBetweenTextAndEnterMs
Dim g_EnableDiagnosticLogging, DIAGNOSTIC_LOG_PATH
Dim g_DiagLogQueue
Dim g_EnableDetailedLogging
Dim g_LastScrapedStatus
Dim g_BaseScriptPath
Dim g_ShouldAbort, g_AbortReason
Dim g_EmployeeNumber, g_EmployeeNameConfirm
Dim g_StartSequenceNumber, g_EndSequenceNumber
Dim g_DebugDelayFactor
Dim MainPromptLine
Dim LEGACY_BASE_PATH, LEGACY_CSV_PATH, LEGACY_LOG_PATH, LEGACY_DIAG_LOG_PATH, LEGACY_COMMONLIB_PATH
Dim g_CurrentCriticality, g_CurrentVerbosity
Dim g_SessionDateLogged
Dim g_LastSuccessfulLine
Dim g_NoPromptCount
Dim g_ProcessPromptSequenceTimeoutMsOverride
Dim g_ProcessPromptSequenceMaxNoPromptIterationsOverride
Dim g_ProcessPromptSequenceNoPromptRetryWaitMsOverride
Dim g_CloseoutConfirmDelayMs
Dim g_StabilityPause
Dim g_ReviewedROCount
Dim g_FiledROCount
Dim g_BlacklistTermsRaw
Dim g_CloseoutTriggers
Dim g_SkipBlacklistCount
Dim g_SkipStatusOpenCount
Dim g_SkipStatusPreassignedCount
Dim g_SkipStatusOtherCount
Dim g_SkipOtherStates
Dim g_SkipRoListRaw
Dim g_SkipRoLookup
Dim g_SkipConfiguredCount
Dim g_SkipWarrantyCount
Dim g_SkipWchEnabled
Dim g_SkipPartsOrderNeededCount
Dim g_PartsOrderKeywords
Dim g_PartsOrderNegators
Dim g_AllowedTechCodes
Dim g_arrCDKExceptions
Dim g_arrCDKDescriptionExceptions
Dim g_SkipTechCodeCount
Dim g_ClosedRoCount
Dim g_NotOnFileRoCount
Dim g_SkipVehidNotOnFileCount
Dim g_SkipNoCloseoutTextCount
Dim g_SkipNoPartsChargedCount
Dim g_LeftOpenManualCount
Dim g_FcaMissingPartFlagCount
Dim g_FcaHandlerNotConfiguredCount
Dim g_ErrorInMainCount
Dim g_NoResultRecordedCount
Dim g_SummaryOtherOutcomeCount
Dim g_SummaryOtherOutcomeBreakdown
Dim g_SummaryOtherOutcomeRawBreakdown
Dim g_OverwriteLogOnStart
Dim g_PreviousNormalizedRo
Dim g_PreviousSequenceNumber
Dim LEGACY_TRIGGER_LIST_PATH
Dim g_OlderRoThresholdDays
Dim g_OlderRoStatuses
Dim g_OlderRoFiledCount
Dim g_OlderRoAttemptCount

MainPromptLine = 23

' Criticality constants (higher numbers = higher priority)
Const CRIT_COMMON = 0
Const CRIT_MINOR = 1
Const CRIT_MAJOR = 2
Const CRIT_CRITICAL = 3

' Verbosity constants (higher numbers = more detail)
Const VERB_LOW = 0
Const VERB_MEDIUM = 1
Const VERB_HIGH = 2
Const VERB_MAX = 3
Const DEBUG_SCREEN_LINES = 3
Const g_DiagLogQueueSize = 5
' LEGACY_BASE_PATH is now dynamic (derived from .cdkroot)
' Simplified timeout logic - no midnight handling needed for current debugging

' --- EARLY LOGGING: Force maximum logging for startup ---
g_CurrentCriticality = CRIT_COMMON ' Log all criticality levels
g_CurrentVerbosity = VERB_MAX ' Show maximum detail during startup
g_SessionDateLogged = False


LEGACY_BASE_PATH = g_fso.BuildPath(GetRepoRoot(), "apps\post_final_charges")
LEGACY_CSV_PATH = GetConfigPath("PostFinalCharges", "CSV")
LEGACY_LOG_PATH = GetConfigPath("PostFinalCharges", "Log")
LEGACY_DIAG_LOG_PATH = GetConfigPath("PostFinalCharges", "DiagnosticLog")
LEGACY_COMMONLIB_PATH = GetConfigPath("PostFinalCharges", "CommonLib")
LEGACY_TRIGGER_LIST_PATH = GetConfigPath("PostFinalCharges", "TriggerList")



' Bootstrap defaults so logging works before config initialization
g_BaseScriptPath = LEGACY_BASE_PATH
CSV_FILE_PATH = LEGACY_CSV_PATH
LOG_FILE_PATH = LEGACY_LOG_PATH
commonLibLoaded = False
g_ShouldAbort = False

' Initialize session header and perform log trimming once at startup
Dim sessionHeaderFailed, headerErrorNumber, headerErrorDescription
sessionHeaderFailed = False
headerErrorNumber = 0
headerErrorDescription = ""

On Error Resume Next
Call WriteSessionHeader()
If Err.Number <> 0 Then
    sessionHeaderFailed = True
    headerErrorNumber = Err.Number
    headerErrorDescription = Err.Description
    Err.Clear
End If
On Error GoTo 0

If sessionHeaderFailed Then
    Call LogEvent("comm", "low", "Session header write failed (" & headerErrorNumber & "): " & headerErrorDescription, "Startup", "", "")
End If

' --- STARTUP LOGGING: Script Startup and Path Resolution ---
Call LogEvent("comm", "low", "Script entrypoint reached", "Startup", "", "")
Call LogEvent("comm", "low", "About to resolve g_BaseScriptPath", "Startup", "", "")
Call LogEvent("comm", "low", "g_BaseScriptPath set to: " & g_BaseScriptPath, "Startup", "", "")
'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** ProcessPromptSequence
' **DATE CREATED:** 2025-11-25
' **AUTHOR:** GitHub Copilot
' **MODIFIED:** 2025-11-25 by Gemini
' 
' **FUNCTIONALITY:**
' Processes a sequence of prompts using a state machine approach. Scans the screen
' for known prompts from a given dictionary and responds accordingly.
'-----------------------------------------------------------------------------------

' Defines the data structure for a single screen prompt and its corresponding action.
Class Prompt
    Public TriggerText    ' Pattern to match on screen (can be literal text or regex)
    Public ResponseText   ' Text to send when prompt is detected (empty if accepting default)
    Public KeyPress       ' Key to press after response text (e.g. "<NumpadEnter>")
    Public IsSuccess      ' True if this prompt indicates successful completion
    Public AcceptDefault  ' True to accept default values shown in parentheses
    Public IsRegex        ' True when TriggerText should be evaluated as regex
End Class

Function InferRegexPattern(pattern)
    InferRegexPattern = False
    If Left(pattern, 1) = "^" Or InStr(pattern, "(") > 0 Or InStr(pattern, "[") > 0 Or InStr(pattern, ".*") > 0 Or InStr(pattern, "\d") > 0 Then
        InferRegexPattern = True
    End If
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** AddPromptToDict
' **FUNCTIONALITY:** 
' Creates a Prompt object with AcceptDefault=False and adds it to the dictionary.
' Use this for prompts that should always send the specified ResponseText,
' regardless of any default values shown on screen.
' 
' **PARAMETERS:**
' dict - Dictionary to add the prompt to
' trigger - Text pattern to match (can be literal or regex)
' response - Text to send when prompt is matched
' key - Keystroke to send after response text
' isSuccess - True if this prompt indicates completion of the sequence
'-----------------------------------------------------------------------------------
Sub AddPromptToDict(dict, trigger, response, key, isSuccess)
    Dim p
    Set p = New Prompt
    p.TriggerText = trigger
    p.ResponseText = response
    p.KeyPress = key
    p.IsSuccess = isSuccess
    p.AcceptDefault = False  ' Always send response text, ignore screen defaults
    p.IsRegex = InferRegexPattern(trigger)
    dict.Add trigger, p
End Sub

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** AddPromptToDictEx  
' **FUNCTIONALITY:**
' Enhanced version with AcceptDefault support. Creates a Prompt object that can
' optionally detect and accept default values shown in parentheses on screen.
' When AcceptDefault=True and a default value is detected (e.g. "TECHNICIAN (72925)?"),
' the system will send only the keystroke without response text to preserve the default.
'
' **PARAMETERS:**
' dict - Dictionary to add the prompt to
' trigger - Text pattern to match (can be literal or regex)
' response - Text to send when NO default is present or AcceptDefault=False
' key - Keystroke to send (with or without response text)
' isSuccess - True if this prompt indicates completion of the sequence
' acceptDefault - True to detect/accept defaults in parentheses, False to always send response
'
' **EXAMPLES:**
' ' Always send "99" regardless of screen defaults:
' Call AddPromptToDict(dict, "TECHNICIAN", "99", "<NumpadEnter>", False)
'
' ' Accept screen defaults when present, send "99" when no default shown:
' Call AddPromptToDictEx(dict, "TECHNICIAN \([A-Za-z0-9]+\)\?", "99", "<NumpadEnter>", False, True)
'-----------------------------------------------------------------------------------
Sub AddPromptToDictEx(dict, trigger, response, key, isSuccess, acceptDefault)
    Dim p
    Set p = New Prompt
    p.TriggerText = trigger
    p.ResponseText = response
    p.KeyPress = key
    p.IsSuccess = isSuccess
    p.AcceptDefault = acceptDefault
    p.IsRegex = InferRegexPattern(trigger)
    dict.Add trigger, p
End Sub

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** CreateLineItemPromptDictionary
' **FUNCTIONALITY:**
' Creates and returns the prompt dictionary for the line item processing sequence.
' Demonstrates both AddPromptToDict (always sends response) and AddPromptToDictEx 
' (can accept defaults) usage patterns.
'-----------------------------------------------------------------------------------
Function CreateLineItemPromptDictionary()
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")
    ' Handle end-of-sequence error
    Call AddPromptToDict(dict, "SEQUENCE NUMBER \d+ DOES NOT EXIST", "", "", True)
    
    ' OPERATION CODE FOR LINE: Comprehensive pattern handles all formats
    ' When screen shows "OPERATION CODE FOR LINE A, L1 (I)?", accepts default "I"
    ' When screen shows "OPERATION CODE FOR LINE A, L1 ()?", sends "I" (empty parens = no default)
    ' When screen shows "OPERATION CODE FOR LINE A, L1?", sends "I" (no parentheses)
    Call AddPromptToDictEx(dict, "OPERATION CODE FOR LINE.*(\([A-Za-z0-9]*\))?\?", "I", "<NumpadEnter>", False, True)
    ' Narrow watcher for the rare no-parenthesis variant - avoid matching lines that include '('
    Call AddPromptToDict(dict, "OPERATION CODE FOR LINE[^\(]*\?", "I", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "COMMAND:\(SEQ#/E/N/B/\?\)", "", "", True)
    ' COMMAND: prompt removed - handled by legacy WaitForPrompt in ProcessLineItems
    Call AddPromptToDict(dict, "This OpCode was performed in the last 270 days.", "", "", False)
    
    ' LINE CODE: Uses AddPromptToDictEx with AcceptDefault=False (always press Enter, never send text)
    Call AddPromptToDictEx(dict, "LINE CODE [A-Z] IS NOT ON FILE", "", "<Enter>", True, False)
    Call AddPromptToDict(dict, "LABOR TYPE FOR LINE", "", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "DESC:", "", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "Enter a technician number", "", "<F3>", False)
    
    ' TECHNICIAN prompts: Uses AddPromptToDictEx with AcceptDefault=True
    ' Accepts defaults like "TECHNICIAN (72925)?" or sends "99" for "TECHNICIAN?"
    Call AddPromptToDictEx(dict, "TECHNICIAN \(\d+\)", "99", "<NumpadEnter>", False, True)
    Call AddPromptToDictEx(dict, "TECHNICIAN?", "99", "<NumpadEnter>", False, True)
    Call AddPromptToDictEx(dict, "TECHNICIAN \([A-Za-z0-9]+\)\?", "99", "<NumpadEnter>", False, True)
    
    ' Specific prompt for "TECHNICIAN FINISHING WORK" - sends technician number
    Call AddPromptToDict(dict, "TECHNICIAN FINISHING WORK", "99", "<NumpadEnter>", False)
    
    ' Handle technician assignment + line completion prompts
    Call AddPromptToDict(dict, "IS ASSIGNED TO LINE", "Y", "<NumpadEnter>", False)
    
    ' Handle line completion confirmation prompt (always send "Y", no defaults to check)
    ' Make this pattern more specific to avoid matching other TECHNICIAN prompts
    Call AddPromptToDictEx(dict, "TECHNICIAN \(Y/N\)", "Y", "<NumpadEnter>", True, False)
    
    ' HOURS prompts: Uses AddPromptToDictEx with AcceptDefault=True
    ' Accepts defaults like "ACTUAL HOURS (117)?" or sends "0" for "ACTUAL HOURS?"
    Call AddPromptToDictEx(dict, "ACTUAL HOURS \(\d+\)", "0", "<NumpadEnter>", False, True)
    ' SOLD HOURS: Advanced pattern handles both with and without parentheses
    ' "SOLD HOURS (10)?" -> accepts default 10, "SOLD HOURS?" -> sends "0"
    Call AddPromptToDictEx(dict, "SOLD HOURS( \(\d+\))?\?", "0", "<NumpadEnter>", False, True)
    Call AddPromptToDict(dict, "ADD A LABOR OPERATION( \(N\)\?)?", "N", "<NumpadEnter>", True)
    ' Fallback watcher: some terminals render spacing/default text variably on this prompt.
    ' Enter accepts the default (N) and safely returns to COMMAND.
    Call AddPromptToDict(dict, "ADD A LABOR OPERATION", "", "<Enter>", True)
    ' Note: COMMAND: success condition removed - handled by checking MainPromptLine specifically
    ' to avoid false positives when COMMAND appears elsewhere on screen
    Call AddPromptToDict(dict, "Is this a comeback \(Y/N\)\.\.\.", "Y", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "NOT ON FILE", "", "<Enter>", True)
    Call AddPromptToDict(dict, "NOT AVAILABLE", "", "<Enter>", True)
    Call AddPromptToDict(dict, "NOT ALL LINES HAVE A COMPLETE STATUS", "", "<Enter>", True)
    Call AddPromptToDict(dict, "PRESS RETURN TO CONTINUE", "", "<Enter>", False)
    Call AddPromptToDict(dict, "Press F3 to exit.", "", "<F3>", False)

    Set CreateLineItemPromptDictionary = dict
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** CreateFnlPromptDictionary
' **DATE CREATED:** 2026-01-01
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Creates a dictionary of specific prompts that might appear after FNL commands.
' This replaces the generic colon detection with explicit prompt handling.
'-----------------------------------------------------------------------------------
Function CreateFnlPromptDictionary()
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Technician prompts that may appear after FNL commands
    Call AddPromptToDict(dict, "TECHNICIAN FINISHING WORK ?", "99", "<NumpadEnter>", False)
    Call AddPromptToDictEx(dict, "TECHNICIAN \([A-Za-z0-9]+\)\?", "99", "<NumpadEnter>", False, True)
    Call AddPromptToDictEx(dict, "TECHNICIAN\?", "99", "<NumpadEnter>", False, True)
    
    ' Handle technician assignment + line completion as a combined prompt
    ' First line: "TECHNICIAN 12051 IS ASSIGNED TO LINE 'A'"  
    ' Second line: "Is Line 'A' completed (Y/N):"
    ' We detect the first line but respond to the second line's question
    Call AddPromptToDict(dict, "IS ASSIGNED TO LINE", "Y", "<NumpadEnter>", False)
    
    ' Also handle direct line completion prompts (in case they appear standalone)
    Call AddPromptToDictEx(dict, "Is Line '[A-Z]' completed \(Y/N\)", "Y", "<NumpadEnter>", False, False)
    
    ' Hours prompts that may appear after FNL commands
    Call AddPromptToDictEx(dict, "ACTUAL HOURS \(\d+\)", "0", "<NumpadEnter>", False, True)
    Call AddPromptToDictEx(dict, "SOLD HOURS( \(\d+\))?\?", "0", "<NumpadEnter>", False, True)
    
    ' Success condition - back to command prompt
    Call AddPromptToDict(dict, "COMMAND:", "", "", True)
    
    Set CreateFnlPromptDictionary = dict
End Function

' Creates and returns the prompt dictionary for the final closeout sequence.
Function CreateCloseoutPromptDictionary()
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")

    Call AddPromptToDict(dict, "COMMAND:", "", "", True) ' Success
    Call AddPromptToDict(dict, "ALL LABOR POSTED", "Y", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "MILEAGE OUT", "", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "Current Mileage less than Previous Mileage", "Y", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "MILEAGE IN", "", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "O.K. TO CLOSE RO", "Y", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "INVOICE PRINTER", "2", "<NumpadEnter>", True) ' Success
    Call AddPromptToDict(dict, "Press F3 to exit.", "", "<F3>", False)

    Set CreateCloseoutPromptDictionary = dict
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** WaitForScreenStable
' **DATE CREATED:** 2026-01-01
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Waits for the screen content to stabilize by checking that the main prompt line
' doesn't change for a specified period. This ensures we don't read the screen
' during a transition state.
' 
' **PARAMETERS:**
' timeoutMs (Integer): Maximum time to wait for stability (default: 2000ms)
' stabilityMs (Integer): How long the screen must remain unchanged (default: 200ms)
' 
' **RETURN VALUE:**
' (Boolean) Returns True if screen stabilized, False if timeout occurred
'-----------------------------------------------------------------------------------
Function WaitForScreenStable(timeoutMs, stabilityMs)
    If timeoutMs <= 0 Then timeoutMs = 2000
    If stabilityMs <= 0 Then stabilityMs = 200
    
    Call LogEvent("comm", "max", "WaitForScreenStable starting", "WaitForScreenStable", "Timeout: " & timeoutMs & "ms, Stability: " & stabilityMs & "ms", "")
    
    Dim waitStart, waitElapsed, lastContent, stableStart, stableElapsed
    Dim currentContent
    
    waitStart = Timer
    lastContent = ""
    stableStart = 0
    
    Do
        currentContent = GetScreenLine(MainPromptLine)
        
        ' Check if content has changed
        If currentContent <> lastContent Then
            ' Content changed, restart stability timer
            Call LogEvent("comm", "max", "Screen content changed", "WaitForScreenStable", "New: '" & Left(currentContent, 50) & "'", "Old: '" & Left(lastContent, 50) & "'")
            lastContent = currentContent
            stableStart = Timer
        Else
            ' Content same, check if we've been stable long enough
            If stableStart > 0 Then
                stableElapsed = (Timer - stableStart) * 1000
                If stableElapsed >= stabilityMs Then
                    ' Screen has been stable for required period
                    Call LogEvent("comm", "high", "Screen stable for required period", "WaitForScreenStable", Int(stableElapsed) & "ms >= " & stabilityMs & "ms", "")
                    WaitForScreenStable = True
                    Exit Function
                End If
            Else
                ' First iteration with this content, start stability timer
                stableStart = Timer
            End If
        End If
        
        ' Check overall timeout
        waitElapsed = (Timer - waitStart) * 1000
        If waitElapsed > timeoutMs Then
            ' Overall timeout exceeded
            Call LogEvent("min", "med", "WaitForScreenStable timeout exceeded", "WaitForScreenStable", Int(waitElapsed) & "ms > " & timeoutMs & "ms", "Last content: '" & Left(currentContent, 50) & "'")
            WaitForScreenStable = False
            Exit Function
        End If
        
        Call WaitMs(50) ' Small delay between checks
    Loop
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** TimerToClockTime
' **FUNCTIONALITY:** Converts Timer value (seconds since midnight) to HH:MM:SS format
'-----------------------------------------------------------------------------------
Function TimerToClockTime(timerValue)
    Dim hours, minutes, seconds, totalSeconds
    totalSeconds = Int(timerValue)
    hours = totalSeconds \ 3600
    minutes = (totalSeconds Mod 3600) \ 60
    seconds = totalSeconds Mod 60
    TimerToClockTime = Right("0" & hours, 2) & ":" & Right("0" & minutes, 2) & ":" & Right("0" & seconds, 2) & "." & Right("0" & Int((timerValue - totalSeconds) * 100), 2)
End Function

' Generic state machine to process a sequence of prompts from a given dictionary.
Sub ResetProcessPromptSequenceOverrides()
    g_ProcessPromptSequenceTimeoutMsOverride = Empty
    g_ProcessPromptSequenceMaxNoPromptIterationsOverride = Empty
    g_ProcessPromptSequenceNoPromptRetryWaitMsOverride = Empty
End Sub

Sub ProcessPromptSequence(prompts)
    Dim finished, promptKey, promptDetails, bestMatchKey, bestMatchLength
    Dim sequenceStartTime, sequenceElapsed
    Dim resolvedSequenceTimeoutMs, resolvedMaxNoPromptIterations, resolvedNoPromptRetryWaitMs
    finished = False
    sequenceStartTime = Now() ' Use actual date/time instead of Timer

    resolvedSequenceTimeoutMs = 30000
    resolvedMaxNoPromptIterations = 20
    resolvedNoPromptRetryWaitMs = 1000

    On Error Resume Next
    If Not IsEmpty(g_ProcessPromptSequenceTimeoutMsOverride) Then
        resolvedSequenceTimeoutMs = CLng(g_ProcessPromptSequenceTimeoutMsOverride)
        If Err.Number <> 0 Or resolvedSequenceTimeoutMs <= 0 Then
            resolvedSequenceTimeoutMs = 30000
            Err.Clear
        End If
    End If
    If Not IsEmpty(g_ProcessPromptSequenceMaxNoPromptIterationsOverride) Then
        resolvedMaxNoPromptIterations = CInt(g_ProcessPromptSequenceMaxNoPromptIterationsOverride)
        If Err.Number <> 0 Or resolvedMaxNoPromptIterations <= 0 Then
            resolvedMaxNoPromptIterations = 20
            Err.Clear
        End If
    End If
    If Not IsEmpty(g_ProcessPromptSequenceNoPromptRetryWaitMsOverride) Then
        resolvedNoPromptRetryWaitMs = CLng(g_ProcessPromptSequenceNoPromptRetryWaitMsOverride)
        If Err.Number <> 0 Or resolvedNoPromptRetryWaitMs <= 0 Then
            resolvedNoPromptRetryWaitMs = 1000
            Err.Clear
        End If
    End If
    On Error GoTo 0
    
    ' DIAGNOSTIC: Log Timer behavior at start
    Call LogEvent("comm", "high", "ProcessPromptSequence started", "ProcessPromptSequence", "Timer diagnostics", "sequenceStartTime=" & sequenceStartTime & " (Now() at start)")
    Call LogEvent("comm", "med", "ProcessPromptSequence retry policy", "ProcessPromptSequence", "timeoutMs=" & resolvedSequenceTimeoutMs & " maxNoPromptIterations=" & resolvedMaxNoPromptIterations, "noPromptRetryWaitMs=" & resolvedNoPromptRetryWaitMs)
    
    ' Initialize no-prompt counter for this sequence
    g_NoPromptCount = 0
    
    ' Wait for screen to stabilize before starting scan
    If Not WaitForScreenStable(3000, 200) Then
        Call LogEvent("maj", "med", "Screen did not stabilize before starting prompt sequence", "ProcessPromptSequence", "May affect automation reliability", "WaitForScreenStable(3000,200) returned False")
    End If

    Do While Not finished
        ' DIAGNOSTIC: Log Timer values with clock time conversion to expose stale start times
        Dim currentTimer, currentTimerClock, sequenceStartTimer, sequenceStartClock
        currentTimer = Timer
        currentTimerClock = TimerToClockTime(currentTimer)
        
        ' Extract Timer value from sequenceStartTime (might be Now() date or Timer value)
        On Error Resume Next
        If IsDate(sequenceStartTime) Then
            ' If sequenceStartTime is a Date from Now(), calculate equivalent Timer value
            Dim startHour, startMin, startSec
            startHour = Hour(sequenceStartTime)
            startMin = Minute(sequenceStartTime) 
            startSec = Second(sequenceStartTime)
            sequenceStartTimer = startHour * 3600 + startMin * 60 + startSec
        Else
            ' If sequenceStartTime is already a Timer value
            sequenceStartTimer = CDbl(sequenceStartTime)
        End If
        On Error GoTo 0
        
        sequenceStartClock = TimerToClockTime(sequenceStartTimer)
        
        Call LogEvent("comm", "high", "TIMER DIAGNOSTIC - Current vs Start", "ProcessPromptSequence", "currentTimer=" & currentTimer & " (" & currentTimerClock & ") sequenceStartTime=" & sequenceStartTime & " (" & sequenceStartClock & ")", "")
        
        ' Check for timeout - using real wall clock time
        sequenceElapsed = DateDiff("s", sequenceStartTime, Now()) * 1000
        
        ' DIAGNOSTIC: Log the calculation step by step
        Call LogEvent("comm", "high", "Timeout calculation details", "ProcessPromptSequence", "Now()=" & Now() & " minus " & sequenceStartTime & " equals " & DateDiff("s", sequenceStartTime, Now()) & " seconds = " & sequenceElapsed & "ms", "")
        
        ' DIAGNOSTIC: Show Timer calculation for comparison to expose discrepancy
        Dim timerElapsed
        timerElapsed = (currentTimer - sequenceStartTimer) 
        Call LogEvent("comm", "high", "TIMER vs CLOCK COMPARISON", "ProcessPromptSequence", "Timer elapsed: " & timerElapsed & " seconds", "Clock: " & currentTimerClock & " - " & sequenceStartClock)
        
        If sequenceElapsed > resolvedSequenceTimeoutMs Then
            Call LogEvent("crit", "low", "ProcessPromptSequence timed out", "ProcessPromptSequence", "Automation stopped", "Now()=" & Now() & " sequenceStartTime=" & sequenceStartTime & " calculated=" & sequenceElapsed & "ms > " & resolvedSequenceTimeoutMs & "ms")
            SafeMsg "ProcessPromptSequence timed out after " & Int(resolvedSequenceTimeoutMs / 1000) & " seconds." & vbCrLf & "Automation stopped.", True, "Sequence Timeout"
            g_ShouldAbort = True
            Exit Sub
        End If
        
        ' Wait for screen to be stable before scanning
        Call WaitForScreenStable(1000, 100)
        
        Dim mainPromptText
        mainPromptText = GetScreenLine(MainPromptLine)
        
        ' DIAGNOSTIC: Log detailed prompt matching information at Max verbosity
        Call LogEvent("comm", "max", "Prompt detection analysis", "ProcessPromptSequence", "MainPromptLine=" & MainPromptLine & " Text='" & mainPromptText & "'", "Length=" & Len(mainPromptText))
        
        If Len(mainPromptText) > 0 Then
            ' Log all dictionary keys being checked against
            Dim diagKeys, diagKeyList, diagKey
            diagKeys = prompts.Keys
            diagKeyList = ""
            For Each diagKey In diagKeys
                If diagKeyList <> "" Then diagKeyList = diagKeyList & ", "
                diagKeyList = diagKeyList & "'" & diagKey & "'"
            Next
            Call LogEvent("comm", "max", "Checking against prompt dictionary", "ProcessPromptSequence", "Dictionary contains " & prompts.Count & " entries", "Keys: " & diagKeyList)
            
            If Not IsPromptInConfig(mainPromptText, prompts) Then
                Call LogEvent("comm", "max", "Prompt matching failed", "ProcessPromptSequence", "No dictionary entry matched '" & mainPromptText & "'", "Checked " & prompts.Count & " patterns")
                Call LogEvent("crit", "low", "Unknown prompt on line " & MainPromptLine & ": '" & mainPromptText & "'", "ProcessPromptSequence", "Aborting script - requires manual review", "prompt not found in prompts dictionary")
                SafeMsg "Unknown prompt detected on line " & MainPromptLine & ": '" & mainPromptText & vbCrLf & "Automation stopped for manual review.", True, "Unknown Prompt Error"
                g_ShouldAbort = True
                Exit Sub
            End If
        End If

        ' --- Find the longest (most specific) matching prompt ---
        bestMatchKey = ""
        bestMatchLength = 0
        Dim bestMatchDistance
        bestMatchDistance = 999
        
        ' Priority-matched scanning: check active prompt line first to avoid stale-text collisions
        Dim lineToCheck, lineText, linesToCheck, primaryLines
        primaryLines = Array(MainPromptLine, MainPromptLine - 1, MainPromptLine + 1)
        linesToCheck = Array(1, 2, 3, 4, 5, 20, 21, 22, 23, 24) ' Broad fallback scan
        
        For Each lineToCheck In primaryLines
            If lineToCheck >= 1 And lineToCheck <= 24 Then
                lineText = GetScreenLine(lineToCheck)
                If Len(lineText) > 0 Then
                    Dim currentDistance
                    currentDistance = Abs(lineToCheck - MainPromptLine)
                    For Each promptKey In prompts.Keys
                        Dim isRegex, re, regexError
                        Set promptDetails = prompts.Item(promptKey)
                        isRegex = promptDetails.IsRegex
                        regexError = False
                        If isRegex Then
                            On Error Resume Next
                            Set re = CreateObject("VBScript.RegExp")
                            re.Pattern = promptKey
                            re.IgnoreCase = True
                            re.Global = False
                            If Err.Number <> 0 Then
                                regexError = True
                                Err.Clear
                            End If
                            If Not regexError Then
                                If re.Test(lineText) Then
                                    If currentDistance < bestMatchDistance Or (currentDistance = bestMatchDistance And Len(promptKey) > bestMatchLength) Then
                                        bestMatchKey = promptKey
                                        bestMatchLength = Len(promptKey)
                                        bestMatchDistance = currentDistance
                                    End If
                                End If
                            End If
                            On Error GoTo 0
                        End If
                        If Not isRegex Then
                            If InStr(1, lineText, promptKey, vbTextCompare) > 0 Then
                                If currentDistance < bestMatchDistance Or (currentDistance = bestMatchDistance And Len(promptKey) > bestMatchLength) Then
                                    bestMatchKey = promptKey
                                    bestMatchLength = Len(promptKey)
                                    bestMatchDistance = currentDistance
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        Next

        ' Fallback to broad scan only when no active-line prompt match was found
        If bestMatchLength = 0 Then
            For Each lineToCheck In linesToCheck
                lineText = GetScreenLine(lineToCheck)
                If Len(lineText) > 0 Then
                    For Each promptKey In prompts.Keys
                        Set promptDetails = prompts.Item(promptKey)
                        isRegex = promptDetails.IsRegex
                        regexError = False
                        If isRegex Then
                            On Error Resume Next
                            Set re = CreateObject("VBScript.RegExp")
                            re.Pattern = promptKey
                            re.IgnoreCase = True
                            re.Global = False
                            If Err.Number <> 0 Then
                                regexError = True
                                Err.Clear
                            End If
                            If Not regexError Then
                                If re.Test(lineText) Then
                                    If Len(promptKey) > bestMatchLength Then
                                        bestMatchKey = promptKey
                                        bestMatchLength = Len(promptKey)
                                    End If
                                End If
                            End If
                            On Error GoTo 0
                        End If
                        If Not isRegex Then
                            If InStr(1, lineText, promptKey, vbTextCompare) > 0 Then
                                If Len(promptKey) > bestMatchLength Then
                                    bestMatchKey = promptKey
                                    bestMatchLength = Len(promptKey)
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        End If

        ' --- If a prompt was found, handle it ---
        If bestMatchLength > 0 Then
            ' Reset the no-prompt counter since we found a prompt
            g_NoPromptCount = 0
            
            ' CRITICAL FIX: Reset timer for each individual prompt
            ' Each prompt gets its own 30-second timeout window
            sequenceStartTime = Now()
            Call LogEvent("comm", "high", "TIMER RESET for new prompt", "ProcessPromptSequence", "Individual prompt timeout starts now", "sequenceStartTime=" & sequenceStartTime)
            
            Set promptDetails = prompts.Item(bestMatchKey)
            Call LogEvent("comm", "high", "Matched prompt: '" & bestMatchKey & "'", "ProcessPromptSequence", "Found most specific match", "match length=" & bestMatchLength)
            Call LogEvent("comm", "max", "Prompt details: '" & bestMatchKey & "'", "ProcessPromptSequence", "", "ResponseText='" & promptDetails.ResponseText & "' KeyPress='" & promptDetails.KeyPress & "' AcceptDefault=" & promptDetails.AcceptDefault & " IsSuccess=" & promptDetails.IsSuccess)

            ' Log the screen before responding
            Call LogEvent("comm", "high", "Screen before responding to '" & bestMatchKey & "'", "ProcessPromptSequence", "", GetScreenSnapshot(3))

            ' Check if this prompt should accept default values and if one is present
            Dim shouldAcceptDefault
            shouldAcceptDefault = False
            If promptDetails.AcceptDefault Then
                Call LogEvent("comm", "high", "Checking for default values on screen", "ProcessPromptSequence", "Prompt has AcceptDefault=True", "")
                
                ' Find the line that contains the matched prompt for default value checking
                Dim matchedLineContent
                matchedLineContent = ""
                For Each lineToCheck In linesToCheck
                    lineText = GetScreenLine(lineToCheck)
                    
                    ' Check if the bestMatchKey pattern matches this line's content
                    ' Use regex matching for regex patterns, InStr for plain text patterns
                    Dim lineMatches, lineMatchFound
                    lineMatchFound = False
                    
                    If promptDetails.IsRegex Then
                        ' Regex pattern - use proper regex matching
                        On Error Resume Next
                        Dim lineRe
                        Set lineRe = CreateObject("VBScript.RegExp")
                        lineRe.Pattern = bestMatchKey
                        lineRe.IgnoreCase = True
                        lineRe.Global = False
                        Set lineMatches = lineRe.Execute(lineText)
                        If Err.Number = 0 And lineMatches.Count > 0 Then
                            lineMatchFound = True
                        End If
                        On Error GoTo 0
                    Else
                        ' Plain text pattern - use InStr
                        If InStr(1, lineText, bestMatchKey, vbTextCompare) > 0 Then
                            lineMatchFound = True
                        End If
                    End If
                    
                    If lineMatchFound Then
                        matchedLineContent = lineText
                        Call LogEvent("comm", "high", "Found matched line " & lineToCheck, "ProcessPromptSequence", "Contains prompt pattern", "'" & matchedLineContent & "'")
                        Exit For
                    End If
                Next
                
                If matchedLineContent = "" Then
                    Call LogEvent("min", "med", "No line found containing prompt pattern", "ProcessPromptSequence", "Pattern: '" & bestMatchKey & "'", "")
                End If
                
                shouldAcceptDefault = HasDefaultValueInPrompt(bestMatchKey, matchedLineContent)
                If shouldAcceptDefault Then
                    Call LogEvent("comm", "high", "Default value detected in prompt", "ProcessPromptSequence", "Accepting by sending only key press", "")
                Else
                    Call LogEvent("comm", "high", "No valid default value detected", "ProcessPromptSequence", "Will send ResponseText", "")
                End If
            Else
                Call LogEvent("comm", "high", "Prompt has AcceptDefault=False", "ProcessPromptSequence", "Will always send ResponseText if provided", "")
            End If

            If promptDetails.ResponseText <> "" And Not shouldAcceptDefault Then
                Call FastText(promptDetails.ResponseText)
                Call LogEvent("comm", "high", "Sent ResponseText", "ProcessPromptSequence", "'" & promptDetails.ResponseText & "'", "")
            Else
                Call LogEvent("comm", "high", "No ResponseText to send", "ProcessPromptSequence", "Empty or accepting default", "")
            End If
            
            Call FastKey(promptDetails.KeyPress)

            If InStr(1, bestMatchKey, "O.K. TO CLOSE RO", vbTextCompare) > 0 Then
                Call LogEvent("comm", "high", "Applied closeout confirm delay", "ProcessPromptSequence", "Waiting " & g_CloseoutConfirmDelayMs & "ms before next scan", "")
                Call WaitMs(g_CloseoutConfirmDelayMs)
            End If
            
            ' Add extra logging for problematic prompts
            If InStr(bestMatchKey, "ADD A LABOR OPERATION") > 0 Then
                Call LogEvent("comm", "high", "Responded to ADD A LABOR OPERATION prompt", "ProcessPromptSequence", "Waiting for screen to stabilize", "")
                Call WaitMs(2000) ' Extra wait for this specific prompt
                Call LogEvent("comm", "high", "Screen after ADD A LABOR OPERATION response", "ProcessPromptSequence", "", GetScreenSnapshot(5))
                
                ' Check if we're back at COMMAND prompt
                If IsTextPresent("COMMAND:") Then
                    Call LogEvent("comm", "high", "Successfully returned to COMMAND prompt", "ProcessPromptSequence", "After ADD A LABOR OPERATION", "")
                    finished = True
                Else
                    Call LogEvent("min", "med", "Not at COMMAND prompt after ADD A LABOR OPERATION", "ProcessPromptSequence", "Continuing to wait", "")
                End If
            End If
            
            
            ' Wait a moment for the response to take effect
            Call WaitMs(800)
            Call LogEvent("comm", "high", "MainPromptLine (" & MainPromptLine & ") after response", "ProcessPromptSequence", "'" & GetScreenLine(MainPromptLine) & "'", "")
            Call LogEvent("comm", "max", "Screen lines 22-24 after response", "ProcessPromptSequence", "", "[22]='" & GetScreenLine(22) & "' [23]='" & GetScreenLine(23) & "' [24]='" & GetScreenLine(24) & "'")
            
            If InStr(bestMatchKey, "SOLD HOURS") > 0 Then
                Call LogEvent("comm", "high", "Responded to SOLD HOURS prompt", "ProcessPromptSequence", "Waiting for screen to stabilize", "")
                Call WaitMs(1500) ' Extra wait for this specific prompt
                Call LogEvent("comm", "high", "Screen after SOLD HOURS response", "ProcessPromptSequence", "", GetScreenSnapshot(5))
            End If

            If promptDetails.IsSuccess Then
                finished = True
                Call LogEvent("comm", "high", "Success prompt reached", "ProcessPromptSequence", bestMatchKey, "")
            End If

            ' TRACE: Log screen snapshot after key send
            Call LogScreenSnapshot("AfterKeySend")

            ' Wait for the prompt to clear before rescanning
            Dim clearStart, clearElapsed
            If Len(Trim(CStr(POST_PROMPT_WAIT_MS))) = 0 Or POST_PROMPT_WAIT_MS <= 0 Then POST_PROMPT_WAIT_MS = 1000
            clearStart = Timer
            Do While IsTextPresent(bestMatchKey)
                Call WaitMs(500)
                clearElapsed = (Timer - clearStart) * 1000
                If clearElapsed < 0 Then clearElapsed = clearElapsed + 86400000 ' Handle midnight rollover
                If clearElapsed > POST_PROMPT_WAIT_MS Then ' Configurable prompt clear timeout
                    Call LogEvent("min", "med", "Prompt '" & bestMatchKey & "' did not clear within " & POST_PROMPT_WAIT_MS & " ms.", "ProcessPromptSequence", "Timeout reached", "clearElapsed=" & Int(clearElapsed) & "ms")
                    Exit Do
                End If
            Loop

            If promptDetails.IsSuccess Then
                finished = True
                Call LogDebug("Success prompt reached: " & bestMatchKey, "ProcessPromptSequence")
                Call LogTrace("Exiting ProcessPromptSequence on success.", "ProcessPromptSequence")
            End If
            ' The loop will now naturally restart and rescan for the next prompt
        Else
            ' No prompt found, wait a moment before trying again
            
            ' Check if we're back at the COMMAND prompt by examining MainPromptLine specifically
            ' This prevents false positives when "COMMAND" appears elsewhere on screen
            If InStr(1, mainPromptText, "COMMAND:", vbTextCompare) > 0 Then
                Call LogEvent("comm", "high", "Detected return to COMMAND prompt on MainPromptLine", "ProcessPromptSequence", "Line processing complete", "")
                finished = True
            Else
                ' Track consecutive "no prompt" iterations to prevent infinite loops
                g_NoPromptCount = g_NoPromptCount + 1
                
                If g_NoPromptCount > resolvedMaxNoPromptIterations Then
                    Call LogEvent("crit", "low", "Too many consecutive iterations with no prompt detected", "ProcessPromptSequence", "Possible infinite loop - aborting", "noPromptCount=" & g_NoPromptCount & " line=" & MainPromptLine & " text='" & mainPromptText & "'")
                    Call SafeMsg("Automation appears stuck - too many iterations with no prompt detected." & vbCrLf & "Line " & MainPromptLine & ": '" & mainPromptText & "'" & vbCrLf & "Stopping automation.", True, "Infinite Loop Detection")
                    g_ShouldAbort = True
                    finished = True
                Else
                    Call LogEvent("min", "med", "No prompt found - waiting and retrying", "ProcessPromptSequence", "Attempt " & g_NoPromptCount & " of " & resolvedMaxNoPromptIterations, "line=" & MainPromptLine & " text='" & mainPromptText & "' waitMs=" & resolvedNoPromptRetryWaitMs)
                    Call WaitMs(resolvedNoPromptRetryWaitMs)
                End If
            End If
        End If
    Loop
End Sub

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** WaitForScreenTransition
' **DATE CREATED:** 2025-12-29
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Generic function to wait for a screen transition by looking for specific text.
' Provides consistent waiting behavior across all screen transitions.
' 
' **PARAMETERS:**
' expectedText (String): The text to look for that indicates the target screen has loaded
' timeoutMs (Integer): Maximum time to wait in milliseconds (default: 3000)
' description (String): Description of what we're waiting for (for logging)
' 
' **RETURN VALUE:**
' (Boolean) Returns True if the expected text was found, False if timeout occurred
'-----------------------------------------------------------------------------------
Function WaitForScreenTransition(expectedText, timeoutMs, description)
    If timeoutMs <= 0 Then timeoutMs = 3000 ' Default 3 second timeout
    If Len(description) = 0 Then description = "screen transition"
    
    Dim waitStart, waitElapsed, found
    waitStart = Timer
    found = False
    
    Call LogEvent("comm", "max", "Waiting for " & description, "WaitForScreenTransition", "Looking for: '" & expectedText & "'", "")
    
    Do
        If IsTextPresent(expectedText) Then
            found = True
            Call LogEvent("comm", "max", description & " detected", "WaitForScreenTransition", "Found: '" & expectedText & "'", "")
            Exit Do
        End If
        
        Call WaitMs(50) ' Fast polling for quick detection
        waitElapsed = (Timer - waitStart) * 1000
        
        If waitElapsed > timeoutMs Then
            Call LogEvent("min", "med", "Timeout waiting for " & description, "WaitForScreenTransition", "After " & timeoutMs & "ms - Expected: '" & expectedText & "'", "")
            Exit Do
        End If
    Loop
    
    ' Log explicit result for clarity at point of return
    If found Then
        Call LogEvent("comm", "max", "Successfully found " & description, "WaitForScreenTransition", "Within " & Int(waitElapsed) & "ms", "")
    Else
        Call LogEvent("comm", "max", "Failed to find " & description, "WaitForScreenTransition", "Timeout after " & Int(waitElapsed) & "ms", "")
    End If
    
    WaitForScreenTransition = found
End Function

' Helper to get the text from any line of the screen (1-based)
Function GetScreenLine(lineNum)
    Dim screenContentBuffer, lineText
    On Error Resume Next
    g_bzhao.ReadScreen screenContentBuffer, 80, lineNum, 1
    If Err.Number <> 0 Then
        GetScreenLine = ""
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    lineText = Trim(screenContentBuffer)
    GetScreenLine = lineText
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** HasPartsCharged
' **DATE CREATED:** 2026-04-09
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Scans rows 9-22 of the current RO detail screen for parts lines (P1, P2, etc.)
' and checks whether any carry a non-zero SALE AMT.
'
' **DETECTION LOGIC:**
' - Reads each row via g_bzhao.ReadScreen (80 chars, col 1) to preserve column
'   positions. GetScreenLine Trims the buffer and would shift column indices.
' - Parts-line indicator: screen column 6 = "P", column 7 is a digit (P1, P2...).
' - SALE AMT field occupies the area around columns 70-80; the value is extracted
'   with Mid(buf, 70, 11), trimmed, checked with IsNumeric(), and converted with CDbl().
'
' **RETURNS:** True if at least one P-line with a sale amount > 0 is found.
'-----------------------------------------------------------------------------------
Function HasPartsCharged()
    Dim row, buf, amtRaw, amtVal
    HasPartsCharged = False
    For row = 9 To 22
        buf = ""
        On Error Resume Next
        g_bzhao.ReadScreen buf, 80, row, 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Len(buf) >= 80 Then
            If Mid(buf, 6, 1) = "P" And IsNumeric(Mid(buf, 7, 1)) Then
                amtRaw = Trim(Mid(buf, 70, 11))
                amtVal = 0
                If IsNumeric(amtRaw) Then amtVal = CDbl(amtRaw)
                If amtVal > 0 Then
                    HasPartsCharged = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** IsCdkLaborOnlyExceptionTech
' **DATE CREATED:** 2026-04-15
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Returns True when the provided labor type code is present in g_arrCDKExceptions.
' Exception LTYPE codes are configured via [PostFinalCharges] CDKLaborOnlyLTypeExceptions.
'-----------------------------------------------------------------------------------
Function IsCdkLaborOnlyExceptionTech(techCode)
    IsCdkLaborOnlyExceptionTech = False
    If Not IsArray(g_arrCDKExceptions) Then Exit Function

    Dim i, normalized
    normalized = UCase(Trim(CStr(techCode)))
    If Len(normalized) = 0 Then Exit Function

    For i = 0 To UBound(g_arrCDKExceptions)
        If normalized = g_arrCDKExceptions(i) Then
            IsCdkLaborOnlyExceptionTech = True
            Exit Function
        End If
    Next
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** IsCdkLaborOnlyExceptionDesc
' **DATE CREATED:** 2026-04-15
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Returns True when descText contains any configured labor-only description clue
' from g_arrCDKDescriptionExceptions (case-insensitive substring match).
'-----------------------------------------------------------------------------------
Function IsCdkLaborOnlyExceptionDesc(descText)
    IsCdkLaborOnlyExceptionDesc = False
    If Not IsArray(g_arrCDKDescriptionExceptions) Then Exit Function

    Dim i, lowerDesc
    lowerDesc = LCase(Trim(CStr(descText)))
    If Len(lowerDesc) = 0 Then Exit Function

    For i = 0 To UBound(g_arrCDKDescriptionExceptions)
        If Len(g_arrCDKDescriptionExceptions(i)) > 0 Then
            If InStr(1, lowerDesc, g_arrCDKDescriptionExceptions(i), vbTextCompare) > 0 Then
                IsCdkLaborOnlyExceptionDesc = True
                Exit Function
            End If
        End If
    Next
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** EvaluatePartsChargedGate
' **DATE CREATED:** 2026-04-15
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Evaluates the parts-charged closeout gate across all RO detail pages.
' - Passes immediately when a charged P-line (SALE AMT > 0) is found.
' - When NO P-lines exist across all pages, allows labor-only exceptions based on
'   configured header LTYPE codes (g_arrCDKExceptions).
' - Otherwise returns False with an explicit skip reason.
'
' **PAGINATION:**
' Uses N + <NumpadEnter> to advance detail pages and B + <NumpadEnter> to return
' to page 1, preserving downstream screen-state expectations.
'
' **PARAMETERS:**
' skipReason (ByRef) - Human-readable skip reason when the gate fails.
'
' **RETURNS:**
' True when closeout may proceed; otherwise False.
'-----------------------------------------------------------------------------------
Function EvaluatePartsChargedGate(ByRef skipReason)
    Dim row, buf, amtRaw, amtVal
    Dim pageIndicator
    Dim doneScanning, pagesAdvanced, p
    Dim preSig, postSig, preSig2, postSig2, preMarker, postMarker
    Dim maxPageAdvances
    Dim hasAnyPartLine, hasChargedPart
    Dim firstExceptionEvidence, firstNonExceptionTech
    Dim firstChar, techCode, lineDesc
    Dim hasTechException, hasDescException

    EvaluatePartsChargedGate = False
    skipReason = "Skipped - No parts charged"

    doneScanning = False
    pagesAdvanced = 0
    maxPageAdvances = 50
    hasAnyPartLine = False
    hasChargedPart = False
    firstExceptionEvidence = ""
    firstNonExceptionTech = ""

    Do While Not doneScanning
        For row = 9 To 22
            buf = ""
            On Error Resume Next
            g_bzhao.ReadScreen buf, 80, row, 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0

            If Len(buf) >= 80 Then
                If Mid(buf, 6, 1) = "P" And IsNumeric(Mid(buf, 7, 1)) Then
                    hasAnyPartLine = True
                    amtRaw = Trim(Mid(buf, 70, 11))
                    amtVal = 0
                    If IsNumeric(amtRaw) Then amtVal = CDbl(amtRaw)
                    If amtVal > 0 Then
                        hasChargedPart = True
                        Exit For
                    End If
                End If
            End If

            If Len(buf) >= 44 Then
                firstChar = Mid(buf, 1, 1)
                If firstChar >= "A" And firstChar <= "Z" Then
                    techCode = UCase(Trim(Mid(buf, 42, 8)))
                    lineDesc = Trim(Mid(buf, 4, 38))
                    hasTechException = (Len(techCode) > 0 And IsCdkLaborOnlyExceptionTech(techCode))
                    hasDescException = IsCdkLaborOnlyExceptionDesc(lineDesc)

                    If hasTechException Or hasDescException Then
                        If Len(firstExceptionEvidence) = 0 Then
                            If hasTechException Then
                                firstExceptionEvidence = "Line " & firstChar & " tech code " & techCode
                            Else
                                firstExceptionEvidence = "Line " & firstChar & " description """ & lineDesc & """"
                            End If
                        End If
                    Else
                        If Len(techCode) > 0 Then
                            If Len(firstNonExceptionTech) = 0 Then firstNonExceptionTech = techCode
                        Else
                            If Len(firstNonExceptionTech) = 0 Then firstNonExceptionTech = "Line " & firstChar
                        End If
                    End If
                End If
            End If
        Next

        If hasChargedPart Then
            doneScanning = True
        Else
            pageIndicator = ""
            On Error Resume Next
            g_bzhao.ReadScreen pageIndicator, 80, 22, 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0

            If InStr(1, pageIndicator, "(END OF DISPLAY)", vbTextCompare) > 0 Then
                doneScanning = True
            ElseIf InStr(1, pageIndicator, "(MORE ON NEXT SCREEN)", vbTextCompare) > 0 Then
                preMarker = pageIndicator
                preSig = ""
                preSig2 = ""

                On Error Resume Next
                g_bzhao.ReadScreen preSig, 80, 9, 1
                If Err.Number <> 0 Then Err.Clear
                g_bzhao.ReadScreen preSig2, 80, 10, 1
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0

                On Error Resume Next
                g_bzhao.SendKey "N"
                g_bzhao.SendKey "<NumpadEnter>"
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                g_bzhao.Pause 500

                postMarker = ""
                postSig = ""
                postSig2 = ""

                On Error Resume Next
                g_bzhao.ReadScreen postMarker, 80, 22, 1
                If Err.Number <> 0 Then Err.Clear
                g_bzhao.ReadScreen postSig, 80, 9, 1
                If Err.Number <> 0 Then Err.Clear
                g_bzhao.ReadScreen postSig2, 80, 10, 1
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0

                If postMarker = preMarker And postSig = preSig And postSig2 = preSig2 Then
                    doneScanning = True
                Else
                    pagesAdvanced = pagesAdvanced + 1
                    If pagesAdvanced >= maxPageAdvances Then
                        doneScanning = True
                    End If
                End If
            Else
                doneScanning = True
            End If
        End If
    Loop

    ' Return to page 1 so downstream flow remains on a known screen state.
    If pagesAdvanced > 0 Then
        For p = 1 To pagesAdvanced
            On Error Resume Next
            g_bzhao.SendKey "B"
            g_bzhao.SendKey "<NumpadEnter>"
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            g_bzhao.Pause 500
        Next
    End If

    If hasChargedPart Then
        EvaluatePartsChargedGate = True
        Exit Function
    End If

    ' Only bypass for labor-only exceptions when there are no P-lines anywhere.
    If Not hasAnyPartLine Then
        If Len(firstNonExceptionTech) > 0 Then
            skipReason = "Skipped - No parts charged: " & firstNonExceptionTech
            Exit Function
        End If

        If Len(firstExceptionEvidence) > 0 Then
            EvaluatePartsChargedGate = True
            Call LogEvent("comm", "med", "Labor-only exception matched - bypassing no-parts skip", "Closeout_Ro", firstExceptionEvidence, "")
            Exit Function
        End If
    End If

    If Len(firstNonExceptionTech) > 0 Then
        skipReason = "Skipped - No parts charged: " & firstNonExceptionTech
    End If
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** HasWchOnAnyDetailPage
' **DATE CREATED:** 2026-04-15
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Scans RO detail pages for WCH labor type with pagination awareness.
' Stops scanning as soon as WCH is found or when END OF DISPLAY is reached.
' Uses row 22 pagination markers and advances via N + <NumpadEnter>.
'
' **RETURNS:** True when WCH is found on any page; otherwise False.
'-----------------------------------------------------------------------------------
Function HasWchOnAnyDetailPage()
    Dim row, buf, pageIndicator
    Dim foundWch, doneScanning
    Dim pagesAdvanced, p
    Dim preSig, postSig, preSig2, postSig2, preMarker, postMarker
    Dim maxPageAdvances

    HasWchOnAnyDetailPage = False
    pagesAdvanced = 0
    doneScanning = False
    maxPageAdvances = 50

    Do While Not doneScanning
        foundWch = False

        ' Scan current detail page (rows 9-22)
        On Error Resume Next
        For row = 9 To 22
            buf = ""
            g_bzhao.ReadScreen buf, 80, row, 1
            If Err.Number <> 0 Then
                Err.Clear
            Else
                If InStr(1, buf, "WCH", vbTextCompare) > 0 Then
                    foundWch = True
                    Exit For
                End If
            End If
        Next
        On Error GoTo 0

        If foundWch Then
            HasWchOnAnyDetailPage = True
            doneScanning = True
        Else
            pageIndicator = ""
            On Error Resume Next
            g_bzhao.ReadScreen pageIndicator, 80, 22, 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0

            If InStr(1, pageIndicator, "(END OF DISPLAY)", vbTextCompare) > 0 Then
                doneScanning = True
            ElseIf InStr(1, pageIndicator, "(MORE ON NEXT SCREEN)", vbTextCompare) > 0 Then
                preMarker = pageIndicator
                preSig = ""
                preSig2 = ""
                On Error Resume Next
                g_bzhao.ReadScreen preSig, 80, 9, 1
                If Err.Number <> 0 Then Err.Clear
                g_bzhao.ReadScreen preSig2, 80, 10, 1
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0

                On Error Resume Next
                g_bzhao.SendKey "N"
                g_bzhao.SendKey "<NumpadEnter>"
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                g_bzhao.Pause 500

                postMarker = ""
                postSig = ""
                postSig2 = ""
                On Error Resume Next
                g_bzhao.ReadScreen postMarker, 80, 22, 1
                If Err.Number <> 0 Then Err.Clear
                g_bzhao.ReadScreen postSig, 80, 9, 1
                If Err.Number <> 0 Then Err.Clear
                g_bzhao.ReadScreen postSig2, 80, 10, 1
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0

                If postMarker = preMarker And postSig = preSig And postSig2 = preSig2 Then
                    doneScanning = True
                Else
                    pagesAdvanced = pagesAdvanced + 1
                    If pagesAdvanced >= maxPageAdvances Then
                        doneScanning = True
                    End If
                End If
            Else
                doneScanning = True
            End If
        End If
    Loop

    ' Return to page 1 so downstream flow remains on a known screen state.
    If pagesAdvanced > 0 Then
        For p = 1 To pagesAdvanced
            On Error Resume Next
            g_bzhao.SendKey "B"
            g_bzhao.SendKey "<NumpadEnter>"
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            g_bzhao.Pause 500
        Next
    End If
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** IsWchLine
' **DATE CREATED:** 2026-04-10
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Scans rows 9-22 for L-operations belonging to the given line letter and checks
' whether any carry a LABOR TYPE of "WCH" (col 50-55, 1-indexed).
'
' Screen layout (from RO DETAIL header row):
'   LC DESCRIPTION                           TECH... LTYPE    ACT   SOLD    SALE AMT
' Line letter headers have the letter at col 1. L-operation rows have:
'   - col 1 = " " (space-indented)
'   - col 4 = "L", col 5 = digit (e.g., L1, L2)
'   - col 50-55 = LTYPE value (e.g., "WCH", "I", "B")
'
' **PARAMETERS:**
' lineLetterChar - Single uppercase letter identifying the line (e.g., "A")
'
' **RETURNS:** True if any L-row for the given line has LTYPE = "WCH".
'-----------------------------------------------------------------------------------
Function IsWchLine(lineLetterChar)
    IsWchLine = False
    Dim row, buf, inTargetLine, firstChar
    inTargetLine = False
    For row = 9 To 22
        buf = ""
        On Error Resume Next
        g_bzhao.ReadScreen buf, 80, row, 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Len(buf) >= 55 Then
            firstChar = Mid(buf, 1, 1)
            ' Line letter headers: uppercase A-Z in col 1
            If firstChar >= "A" And firstChar <= "Z" Then
                If inTargetLine Then Exit For  ' Past our target line - done scanning
                If firstChar = lineLetterChar Then inTargetLine = True
            End If
            ' Within target line, check L-rows (col 4 = "L", col 5 = digit) for LTYPE
            If inTargetLine And Mid(buf, 4, 1) = "L" And IsNumeric(Mid(buf, 5, 1)) Then
                If Trim(Mid(buf, 50, 6)) = "WCH" Then
                    IsWchLine = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** ExtractPartNumberForFca
' **DATE CREATED:** 2026-04-10
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Scans rows 9-22 for the first P-line (col 6 = "P", col 7 = digit) and extracts
' the part number from cols 9 onwards. Intended to be called BEFORE the FNL
' command is issued (while the detail screen is clean, no dialog overlay), since
' IsWchLine() can predict when the FCA dialog will appear.
'
' **RETURNS:** Part number string (e.g., "BBH6A001AA"), or "" if not found.
'-----------------------------------------------------------------------------------
Function ExtractPartNumberForFca()
    Dim row, buf, partToken, spacePos
    ExtractPartNumberForFca = ""
    For row = 9 To 22
        buf = ""
        On Error Resume Next
        g_bzhao.ReadScreen buf, 80, row, 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Len(buf) >= 20 Then
            If Mid(buf, 6, 1) = "P" And IsNumeric(Mid(buf, 7, 1)) Then
                partToken = Trim(Mid(buf, 9, 20))
                spacePos = InStr(1, partToken, " ")
                If spacePos > 1 Then partToken = Left(partToken, spacePos - 1)
                If Len(partToken) > 0 Then
                    ExtractPartNumberForFca = partToken
                    Exit Function
                End If
            End If
        End If
    Next
    Call LogWarn("No P-line found on screen - cannot extract part number for FCA dialog", "ExtractPartNumberForFca")
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** DescMatchesPartsKeyword
' **DATE CREATED:** 2026-04-14
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Returns True if descText contains any of g_PartsOrderKeywords (case-insensitive
' substring match) AND does NOT contain any of g_PartsOrderNegators.
' Both arrays are pre-lowercased by InitializeConfig.
'
' **PARAMETERS:**
' descText - Raw labor-line description text (trimmed, any case)
'
' **RETURNS:** True if a keyword fires and no negator is present.
'-----------------------------------------------------------------------------------
Function DescMatchesPartsKeyword(descText)
    DescMatchesPartsKeyword = False
    If Not IsArray(g_PartsOrderKeywords) Or Not IsArray(g_PartsOrderNegators) Then Exit Function
    Dim i, lowerDesc
    lowerDesc = LCase(descText)
    For i = 0 To UBound(g_PartsOrderNegators)
        If Len(g_PartsOrderNegators(i)) > 0 Then
            If InStr(1, lowerDesc, g_PartsOrderNegators(i), vbTextCompare) > 0 Then Exit Function
        End If
    Next
    For i = 0 To UBound(g_PartsOrderKeywords)
        If Len(g_PartsOrderKeywords(i)) > 0 Then
            If InStr(1, lowerDesc, g_PartsOrderKeywords(i), vbTextCompare) > 0 Then
                DescMatchesPartsKeyword = True
                Exit Function
            End If
        End If
    Next
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** GetPartsNeededLaborDesc
' **DATE CREATED:** 2026-04-14
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Scans rows 9-23 of the current RO detail screen and returns the description of
' the first L-line that has NO P-line immediately following it and whose
' description matches a parts-order keyword (via DescMatchesPartsKeyword).
'
' Column layout (1-indexed, matches HasPartsCharged / IsWchLine conventions):
'   L-operation rows : col 4 = "L", col 5 = digit
'   P-operation rows : col 6 = "P", col 7 = digit
'   Description field: cols 7-41 -> Mid(buf, 7, 35), then Trimmed
'
' Row 23 (COMMAND:) is read solely to flush the last pending L-row; it cannot
' be mistaken for a P-row (col 6 = "A" from "COMMAND").
'
' **RETURNS:** First matching description string, or "" if none found.
'-----------------------------------------------------------------------------------
Function GetPartsNeededLaborDesc()
    GetPartsNeededLaborDesc = ""
    If Not IsArray(g_PartsOrderKeywords) Then Exit Function
    If UBound(g_PartsOrderKeywords) < 0 Then Exit Function

    Dim row, buf, pendingDesc, pendingRow
    pendingDesc = ""
    pendingRow = -1

    For row = 9 To 23
        buf = ""
        On Error Resume Next
        g_bzhao.ReadScreen buf, 80, row, 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0

        If pendingRow > 0 Then
            ' Check whether this row is a P-line (parts attached to the pending L-row)
            If Len(buf) >= 7 And Mid(buf, 6, 1) = "P" And IsNumeric(Mid(buf, 7, 1)) Then
                ' P-line found: pending L-row has parts, clear without evaluating
                pendingDesc = ""
                pendingRow = -1
            Else
                ' No P-line: evaluate the pending L-row now
                If DescMatchesPartsKeyword(pendingDesc) Then
                    GetPartsNeededLaborDesc = pendingDesc
                    Exit Function
                End If
                pendingDesc = ""
                pendingRow = -1
                ' Current row may itself be a new L-row
                If Len(buf) >= 5 And Mid(buf, 4, 1) = "L" And IsNumeric(Mid(buf, 5, 1)) Then
                    pendingDesc = Trim(Mid(buf, 7, 35))
                    pendingRow = row
                End If
            End If
        Else
            If Len(buf) >= 5 And Mid(buf, 4, 1) = "L" And IsNumeric(Mid(buf, 5, 1)) Then
                pendingDesc = Trim(Mid(buf, 7, 35))
                pendingRow = row
            End If
        End If
    Next
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** GetFirstNonCompliantLineTech
' **DATE CREATED:** 2026-04-14
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Scans rows 9-22 of the RO detail screen for line-letter header rows
' (col 1 = uppercase A-Z) and extracts the tech/status code at col 42
' (1-indexed). Returns a description of the first line whose code is not
' found in g_AllowedTechCodes, or "" if every line passes.
'
' Screen layout (1-indexed):
'   Col 1     = line letter (A, B, C ...)
'   Cols 4-41 = description (38 chars)
'   Cols 42+  = tech code (e.g. "C92", "I91")
'
' An empty or blank tech code on a header row is treated as compliant
' (no tech assigned yet, not an unauthorized code).
'
' **RETURNS:**
'   "" if all lines are compliant or AllowedTechCodes is not configured.
'   "Line X: <code>" (e.g. "Line B: I91") for the first non-compliant line.
'-----------------------------------------------------------------------------------
Function GetFirstNonCompliantLineTech()
    GetFirstNonCompliantLineTech = ""
    If Not IsArray(g_AllowedTechCodes) Then Exit Function
    If UBound(g_AllowedTechCodes) < 0 Then Exit Function
    If Len(Trim(g_AllowedTechCodes(0))) = 0 Then Exit Function  ' empty config = gate disabled

    Dim row, buf, firstChar, techCode, i, isAllowed
    For row = 9 To 22
        buf = ""
        On Error Resume Next
        g_bzhao.ReadScreen buf, 80, row, 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Len(buf) >= 44 Then
            firstChar = Mid(buf, 1, 1)
            If firstChar >= "A" And firstChar <= "Z" Then
                techCode = UCase(Trim(Mid(buf, 42, 8)))
                If Len(techCode) > 0 Then
                    isAllowed = False
                    For i = 0 To UBound(g_AllowedTechCodes)
                        If techCode = g_AllowedTechCodes(i) Then
                            isAllowed = True
                            Exit For
                        End If
                    Next
                    If Not isAllowed Then
                        GetFirstNonCompliantLineTech = "Line " & firstChar & ": " & techCode
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** CreateFcaPromptDictionary
' **DATE CREATED:** 2026-04-10
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Builds the prompt dictionary for the FCA Global Claims Information dialog.
' Uses the same AddPromptToDict pattern as CreateLineItemPromptDictionary().
'
' The Condition Code footer prompt is confirmed from screen capture.
' Remaining field footer prompts are marked TODO and require a live WCH POC
' session to determine exact text before those entries can be activated.
'
' **PARAMETERS:**
' partNumber   - Failed Part Number extracted from P1 line on RO detail screen
' condCode     - Condition Code value ("1", "2", or "3") from config
' causalLop    - Causal LOP answer ("Y" or "N") from config
' calEmissions - Cal. Emissions answer ("Y" or "N") from config
'-----------------------------------------------------------------------------------
Function CreateFcaPromptDictionary(partNumber, condCode, causalLop, calEmissions)
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")
    ' Condition Code: footer confirmed from screen capture (2026-04-10)
    Call AddPromptToDict(dict, "Enter 1, 2, or 3 for the Condition Code.", condCode, "<NumpadEnter>", False)
    ' TODO POC: Add footer prompts for Causal LOP, Cal. Emissions, Failure Code,
    '           and Failed Part Number once verified via live WCH session.
    ' Example (uncomment and update text after POC):
    '   Call AddPromptToDict(dict, "<causal lop footer text>", causalLop, "<NumpadEnter>", False)
    '   Call AddPromptToDict(dict, "<cal emissions footer text>", calEmissions, "<NumpadEnter>", False)
    '   Call AddPromptToDict(dict, "<failed part number footer text>", partNumber, "<NumpadEnter>", False)
    '   Call AddPromptToDict(dict, "<failure code footer text>", "", "<NumpadEnter>", False)
    Set CreateFcaPromptDictionary = dict
End Function

'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** HandleFcaDialog
' **DATE CREATED:** 2026-04-10
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Detects and handles the FCA Global Claims Information dialog that appears when
' a WCH labor-type line is finalized. Extracts the Failed Part Number from the
' background RO detail screen (left portion, unobscured by dialog overlay), reads
' config values for Condition Code and Y/N fields, then drives field entry via
' ProcessPromptSequence.
'
' If no part number is found, the RO is flagged for manual review and the script
' is halted (g_ShouldAbort = True) so the operator can intervene.
'-----------------------------------------------------------------------------------
Sub HandleFcaDialog(prePartNumber)
    Call LogInfo("FCA warranty dialog detected", "HandleFcaDialog")

    ' Feature flag: disabled until field values are confirmed with management
    If LCase(Trim(GetIniSetting("PostFinalCharges", "FcaDialogEnabled", "false"))) <> "true" Then
        Call LogWarn("FCA dialog handler is disabled (FcaDialogEnabled=false). WCH RO requires manual review.", "HandleFcaDialog")
        lastRoResult = "Skipped - FCA dialog handler not yet configured"
        Exit Sub
    End If

    Call LogInfo("Beginning automated FCA field entry", "HandleFcaDialog")

    ' Read config values
    Dim condCode, causalLop, calEmissions
    condCode     = Trim(GetIniSetting("PostFinalCharges", "FcaConditionCode", "1"))
    causalLop    = Trim(GetIniSetting("PostFinalCharges", "FcaCausalLop", "Y"))
    calEmissions = Trim(GetIniSetting("PostFinalCharges", "FcaCalEmissions", "N"))

    ' Use pre-captured part number (extracted before FNL, no dialog overlay).
    ' prePartNumber is supplied by the caller who called IsWchLine() + ExtractPartNumberForFca()
    ' before issuing the FNL command.
    If Len(prePartNumber) = 0 Then
        Call LogWarn("Cannot automate FCA dialog: no part number was pre-captured. Flagging for manual review.", "HandleFcaDialog")
        lastRoResult = "Flagged - Missing part number for FCA dialog"
        g_ShouldAbort = True
        Exit Sub
    End If
    Call LogInfo("Using pre-captured part number for FCA dialog: " & prePartNumber, "HandleFcaDialog")

    ' Drive field entry via prompt sequence
    Call ProcessPromptSequence(CreateFcaPromptDictionary(prePartNumber, condCode, causalLop, calEmissions))

    Call LogInfo("FCA dialog processing complete", "HandleFcaDialog")
End Sub

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** GetScreenLines
' **DATE CREATED:** 2026-02-13
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Read multiple lines from the BlueZone screen and return as delimited string.
' Used for screen capture/snapshot for logging and debugging.
' 
' **PARAMETERS:**
' startLine (Integer): Starting line number (1-based)
' numLines (Integer): Number of lines to read
' 
' **RETURN VALUE:**
' (String) Lines formatted as "Line N: [content]" separated by | delimiter
'-----------------------------------------------------------------------------------
Function GetScreenLines(startLine, numLines)
    Dim lineNum, screenContentBuffer, lineContent, result
    Dim maxLine
    
    result = ""
    If numLines <= 0 Then numLines = 1
    If startLine <= 0 Then startLine = 1
    
    maxLine = startLine + numLines - 1
    If maxLine > 24 Then maxLine = 24 ' Don't exceed screen height
    
    On Error Resume Next
    For lineNum = startLine To maxLine
        screenContentBuffer = ""
        g_bzhao.ReadScreen screenContentBuffer, 80, lineNum, 1
        If Err.Number = 0 Then
            lineContent = Trim(screenContentBuffer)
            If Len(lineContent) > 0 Then
                If Len(result) > 0 Then result = result & " | "
                result = result & "L" & lineNum & ":[" & lineContent & "]"
            End If
        End If
        Err.Clear
    Next
    On Error GoTo 0
    
    GetScreenLines = result
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** GetScreenSnapshot
' **DATE CREATED:** 2026-02-13
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Capture a snapshot of the current BlueZone screen for logging and debugging.
' Reads the specified number of lines starting from line 1 and returns formatted string.
' 
' **PARAMETERS:**
' numLines (Integer): Number of lines to capture (default: 24 for full screen)
' 
' **RETURN VALUE:**
' (String) Screen snapshot formatted as "L1:[...] | L2:[...] | ..." for easy logging
'-----------------------------------------------------------------------------------
Function GetScreenSnapshot(numLines)
    If numLines <= 0 Then numLines = 24 ' Default to full screen
    If numLines > 24 Then numLines = 24 ' Don't exceed screen height
    
    GetScreenSnapshot = GetScreenLines(1, numLines)
End Function

' IsTextPresent — provided by framework\BZHelper.vbs
' BZH_GetMatchedBlacklistTerm — provided by framework\BZHelper.vbs (paginated)

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** NormalizeRoIdentifier
' **FUNCTIONALITY:**
' Extracts numeric RO characters from mixed text and returns normalized RO ID.
' Returns empty string when no numeric content exists.
'-----------------------------------------------------------------------------------
Function NormalizeRoIdentifier(rawValue)
    Dim textValue, i, ch, normalized
    textValue = Trim(CStr(rawValue))
    normalized = ""

    For i = 1 To Len(textValue)
        ch = Mid(textValue, i, 1)
        If ch >= "0" And ch <= "9" Then
            normalized = normalized & ch
        End If
    Next

    NormalizeRoIdentifier = normalized
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** LoadCloseoutTriggers
' **FUNCTIONALITY:**
' Loads closeout trigger strings from a configured CSV/text file.
' The file uses one trigger per line; blank lines and comment lines (# or ;) are ignored.
' Fails fast if the configured file is missing or contains no usable triggers.
'-----------------------------------------------------------------------------------
Function LoadCloseoutTriggers(triggerListPath, ByRef triggerArray)
    Dim fullPath, fileHandle, lineText, cleanedValue
    Dim loadedCount

    LoadCloseoutTriggers = False
    ReDim triggerArray(-1)

    triggerListPath = Trim(CStr(triggerListPath))
    If Len(triggerListPath) = 0 Then
        Call LogEvent("crit", "low", "TriggerList config missing", "LoadCloseoutTriggers", "PostFinalCharges.TriggerList", "Configure TriggerList in config.ini")
        Exit Function
    End If

    If IsAbsolutePath(triggerListPath) Then
        fullPath = triggerListPath
    Else
        fullPath = g_fso.BuildPath(GetRepoRoot(), triggerListPath)
    End If

    If Not g_fso.FileExists(fullPath) Then
        Call LogEvent("crit", "low", "Configured TriggerList file not found", "LoadCloseoutTriggers", fullPath, "Check PostFinalCharges.TriggerList in config.ini")
        Exit Function
    End If

    loadedCount = -1
    Set fileHandle = g_fso.OpenTextFile(fullPath, 1)
    Do While Not fileHandle.AtEndOfStream
        lineText = Trim(CStr(fileHandle.ReadLine))

        If Len(lineText) = 0 Then
            ' ignore blank lines
        ElseIf Left(lineText, 1) = "#" Or Left(lineText, 1) = ";" Then
            ' ignore comment lines
        Else
            cleanedValue = lineText
            If Left(cleanedValue, 1) = Chr(34) And Right(cleanedValue, 1) = Chr(34) And Len(cleanedValue) >= 2 Then
                cleanedValue = Mid(cleanedValue, 2, Len(cleanedValue) - 2)
            End If

            If Len(Trim(CStr(cleanedValue))) > 0 Then
                loadedCount = loadedCount + 1
                ReDim Preserve triggerArray(loadedCount)
                triggerArray(loadedCount) = cleanedValue
            End If
        End If
    Loop
    fileHandle.Close

    If loadedCount < 0 Then
        Call LogEvent("crit", "low", "Configured TriggerList is empty", "LoadCloseoutTriggers", fullPath, "Add at least one trigger entry")
        Exit Function
    End If

    Call LogEvent("comm", "med", "Loaded closeout triggers", "LoadCloseoutTriggers", "Count: " & (UBound(triggerArray) + 1), "Source: " & triggerListPath)
    LoadCloseoutTriggers = True
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** LoadSkipRoLookup
' **FUNCTIONALITY:**
' Loads RO numbers from one or more configured CSV files into a dictionary.
' Config value is a comma-separated list of file paths (relative to repo root or absolute).
' Fails fast if any configured file is missing.
'-----------------------------------------------------------------------------------
Function LoadSkipRoLookup(skipRoListCsvPaths, ByRef lookupDict)
    Dim pathEntries, i, configuredPath, fullPath
    Dim fileHandle, lineText, normalizedRo

    LoadSkipRoLookup = True

    If Not IsObject(lookupDict) Then
        Set lookupDict = CreateObject("Scripting.Dictionary")
    End If

    skipRoListCsvPaths = Trim(CStr(skipRoListCsvPaths))
    If Len(skipRoListCsvPaths) = 0 Then Exit Function

    pathEntries = Split(skipRoListCsvPaths, ",")
    For i = LBound(pathEntries) To UBound(pathEntries)
        configuredPath = Trim(CStr(pathEntries(i)))
        If Len(configuredPath) > 0 Then
            If IsAbsolutePath(configuredPath) Then
                fullPath = configuredPath
            Else
                fullPath = g_fso.BuildPath(GetRepoRoot(), configuredPath)
            End If

            If Not g_fso.FileExists(fullPath) Then
                Call LogEvent("crit", "low", "Configured SkipRoList file not found", "LoadSkipRoLookup", fullPath, "Check PostFinalCharges.SkipRoList in config.ini")
                LoadSkipRoLookup = False
                Exit Function
            End If

            Set fileHandle = g_fso.OpenTextFile(fullPath, 1)
            Do While Not fileHandle.AtEndOfStream
                lineText = Trim(CStr(fileHandle.ReadLine))

                If Len(lineText) = 0 Then
                    ' ignore blank lines
                ElseIf Left(lineText, 1) = "#" Or Left(lineText, 1) = ";" Then
                    ' ignore comment lines
                Else
                    normalizedRo = NormalizeRoIdentifier(lineText)
                    If Len(normalizedRo) > 0 Then
                        If Not lookupDict.Exists(normalizedRo) Then
                            lookupDict.Add normalizedRo, True
                        End If
                    End If
                End If
            Loop
            fileHandle.Close
        End If
    Next

    Call LogEvent("comm", "med", "Loaded SkipRoList entries", "LoadSkipRoLookup", "Count: " & lookupDict.Count, "Sources: " & skipRoListCsvPaths)
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** ShouldSkipRo
' **FUNCTIONALITY:**
' Returns True when current RO identifier exists in configured SkipRoList.
'-----------------------------------------------------------------------------------
Function ShouldSkipRo(roValue)
    Dim normalized
    ShouldSkipRo = False

    normalized = NormalizeRoIdentifier(roValue)
    If Len(normalized) = 0 Then Exit Function

    If IsObject(g_SkipRoLookup) Then
        If g_SkipRoLookup.Exists(normalized) Then
            ShouldSkipRo = True
        End If
    End If
End Function

' WaitMs — provided by framework\BZHelper.vbs

' WaitForPrompt — provided by framework\BZHelper.vbs

' FastKey - Send a key press to the terminal
' Supports special keys like <NumpadEnter>, <Enter>, or regular characters
Sub FastKey(keyValue)
    On Error Resume Next
    If Len(keyValue) = 0 Then
        Call LogEvent("warn", "low", "FastKey: Empty key value", "FastKey", "", "")
        On Error GoTo 0
        Exit Sub
    End If
    
    Call LogEvent("comm", "high", "FastKey: Sending '" & keyValue & "'", "FastKey", "", "")
    g_bzhao.SendKey keyValue
    
    If Err.Number <> 0 Then
        Call LogEvent("maj", "med", "FastKey: Failed to send '" & keyValue & "'", "FastKey", "Error: " & Err.Description, "")
        Err.Clear
    Else
        Call LogEvent("comm", "high", "FastKey: Sent '" & keyValue & "' successfully", "FastKey", "", "")
    End If
    On Error GoTo 0
End Sub

' FastText - Send text to the terminal (convenience wrapper for SendKey)
Sub FastText(textValue)
    On Error Resume Next
    If Len(textValue) = 0 Then
        Call LogEvent("warn", "low", "FastText: Empty text value", "FastText", "", "")
        On Error GoTo 0
        Exit Sub
    End If
    
    Call LogEvent("comm", "high", "FastText: Sending '" & textValue & "'", "FastText", "", "")
    g_bzhao.SendKey textValue
    
    If Err.Number <> 0 Then
        Call LogEvent("maj", "med", "FastText: Failed to send '" & textValue & "'", "FastText", "Error: " & Err.Description, "")
        Err.Clear
    Else
        Call LogEvent("comm", "high", "FastText: Sent '" & textValue & "' successfully", "FastText", "", "")
    End If
    On Error GoTo 0
End Sub

' Helper to check if a prompt is in the prompts dictionary
Function IsPromptInConfig(promptText, promptsDict)
    Dim key, trimmedPromptText
    trimmedPromptText = Trim(promptText)
    
    ' Special handling for COMMAND prompts that may show the last entered command
    ' e.g., "COMMAND: R A" should match "COMMAND:" in the dictionary
    If InStr(1, trimmedPromptText, "COMMAND:", vbTextCompare) = 1 Then
        ' Check if any COMMAND-related keys exist in the dictionary (also starting with COMMAND:)
        For Each key In promptsDict.Keys
            If InStr(1, key, "COMMAND:", vbTextCompare) = 1 Then
                IsPromptInConfig = True
                Exit Function
            End If
        Next
    End If
    
    For Each key In promptsDict.Keys
        Dim isRegex, re, promptDetails
        isRegex = False
        On Error Resume Next
        Set promptDetails = promptsDict.Item(key)
        If Err.Number = 0 Then
            isRegex = CBool(promptDetails.IsRegex)
        End If
        Err.Clear
        On Error GoTo 0

        If Not isRegex Then
            isRegex = InferRegexPattern(key)
        End If

        If isRegex Then
            On Error Resume Next
            Set re = CreateObject("VBScript.RegExp")
            re.Pattern = key
            re.IgnoreCase = True
            re.Global = False
            If Err.Number = 0 Then
                If re.Test(trimmedPromptText) Then
                    IsPromptInConfig = True
                    Exit Function
                End If
            End If
            Err.Clear
            On Error GoTo 0
        Else
            ' For non-regex patterns, check for exact match or substring match
            If StrComp(trimmedPromptText, Trim(key), vbTextCompare) = 0 Or InStr(1, trimmedPromptText, key, vbTextCompare) > 0 Then
                IsPromptInConfig = True
                Exit Function
            End If
        End If
    Next
    IsPromptInConfig = False
End Function

' Checks if a prompt contains a default value in parentheses that should be accepted
Function HasDefaultValueInPrompt(promptPattern, screenContent)
    HasDefaultValueInPrompt = False
    
    Call LogEvent("comm", "max", "Analyzing prompt for default values", "HasDefaultValueInPrompt", "Pattern: '" & promptPattern & "'", "Content: '" & screenContent & "'")
    
    ' Use a more robust approach - look for any text followed by parentheses containing alphanumeric content
    ' This handles all prompt types without hardcoding specific patterns
    On Error Resume Next
    Dim re, matches, match, parenContent
    Set re = CreateObject("VBScript.RegExp")
    
    ' Universal pattern: any text followed by parentheses containing non-empty alphanumeric content
    ' Examples: TECHNICIAN(12345), ACTUAL HOURS (8), SOLD HOURS (10), OPERATION CODE FOR LINE A, L1 (I)
    ' Updated pattern to handle any content before parentheses
    re.Pattern = ".*\(([A-Za-z0-9]+)\)"
    re.IgnoreCase = True
    re.Global = False
    
    Call LogEvent("comm", "max", "Using regex pattern", "HasDefaultValueInPrompt", "'" & re.Pattern & "'", "")
    
    If Err.Number = 0 Then
        Set matches = re.Execute(screenContent)
        Call LogEvent("comm", "max", "Regex execution completed", "HasDefaultValueInPrompt", "Match count: " & matches.Count, "")
        
        If matches.Count > 0 Then
            Set match = matches(0)
            Call LogEvent("comm", "max", "First match found", "HasDefaultValueInPrompt", "'" & match.Value & "'", "SubMatches count: " & match.SubMatches.Count)
            
            If match.SubMatches.Count > 0 Then
                parenContent = Trim(match.SubMatches(0))
                Call LogEvent("comm", "max", "Extracted parentheses content", "HasDefaultValueInPrompt", "'" & parenContent & "'", "")
                
                ' If there's content in parentheses and it's not empty or just question marks
                If Len(parenContent) > 0 And parenContent <> "?" And parenContent <> "" Then
                    HasDefaultValueInPrompt = True
                    Call LogEvent("comm", "max", "FOUND VALID DEFAULT VALUE: " & parenContent, "HasDefaultValueInPrompt", "Will accept default", "")
                Else
                    Call LogEvent("comm", "max", "Invalid parentheses content", "HasDefaultValueInPrompt", "Empty, '?' or whitespace only - no default to accept", "")
                End If
            Else
                Call LogEvent("comm", "max", "No submatches found in regex result", "HasDefaultValueInPrompt", "", "")
            End If
        Else
            Call LogEvent("comm", "max", "No regex matches found in screen content", "HasDefaultValueInPrompt", "", "")
        End If
    Else
        Call LogEvent("min", "med", "Regex execution error", "HasDefaultValueInPrompt", Err.Description & " (" & Err.Number & ")", "")
    End If
    
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    
    Call LogEvent("comm", "max", "Final result: HasDefaultValueInPrompt = " & HasDefaultValueInPrompt, "HasDefaultValueInPrompt", "", "")
End Function

    ' Bootstrap defaults so logging works before config initialization
    g_BaseScriptPath = ""
    CSV_FILE_PATH = ResolvePath("CashoutRoList.csv", LEGACY_CSV_PATH, True)
    LOG_FILE_PATH = ResolvePath("PostFinalCharges.log", LEGACY_LOG_PATH, False)
    commonLibLoaded = False
    g_ShouldAbort = False

' LogResult adapter for the library
Sub LogResult(logMsg)
    ' This adapter is for messages coming from the legacy CommonLib.
    ' Parse ERROR prefix and route to appropriate criticality level.
    If Left(UCase(Trim(logMsg)), 6) = "ERROR:" Then
        Call LogEvent("maj", "med", "ADAPTER: " & logMsg, "CommonLib", "", "")
    Else
        Call LogEvent("comm", "high", "ADAPTER: " & logMsg, "CommonLib", "", "")
    End If
End Sub

' --- Common Library for BlueZone Scripts ---
' Contains reusable functions for terminal interaction.



Sub RunMainProcess()
    '------------------------------
    ' Initialization and main flow
    '------------------------------
    g_AbortReason = ""
    Call InitializeObjects()
    If commonLibLoaded Then
        Call LogEvent("comm", "med", "CommonLib loaded successfully", "RunMainProcess", "In PostFinalCharges.vbs", "")
    End If
    If ConnectBlueZone() Then
        ProcessRONumbers()
        If Not g_IsTestMode Then
            Dim summaryMsg
            Dim accountedTotal
            Dim otherOutcomeDetails

            accountedTotal = g_FiledROCount + _
                g_SkipConfiguredCount + _
                g_SkipWarrantyCount + _
                g_SkipTechCodeCount + _
                g_SkipPartsOrderNeededCount + _
                g_SkipBlacklistCount + _
                g_SkipStatusOpenCount + _
                g_SkipStatusPreassignedCount + _
                g_SkipStatusOtherCount + _
                g_ClosedRoCount + _
                g_NotOnFileRoCount + _
                g_SkipVehidNotOnFileCount + _
                g_SkipNoCloseoutTextCount + _
                g_SkipNoPartsChargedCount + _
                g_LeftOpenManualCount + _
                g_FcaMissingPartFlagCount + _
                g_FcaHandlerNotConfiguredCount + _
                g_ErrorInMainCount + _
                g_NoResultRecordedCount + _
                g_SummaryOtherOutcomeCount

            summaryMsg = "DONE" & vbCrLf & _
                "ROs Reviewed: " & g_ReviewedROCount & vbCrLf & _
                "ROs Posted: " & g_FiledROCount & vbCrLf & _
                "Skips - Specific ROs: " & g_SkipConfiguredCount & vbCrLf & _
                "Skips - Warranty (WCH): " & g_SkipWarrantyCount & vbCrLf & _
                "Skips - Non-compliant tech code: " & g_SkipTechCodeCount & vbCrLf & _
                "Skips - Parts Order Needed: " & g_SkipPartsOrderNeededCount & vbCrLf & _
                "Skips - Other Terms: " & g_SkipBlacklistCount & vbCrLf & _
                "Skips - Open: " & g_SkipStatusOpenCount & vbCrLf & _
                "Skips - Pre-Assigned: " & g_SkipStatusPreassignedCount & vbCrLf & _
                "Skips - Other Statuses: " & g_SkipStatusOtherCount & vbCrLf & _
                "Closed (already): " & g_ClosedRoCount & vbCrLf & _
                "Not On File: " & g_NotOnFileRoCount & vbCrLf & _
                "Skipped - VEHID not on file: " & g_SkipVehidNotOnFileCount & vbCrLf & _
                "Skipped - No closeout text: " & g_SkipNoCloseoutTextCount & vbCrLf & _
                "Skipped - No parts charged: " & g_SkipNoPartsChargedCount & vbCrLf & _
                "Left Open for manual closing: " & g_LeftOpenManualCount & vbCrLf & _
                "Flagged - Missing FCA part #: " & g_FcaMissingPartFlagCount & vbCrLf & _
                "Skipped - FCA handler not configured: " & g_FcaHandlerNotConfiguredCount & vbCrLf & _
                "Errors in Main: " & g_ErrorInMainCount & vbCrLf & _
                "No result recorded: " & g_NoResultRecordedCount & vbCrLf & _
                "Other Outcomes: " & g_SummaryOtherOutcomeCount & vbCrLf & _
                "Accounted Total: " & accountedTotal & vbCrLf & _
                "Older ROs Attempted (subset): " & g_OlderRoAttemptCount

            If g_SummaryOtherOutcomeCount > 0 Then
                otherOutcomeDetails = BuildOtherOutcomeBreakdown(8)
                If Len(Trim(CStr(otherOutcomeDetails))) > 0 Then
                    summaryMsg = summaryMsg & vbCrLf & "Other Outcome Breakdown:" & vbCrLf & otherOutcomeDetails
                End If

                otherOutcomeDetails = BuildOtherOutcomeRawBreakdown(12)
                If Len(Trim(CStr(otherOutcomeDetails))) > 0 Then
                    summaryMsg = summaryMsg & vbCrLf & "Other Outcome Raw Results:" & vbCrLf & otherOutcomeDetails
                End If
            End If

            Call SafeMsg(summaryMsg, False, "PostFinalCharges")
        End If
    Else
        SafeMsg "Unable to connect to BlueZone. Check that itG��s open and logged in.", True, "Connection Error"
    End If

    ' Cleanup
    ' Guard object cleanup with IsObject to avoid 'Object required' when variables are Empty
    If IsObject(g_bzhao) Then
        On Error Resume Next
        g_bzhao.Disconnect
        Set g_bzhao = Nothing
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If
End Sub

'----------------------------------------------------
' Includes a VBScript file into the global scope
'----------------------------------------------------
Function IncludeFile(filePath)
    On Error Resume Next
    Dim fsoInclude, fileContent, includeStream

    Set fsoInclude = CreateObject("Scripting.FileSystemObject")

    If Not fsoInclude.FileExists(filePath) Then
        Call LogEvent("crit", "low", "IncludeFile - File not found", "IncludeFile", filePath, "")
        IncludeFile = False
        Exit Function
    End If

    Set includeStream = fsoInclude.OpenTextFile(filePath, 1)
    fileContent = includeStream.ReadAll
    includeStream.Close
    Set includeStream = Nothing

    ExecuteGlobal fileContent

    If Err.Number <> 0 Then
        Call LogEvent("crit", "med", "IncludeFile - Error executing file '" & filePath & "'", "IncludeFile", Err.Description, "Error " & Err.Number & ", " & Err.Source)
        Err.Clear
        IncludeFile = False
        Exit Function
    End If
    On Error GoTo 0
    IncludeFile = True
End Function




Function ResolvePath(targetPath, defaultPath, mustExist)
    Call LogEvent("comm", "max", "ResolvePath starting", "ResolvePath", "Target: '" & targetPath & "'", "Default: '" & defaultPath & "', mustExist: " & mustExist)
    
    Dim fsoLocal, basePath, candidate, hasDefault, requireExists
    Set fsoLocal = CreateObject("Scripting.FileSystemObject")

    hasDefault = (Len(CStr(defaultPath)) > 0)
    requireExists = CBool(mustExist)

    If IsAbsolutePath(targetPath) Then
        candidate = targetPath
        Call LogEvent("comm", "max", "Using absolute path", "ResolvePath", "Candidate: '" & candidate & "'", "")
    Else
        basePath = GetBaseScriptPath()
        Call LogEvent("comm", "max", "Resolving relative path", "ResolvePath", "BasePath: '" & basePath & "'", "")
        If Len(basePath) > 0 Then
            On Error Resume Next
            candidate = fsoLocal.BuildPath(basePath, targetPath)
            If Err.Number <> 0 Then
                Call LogEvent("min", "med", "BuildPath failed", "ResolvePath", "Error: " & Err.Description, "")
                candidate = ""
                Err.Clear
            Else
                Call LogEvent("comm", "max", "BuildPath succeeded", "ResolvePath", "Candidate: '" & candidate & "'", "")
            End If
            On Error GoTo 0
        End If

        If Len(candidate) = 0 Then
            candidate = fsoLocal.GetAbsolutePathName(targetPath)
        End If
    End If

    If requireExists Then
        If Not PathExists(fsoLocal, candidate) Then
            Call LogEvent("min", "med", "Required path does not exist", "ResolvePath", "Candidate: '" & candidate & "'", "")
            If hasDefault Then
                Call LogEvent("comm", "high", "Using default path", "ResolvePath", "Default: '" & defaultPath & "'", "")
                ResolvePath = defaultPath
            Else
                Call LogEvent("comm", "high", "No default available, returning non-existent path", "ResolvePath", "Candidate: '" & candidate & "'", "")
                ResolvePath = candidate
            End If
            Set fsoLocal = Nothing
            Exit Function
        End If
    End If

    If hasDefault And Len(candidate) = 0 Then
        ResolvePath = defaultPath
    ElseIf Len(candidate) > 0 Then
        ResolvePath = candidate
    ElseIf hasDefault Then
        ResolvePath = defaultPath
    Else
        ResolvePath = targetPath
    End If

    Set fsoLocal = Nothing
End Function

Function PathExists(fs, pathValue)
    On Error Resume Next
    If Len(pathValue) = 0 Then
        PathExists = False
        Exit Function
    End If

    If fs.FileExists(pathValue) Then
        PathExists = True
    ElseIf fs.FolderExists(pathValue) Then
        PathExists = True
    Else
        PathExists = False
    End If
    On Error GoTo 0
End Function

Function IsAbsolutePath(pathValue)
    If Len(pathValue) >= 2 Then
        If Mid(pathValue, 2, 1) = ":" Then
            IsAbsolutePath = True
            Exit Function
        End If
        If Left(pathValue, 2) = "\\" Then
            IsAbsolutePath = True
            Exit Function
        End If
    End If
    IsAbsolutePath = False
End Function

Function GetBaseScriptPath()
    ' Use repo root from CDK_BASE / GetRepoRoot() via PathHelper
    ' This ensures paths resolve correctly regardless of deployment location
    Dim repoRoot
    On Error Resume Next
    repoRoot = GetRepoRoot()
    If Err.Number <> 0 Or Len(repoRoot) = 0 Then
        Call LogEvent("maj", "low", "GetBaseScriptPath: GetRepoRoot failed, falling back to PostFinalCharges subfolder", "GetBaseScriptPath", Err.Description, "")
        Err.Clear
        repoRoot = g_fso.BuildPath(g_fso.GetParentFolderName(WScript.ScriptFullName), "..")
    End If
    On Error GoTo 0
    
    g_BaseScriptPath = g_fso.BuildPath(repoRoot, "apps\post_final_charges")
    Call LogEvent("comm", "high", "GetBaseScriptPath resolved to: " & g_BaseScriptPath, "GetBaseScriptPath", "", "")
    GetBaseScriptPath = g_BaseScriptPath
End Function


'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** GetIniSetting
' **DATE CREATED:** 2025-11-28
' **AUTHOR:** Gemini
' 
' **FUNCTIONALITY:**
' Reads a specific value from the config.ini file.
' 
' **PARAMETERS:**
' section (String): The [SectionName] in the INI file.
' key (String): The Key=Value pair to find.
' defaultValue (Variant): The value to return if the key is not found.
'-----------------------------------------------------------------------------------
Function GetIniSetting(section, key, defaultValue)
    Dim fso, file, line, inSection, result, configPath
    result = defaultValue ' Start with the default value
    inSection = False
    
    ' config.ini is now in config/ subfolder relative to repo root
    On Error Resume Next
    configPath = g_fso.BuildPath(GetRepoRoot(), "config\config.ini")
    If Err.Number <> 0 Then
        configPath = ""
        Err.Clear
    End If
    On Error GoTo 0
    
    Call LogEvent("comm", "high", "Reading INI: " & configPath, "GetIniSetting", "", "")

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(configPath) Then
        Call LogEvent("comm", "high", "INI file not found at: " & configPath, "GetIniSetting", "", "")
        GetIniSetting = defaultValue
        Exit Function
    End If
    
    Set file = fso.OpenTextFile(configPath, 1)
    If Err.Number <> 0 Then
        Call LogEvent("min", "low", "Could not open INI file: " & configPath, "GetIniSetting", "Check file permissions", "")
        GetIniSetting = defaultValue
        Exit Function
    End If
    On Error GoTo 0

    Do While Not file.AtEndOfStream
        line = Trim(file.ReadLine)
        If Left(line, 1) = "[" And Right(line, 1) = "]" Then
            ' It's a section header
            Dim currentSection
            currentSection = LCase(Trim(Mid(line, 2, Len(line) - 2)))
            If currentSection = LCase(section) Then
                inSection = True
                Call LogEvent("comm", "high", "Entered section [" & section & "]", "GetIniSetting", "", "")
            ElseIf inSection Then
                ' We have passed the relevant section, so we can stop.
                Call LogEvent("comm", "max", "Exited section [" & section & "]", "GetIniSetting", "", "")
                Exit Do
            End If
        ElseIf inSection And InStr(line, "=") > 0 And Left(line, 1) <> ";" Then
            ' It's a key-value pair within the correct section (and not a comment)
            Dim parts, currentKey, currentValue
            parts = Split(line, "=", 2)
            currentKey = Trim(parts(0))
            currentValue = Trim(parts(1))
            If LCase(currentKey) = LCase(key) Then
                result = currentValue
                Call LogEvent("comm", "high", "Found key '" & key & "' with value '" & result & "'", "GetIniSetting", "", "")
                Exit Do ' Found the key, no need to read further
            End If
        End If
    Loop

    file.Close
    GetIniSetting = result
End Function

' Initialize required objects (FileSystemObject and BlueZone instance)

'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** InitializeObjects
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' Initializes global objects required for the script's operation.
' This includes creating the FileSystemObject for file operations and the
' BlueZone WhllObj for terminal interaction. It also triggers debug mode
' detection and configuration loading.
'-----------------------------------------------------------------------------------
Sub InitializeObjects()
    Call InitializeConfig
    Call DetermineDebugMode
    
    ' Check for test mode via environment variable
    g_IsTestMode = False
    On Error Resume Next
    Dim shell, testModeEnv
    Set shell = CreateObject("WScript.Shell")
    If Err.Number = 0 Then
        testModeEnv = LCase(Trim(shell.Environment("PROCESS")("PFC_TEST_MODE")))
        If Len(testModeEnv) = 0 Then testModeEnv = LCase(Trim(shell.Environment("USER")("PFC_TEST_MODE")))
        If Len(testModeEnv) = 0 Then testModeEnv = LCase(Trim(shell.Environment("SYSTEM")("PFC_TEST_MODE")))
        
        If testModeEnv = "1" Or testModeEnv = "true" Or testModeEnv = "yes" Then
            g_IsTestMode = True
        End If
    End If
    On Error GoTo 0
    
    If g_IsTestMode Then
        ' Include and use mock g_bzhao for testing
        Dim mockPath
        mockPath = ResolvePath("mocks\MockBzhao.vbs", "", True)
        If IncludeFile(mockPath) Then
            Set g_bzhao = New MockBzhao
            Call LogEvent("comm", "med", "Using MockBzhao for testing", "InitializeObjects", "", "")
            
            ' Setup initial test scenario
            g_bzhao.SetupTestScenario("basic_command_prompt")
        Else
            Call LogEvent("maj", "low", "Could not load MockBzhao.vbs for test mode", "InitializeObjects", "", "")
            g_IsTestMode = False
        End If
    End If
    
    If Not g_IsTestMode Then
        Set g_bzhao = CreateObject("BZWhll.WhllObj")
        If Err.Number <> 0 Then
            Call LogEvent("crit", "med", "Failed to create BZWhll.WhllObj", "InitializeObjects", Err.Description, "")
            Err.Clear
        End If
    End If

    ReDim g_DiagLogQueue(g_DiagLogQueueSize - 1)
End Sub


'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** InitializeConfig
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' **MODIFIED:** 2025-11-28 by Gemini
' 
' **FUNCTIONALITY:**
' Reads configuration settings from the 'config.ini' file, falling back to
' hardcoded defaults if the file or settings are not present.
'-----------------------------------------------------------------------------------
Sub InitializeConfig()
    ' Resolve g_BaseScriptPath dynamically using repo root
    g_BaseScriptPath = GetBaseScriptPath()
    
    ' --- Initialize Logging Configuration from INI file ---
    Dim criticalityValue, verbosityValue
    criticalityValue = LCase(GetIniSetting("Settings", "LogCriticality", "comm"))
    verbosityValue = LCase(GetIniSetting("PostFinalCharges", "log_verbosity", "low"))
    
    ' Set criticality threshold
    Select Case criticalityValue
        Case "crit": g_CurrentCriticality = CRIT_CRITICAL
        Case "maj": g_CurrentCriticality = CRIT_MAJOR
        Case "min": g_CurrentCriticality = CRIT_MINOR
        Case "comm": g_CurrentCriticality = CRIT_COMMON
        Case Else: g_CurrentCriticality = CRIT_COMMON
    End Select
    
    ' Set verbosity threshold
    Select Case verbosityValue
        Case "low": g_CurrentVerbosity = VERB_LOW
        Case "med": g_CurrentVerbosity = VERB_MEDIUM
        Case "high": g_CurrentVerbosity = VERB_HIGH
        Case "max": g_CurrentVerbosity = VERB_MAX
        Case Else: g_CurrentVerbosity = VERB_MEDIUM
    End Select

    ' --- Now load other settings ---
    g_DefaultWait = GetIniSetting("Settings", "DefaultWait", 1000)
    g_PromptWait = GetIniSetting("Settings", "PromptWait", 5000)
    g_DebugDelayFactor = GetIniSetting("Settings", "DebugDelayFactor", 1.0)
    
    Dim startSequenceNumberValue, endSequenceNumberValue
    startSequenceNumberValue = GetIniSetting("PostFinalCharges", "StartSequenceNumber", "")
    endSequenceNumberValue = GetIniSetting("PostFinalCharges", "EndSequenceNumber", "")

    If startSequenceNumberValue = "" Or endSequenceNumberValue = "" Then
        Call LogEvent("crit", "low", "Critical config missing: 'StartSequenceNumber' and/or 'EndSequenceNumber' not found in config.ini", "InitializeConfig", "Script cannot continue without sequence range", "")
        g_ShouldAbort = True
        g_StartSequenceNumber = 1 ' Set loop to be non-executing
        g_EndSequenceNumber = 0
    Else
        g_StartSequenceNumber = CInt(startSequenceNumberValue)
        g_EndSequenceNumber = CInt(endSequenceNumberValue)
    End If

    ' --- Deprecated settings, kept for compatibility ---
    CSV_FILE_PATH = GetConfigPath("PostFinalCharges", "CSV")
    LOG_FILE_PATH = GetConfigPath("PostFinalCharges", "Log")

    Dim overwriteLogOnStartValue
    overwriteLogOnStartValue = LCase(Trim(GetIniSetting("PostFinalCharges", "OverwriteLogOnStart", "false")))
    g_OverwriteLogOnStart = False
    If overwriteLogOnStartValue = "1" Or overwriteLogOnStartValue = "true" Or overwriteLogOnStartValue = "yes" Then
        g_OverwriteLogOnStart = True
    End If
    Call ApplyLogStartupMode()

    g_LongWait = 2000
    g_SendRetryCount = 2
    g_DelayBetweenTextAndEnterMs = 2000
    POST_PROMPT_WAIT_MS = Int(1000 * g_DebugDelayFactor)  ' Base 1000ms scaled by debug delay factor
    Dim closeoutDelayValue
    closeoutDelayValue = GetIniSetting("PostFinalCharges", "CloseoutConfirmDelayMs", "1200")
    On Error Resume Next
    g_CloseoutConfirmDelayMs = CInt(closeoutDelayValue)
    If Err.Number <> 0 Or g_CloseoutConfirmDelayMs < 0 Then
        g_CloseoutConfirmDelayMs = 1200
        Err.Clear
    End If
    On Error GoTo 0

    Dim stabilityPauseValue
    stabilityPauseValue = GetIniSetting("PostFinalCharges", "StabilityPause", "1000")
    On Error Resume Next
    g_StabilityPause = CInt(stabilityPauseValue)
    If Err.Number <> 0 Or g_StabilityPause < 0 Then
        g_StabilityPause = 1000
        Err.Clear
    End If
    On Error GoTo 0

    g_BlacklistTermsRaw = GetIniSetting("PostFinalCharges", "blacklist_terms", "")
    If Not LoadCloseoutTriggers(GetConfigPath("PostFinalCharges", "TriggerList"), g_CloseoutTriggers) Then
        g_ShouldAbort = True
    End If
    g_SkipRoListRaw = GetIniSetting("PostFinalCharges", "SkipRoList", "")
    Set g_SkipRoLookup = CreateObject("Scripting.Dictionary")
    If Not LoadSkipRoLookup(g_SkipRoListRaw, g_SkipRoLookup) Then
        g_ShouldAbort = True
    End If

    g_EnableDiagnosticLogging = False
    DIAGNOSTIC_LOG_PATH = GetConfigPath("PostFinalCharges", "DiagnosticLog")

    Dim olderRoThresholdValue
    olderRoThresholdValue = GetIniSetting("PostFinalCharges", "OlderRoThresholdDays", "30")
    On Error Resume Next
    g_OlderRoThresholdDays = CInt(olderRoThresholdValue)
    If Err.Number <> 0 Or g_OlderRoThresholdDays < 0 Then
        g_OlderRoThresholdDays = 30
        Err.Clear
    End If
    On Error GoTo 0
    Dim olderRoStatusesRaw, olderStatusArr, si
    olderRoStatusesRaw = GetIniSetting("PostFinalCharges", "OlderRoStatuses", "OPENED,OPEN")
    olderStatusArr = Split(olderRoStatusesRaw, ",")
    For si = 0 To UBound(olderStatusArr)
        olderStatusArr(si) = UCase(Trim(olderStatusArr(si)))
    Next
    g_OlderRoStatuses = olderStatusArr
    Call LogEvent("comm", "high", "Older RO threshold configured", "InitializeConfig", "Days: " & g_OlderRoThresholdDays & " Statuses: " & olderRoStatusesRaw, "")

    g_EmployeeNumber = GetIniSetting("PostFinalCharges", "EmployeeNumber", "")
    g_EmployeeNameConfirm = GetIniSetting("PostFinalCharges", "EmployeeNameConfirm", "")

    Dim skipWchValue
    skipWchValue = LCase(Trim(GetIniSetting("PostFinalCharges", "SkipWchLabor", "true")))
    g_SkipWchEnabled = (skipWchValue = "true" Or skipWchValue = "1" Or skipWchValue = "yes")
    Dim wchGateState : wchGateState = "disabled"
    If g_SkipWchEnabled Then wchGateState = "enabled"
    Call LogEvent("comm", "high", "WCH skip gate: " & wchGateState, "InitializeConfig", "", "")

    Dim partsOrderKeywordsRaw, partsOrderNegatorsRaw, ki
    partsOrderKeywordsRaw = GetIniSetting("PostFinalCharges", "PartsOrderKeywords", "")
    partsOrderNegatorsRaw = GetIniSetting("PostFinalCharges", "PartsOrderNegators", "")
    g_PartsOrderKeywords = Split(partsOrderKeywordsRaw, ",")
    g_PartsOrderNegators = Split(partsOrderNegatorsRaw, ",")
    For ki = 0 To UBound(g_PartsOrderKeywords)
        g_PartsOrderKeywords(ki) = LCase(Trim(g_PartsOrderKeywords(ki)))
    Next
    For ki = 0 To UBound(g_PartsOrderNegators)
        g_PartsOrderNegators(ki) = LCase(Trim(g_PartsOrderNegators(ki)))
    Next
    Call LogEvent("comm", "high", "Parts-order keyword scan configured", "InitializeConfig", "Keywords: " & partsOrderKeywordsRaw & " | Negators: " & partsOrderNegatorsRaw, "")

    Dim allowedTechCodesRaw, ti
    allowedTechCodesRaw = GetIniSetting("PostFinalCharges", "AllowedTechCodes", "C92,C93")
    g_AllowedTechCodes = Split(allowedTechCodesRaw, ",")
    For ti = 0 To UBound(g_AllowedTechCodes)
        g_AllowedTechCodes(ti) = UCase(Trim(g_AllowedTechCodes(ti)))
    Next
    Call LogEvent("comm", "high", "Closure readiness tech-code gate configured", "InitializeConfig", "AllowedTechCodes: " & allowedTechCodesRaw, "")

    Dim cdkLaborExceptionRaw, ei
    cdkLaborExceptionRaw = GetIniSetting("PostFinalCharges", "CDKLaborOnlyLTypeExceptions", "WCH,WT,WF")
    g_arrCDKExceptions = Split(cdkLaborExceptionRaw, ",")
    For ei = 0 To UBound(g_arrCDKExceptions)
        g_arrCDKExceptions(ei) = UCase(Trim(g_arrCDKExceptions(ei)))
    Next
    Call LogEvent("comm", "high", "CDK labor-only exception LTYPE codes configured", "InitializeConfig", "CDKLaborOnlyLTypeExceptions: " & cdkLaborExceptionRaw, "")

    Dim cdkLaborDescExceptionRaw, di
    cdkLaborDescExceptionRaw = GetIniSetting("PostFinalCharges", "CDKLaborOnlyDescriptionExceptions", "check and adjust")
    g_arrCDKDescriptionExceptions = Split(cdkLaborDescExceptionRaw, ",")
    For di = 0 To UBound(g_arrCDKDescriptionExceptions)
        g_arrCDKDescriptionExceptions(di) = LCase(Trim(g_arrCDKDescriptionExceptions(di)))
    Next
    Call LogEvent("comm", "high", "CDK labor-only description exceptions configured", "InitializeConfig", "CDKLaborOnlyDescriptionExceptions: " & cdkLaborDescExceptionRaw, "")
End Sub

Sub ApplyLogStartupMode()
    If Not g_OverwriteLogOnStart Then Exit Sub

    Dim logFSO, logFile, logFolder
    Set logFSO = CreateObject("Scripting.FileSystemObject")
    logFolder = logFSO.GetParentFolderName(LOG_FILE_PATH)
    If Len(logFolder) > 0 Then
        Call EnsureFolderExists(logFSO, logFolder)
    End If

    On Error Resume Next
    Set logFile = logFSO.OpenTextFile(LOG_FILE_PATH, 2, True)
    If Err.Number <> 0 Then
        Err.Clear
        If LOG_FILE_PATH <> LEGACY_LOG_PATH Then
            LOG_FILE_PATH = LEGACY_LOG_PATH
            logFolder = logFSO.GetParentFolderName(LOG_FILE_PATH)
            If Len(logFolder) > 0 Then
                Call EnsureFolderExists(logFSO, logFolder)
            End If
            Set logFile = logFSO.OpenTextFile(LOG_FILE_PATH, 2, True)
        End If
    End If

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Call LogEvent("maj", "low", "Failed to overwrite log file at startup", "ApplyLogStartupMode", LOG_FILE_PATH, "")
        Exit Sub
    End If

    logFile.Close
    Set logFile = Nothing
    Set logFSO = Nothing
    On Error GoTo 0

    g_SessionDateLogged = False
    Call WriteSessionHeader()
    Call LogEvent("comm", "low", "Log file overwritten at startup per config toggle", "ApplyLogStartupMode", "", "")
End Sub


'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** DetermineDebugMode
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' **MODIFIED:** 2025-11-28 by Gemini
' 
' **FUNCTIONALITY:**
' Overrides configuration settings with environment variables for debugging purposes.
' This allows for dynamic changes without editing the config file.
'-----------------------------------------------------------------------------------
Sub DetermineDebugMode()
    ' This function handles environment variable overrides for dual-axis logging.
    ' Default criticality and verbosity are set in InitializeConfig from INI file.
    
    DEBUG_LOGGING = False
    g_EnableDetailedLogging = False

    On Error Resume Next
    Dim shell, envValue
    Set shell = CreateObject("WScript.Shell")
    If Err.Number <> 0 Then Exit Sub ' Cannot proceed without shell object

    ' --- PFC_DEBUG override ---
    ' Enables maximum logging for deep diagnostics.
    envValue = LCase(Trim(shell.Environment("PROCESS")("PFC_DEBUG")))
    If Len(envValue) = 0 Then envValue = LCase(Trim(shell.Environment("USER")("PFC_DEBUG")))
    If Len(envValue) = 0 Then envValue = LCase(Trim(shell.Environment("SYSTEM")("PFC_DEBUG")))

    If envValue = "1" Or envValue = "true" Or envValue = "yes" Then
        DEBUG_LOGGING = True
        g_EnableDetailedLogging = True
        g_CurrentCriticality = CRIT_COMMON ' Log all events
        g_CurrentVerbosity = VERB_MAX ' Maximum detail
    End If
    
    ' --- PFC_LOG_CRITICALITY override ---
    Dim criticalityEnvValue
    criticalityEnvValue = LCase(Trim(shell.Environment("PROCESS")("PFC_LOG_CRITICALITY")))
    If Len(criticalityEnvValue) = 0 Then criticalityEnvValue = LCase(Trim(shell.Environment("USER")("PFC_LOG_CRITICALITY")))
    If Len(criticalityEnvValue) = 0 Then criticalityEnvValue = LCase(Trim(shell.Environment("SYSTEM")("PFC_LOG_CRITICALITY")))
    
    If Len(criticalityEnvValue) > 0 Then
        Select Case criticalityEnvValue
            Case "crit": g_CurrentCriticality = CRIT_CRITICAL
            Case "maj": g_CurrentCriticality = CRIT_MAJOR
            Case "min": g_CurrentCriticality = CRIT_MINOR
            Case "comm": g_CurrentCriticality = CRIT_COMMON
        End Select
    End If
    
    ' --- PFC_LOG_VERBOSITY override ---
    Dim verbosityEnvValue
    verbosityEnvValue = LCase(Trim(shell.Environment("PROCESS")("PFC_LOG_VERBOSITY")))
    If Len(verbosityEnvValue) = 0 Then verbosityEnvValue = LCase(Trim(shell.Environment("USER")("PFC_LOG_VERBOSITY")))
    If Len(verbosityEnvValue) = 0 Then verbosityEnvValue = LCase(Trim(shell.Environment("SYSTEM")("PFC_LOG_VERBOSITY")))
    
    If Len(verbosityEnvValue) > 0 Then
        Select Case verbosityEnvValue
            Case "low": g_CurrentVerbosity = VERB_LOW
            Case "med": g_CurrentVerbosity = VERB_MEDIUM
            Case "high": g_CurrentVerbosity = VERB_HIGH
            Case "max": g_CurrentVerbosity = VERB_MAX
        End Select
    End If

    Set shell = Nothing
    On Error GoTo 0
End Sub


'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** ConnectBlueZone
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' Establishes a connection to the BlueZone terminal emulator session.
' It uses the global g_bzhao object and attempts to connect to the default
' session. It logs the outcome of the connection attempt.
' 
' 
' **RETURN VALUE:**
' (Boolean) Returns True if the connection is successful, False otherwise.
'-----------------------------------------------------------------------------------
Function ConnectBlueZone()
    On Error Resume Next
    If g_bzhao Is Nothing Then
        Call LogEvent("crit", "low", "BlueZone object is not available", "ConnectBlueZone", "CreateObject failed", "")
        ConnectBlueZone = False
        Exit Function
    End If
    
    g_bzhao.Connect ""
    If Err.Number <> 0 Then
        Call LogEvent("crit", "med", "BlueZone connection failed", "ConnectBlueZone", Err.Description, "")
        Err.Clear
        ConnectBlueZone = False
    Else
        Call LogEvent("comm", "med", "Connected to BlueZone", "ConnectBlueZone", "", "")
        ConnectBlueZone = True
    End If
    On Error GoTo 0
End Function




'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** ProcessRONumbers
' **DATE CREATED:** 2025-11-07
' **AUTHOR:** Gemini
' 
' 
' **FUNCTIONALITY:**
' Iterates through a range of Repair Order (RO) numbers. For each RO,
' it orchestrates the main processing logic by calling the Main subroutine.
' It logs progress and final results for each RO.
'-----------------------------------------------------------------------------------
Sub ProcessRONumbers()
    If g_ShouldAbort Then Exit Sub ' Exit if config is invalid

    ' Ensure we are at a clean COMMAND: prompt before starting the session
    ' This addresses the "E before prompt" bug on the first sequence
    Call LogEvent("comm", "low", "Ensuring terminal is at COMMAND: prompt before starting", "ProcessRONumbers", "", "")
    If Not WaitForPrompt("COMMAND:", "", False, 10000, "Initial Command Prompt") Then
        Call LogEvent("crit", "low", "Terminal not at COMMAND: prompt. Manual intervention required.", "ProcessRONumbers", "", "Stopping script - fail fast requested")
        g_ShouldAbort = True
        Exit Sub
    End If

    Dim roNumber
    Dim lineCount
    Dim sequenceLabel
    lineCount = 0
    g_ReviewedROCount = 0
    g_FiledROCount = 0
    g_SkipBlacklistCount = 0
    g_SkipStatusOpenCount = 0
    g_SkipStatusPreassignedCount = 0
    g_SkipStatusOtherCount = 0
    g_SkipConfiguredCount = 0
    g_SkipWarrantyCount = 0
    g_SkipPartsOrderNeededCount = 0
    g_SkipTechCodeCount = 0
    g_ClosedRoCount = 0
    g_NotOnFileRoCount = 0
    g_SkipVehidNotOnFileCount = 0
    g_SkipNoCloseoutTextCount = 0
    g_SkipNoPartsChargedCount = 0
    g_LeftOpenManualCount = 0
    g_FcaMissingPartFlagCount = 0
    g_FcaHandlerNotConfiguredCount = 0
    g_ErrorInMainCount = 0
    g_NoResultRecordedCount = 0
    g_SummaryOtherOutcomeCount = 0
    Set g_SummaryOtherOutcomeBreakdown = CreateObject("Scripting.Dictionary")
    Set g_SummaryOtherOutcomeRawBreakdown = CreateObject("Scripting.Dictionary")
    g_OlderRoAttemptCount = 0
    g_OlderRoFiledCount = 0
    Set g_SkipOtherStates = CreateObject("Scripting.Dictionary")
    g_PreviousNormalizedRo = ""
    g_PreviousSequenceNumber = ""
    
    ' In test mode, only process one RO
    If g_IsTestMode Then
        roNumber = 900
        Call LogROHeader(roNumber)
        sequenceLabel = "Sequence " & roNumber
        
        lastRoResult = ""
        Call Main(roNumber)
        g_ReviewedROCount = 1
        If InStr(1, lastRoResult, "Successfully filed", vbTextCompare) > 0 Then
            g_FiledROCount = g_FiledROCount + 1
        End If
        If Len(Trim(CStr(lastRoResult))) = 0 Then lastRoResult = "No result recorded"
        Call TrackPrimaryOutcomeCounters(lastRoResult)
        Call TrackOtherOutcome(lastRoResult)
        
        Call LogEvent("comm", "med", sequenceLabel & " - Result: " & lastRoResult, "ProcessRONumbers", "", "")
        Call LogEvent("comm", "med", "Test mode: Processed single RO " & roNumber, "ProcessRONumbers", "", "")
        Exit Sub
    End If

    For roNumber = g_StartSequenceNumber To g_EndSequenceNumber
        lineCount = lineCount + 1
        g_ReviewedROCount = lineCount
        'WaitMs(2000)
        Call LogROHeader(roNumber)
        sequenceLabel = "Sequence " & roNumber

        ' Start performance timing for this RO
        Dim roStartTime
        roStartTime = Now

        lastRoResult = ""
        Call Main(roNumber)
        If InStr(1, lastRoResult, "Successfully filed", vbTextCompare) > 0 Then
            g_FiledROCount = g_FiledROCount + 1
        End If

        ' Check for end-of-sequence error
        If IsTextPresent("SEQUENCE NUMBER " & roNumber & " DOES NOT EXIST") Then
            Call LogEvent("maj", "low", "End of sequence detected", "ProcessRONumbers", "SEQUENCE NUMBER " & roNumber & " DOES NOT EXIST", "Stopping script")
            Exit Sub
        End If

        ' Calculate and log performance timing for successful closures only
        If InStr(1, LCase(lastRoResult), "successfully closed") > 0 Then
            Dim roEndTime, roDuration
            roEndTime = Now
            roDuration = DateDiff("s", roStartTime, roEndTime)
            Call LogEvent("comm", "med", sequenceLabel & " - E2E Duration: " & roDuration & " seconds", "ProcessRONumbers", "", "")
        End If

        If Err.Number <> 0 Then
            lastRoResult = "Error in Main: " & Err.Description
            ' Prefer the scraped RO for error/result logging when available
            Dim displayId
            If Len(Trim(CStr(currentRODisplay))) > 0 Then
                displayId = currentRODisplay
            Else
                displayId = roNumber
            End If
            Dim errorLabel
            errorLabel = sequenceLabel
            If Len(Trim(CStr(displayId))) > 0 And CStr(displayId) <> CStr(roNumber) Then
                errorLabel = errorLabel & " (RO " & displayId & ")"
            End If
            Call LogEvent("maj", "low", errorLabel & " - " & lastRoResult, "ProcessRONumbers", "", "")
            Err.Clear
        End If

        ' Ensure there's always a final result logged for the RO
        If Len(Trim(CStr(lastRoResult))) = 0 Then lastRoResult = "No result recorded"
        Call TrackPrimaryOutcomeCounters(lastRoResult)
        Call TrackOtherOutcome(lastRoResult)
        Dim finalMessage
        finalMessage = sequenceLabel & " - Result: " & lastRoResult
        Call LogEvent("comm", "low", finalMessage, "ProcessRONumbers", "", "")
        ' Always write the scraped RO status to the core log for troubleshooting
        Dim statusForLog
        statusForLog = Trim(CStr(g_LastScrapedStatus))
        If Len(statusForLog) = 0 Then statusForLog = "(none)"
        Call LogEvent("comm", "low", "RO STATUS FOUND: " & statusForLog, "ProcessRONumbers", "", "")

        If (lineCount Mod 10) = 0 Then
            Call LogEvent("comm", "med", "Processed " & lineCount & " ROs...", "ProcessRONumbers", "", "")
        End If

        If g_ShouldAbort Then
            Call LogEvent("crit", "low", "Aborting sequence processing", "ProcessRONumbers", "Reason: " & g_AbortReason, "")
            Exit Sub
        End If
    Next
End Sub

Sub TrackPrimaryOutcomeCounters(resultText)
    Dim normalized
    normalized = UCase(Trim(CStr(resultText)))

    If Len(normalized) = 0 Then Exit Sub

    If normalized = "CLOSED" Then
        g_ClosedRoCount = g_ClosedRoCount + 1
        Exit Sub
    End If
    If normalized = "NOT ON FILE" Then
        g_NotOnFileRoCount = g_NotOnFileRoCount + 1
        Exit Sub
    End If
    If InStr(1, normalized, "SKIPPED - VEHID NOT ON FILE", vbTextCompare) = 1 Then
        g_SkipVehidNotOnFileCount = g_SkipVehidNotOnFileCount + 1
        Exit Sub
    End If
    If InStr(1, normalized, "SKIPPED - NO CLOSEOUT TEXT FOUND", vbTextCompare) = 1 Then
        g_SkipNoCloseoutTextCount = g_SkipNoCloseoutTextCount + 1
        Exit Sub
    End If
    If InStr(1, normalized, "SKIPPED - NO PARTS CHARGED", vbTextCompare) = 1 Then
        g_SkipNoPartsChargedCount = g_SkipNoPartsChargedCount + 1
        Exit Sub
    End If
    If InStr(1, normalized, "LEFT OPEN FOR MANUAL CLOSING", vbTextCompare) = 1 Then
        g_LeftOpenManualCount = g_LeftOpenManualCount + 1
        Exit Sub
    End If
    If InStr(1, normalized, "FLAGGED - MISSING PART NUMBER FOR FCA DIALOG", vbTextCompare) = 1 Then
        g_FcaMissingPartFlagCount = g_FcaMissingPartFlagCount + 1
        Exit Sub
    End If
    If InStr(1, normalized, "SKIPPED - FCA DIALOG HANDLER NOT YET CONFIGURED", vbTextCompare) = 1 Then
        g_FcaHandlerNotConfiguredCount = g_FcaHandlerNotConfiguredCount + 1
        Exit Sub
    End If
    If InStr(1, normalized, "ERROR IN MAIN:", vbTextCompare) = 1 Then
        g_ErrorInMainCount = g_ErrorInMainCount + 1
        Exit Sub
    End If
    If InStr(1, normalized, "NO RESULT RECORDED", vbTextCompare) = 1 Then
        g_NoResultRecordedCount = g_NoResultRecordedCount + 1
        Exit Sub
    End If
End Sub

Sub TrackOtherOutcome(resultText)
    Dim bucket
    Dim rawKey

    If IsResultRepresentedInSummary(resultText) Then Exit Sub

    bucket = GetOtherOutcomeBucket(resultText)
    If Len(Trim(CStr(bucket))) = 0 Then bucket = "Other/Unknown"

    g_SummaryOtherOutcomeCount = g_SummaryOtherOutcomeCount + 1

    If Not IsObject(g_SummaryOtherOutcomeBreakdown) Then
        Set g_SummaryOtherOutcomeBreakdown = CreateObject("Scripting.Dictionary")
    End If

    If g_SummaryOtherOutcomeBreakdown.Exists(bucket) Then
        g_SummaryOtherOutcomeBreakdown(bucket) = CLng(g_SummaryOtherOutcomeBreakdown(bucket)) + 1
    Else
        g_SummaryOtherOutcomeBreakdown.Add bucket, 1
    End If

    rawKey = BuildRawOtherOutcomeKey(resultText)
    If Len(Trim(CStr(rawKey))) = 0 Then rawKey = "(blank result)"

    If Not IsObject(g_SummaryOtherOutcomeRawBreakdown) Then
        Set g_SummaryOtherOutcomeRawBreakdown = CreateObject("Scripting.Dictionary")
    End If

    If g_SummaryOtherOutcomeRawBreakdown.Exists(rawKey) Then
        g_SummaryOtherOutcomeRawBreakdown(rawKey) = CLng(g_SummaryOtherOutcomeRawBreakdown(rawKey)) + 1
    Else
        g_SummaryOtherOutcomeRawBreakdown.Add rawKey, 1
    End If
End Sub

Function BuildRawOtherOutcomeKey(resultText)
    Dim keyText
    keyText = Trim(CStr(resultText))

    If Len(keyText) = 0 Then
        BuildRawOtherOutcomeKey = ""
        Exit Function
    End If

    ' Normalize whitespace and cap size for MsgBox readability.
    keyText = Replace(keyText, vbCr, " ")
    keyText = Replace(keyText, vbLf, " ")
    Do While InStr(1, keyText, "  ", vbTextCompare) > 0
        keyText = Replace(keyText, "  ", " ")
    Loop

    If Len(keyText) > 90 Then
        keyText = Left(keyText, 87) & "..."
    End If

    BuildRawOtherOutcomeKey = keyText
End Function

Function GetOtherOutcomeBucket(resultText)
    Dim normalized
    normalized = UCase(Trim(CStr(resultText)))

    GetOtherOutcomeBucket = "Other/Unknown"
    If Len(normalized) = 0 Then Exit Function

    If normalized = "CLOSED" Then
        GetOtherOutcomeBucket = "Already Closed"
        Exit Function
    End If
    If normalized = "NOT ON FILE" Then
        GetOtherOutcomeBucket = "Not On File"
        Exit Function
    End If
    If InStr(1, normalized, "SKIPPED - VEHID NOT ON FILE", vbTextCompare) = 1 Then
        GetOtherOutcomeBucket = "Skipped - VEHID not on file"
        Exit Function
    End If
    If InStr(1, normalized, "SKIPPED - NO CLOSEOUT TEXT FOUND", vbTextCompare) = 1 Then
        GetOtherOutcomeBucket = "Skipped - No closeout text"
        Exit Function
    End If
    If InStr(1, normalized, "SKIPPED - NO PARTS CHARGED", vbTextCompare) = 1 Then
        GetOtherOutcomeBucket = "Skipped - No parts charged"
        Exit Function
    End If
    If InStr(1, normalized, "LEFT OPEN FOR MANUAL CLOSING", vbTextCompare) = 1 Then
        GetOtherOutcomeBucket = "Left open for manual closing"
        Exit Function
    End If
    If InStr(1, normalized, "FLAGGED - MISSING PART NUMBER FOR FCA DIALOG", vbTextCompare) = 1 Then
        GetOtherOutcomeBucket = "Flagged - Missing FCA part number"
        Exit Function
    End If
    If InStr(1, normalized, "SKIPPED - FCA DIALOG HANDLER NOT YET CONFIGURED", vbTextCompare) = 1 Then
        GetOtherOutcomeBucket = "Skipped - FCA dialog handler not configured"
        Exit Function
    End If
    If InStr(1, normalized, "ERROR IN MAIN:", vbTextCompare) = 1 Then
        GetOtherOutcomeBucket = "Error in Main"
        Exit Function
    End If
    If InStr(1, normalized, "NO RESULT RECORDED", vbTextCompare) = 1 Then
        GetOtherOutcomeBucket = "No result recorded"
        Exit Function
    End If

    GetOtherOutcomeBucket = "Other: " & Left(Trim(CStr(resultText)), 60)
End Function

Function BuildOtherOutcomeBreakdown(maxLines)
    Dim key, countValue, linesAdded, hiddenCategories
    Dim output

    If maxLines <= 0 Then maxLines = 8
    output = ""
    linesAdded = 0
    hiddenCategories = 0

    If Not IsObject(g_SummaryOtherOutcomeBreakdown) Then
        BuildOtherOutcomeBreakdown = ""
        Exit Function
    End If

    For Each key In g_SummaryOtherOutcomeBreakdown.Keys
        If linesAdded < maxLines Then
            countValue = CLng(g_SummaryOtherOutcomeBreakdown(key))
            output = output & "  - " & CStr(key) & ": " & CStr(countValue) & vbCrLf
            linesAdded = linesAdded + 1
        Else
            hiddenCategories = hiddenCategories + 1
        End If
    Next

    If hiddenCategories > 0 Then
        output = output & "  - (+" & hiddenCategories & " more categories)"
    ElseIf Len(output) > 0 Then
        output = Left(output, Len(output) - Len(vbCrLf))
    End If

    BuildOtherOutcomeBreakdown = output
End Function

Function BuildOtherOutcomeRawBreakdown(maxLines)
    Dim key, countValue, linesAdded, hiddenCategories
    Dim output

    If maxLines <= 0 Then maxLines = 12
    output = ""
    linesAdded = 0
    hiddenCategories = 0

    If Not IsObject(g_SummaryOtherOutcomeRawBreakdown) Then
        BuildOtherOutcomeRawBreakdown = ""
        Exit Function
    End If

    For Each key In g_SummaryOtherOutcomeRawBreakdown.Keys
        If linesAdded < maxLines Then
            countValue = CLng(g_SummaryOtherOutcomeRawBreakdown(key))
            output = output & "  - " & CStr(key) & ": " & CStr(countValue) & vbCrLf
            linesAdded = linesAdded + 1
        Else
            hiddenCategories = hiddenCategories + 1
        End If
    Next

    If hiddenCategories > 0 Then
        output = output & "  - (+" & hiddenCategories & " more raw results)"
    ElseIf Len(output) > 0 Then
        output = Left(output, Len(output) - Len(vbCrLf))
    End If

    BuildOtherOutcomeRawBreakdown = output
End Function

Function IsResultRepresentedInSummary(resultText)
    Dim normalized
    normalized = UCase(Trim(CStr(resultText)))

    IsResultRepresentedInSummary = False
    If Len(normalized) = 0 Then Exit Function

    If InStr(1, normalized, "SUCCESSFULLY FILED", vbTextCompare) > 0 Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If

    If InStr(1, normalized, "SKIPPED - CONFIGURED RO SKIP LIST", vbTextCompare) = 1 Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
    If InStr(1, normalized, "SKIPPED - WCH LABOR TYPE", vbTextCompare) = 1 Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
    If InStr(1, normalized, "SKIPPED - NON-COMPLIANT TECH CODE:", vbTextCompare) = 1 Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
    If InStr(1, normalized, "SKIPPED - PARTS ORDER NEEDED:", vbTextCompare) = 1 Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
    If InStr(1, normalized, "SKIPPED - BLACKLISTED TERM:", vbTextCompare) = 1 Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
    If InStr(1, normalized, "SKIPPED - STATUS NOT READY", vbTextCompare) = 1 Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
    If normalized = "CLOSED" Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
    If normalized = "NOT ON FILE" Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
    If InStr(1, normalized, "SKIPPED - VEHID NOT ON FILE", vbTextCompare) = 1 Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
    If InStr(1, normalized, "SKIPPED - NO CLOSEOUT TEXT FOUND", vbTextCompare) = 1 Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
    If InStr(1, normalized, "SKIPPED - NO PARTS CHARGED", vbTextCompare) = 1 Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
    If InStr(1, normalized, "LEFT OPEN FOR MANUAL CLOSING", vbTextCompare) = 1 Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
    If InStr(1, normalized, "FLAGGED - MISSING PART NUMBER FOR FCA DIALOG", vbTextCompare) = 1 Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
    If InStr(1, normalized, "SKIPPED - FCA DIALOG HANDLER NOT YET CONFIGURED", vbTextCompare) = 1 Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
    If InStr(1, normalized, "ERROR IN MAIN:", vbTextCompare) = 1 Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
    If InStr(1, normalized, "NO RESULT RECORDED", vbTextCompare) = 1 Then
        IsResultRepresentedInSummary = True
        Exit Function
    End If
End Function


'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** LogROHeader
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' Writes a formatted header section to the log file for a given Repair Order (RO)
' number. This improves log readability by clearly separating the log entries
' for each RO being processed.
' 
' 
' **PARAMETERS:**
' **ro** (String): The Repair Order number to include in the log header.
'-----------------------------------------------------------------------------------
Sub LogROHeader(ro)
    Dim logFSO, logFile, sep
    sep = "==================="
    Set logFSO = CreateObject("Scripting.FileSystemObject")
    Dim logFolder
    logFolder = logFSO.GetParentFolderName(LOG_FILE_PATH)
    If Len(logFolder) > 0 Then
        Call EnsureFolderExists(logFSO, logFolder)
    End If
    On Error Resume Next
    Set logFile = logFSO.OpenTextFile(LOG_FILE_PATH, 8, True)
    If Err.Number <> 0 Then
        Err.Clear
        If LOG_FILE_PATH <> LEGACY_LOG_PATH Then
            LOG_FILE_PATH = LEGACY_LOG_PATH
        End If
        Set logFile = logFSO.OpenTextFile(LOG_FILE_PATH, 8, True)
        If Err.Number <> 0 Then
            Err.Clear
            ' If both attempts fail, exit the function gracefully
            Call LogEvent("maj", "low", "Failed to open log file", "LogROHeader", LOG_FILE_PATH, "")
            Exit Sub
        End If
    End If
    On Error GoTo 0
    logFile.WriteLine sep
    logFile.WriteLine "Sequence: " & CStr(ro)
    logFile.WriteLine sep
    logFile.Close
    Set logFile = Nothing
    Set logFSO = Nothing
End Sub


'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** Main
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' This is the core processing logic for a single Repair Order (RO). It enters
' the RO number, checks its status (e.g., "closed", "NOT ON FILE"), and verifies
' if it is "READY TO POST". If ready, it looks for specific trigger text on the
' screen to determine if it should proceed with the closeout process.
' 
' 
' **PARAMETERS:**
' **roNumber** (String): The Repair Order number to process.
'-----------------------------------------------------------------------------------
Sub Main(roNumber)
    ' Enter the number and send Enter (use single helper)
    Dim send_enter_key

    send_enter_key = True
    Call WaitForPrompt("COMMAND:", roNumber, send_enter_key, 10000, "")
    
    ' Wait for the new sequence's RO to appear on screen.
    ' WaitForScreenTransition("RO STATUS:") would return immediately if the previous
    ' sequence's screen is still showing, so we poll until the displayed RO changes.
    Dim actualRO, roWaitStart, roWaitElapsed, roCandidate, roCandidateNorm, mainPromptText
    roWaitStart = Timer
    actualRO = ""
    Do
        mainPromptText = GetScreenLine(MainPromptLine)
        If InStr(1, mainPromptText, "Process is locked by", vbTextCompare) > 0 Then
            currentRODisplay = roNumber
            Call LogEvent("comm", "med", "RO locked by another user during RO load - returning to command", "Main", "RO: " & currentRODisplay, "Line " & MainPromptLine & ": '" & mainPromptText & "'")
            Call FastKey("<Enter>")
            Call WaitForPrompt("COMMAND:", "", False, 5000, "Process Lock Recovery")
            lastRoResult = "Skipped - Process locked by another user"
            Exit Sub
        End If

        roCandidate = GetROFromScreen()
        roCandidateNorm = NormalizeRoIdentifier(roCandidate)
        If Len(roCandidateNorm) > 0 And roCandidateNorm <> g_PreviousNormalizedRo Then
            actualRO = roCandidate
            Exit Do
        End If
        Call WaitMs(150)
        roWaitElapsed = (Timer - roWaitStart) * 1000
    Loop While roWaitElapsed < 15000

    If Len(actualRO) = 0 Then
        Call LogEvent("maj", "low", "Screen did not show a new RO within 15s; reading current screen", "Main", "", "")
        actualRO = GetROFromScreen()
    End If

    If Len(Trim(CStr(actualRO))) > 0 Then
        currentRODisplay = actualRO
    Else
        currentRODisplay = roNumber
        Call LogEvent("maj", "low", "RO not found on screen, using sequence: " & roNumber, "Main", "", "")
    End If

    If Len(Trim(CStr(currentRODisplay))) > 0 Then
        Call LogEvent("comm", "med", "Sent RO to BlueZone", "Main", "", "")
    Else
        Call LogEvent("comm", "med", roNumber & " - Sent RO to BlueZone", "Main", "RO: (unknown) - will use sequence number for checks", "")
    End If

    g_PreviousNormalizedRo = NormalizeRoIdentifier(currentRODisplay)
    g_PreviousSequenceNumber = CStr(roNumber)

    Call LogEvent("comm", "low", "Sequence " & roNumber & " (RO " & currentRODisplay & ") - Processing", "ProcessRONumbers", "", "")

    ' Check for "closed" response
    If IsTextPresent("Repair Order " & currentRODisplay & " is closed.") Then
        Call LogEvent("comm", "med", "Repair Order Closed", "Main", "", "")
        lastRoResult = "Closed"
        Exit Sub
    End If
    
    ' Check for "NOT ON FILE" response
    If IsTextPresent("NOT ON FILE") Then
        Call LogEvent("comm", "med", "Not On File", "Main", "", "")
        lastRoResult = "Not On File"
        Exit Sub
    End If

    ' App rule: line 23 may show process lock when another user has this RO open.
    ' Send Enter once to return to COMMAND: and skip this RO for now.
    mainPromptText = GetScreenLine(MainPromptLine)
    If InStr(1, mainPromptText, "Process is locked by", vbTextCompare) > 0 Then
        Call LogEvent("comm", "med", "RO locked by another user - returning to command", "Main", "RO: " & currentRODisplay, "Line " & MainPromptLine & ": '" & mainPromptText & "'")
        Call FastKey("<Enter>")
        Call WaitForPrompt("COMMAND:", "", False, 5000, "Process Lock Recovery")
        lastRoResult = "Skipped - Process locked by another user"
        Exit Sub
    End If

    ' Check for "PRESS RETURN TO CONTINUE" (VEHID not on file)
    If IsTextPresent("PRESS RETURN TO CONTINUE") Then
        Call LogEvent("comm", "med", "VEHID not on file - attempting recovery", "Main", "RO: " & currentRODisplay, "")
        If Not BZH_RecoverFromVehidError(g_EmployeeNumber, g_EmployeeNameConfirm, "2") Then
            Call LogEvent("crit", "low", "VEHID recovery failed - terminal state unknown", "Main", "", "")
            g_ShouldAbort = True
        End If
        lastRoResult = "Skipped - VEHID not on file"
        Exit Sub
    End If
    
    ' Otherwise, assume repair order is open G�� prefer the scraped RO for logging
    If Len(Trim(CStr(currentRODisplay))) > 0 Then
        Call LogEvent("comm", "med", "Repair Order Open", "Main", "", "")
    Else
        Call LogEvent("comm", "med", roNumber & " - Repair Order Open", "Main", "", "")
    End If
    
    ' Ensure the RO screen is fully drawn and interactive by waiting for the bottom prompt.
    ' This addresses the "E before prompt" issue especially on first sequence.
    If Not WaitForPrompt("COMMAND:", "", False, 5000, "RO Screen Ready") Then
        Call LogEvent("crit", "low", "COMMAND prompt did not appear on RO screen. Manual intervention required.", "Main", "", "Stopping script - fail fast requested")
        g_ShouldAbort = True
        Exit Sub
    End If

    If ShouldSkipRo(currentRODisplay) Then
        g_SkipConfiguredCount = g_SkipConfiguredCount + 1
        Call LogEvent("comm", "med", "Configured SkipRoList match - skipping RO", "Main", "RO: " & currentRODisplay, "")
        Call FastText("E")
        Call FastKey("<NumpadEnter>")
        Call WaitForPrompt("COMMAND:", "", False, 5000, "")
        lastRoResult = "Skipped - Configured RO skip list"
        Exit Sub
    End If

    ' --- WARRANTY SKIP GATE ---
    ' Allow 1000ms for RO detail lines (including LTYPE) to fully render before scanning.
    Call WaitMs(1000)
    If g_SkipWchEnabled And HasWchOnAnyDetailPage() Then
        g_SkipWarrantyCount = g_SkipWarrantyCount + 1
        Call LogEvent("comm", "med", "Warranty labor type detected - skipping RO", "Main", "WCH found on RO: " & currentRODisplay, "")
        Call FastText("E")
        Call FastKey("<NumpadEnter>")
        Call WaitForPrompt("COMMAND:", "", False, 5000, "")
        lastRoResult = "Skipped - WCH labor type"
        Exit Sub
    End If

    ' Additional render gate: ensure key RO detail sections are present before blacklist scan.
    ' This helps avoid scanning too early when COMMAND is visible but details are still painting.
    If Not WaitForRODetailReady(3000) Then
        Call LogEvent("min", "med", "RO detail markers not fully detected before blacklist scan", "Main", "Proceeding with current screen", "Timeout waiting for REPAIR ORDER/LC DESCRIPTION markers")
    End If
    
    ' Give detail lines a brief settle window; some RO lines render slightly after headers.
    Call WaitForScreenStable(1200, 150)

    Dim matchedBlacklistTerm
    matchedBlacklistTerm = BZH_GetMatchedBlacklistTerm(g_BlacklistTermsRaw, g_StabilityPause)
    If Len(Trim(CStr(matchedBlacklistTerm))) = 0 And Len(Trim(CStr(g_BlacklistTermsRaw))) > 0 Then
        ' Retry once after an additional short stabilization window to catch late-painted lines.
        Call WaitForScreenStable(1000, 150)
        matchedBlacklistTerm = BZH_GetMatchedBlacklistTerm(g_BlacklistTermsRaw, g_StabilityPause)
    End If
    If Len(Trim(CStr(matchedBlacklistTerm))) > 0 Then
        g_SkipBlacklistCount = g_SkipBlacklistCount + 1
        Call LogEvent("comm", "med", "Blacklisted term found - skipping closeout", "Main", matchedBlacklistTerm, "")
        Call FastText("E")
        Call FastKey("<NumpadEnter>")
        Call WaitForPrompt("COMMAND:", "", False, 5000, "")
        lastRoResult = "Skipped - Blacklisted term: " & matchedBlacklistTerm
        Exit Sub
    End If

    ' After opening an RO, ensure it has the expected READY TO POST status.
    If Not IsStatusReady() Then
        ' Secondary gate: check if this is an older RO with an eligible status for closeout
        Dim normalizedSkipStatus
        normalizedSkipStatus = UCase(Trim(CStr(g_LastScrapedStatus)))
        If IsOlderRoEligibleStatus(normalizedSkipStatus) And IsOlderRo() Then
            g_OlderRoAttemptCount = g_OlderRoAttemptCount + 1
            ' Undo the skip counter that IsStatusReady() already incremented,
            ' since this RO will be processed (not skipped).
            Select Case normalizedSkipStatus
                Case "OPEN", "OPENED"
                    g_SkipStatusOpenCount = g_SkipStatusOpenCount - 1
                Case "PREASSIGNED", "PRE-ASSIGNED"
                    g_SkipStatusPreassignedCount = g_SkipStatusPreassignedCount - 1
            End Select
            Call LogEvent("comm", "med", "Older RO qualifies for closeout", "Main", "Status: " & g_LastScrapedStatus & " RO: " & currentRODisplay, "")
            Call Closeout_Ro(g_LastScrapedStatus)
            If InStr(1, lastRoResult, "Successfully filed", vbTextCompare) > 0 Then
                g_OlderRoFiledCount = g_OlderRoFiledCount + 1
            End If
            Exit Sub
        End If

        Call FastText("E")
        Call FastKey("<NumpadEnter>")
        ' Wait for the command prompt to return to ensure we are in a known state
        Call WaitForPrompt("COMMAND:", "", False, 5000, "")

        lastRoResult = "Skipped - Status not ready"
        Exit Sub
    Else
        Dim currentStatus
        currentStatus = Trim(CStr(g_LastScrapedStatus))
        Call LogEvent("comm", "med", "RO STATUS: " & currentStatus & " (Ready for processing)", "Main", "", "")
    End If
    
    ' Snapshot the scraped status now to avoid timing races, then detect triggers.
    Dim trigger, roStatusForDecision
    roStatusForDecision = Trim(CStr(g_LastScrapedStatus))
    Call LogEvent("comm", "high", "Pre-trigger check", "Main", "Scraped status: '" & roStatusForDecision & "'", "")

    ' --- CLOSURE READINESS GATE ---
    ' Skip ROs where any line-letter header row carries a tech code not in
    ' AllowedTechCodes (default: C92, C93). Configure or clear in config.ini.
    Dim nonCompliantLine
    nonCompliantLine = GetFirstNonCompliantLineTech()
    If Len(Trim(nonCompliantLine)) > 0 Then
        g_SkipTechCodeCount = g_SkipTechCodeCount + 1
        Call LogEvent("comm", "med", "Closure readiness check failed - non-compliant tech code", "Main", nonCompliantLine & " | RO: " & currentRODisplay, "")
        Call FastText("E")
        Call FastKey("<NumpadEnter>")
        Call WaitForPrompt("COMMAND:", "", False, 5000, "")
        lastRoResult = "Skipped - Non-compliant tech code: " & nonCompliantLine
        Exit Sub
    End If

    ' --- PARTS ORDER NEEDED GATE ---
    ' Skip ROs where an L-line has no following P-line but a keyword suggests
    ' parts will need to be ordered. Keywords and negators come from
    ' [PostFinalCharges] PartsOrderKeywords / PartsOrderNegators in config.ini.
    Dim matchedPartsDesc
    matchedPartsDesc = GetPartsNeededLaborDesc()
    If Len(Trim(matchedPartsDesc)) > 0 Then
        g_SkipPartsOrderNeededCount = g_SkipPartsOrderNeededCount + 1
        Call LogEvent("comm", "med", "Parts likely needed on RO - skipping closeout", "Main", "Matched labor: " & matchedPartsDesc & " | RO: " & currentRODisplay, "")
        Call FastText("E")
        Call FastKey("<NumpadEnter>")
        Call WaitForPrompt("COMMAND:", "", False, 5000, "")
        lastRoResult = "Skipped - Parts order needed: " & matchedPartsDesc
        Exit Sub
    End If

    trigger = FindTrigger()
    If trigger <> "" Then
        Call LogEvent("comm", "med", "Trigger found: " & trigger, "Main", "Proceeding to Closeout", "")
        Call Closeout_Ro(roStatusForDecision)
        ' Closeout_Ro should set lastRoResult appropriately
    Else
        ' If no trigger text found, but the scraped RO status is valid for closeout,
        ' proceed to closeout anyway (status supersedes trigger text).
        If IsValidCloseoutStatus(roStatusForDecision) Then
            Call LogEvent("comm", "med", "No closeout trigger text found", "Main", "RO STATUS is " & roStatusForDecision & " G�� proceeding to Closeout", "")
            Call Closeout_Ro(roStatusForDecision)
        Else
            Call LogEvent("comm", "med", "No Closeout Text Found - Skipping Closeout", "Main", "", "")
            Call FastText("E")
            Call FastKey("<NumpadEnter>")
            ' Wait for the command prompt to return to ensure we are in a known state
            Call WaitForPrompt("COMMAND:", "", False, 5000, "")
            lastRoResult = "Skipped - No closeout text found"
        End If
    End If
End Sub

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** WaitForRODetailReady
' **DATE CREATED:** 2026-03-16
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Waits for key RO detail markers to appear after entering sequence number.
' Requires COMMAND prompt plus detail headers to reduce early-screen scan races.
'-----------------------------------------------------------------------------------
Function WaitForRODetailReady(timeoutMs)
    If timeoutMs <= 0 Then timeoutMs = 3000

    Dim waitStart, waitElapsed
    waitStart = Timer

    Do
        If IsTextPresent("COMMAND:") And IsTextPresent("REPAIR ORDER #") And IsTextPresent("LC DESCRIPTION") Then
            WaitForRODetailReady = True
            Exit Function
        End If

        Call WaitMs(50)
        waitElapsed = (Timer - waitStart) * 1000
        If waitElapsed > timeoutMs Then Exit Do
    Loop

    WaitForRODetailReady = False
End Function




'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** LogResultWithSource
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' Writes a structured log message to the log file, including a timestamp,
' log level (e.g., "INFO", "ERROR"), the source procedure name, and the message.
' This provides more context than the basic LogResult function.
' 
' 
' **PARAMETERS:**
' **level** (String): The severity level of the log entry (e.g., "INFO", "DEBUG").
' **message** (String): The main content of the log message.
' **source** (String): The name of the procedure where the log entry originated.
'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** LogEvent
' **DATE CREATED:** 2026-01-02
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' New dual-axis logging system with criticality and verbosity dimensions.
' Criticality filters event importance, verbosity controls detail expansion.
' 
' **PARAMETERS:**
' criticality (String): "crit", "maj", "min", "comm" - event importance
' verbosity (String): "low", "med", "high", "max" - detail level  
' headline (String): Main message (always shown)
' stage (String): Context/stage info (always shown in [source] bracket)
' reason (String): Specific reason/cause (shown at high+ verbosity)
' technical (String): Technical details (shown at max verbosity only)
'-----------------------------------------------------------------------------------
Sub LogEvent(criticality, verbosity, headline, stage, reason, technical)
    Dim critValue, verbValue
    
    ' Convert criticality to numeric value
    Select Case LCase(Trim(criticality))
        Case "crit": critValue = CRIT_CRITICAL
        Case "maj": critValue = CRIT_MAJOR
        Case "min": critValue = CRIT_MINOR
        Case "comm": critValue = CRIT_COMMON
        Case Else: critValue = CRIT_COMMON
    End Select
    
    ' Convert verbosity to numeric value  
    Select Case LCase(Trim(verbosity))
        Case "low": verbValue = VERB_LOW
        Case "med": verbValue = VERB_MEDIUM
        Case "high": verbValue = VERB_HIGH
        Case "max": verbValue = VERB_MAX
        Case Else: verbValue = VERB_MEDIUM
    End Select
    
    ' Check if message should be logged (both thresholds must be met)
    If critValue >= g_CurrentCriticality And verbValue <= g_CurrentVerbosity Then
        Call WriteLogEntry(criticality, verbosity, headline, stage, reason, technical)
    End If
End Sub

'-----------------------------------------------------------------------------------
' **CONVENIENCE FUNCTIONS:** New dual-axis logging shortcuts
'-----------------------------------------------------------------------------------
Sub LogCritical(headline, stage, reason, technical)
    Call LogEvent("crit", "low", headline, stage, reason, technical)
End Sub

Sub LogMajor(headline, stage, reason, technical)
    Call LogEvent("maj", "low", headline, stage, reason, technical)
End Sub

Sub LogMinor(headline, stage, reason, technical)
    Call LogEvent("min", "low", headline, stage, reason, technical)
End Sub

Sub LogCommon(headline, stage, reason, technical)
    Call LogEvent("comm", "low", headline, stage, reason, technical)
End Sub

' Backward compatibility adapters - map old calls to new system
Sub LogCore(message, source)
    Call LogEvent("crit", "low", message, source, "", "")
End Sub

Sub LogError(message, source)
    Call LogEvent("maj", "low", message, source, "", "")
End Sub

Sub LogWarn(message, source)
    Call LogEvent("min", "low", message, source, "", "")
End Sub

Sub LogInfo(message, source)
    Call LogEvent("comm", "low", message, source, "", "")
End Sub

Sub LogDebug(message, source)
    Call LogEvent("comm", "high", message, source, "", "")
End Sub

Sub LogTrace(message, source)
    Call LogEvent("comm", "max", message, source, "", "")
End Sub

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** LogResultWithSource (DEPRECATED)
' **DATE CREATED:** 2025-11-19
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Legacy function - now delegates to new Log function.
' Kept for backward compatibility.
'-----------------------------------------------------------------------------------
Sub LogResultWithSource(level, message, source)
    ' Legacy adapter - map old level to new criticality/verbosity
    Dim criticality
    Select Case level
        Case 0, 1: criticality = "crit"  ' LOG_LEVEL_CORE, LOG_LEVEL_ERROR
        Case 2: criticality = "min"      ' LOG_LEVEL_WARN
        Case 3: criticality = "comm"     ' LOG_LEVEL_INFO
        Case 4, 5: criticality = "comm"  ' LOG_LEVEL_DEBUG, LOG_LEVEL_TRACE
        Case Else: criticality = "comm"
    End Select
    Call LogEvent(criticality, "med", message, source, "", "")
End Sub

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** LogDetailed (DEPRECATED)
' **DATE CREATED:** 2025-11-19
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Legacy function - now delegates to new Log function.
' Kept for backward compatibility.
'-----------------------------------------------------------------------------------
Sub LogDetailed(level, message, source)
    ' Legacy adapter - map old level to new criticality/verbosity
    Dim criticality
    Select Case level
        Case 0, 1: criticality = "crit"  ' LOG_LEVEL_CORE, LOG_LEVEL_ERROR
        Case 2: criticality = "min"      ' LOG_LEVEL_WARN
        Case 3: criticality = "comm"     ' LOG_LEVEL_INFO
        Case 4, 5: criticality = "comm"  ' LOG_LEVEL_DEBUG, LOG_LEVEL_TRACE
        Case Else: criticality = "comm"
    End Select
    Call LogEvent(criticality, "high", message, source, "", "")
End Sub

Sub WriteLogEntry(criticality, verbosity, headline, stage, reason, technical)
    Dim logFSO, logFile, logLine, logFolder, message, source
    
    ' Build message based on verbosity level
    message = headline
    ' Note: Stage is already shown in [source] field, no need to duplicate it in message
    If Len(Trim(reason)) > 0 And (verbosity = "high" Or verbosity = "max") Then
        message = message & " - Reason: " & reason
    End If
    If Len(Trim(technical)) > 0 And verbosity = "max" Then
        message = message & " | Tech: " & technical
    End If
    
    ' Use stage as source, truncate to 16 chars for compact display
    If Len(Trim(stage)) > 0 Then
        source = Left(Trim(stage) & "                ", 16) ' Pad to 16 chars
    Else
        source = "General         "
    End If
    
    ' Build compact log line: HH:MM:SS[crit/verb][source]Message
    Dim timeStamp, currentTime
    currentTime = Now
    timeStamp = Right("0" & Hour(currentTime), 2) & ":" & Right("0" & Minute(currentTime), 2) & ":" & Right("0" & Second(currentTime), 2)
    logLine = timeStamp & "[" & criticality & "/" & verbosity & "][" & source & "]" & message

    On Error Resume Next
    Set logFSO = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    logFolder = logFSO.GetParentFolderName(LOG_FILE_PATH)
    If Len(logFolder) > 0 Then
        Call EnsureFolderExists(logFSO, logFolder)
        If Err.Number <> 0 Then
            Err.Clear
            Exit Sub
        End If
    End If

    Set logFile = logFSO.OpenTextFile(LOG_FILE_PATH, 8, True)
    If Err.Number <> 0 Then
        Err.Clear
        If LOG_FILE_PATH <> LEGACY_LOG_PATH Then
            LOG_FILE_PATH = LEGACY_LOG_PATH
            logFolder = logFSO.GetParentFolderName(LOG_FILE_PATH)
            If Len(logFolder) > 0 Then Call EnsureFolderExists(logFSO, logFolder)
            Set logFile = logFSO.OpenTextFile(LOG_FILE_PATH, 8, True)
            If Err.Number <> 0 Then
                Err.Clear
                Set logFile = Nothing
                Set logFSO = Nothing
                On Error GoTo 0
                Exit Sub
            End If
        Else
            Set logFile = Nothing
            Set logFSO = Nothing
            On Error GoTo 0
            Exit Sub
        End If
    End If

    logFile.WriteLine logLine
    logFile.Close
    Set logFile = Nothing
    Set logFSO = Nothing
    On Error GoTo 0
End Sub

'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TrimLogToLimit
' **DATE CREATED:** 2026-01-03
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Uses ratio-based calculation to trim log file to configured size limit.
' Only trims at session boundaries to maintain log integrity.
'-----------------------------------------------------------------------------------
Sub TrimLogToLimit(logFSO)
    ' Realistic character-to-byte ratio for log files with timestamps, brackets, and mixed content
    ' Based on typical log file analysis: ~750 chars/KB accounts for structured log format overhead
    Const CHARS_PER_KB = 750  ' Conservative estimate for log file density
    Dim rotationSize, currentSize, excessKB, charsToRemove
    
    ' Exit if log file doesn't exist yet
    If Not logFSO.FileExists(LOG_FILE_PATH) Then Exit Sub
    
    ' Get rotation size from config (default to 1MB if not set)
    rotationSize = GetIniSetting("Settings", "LogRotationSize", "1048576")
    If Not IsNumeric(rotationSize) Then rotationSize = 1048576
    rotationSize = CLng(rotationSize)
    
    ' Get current file size
    Dim logFileObj
    Set logFileObj = logFSO.GetFile(LOG_FILE_PATH)
    currentSize = logFileObj.Size
    Set logFileObj = Nothing
    
    ' Check if trimming is needed
    If currentSize <= rotationSize Then Exit Sub
    
    ' Calculate excess and characters to remove
    excessKB = (currentSize - rotationSize) \ 1024
    charsToRemove = excessKB * CHARS_PER_KB
    
    ' Add 20% buffer to ensure we get under the limit (keep result as integer)
    charsToRemove = Int(charsToRemove * 1.2)
    
    ' Trim the log file
    Call PerformLogTrim(logFSO, charsToRemove)
End Sub

'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** PerformLogTrim
' **DATE CREATED:** 2026-01-03
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Performs the actual log trimming by reading content and rewriting without head portion.
'-----------------------------------------------------------------------------------
Sub PerformLogTrim(logFSO, charsToRemove)
    Dim logContent, trimPoint, newContent, tempLogPath
    Dim logFile
    
    On Error Resume Next
    
    ' Read current log content
    Set logFile = logFSO.OpenTextFile(LOG_FILE_PATH, 1) ' ForReading
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub ' Cannot read original log file
    End If
    
    logContent = logFile.ReadAll()
    logFile.Close
    Set logFile = Nothing
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub ' Error reading file content
    End If
    
    ' Find trim point at a line boundary
    trimPoint = charsToRemove
    If trimPoint >= Len(logContent) Then
        ' If we'd remove everything, keep last 25% of file
        trimPoint = Int(Len(logContent) * 0.75)
    End If
    
    ' Find next newline after trim point to preserve line boundaries
    Do While trimPoint < Len(logContent) And Mid(logContent, trimPoint, 1) <> vbLf
        trimPoint = trimPoint + 1
    Loop
    
    ' Look for next session boundary if possible (90% threshold ensures we find boundaries in trim area)
    Dim sessionPos
    sessionPos = InStr(trimPoint, logContent, "=== SESSION:")
    If sessionPos > 0 And sessionPos < Int(Len(logContent) * 0.9) Then
        trimPoint = sessionPos - 1
        ' Find start of that line
        Do While trimPoint > 1 And Mid(logContent, trimPoint - 1, 1) <> vbLf
            trimPoint = trimPoint - 1
        Loop
    End If
    
    ' Get content to keep
    newContent = Mid(logContent, trimPoint + 1)
    
    ' Write trimmed content atomically using temp file
    tempLogPath = LOG_FILE_PATH & ".tmp"
    Set logFile = logFSO.CreateTextFile(tempLogPath, True)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub ' Cannot create temp file
    End If
    
    logFile.Write newContent
    If Err.Number <> 0 Then
        logFile.Close
        Set logFile = Nothing
        ' Clean up temp file on write error
        If logFSO.FileExists(tempLogPath) Then logFSO.DeleteFile tempLogPath
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    
    logFile.Close
    Set logFile = Nothing
    If Err.Number <> 0 Then
        ' Clean up temp file on close error
        If logFSO.FileExists(tempLogPath) Then logFSO.DeleteFile tempLogPath
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    
    ' Replace original with trimmed version
    If logFSO.FileExists(LOG_FILE_PATH) Then
        logFSO.DeleteFile LOG_FILE_PATH
        If Err.Number <> 0 Then
            ' Failed to delete original; do not proceed with move
            Err.Clear
            If logFSO.FileExists(tempLogPath) Then
                logFSO.DeleteFile tempLogPath
            End If
            On Error GoTo 0
            Exit Sub
        End If
    End If

    logFSO.MoveFile tempLogPath, LOG_FILE_PATH
    If Err.Number <> 0 Then
        ' Move failed; attempt to remove orphaned temp file
        Err.Clear
        If logFSO.FileExists(tempLogPath) Then
            logFSO.DeleteFile tempLogPath
        End If
    End If
    
    On Error GoTo 0
    ' Important: Don't reset session flag after trimming
    ' The trimmed content may already contain today's session header
End Sub

Sub WriteSessionHeader()
    If g_SessionDateLogged Then Exit Sub
    
    Dim logFSO, logFile, sessionLine, logFolder, currentDate
    
    On Error Resume Next
    Set logFSO = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    
    ' Check if log trimming is needed at start of new session
    Call TrimLogToLimit(logFSO)
    If Err.Number <> 0 Then
        Err.Clear
        ' Continue even if trimming fails
    End If
    
    currentDate = Now
    sessionLine = "=== SESSION: " & Year(currentDate) & "-" & Right("0" & Month(currentDate), 2) & "-" & Right("0" & Day(currentDate), 2) & " ==="

    ' Check if today's session header already exists in the log
    If logFSO.FileExists(LOG_FILE_PATH) Then
        Dim existingContent, checkFile
        Set checkFile = logFSO.OpenTextFile(LOG_FILE_PATH, 1)
        If Err.Number = 0 Then
            existingContent = checkFile.ReadAll
            If Err.Number = 0 Then
                checkFile.Close
                Set checkFile = Nothing
                ' If today's session header already exists, mark as logged and exit
                If InStr(existingContent, sessionLine) > 0 Then
                    g_SessionDateLogged = True
                    Set logFSO = Nothing
                    On Error GoTo 0
                    Exit Sub
                End If
            Else
                ' ReadAll failed, clean up and continue
                checkFile.Close
                Set checkFile = Nothing
                Err.Clear
            End If
        Else
            ' OpenTextFile failed, clear error and continue
            Err.Clear
        End If
    End If

    logFolder = logFSO.GetParentFolderName(LOG_FILE_PATH)
    If Len(logFolder) > 0 Then
        Call EnsureFolderExists(logFSO, logFolder)
        If Err.Number <> 0 Then
            Err.Clear
            Exit Sub
        End If
    End If

    Set logFile = logFSO.OpenTextFile(LOG_FILE_PATH, 8, True)
    If Err.Number <> 0 Then
        Err.Clear
        Set logFile = Nothing
        Set logFSO = Nothing
        On Error GoTo 0
        Exit Sub
    End If

    logFile.WriteLine sessionLine
    logFile.Close
    Set logFile = Nothing
    Set logFSO = Nothing
    ' Only set flag after successfully writing the header
    g_SessionDateLogged = True
    On Error GoTo 0
End Sub



Sub EnsureFolderExists(fs, folderPath)
    If Len(folderPath) = 0 Then Exit Sub
    If fs.FolderExists(folderPath) Then Exit Sub

    Dim parent
    parent = fs.GetParentFolderName(folderPath)
    If Len(parent) > 0 Then
        Call EnsureFolderExists(fs, parent)
    End If

    On Error Resume Next
    fs.CreateFolder folderPath
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub


'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** SafeMsg
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' Displays a message to the user in a host-safe manner. It attempts to use the
' BlueZone host's message box if available, otherwise falls back to the standard
' VBScript MsgBox. It also logs the message. This prevents errors when the
' script is run in an environment that does not support UI dialogs.
' 
' 
' **PARAMETERS:**
' **text** (String): The message text to display.
' **isCritical** (Boolean): If True, the message is logged as an error and displayed with a critical icon.
' **title** (String): The title for the message box window.
'-----------------------------------------------------------------------------------
Sub SafeMsg(text, isCritical, title)
    If Len(Trim(CStr(title))) = 0 Then title = ""
    If isCritical Then
        Call LogEvent("maj", "low", text, "SafeMsg", "", "")
    Else
        Call LogEvent("comm", "med", text, "SafeMsg", "", "")
    End If

    ' Try to show a MsgBox only if MsgBox exists in this host (wrap to avoid errors)
    On Error Resume Next
    ' Prefer BlueZone host message if available
    If Not g_bzhao Is Nothing Then
        g_bzhao.MsgBox text
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit Sub
        Else
            Err.Clear
        End If
    End If

    Dim tmp, msgStyle
    If isCritical Then msgStyle = vbCritical Else msgStyle = vbOKOnly
    tmp = MsgBox(text, msgStyle, title)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub




'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** GetROFromScreen
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' Scans the top three lines of the terminal screen to find and extract a Repair
' Order (RO) number. It uses a regular expression to match the pattern "RO: 123456"
' and returns the numeric part.
' 
' 
' **RETURN VALUE:**
' (String) Returns the extracted RO number, or an empty string if not found.
'-----------------------------------------------------------------------------------
Function GetROFromScreen()
    If g_bzhao Is Nothing Then
        Call LogEvent("maj", "low", "g_bzhao object is not available", "GetROFromScreen", "", "")
        GetROFromScreen = ""
        Exit Function
    End If

    Dim re
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "RO:\s*(\d{4,})"
    re.IgnoreCase = True
    re.Global = True

    Dim screenContentBuffer, screenLength, matches
    Dim attempt, maxAttempts, waitBeforeNextReadMs
    Dim candidateRo, lastObservedRo, stableCount

    screenLength = 5 * 80 ' Read beyond header to tolerate partial screen repaints
    maxAttempts = 5
    candidateRo = ""
    lastObservedRo = ""
    stableCount = 0

    For attempt = 1 To maxAttempts
        On Error Resume Next
        g_bzhao.ReadScreen screenContentBuffer, screenLength, 1, 1
        If Err.Number <> 0 Then
            Call LogEvent("maj", "med", "GetROFromScreen ReadScreen failed", "GetROFromScreen", Err.Description, "")
            Err.Clear
            On Error GoTo 0
            Exit For
        End If
        On Error GoTo 0

        candidateRo = ""
        If re.Test(screenContentBuffer) Then
            Set matches = re.Execute(screenContentBuffer)
            candidateRo = matches(matches.Count - 1).SubMatches(0)
        End If

        If Len(candidateRo) > 0 Then
            If candidateRo = lastObservedRo Then
                stableCount = stableCount + 1
            Else
                stableCount = 1
                lastObservedRo = candidateRo
            End If

            If stableCount >= 2 Then
                GetROFromScreen = candidateRo
                Exit Function
            End If
        End If

        If attempt < maxAttempts Then
            ' Keep the common case fast (stable UI), only adding delay when repaint lag is detected.
            Select Case attempt
                Case 1
                    waitBeforeNextReadMs = 40
                Case 2
                    waitBeforeNextReadMs = 80
                Case 3
                    waitBeforeNextReadMs = 120
                Case Else
                    waitBeforeNextReadMs = 160
            End Select
            Call WaitMs(waitBeforeNextReadMs)
        End If
    Next

    If Len(lastObservedRo) > 0 Then
        GetROFromScreen = lastObservedRo
        Call LogEvent("min", "med", "RO read did not stabilize before timeout; using last observed value", "GetROFromScreen", "RO: " & lastObservedRo, "")
    Else
        GetROFromScreen = ""
        Call LogEvent("crit", "low", "RO not found on screen", "GetROFromScreen", "", "")
    End If
End Function


'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** GetValidCloseoutStatuses
' **DATE CREATED:** 2025-12-29
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Returns an array of RO statuses that are valid for proceeding with closeout.
' Reads from config.ini [PostFinalCharges] ValidCloseoutStatuses setting.
' Falls back to default statuses if not configured.
' 
' **RETURN VALUE:**
' (Array) Array of valid status strings for closeout processing
'-----------------------------------------------------------------------------------
Function GetValidCloseoutStatuses()
    Dim configStatuses, statusArray, i
    
    ' Read from config.ini with fallback to defaults
    configStatuses = GetIniSetting("PostFinalCharges", "ValidCloseoutStatuses", "READY TO POST")
    
    ' Parse comma-separated values and trim whitespace
    statusArray = Split(configStatuses, ",")
    For i = 0 To UBound(statusArray)
        statusArray(i) = Trim(statusArray(i))
    Next
    
    ' Log the configured statuses for transparency
    Call LogEvent("comm", "high", "Valid closeout statuses", "GetValidCloseoutStatuses", configStatuses, "")
    
    GetValidCloseoutStatuses = statusArray
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** IsValidCloseoutStatus
' **DATE CREATED:** 2025-12-29
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Helper function to check if a given status string is valid for closeout.
' 
' **PARAMETERS:**
' statusToCheck (String): The status string to validate
' 
' **RETURN VALUE:**
' (Boolean) Returns True if the status is valid for closeout, False otherwise.
'-----------------------------------------------------------------------------------
Function IsValidCloseoutStatus(statusToCheck)
    Dim validStatuses, i, trimmedStatus
    validStatuses = GetValidCloseoutStatuses()
    trimmedStatus = Trim(CStr(statusToCheck))
    
    For i = 0 To UBound(validStatuses)
        If StrComp(trimmedStatus, validStatuses(i), vbTextCompare) = 0 Then
            IsValidCloseoutStatus = True
            Exit Function
        End If
    Next
    
    IsValidCloseoutStatus = False
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** IsStatusReady
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' **MODIFIED:** 2025-12-29 - Refactored to use centralized status list
' 
' **FUNCTIONALITY:**
' Checks if the current screen indicates that the Repair Order status is
' valid for closeout processing. Uses GetValidCloseoutStatuses() to determine
' which statuses are acceptable.
' 
' **RETURN VALUE:**
' (Boolean) Returns True if the status is valid for closeout, False otherwise.
'-----------------------------------------------------------------------------------
Function IsStatusReady()
    ' Use GetRepairOrderStatus() to scrape the exact RO status from the screen
    ' Caller may choose to add waits before calling if needed
    g_bzhao.pause 1000 ' brief pause to ensure screen is stable
    Dim roStatus, validStatuses, i
    roStatus = GetRepairOrderStatus()
    validStatuses = GetValidCloseoutStatuses()
    
    Dim trimmedStatus
    trimmedStatus = Trim(CStr(roStatus))
    
    ' Check if current status matches any valid closeout status
    For i = 0 To UBound(validStatuses)
        If StrComp(trimmedStatus, validStatuses(i), vbTextCompare) = 0 Then
            IsStatusReady = True
            Exit Function
        End If
    Next

    Dim normalizedStatus
    normalizedStatus = UCase(trimmedStatus)

    Select Case normalizedStatus
        Case "OPEN", "OPENED"
            g_SkipStatusOpenCount = g_SkipStatusOpenCount + 1
        Case "PREASSIGNED", "PRE-ASSIGNED"
            g_SkipStatusPreassignedCount = g_SkipStatusPreassignedCount + 1
        Case Else
            If normalizedStatus <> "READY TO POST" Then
                g_SkipStatusOtherCount = g_SkipStatusOtherCount + 1

                If normalizedStatus = "" Then normalizedStatus = "(BLANK)"
                If Not IsObject(g_SkipOtherStates) Then
                    Set g_SkipOtherStates = CreateObject("Scripting.Dictionary")
                End If

                If g_SkipOtherStates.Exists(normalizedStatus) Then
                    g_SkipOtherStates.Item(normalizedStatus) = g_SkipOtherStates.Item(normalizedStatus) + 1
                Else
                    g_SkipOtherStates.Add normalizedStatus, 1
                End If
            End If
    End Select
    
    IsStatusReady = False
End Function

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** GetRepairOrderStatus
' **DATE CREATED:** 2025-12-17
' **AUTHOR:** GitHub Copilot (modified)
'
' **FUNCTIONALITY:**
' Reads a specific region of the terminal screen (line 5, cols 1-30) to find
' the prefix "RO STATUS: " and returns the following status text (up to 15 chars).
' Returns an empty string if the prefix isn't present or on read error.
'-----------------------------------------------------------------------------------
Function GetRepairOrderStatus()
    On Error Resume Next
    If g_bzhao Is Nothing Then
        Call LogEvent("min", "med", "GetRepairOrderStatus: g_bzhao object not available", "GetRepairOrderStatus", "", "")
        GetRepairOrderStatus = ""
        Exit Function
    End If

    Dim buf, lengthToRead, lineNum, colNum
    lengthToRead = 30
    lineNum = 5
    colNum = 1
    g_bzhao.ReadScreen buf, lengthToRead, lineNum, colNum
    If Err.Number <> 0 Then
        Call LogEvent("min", "med", "GetRepairOrderStatus: ReadScreen failed", "GetRepairOrderStatus", Err.Description, "")
        Err.Clear
        GetRepairOrderStatus = ""
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Call LogEvent("comm", "max", "GetRepairOrderStatus - raw buffer", "GetRepairOrderStatus", "'" & Replace(buf, vbCrLf, " ") & "'", "")

    Dim prefix, pos, raw
    prefix = "RO STATUS: "
    pos = InStr(1, buf, prefix, vbTextCompare)
    If pos = 0 Then
        ' Not found in this slice - clear stale status
        g_LastScrapedStatus = ""
        GetRepairOrderStatus = ""
        Exit Function
    End If

    Dim startPos
    startPos = pos + Len(prefix)
    raw = Mid(buf, startPos, 15) ' take next 15 chars per spec
    raw = Replace(raw, vbCrLf, " ")
    raw = Replace(raw, vbCr, " ")
    raw = Replace(raw, vbLf, " ")
    Dim parsedStatus
    parsedStatus = Trim(raw)
    GetRepairOrderStatus = parsedStatus
    ' Save parsed status for later logging by the caller
    g_LastScrapedStatus = parsedStatus
    Call LogEvent("comm", "high", "GetRepairOrderStatus - parsed status", "GetRepairOrderStatus", "'" & parsedStatus & "'", "")
End Function


'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** GetOpenedDate
' **DATE CREATED:** 2026-03-25
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Reads the OPENED DATE field from the PFC screen (row 4) and returns the raw
' date string in CDK DDMMMYY format (e.g. "04FEB26"). Returns empty string on
' failure or if the prefix is not found.
'-----------------------------------------------------------------------------------
Function GetOpenedDate()
    On Error Resume Next
    If g_bzhao Is Nothing Then
        Call LogEvent("min", "med", "GetOpenedDate: g_bzhao object not available", "GetOpenedDate", "", "")
        GetOpenedDate = ""
        Exit Function
    End If

    ' Read rows 1-6 (480 chars) to capture the OPENED DATE field regardless of exact row
    Dim buf
    g_bzhao.ReadScreen buf, 480, 1, 1
    If Err.Number <> 0 Then
        Call LogEvent("min", "med", "GetOpenedDate: ReadScreen failed", "GetOpenedDate", Err.Description, "")
        Err.Clear
        GetOpenedDate = ""
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Call LogEvent("comm", "max", "GetOpenedDate - raw buffer", "GetOpenedDate", "'" & Replace(buf, vbCrLf, " ") & "'", "")

    ' Use regex to extract date token in either DDMMMYY or M/D/YY format
    Dim regEx, matches
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = False
    regEx.Pattern = "OPENED DATE:\s*(\d{1,2}[A-Z]{3}\d{2,4}|\d{1,2}/\d{1,2}/\d{2,4})"

    If regEx.Test(buf) Then
        Set matches = regEx.Execute(buf)
        GetOpenedDate = Trim(matches(0).SubMatches(0))
        Call LogEvent("comm", "high", "GetOpenedDate - parsed date", "GetOpenedDate", "'" & GetOpenedDate & "'", "")
    Else
        GetOpenedDate = ""
    End If
End Function


'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** ParseCdkDate
' **DATE CREATED:** 2026-03-25
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Parses a CDK date string in DDMMMYY format (e.g. "04FEB26") or slash format
' (e.g. "01/20/26") into a VBScript Date value. Returns Empty if the input
' cannot be parsed.
'-----------------------------------------------------------------------------------
Function ParseCdkDate(dateStr)
    ParseCdkDate = Empty
    Dim cleaned
    cleaned = Trim(dateStr)
    If Len(cleaned) = 0 Then Exit Function

    ' Handle slash format (e.g. "01/20/26", "1/5/26") via VBScript CDate
    If InStr(cleaned, "/") > 0 Then
        On Error Resume Next
        If IsDate(cleaned) Then
            ParseCdkDate = CDate(cleaned)
        End If
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    ' Handle DDMMMYY / DDMMMYYYY format (e.g. "04FEB26", "04FEB2026")
    If Len(cleaned) < 7 Then Exit Function
    cleaned = UCase(cleaned)

    Dim dayPart, monthPart, yearPart
    dayPart = Left(cleaned, 2)
    monthPart = Mid(cleaned, 3, 3)
    yearPart = Mid(cleaned, 6)

    Dim monthNum
    Select Case monthPart
        Case "JAN": monthNum = 1
        Case "FEB": monthNum = 2
        Case "MAR": monthNum = 3
        Case "APR": monthNum = 4
        Case "MAY": monthNum = 5
        Case "JUN": monthNum = 6
        Case "JUL": monthNum = 7
        Case "AUG": monthNum = 8
        Case "SEP": monthNum = 9
        Case "OCT": monthNum = 10
        Case "NOV": monthNum = 11
        Case "DEC": monthNum = 12
        Case Else: Exit Function
    End Select

    On Error Resume Next
    Dim dayNum, yearNum
    dayNum = CInt(dayPart)
    yearNum = CInt(yearPart)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    If Len(yearPart) = 2 Then
        If yearNum >= 70 Then
            yearNum = 1900 + yearNum
        Else
            yearNum = 2000 + yearNum
        End If
    End If

    On Error Resume Next
    ParseCdkDate = DateSerial(yearNum, monthNum, dayNum)
    If Err.Number <> 0 Then
        Err.Clear
        ParseCdkDate = Empty
    End If
    On Error GoTo 0
End Function


'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** IsOlderRo
' **DATE CREATED:** 2026-03-25
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Determines if the current RO on screen qualifies as an "older" RO by comparing
' the OPENED DATE to today. Returns True if the RO has been open for at least
' g_OlderRoThresholdDays days. Returns False on parse failure or if below threshold.
'-----------------------------------------------------------------------------------
Function IsOlderRo()
    IsOlderRo = False

    If g_OlderRoThresholdDays <= 0 Then
        Call LogEvent("comm", "high", "Older RO check disabled (threshold=0)", "IsOlderRo", "", "")
        Exit Function
    End If

    Dim rawDate, openedDate, ageDays
    rawDate = GetOpenedDate()
    If Len(Trim(rawDate)) = 0 Then
        Call LogEvent("min", "med", "Cannot determine RO age - OPENED DATE not found", "IsOlderRo", "", "")
        Exit Function
    End If

    openedDate = ParseCdkDate(rawDate)
    If IsEmpty(openedDate) Then
        Call LogEvent("min", "med", "Cannot determine RO age - date parse failed", "IsOlderRo", "Raw: '" & rawDate & "'", "")
        Exit Function
    End If

    ageDays = DateDiff("d", openedDate, Now())
    Call LogEvent("comm", "high", "RO age calculated", "IsOlderRo", "Opened: " & rawDate & " Age: " & ageDays & " days (threshold: " & g_OlderRoThresholdDays & ")", "")

    If ageDays >= g_OlderRoThresholdDays Then
        IsOlderRo = True
    End If
End Function


'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** IsOlderRoEligibleStatus
' **DATE CREATED:** 2026-03-25
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Checks whether a given RO status is in the config-driven list of statuses
' eligible for age-based closeout (g_OlderRoStatuses from config.ini).
' Returns True if the status is eligible, False otherwise.
'-----------------------------------------------------------------------------------
Function IsOlderRoEligibleStatus(statusToCheck)
    IsOlderRoEligibleStatus = False
    If Not IsArray(g_OlderRoStatuses) Then Exit Function

    Dim normalized, i
    normalized = UCase(Trim(statusToCheck))
    For i = 0 To UBound(g_OlderRoStatuses)
        If normalized = g_OlderRoStatuses(i) Then
            IsOlderRoEligibleStatus = True
            Exit Function
        End If
    Next
End Function


'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** WaitForTextSilent
' **DATE CREATED:** 2025-11-15
' **AUTHOR:** GitHub Copilot
' 
'----------------------------------------------------------------------------------
' FUNCTION NAME: WaitForTextSilent
' DATE CREATED: .....2025-11-15
' AUTHOR:........... GitHub Copilot
' 
' **FUNCTIONALITY:**
' Polls for the specified text without logging timeout entries. Returns True if the
' text is detected within the provided timeout (in milliseconds).
'----------------------------------------------------------------------------------
Function WaitForTextSilent(textToFind, timeoutMs)
    Dim startTime, elapsedMs
    startTime = Timer

    Do
        If IsTextPresent(textToFind) Then
            WaitForTextSilent = True
            Exit Function
        End If

        Call WaitMs(120)
        elapsedMs = (Timer - startTime) * 1000
    Loop While elapsedMs < timeoutMs

    WaitForTextSilent = False
End Function




'----------------------------------------------------------------------------------
' PROCEDURE NAME: ...LogScreenSnapshot
' DATE CREATED: .....2025-11-04
' AUTHOR:........... Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' Captures the top few lines of the terminal screen and writes them to the log
' file with a "DEBUG" level. This is used for debugging to understand the
' state of the screen at a specific point in the script's execution.
' 
' 
' **PARAMETERS:**
' **name** (String): A descriptive name for the snapshot, included in the log entry.
'-----------------------------------------------------------------------------------
Sub LogScreenSnapshot(name)
    If Not DEBUG_LOGGING Then Exit Sub
    Dim screenContentBuffer, screenLength, snippet
    screenLength = DEBUG_SCREEN_LINES * 80
    On Error Resume Next
    g_bzhao.ReadScreen screenContentBuffer, screenLength, 1, 1
    If Err.Number <> 0 Then
        Call LogEvent("min", "med", "LogScreenSnapshot failed to read screen", "LogScreenSnapshot", Err.Description, "")
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' Replace newlines for compact single-line logging
    snippet = Replace(screenContentBuffer, vbCrLf, " ")
    snippet = Replace(snippet, vbCr, " ")
    snippet = Replace(snippet, vbLf, " ")
    ' Trim to 240 chars so logs are compact
    If Len(snippet) > 240 Then snippet = Left(snippet, 240) & "..."
    Call LogEvent("comm", "max", "ScreenSnapshot(" & name & ")", "LogScreenSnapshot", snippet, "")
End Sub




'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** ProcessLineItems
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' Iterates through all possible line items on a Repair Order (A-Z) and processes
' each one. It uses a "loop-until-error" strategy, where it attempts to action
' each line letter and stops when it detects a "NOT ON FILE" error. For each
' valid line, it steps through a series of prompts, accepting the default values.
'-----------------------------------------------------------------------------------
Sub ProcessLineItems()
    Dim lineLetterChar, i, lineItemPrompts
    Set lineItemPrompts = CreateLineItemPromptDictionary()
    
    ' Initialize line tracking
    g_LastSuccessfulLine = ""

    ' Phase 1: Run FNL commands for all lines (A through Z)
    Call LogEvent("comm", "high", "Phase 1: Running FNL commands for all lines", "ProcessLineItems", "", "")
    Dim pliPrePartNum
    For i = 65 To 90 ' ASCII for A to Z
        lineLetterChar = Chr(i)
        Call LogEvent("comm", "high", "Running FNL " & lineLetterChar & " command", "ProcessLineItems", "", "")

        ' Pre-detect WCH labor type BEFORE sending FNL so part number can be captured
        ' from the clean detail screen (no dialog overlay yet).
        pliPrePartNum = ""
        If IsWchLine(lineLetterChar) Then
            Call LogInfo("WCH labor type detected on line " & lineLetterChar & " - pre-capturing part number", "ProcessLineItems")
            pliPrePartNum = ExtractPartNumberForFca()
        End If

        ' Send the FNL command and let ProcessPromptSequence handle any prompts that appear
        Call FastText("FNL " & lineLetterChar)
        Call FastKey("<NumpadEnter>")

        ' Brief wait to let the response appear
        Call WaitMs(500)

        ' Check if line exists BEFORE processing other prompts
        If IsTextPresent("LINE CODE " & lineLetterChar & " IS NOT ON FILE") Then
            Dim screenResponse
            screenResponse = GetScreenSnapshot(24)
            If g_LastSuccessfulLine = "" Then
                Call LogEvent("comm", "low", "No line items found - Line " & lineLetterChar & " does not exist", "ProcessLineItems", "No lines to process", "")
            Else
                Call LogEvent("comm", "low", "Finished processing line items. No more lines found after " & g_LastSuccessfulLine, "ProcessLineItems", "Line " & lineLetterChar & " does not exist", "")
            End If
            Call LogEvent("comm", "high", "FNL " & lineLetterChar & " command response", "ProcessLineItems", screenResponse, "")
            ' System automatically returns to COMMAND prompt without manual ENTER
            Exit For
        End If

        ' Track successful line for better reporting
        g_LastSuccessfulLine = lineLetterChar

        ' Handle FCA warranty dialog if it appeared after FNL (predicted by IsWchLine above)
        If IsTextPresent("FCA GLOBAL CLAIMS INFORMATION") Then
            Call LogInfo("FCA warranty dialog present on line " & lineLetterChar & " - handling", "ProcessLineItems")
            Call HandleFcaDialog(pliPrePartNum)
        End If

        ' Process any other prompts that appear (including technician assignment)
        Call ProcessPromptSequence(lineItemPrompts)
    Next

    Call LogEvent("comm", "low", "Phase 1 completed - All lines finalized", "ProcessLineItems", "", "")

    ' Phase 2: Run R commands and process prompts for all lines
    Call LogEvent("comm", "high", "Phase 2: Processing line prompts with R commands", "ProcessLineItems", "", "")
    For i = 65 To 90 ' ASCII for A to Z
        lineLetterChar = Chr(i)
        Call LogEvent("comm", "high", "Running R " & lineLetterChar & " command", "ProcessLineItems", "", "")
        
        ' Wait for the COMMAND prompt and then enter "R" + the current line letter.
        Call WaitForPrompt("COMMAND:", "R " & lineLetterChar, True, g_PromptWait, "")
        
        ' Brief wait to let the response appear
        Call WaitMs(500)
        
        ' Check if the line exists FIRST - this avoids inappropriate timeout errors
        If IsTextPresent("LINE CODE " & lineLetterChar & " IS NOT ON FILE") Then
            Dim rScreenResponse
            rScreenResponse = GetScreenSnapshot(24)
            Call LogEvent("comm", "high", "R " & lineLetterChar & " command response", "ProcessLineItems", rScreenResponse, "")
            If g_LastSuccessfulLine = "" Then
                Call LogEvent("comm", "low", "Finished processing line items. No line items found to process", "ProcessLineItems", "Line " & lineLetterChar & " does not exist", "")
            Else
                Call LogEvent("comm", "low", "Finished processing line items. No more lines found after " & g_LastSuccessfulLine, "ProcessLineItems", "Line " & lineLetterChar & " does not exist", "")
            End If
            ' System automatically returns to COMMAND prompt without manual ENTER
            Exit For ' Exit the For loop.
        End If
        
        ' Wait for the specific line item screen to appear using generic transition function
        Dim expectedLineText, lineScreenLoaded
        expectedLineText = "LINE " & lineLetterChar & " STORY :"
        lineScreenLoaded = WaitForScreenTransition(expectedLineText, 3000, "line " & lineLetterChar & " screen")
        
        If Not lineScreenLoaded Then
            Call LogEvent("crit", "low", "CRITICAL: Line " & lineLetterChar & " screen failed to load within timeout", "ProcessLineItems", "Screen state uncertain", "")
            Call LogEvent("crit", "low", "Cannot safely continue processing with unknown screen state", "ProcessLineItems", "Exiting line processing", "")
            ' Press Enter to attempt clearing any pending screen state
            Call FastKey("<Enter>")
            Exit For ' Exit the For loop due to critical screen loading failure
        End If
        
        ' Check if the line exists. If not, we are done with line processing.
        If IsTextPresent("LINE CODE " & lineLetterChar & " IS NOT ON FILE") Then
            If i = 65 Then ' First line (A) not found
                Call LogInfo("Finished processing line items. No line A found - no line items to process", "ProcessLineItems")
            Else ' Subsequent line not found
                Call LogInfo("Finished processing line items. No more lines found after " & Chr(i-1), "ProcessLineItems")
            End If
            ' Press Enter to clear the "NOT ON FILE" message from the screen.
            Call FastKey("<Enter>")
            Exit For ' Exit the For loop.
        End If

        ' Use the new state machine method for all prompt handling
        Call LogDebug("Processing line item " & lineLetterChar & " using ProcessSingleLine_Dynamic", "ProcessLineItems")
        Call LogDetailed("INFO", "Processing line item " & lineLetterChar & " using ProcessPromptSequence", "ProcessLineItems")

        ' Process all prompts for this line item using the new state machine
        Call ProcessPromptSequence(lineItemPrompts)
        
        ' Track successful line processing for R commands
        g_LastSuccessfulLine = lineLetterChar
        Call LogInfo("Completed processing line item " & lineLetterChar, "ProcessLineItems")
    Next
    Call LogEvent("comm", "low", "All line charges have been reviewed and updated", "ProcessLineItems", "", "")
End Sub

'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** Closeout_Ro
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' **MODIFIED:** 2025-12-29 - Added status-aware closeout logic
' 
' **FUNCTIONALITY:**
' Automates the sequence of steps required to close out a Repair Order (RO).
' Routes to status-specific closeout procedures based on the RO status.
' 
' **PARAMETERS:**
' roStatus (String): The RO status to determine closeout procedure
'-----------------------------------------------------------------------------------
Sub Closeout_Ro(roStatus)
    Call LogEvent("comm", "med", "Starting closeout procedure", "Closeout_Ro", "Status: " & roStatus, "")

    ' --- PARTS CHARGE GUARD ---
    ' Abort closeout if no charged parts are found unless labor-only exception rules apply.
    Dim noPartsSkipReason
    If Not EvaluatePartsChargedGate(noPartsSkipReason) Then
        Call LogWarn("No charged parts found on RO detail screen - skipping closeout", "Closeout_Ro")
        Call FastText("E")
        Call FastKey("<NumpadEnter>")
        Call WaitForPrompt("COMMAND:", "", False, 5000, "")
        lastRoResult = noPartsSkipReason
        Exit Sub
    End If

    ' Check if status-specific closeout is enabled
    Dim useStatusSpecific
    useStatusSpecific = GetIniSetting("PostFinalCharges", "UseStatusSpecificCloseout", "true")
    
    If LCase(Trim(useStatusSpecific)) = "true" Then
        ' Route to status-specific closeout logic
        Select Case UCase(Trim(roStatus))
            Case "READY TO POST"
                Call Closeout_ReadyToPost()
            Case "PREASSIGNED", "PRE-ASSIGNED"
                Call Closeout_Preassigned()
            Case "OPENED", "OPEN"
                Call Closeout_Open()
            Case Else
                ' Default/fallback closeout for unknown statuses
                Call LogWarn("Unknown status '" & roStatus & "' - using default closeout procedure", "Closeout_Ro")
                Call Closeout_Default()
        End Select
    Else
        Call LogInfo("Using default closeout procedure (status-specific disabled)", "Closeout_Ro")
        Call Closeout_Default()
    End If
End Sub

'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** Closeout_ReadyToPost
' **DATE CREATED:** 2025-12-29
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Handles closeout procedure specifically for "READY TO POST" status ROs.
'-----------------------------------------------------------------------------------
Sub Closeout_ReadyToPost()
    Call LogInfo("Executing READY TO POST closeout procedure", "Closeout_ReadyToPost")
    
    ' For READY TO POST status ROs, process each line individually: R A -> FNL A -> R B -> FNL B, etc.
    Call ProcessLinesSequentially()
    
    ' Send the File (FC) command
    Call LogInfo("Sending file command after READY TO POST processing", "Closeout_ReadyToPost")
    WaitForPrompt "COMMAND:", "FC", True, g_PromptWait, ""
    If HandleCloseoutErrors() Then Exit Sub

    Call PerformFinalCloseout("Closeout_ReadyToPost")
End Sub

'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** Closeout_Preassigned
' **DATE CREATED:** 2025-12-29
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Handles closeout procedure specifically for "PREASSIGNED" status ROs.
' May have different steps or prompts compared to standard closeout.
'-----------------------------------------------------------------------------------
Sub Closeout_Preassigned()
    Call LogInfo("Executing PREASSIGNED closeout procedure", "Closeout_Preassigned")
    
    ' For PREASSIGNED status ROs, process each line individually: R A -> FNL A -> R B -> FNL B, etc.
    Call ProcessLinesSequentially()
    
    ' Send the File (FC) command
    Call LogInfo("Sending file command after PREASSIGNED processing", "Closeout_Preassigned")
    WaitForPrompt "COMMAND:", "FC", True, g_PromptWait, ""
    If HandleCloseoutErrors() Then Exit Sub

    Call PerformFinalCloseout("Closeout_Preassigned")
End Sub

'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** Closeout_Open
' **DATE CREATED:** 2025-12-30
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Handles closeout procedure specifically for "OPEN" status ROs.
' For OPEN status, individual lines (A, B, C, etc.) need to be closed out
' using "FNL X" commands, followed by "R X" processes, and finally "F".
'-----------------------------------------------------------------------------------
Sub Closeout_Open()
    Call LogInfo("Executing OPEN closeout procedure", "Closeout_Open")
    
    ' For OPEN status ROs, process each line individually: R A -> FNL A -> R B -> FNL B, etc.
    Call ProcessLinesSequentially()
    
    ' Finally, send the File (FC) command
    Call LogInfo("Sending file command after OPEN status processing", "Closeout_Open")
    WaitForPrompt "COMMAND:", "FC", True, g_PromptWait, ""
    If HandleCloseoutErrors() Then Exit Sub

    Call PerformFinalCloseout("Closeout_Open")
End Sub

'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** ProcessLinesSequentially
' **DATE CREATED:** 2026-01-07
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Processes OPEN status lines in proper sequence: R A -> FNL A -> R B -> FNL B, etc.
' Each line is reviewed then immediately closed before moving to the next line.
'-----------------------------------------------------------------------------------
Sub ProcessLinesSequentially()
    Call LogInfo("Starting sequential line processing (R then FNL per line)", "ProcessLinesSequentially")
    Call LogEvent("comm", "high", "ProcessLinesSequentially called", "ProcessLinesSequentially", "Processing lines in R->FNL sequence", "Starting line A")
    
    Dim lineLetterChar, i, lineItemPrompts, fnlPrompts
    Set lineItemPrompts = CreateLineItemPromptDictionary()
    Set fnlPrompts = CreateFnlPromptDictionary()
    
    ' Initialize line tracking
    g_LastSuccessfulLine = ""

    For i = 65 To 90 ' ASCII for A to Z
        lineLetterChar = Chr(i)
        Call LogEvent("comm", "high", "Processing line " & lineLetterChar & " - Finish then Review", "ProcessLinesSequentially", "", "")
        
        ' Pre-detect WCH labor type BEFORE sending FNL so part number can be captured
        ' from the clean detail screen (no dialog overlay yet).
        Dim fcaPrePartNum
        fcaPrePartNum = ""
        If IsWchLine(lineLetterChar) Then
            Call LogInfo("WCH labor type detected on line " & lineLetterChar & " - pre-capturing part number", "ProcessLinesSequentially")
            fcaPrePartNum = ExtractPartNumberForFca()
        End If

        ' Step 1: Finish the line with FNL command FIRST (to ensure it's complete before reviewing)
        Call LogEvent("comm", "high", "Running FNL " & lineLetterChar & " command", "ProcessLinesSequentially", "", "")
        Call WaitForPrompt("COMMAND:", "FNL " & lineLetterChar, True, g_PromptWait, "")

        ' Wait for the FNL response
        Call LogEvent("comm", "high", "Waiting for FNL " & lineLetterChar & " response", "ProcessLinesSequentially", "", "")
        Call WaitForScreenStable(2000, 300)  ' Wait up to 2 sec for screen to stabilize

        ' Check if the line exists FIRST
        If IsTextPresent("LINE CODE " & lineLetterChar & " IS NOT ON FILE") Then
            Dim fnlScreenResponse
            fnlScreenResponse = GetScreenSnapshot(24)
            Call LogEvent("comm", "high", "FNL " & lineLetterChar & " command response", "ProcessLinesSequentially", fnlScreenResponse, "")
            If g_LastSuccessfulLine = "" Then
                Call LogEvent("comm", "low", "No line items found to process", "ProcessLinesSequentially", "Line " & lineLetterChar & " does not exist", "")
            Else
                Call LogEvent("comm", "low", "Finished processing lines. No more lines found after " & g_LastSuccessfulLine, "ProcessLinesSequentially", "Line " & lineLetterChar & " does not exist", "")
            End If
            Exit For ' Exit the For loop
        End If

        ' Check if the line is already finished
        If IsTextPresent("LINE " & lineLetterChar & " IS ALREADY FINISHED") Then
            Call LogEvent("comm", "high", "Line " & lineLetterChar & " already finished", "ProcessLinesSequentially", "Skipping FNL processing", "")
        Else
            ' Handle FCA warranty dialog if it appeared after FNL (predicted by IsWchLine above)
            If IsTextPresent("FCA GLOBAL CLAIMS INFORMATION") Then
                Call LogInfo("FCA warranty dialog present on line " & lineLetterChar & " - handling", "ProcessLinesSequentially")
                Call HandleFcaDialog(fcaPrePartNum)
            End If
            ' Process FNL prompts for this line
            Call LogDebug("Processing FNL " & lineLetterChar & " prompts", "ProcessLinesSequentially")
            Call ProcessPromptSequence(fnlPrompts)
        End If
        
        ' Step 2: Review the line with R command (now that it's finished)
        Call LogEvent("comm", "high", "Running R " & lineLetterChar & " command", "ProcessLinesSequentially", "", "")
        Call WaitForPrompt("COMMAND:", "R " & lineLetterChar, True, g_PromptWait, "")
        
        ' Wait for the first prompt to appear (2 second timeout for response to be ready)
        Call LogEvent("comm", "high", "Waiting for R " & lineLetterChar & " response prompts", "ProcessLinesSequentially", "", "")
        Call WaitForScreenStable(2000, 300)  ' Wait up to 2 sec for screen to stabilize
        
        ' Process review prompts for this line (R command produces prompts immediately, not a screen)
        Call LogDebug("Processing R " & lineLetterChar & " prompts", "ProcessLinesSequentially")
        Call ProcessPromptSequence(lineItemPrompts)
        
        ' Track successful line processing
        g_LastSuccessfulLine = lineLetterChar
        Call LogInfo("Completed sequential processing for line " & lineLetterChar, "ProcessLinesSequentially")
    Next
    
    Call LogEvent("comm", "low", "All lines processed sequentially (FNL->R per line)", "ProcessLinesSequentially", "", "")
End Sub

'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** ProcessReviewLines
' **DATE CREATED:** 2026-01-07
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Processes line review for OPEN status ROs using R commands.
' This is Phase 2 extracted from ProcessLineItems for proper workflow ordering.
' Reviews each line (A-Z) without doing FNL commands.
'-----------------------------------------------------------------------------------
Sub ProcessReviewLines()
    Call LogInfo("Starting line review process with R commands", "ProcessReviewLines")
    Call LogEvent("comm", "high", "ProcessReviewLines called", "ProcessReviewLines", "About to process R commands for line review", "Starting line A")
    
    Dim lineLetterChar, i, lineItemPrompts
    Set lineItemPrompts = CreateLineItemPromptDictionary()
    
    ' Initialize line tracking
    g_LastSuccessfulLine = ""

    ' Process R commands and prompts for all lines
    Call LogEvent("comm", "high", "Processing line prompts with R commands", "ProcessReviewLines", "", "")
    For i = 65 To 90 ' ASCII for A to Z
        lineLetterChar = Chr(i)
        Call LogEvent("comm", "high", "Running R " & lineLetterChar & " command", "ProcessReviewLines", "", "")
        
        ' Wait for the COMMAND prompt and then enter "R" + the current line letter.
        Call WaitForPrompt("COMMAND:", "R " & lineLetterChar, True, g_PromptWait, "")
        
        ' Brief wait to let the response appear
        Call WaitMs(500)
        
        ' Check if the line exists FIRST - this avoids inappropriate timeout errors
        If IsTextPresent("LINE CODE " & lineLetterChar & " IS NOT ON FILE") Then
            Dim rScreenResponse
            rScreenResponse = GetScreenSnapshot(24)
            Call LogEvent("comm", "high", "R " & lineLetterChar & " command response", "ProcessReviewLines", rScreenResponse, "")
            If g_LastSuccessfulLine = "" Then
                Call LogEvent("comm", "low", "Finished processing line items. No line items found to process", "ProcessReviewLines", "Line " & lineLetterChar & " does not exist", "")
            Else
                Call LogEvent("comm", "low", "Finished processing line items. No more lines found after " & g_LastSuccessfulLine, "ProcessReviewLines", "Line " & lineLetterChar & " does not exist", "")
            End If
            ' System automatically returns to COMMAND prompt without manual ENTER
            Exit For ' Exit the For loop.
        End If
        
        ' Wait for the specific line item screen to appear using generic transition function
        Dim expectedLineText, lineScreenLoaded
        expectedLineText = "LINE " & lineLetterChar & " STORY :"
        lineScreenLoaded = WaitForScreenTransition(expectedLineText, 3000, "line " & lineLetterChar & " screen")
        
        If Not lineScreenLoaded Then
            Call LogEvent("crit", "low", "CRITICAL: Line " & lineLetterChar & " screen failed to load within timeout", "ProcessReviewLines", "Screen state uncertain", "")
            Call LogEvent("crit", "low", "Cannot safely continue processing with unknown screen state", "ProcessReviewLines", "Exiting line processing", "")
            ' Press Enter to attempt clearing any pending screen state
            Call FastKey("<Enter>")
            Exit For ' Exit the For loop due to critical screen loading failure
        End If
        ' Use the new state machine method for all prompt handling
        Call LogDebug("Processing line item " & lineLetterChar & " using ProcessPromptSequence", "ProcessReviewLines")

        ' Process all prompts for this line item using the new state machine
        Call ProcessPromptSequence(lineItemPrompts)
        
        ' Track successful line processing for R commands
        g_LastSuccessfulLine = lineLetterChar
        Call LogInfo("Completed processing line item " & lineLetterChar, "ProcessReviewLines")
    Next
    Call LogEvent("comm", "low", "All line charges have been reviewed", "ProcessReviewLines", "", "")
End Sub

'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** ProcessOpenStatusLines
' **DATE CREATED:** 2025-12-30
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Processes individual lines (A-Z) for OPEN status ROs by sending FNL X commands
' to close open lines. This runs AFTER ProcessReviewLines() to ensure proper workflow order.
' Stops when encountering "NOT ON FILE" errors or "ALREADY FINISHED" messages.
'-----------------------------------------------------------------------------------
Sub ProcessOpenStatusLines()
    Call LogInfo("Starting OPEN status line processing with FNL commands", "ProcessOpenStatusLines")
    Call LogEvent("comm", "high", "ProcessOpenStatusLines called", "ProcessOpenStatusLines", "About to process FNL commands for open lines", "Starting line A")
    
    Dim lineLetterChar, i
    For i = 65 To 90 ' ASCII for A to Z
        lineLetterChar = Chr(i)
        Call LogInfo("Processing OPEN status for line " & lineLetterChar, "ProcessOpenStatusLines")
        Call LogEvent("comm", "high", "Attempting FNL command", "ProcessOpenStatusLines", "Line " & lineLetterChar, "Sending 'FNL " & lineLetterChar & "' command")
        
        ' First, try to send FNL (Final Line) command for this line
        Call WaitForPrompt("COMMAND:", "FNL " & lineLetterChar, True, g_PromptWait, "")
        
        ' Check if the line exists by looking for "NOT ON FILE" error
        If IsTextPresent("LINE CODE " & lineLetterChar & " IS NOT ON FILE") Then
            Call LogEvent("comm", "high", "LINE NOT ON FILE detected", "ProcessOpenStatusLines", "Line " & lineLetterChar & " does not exist", "Exiting FNL processing")
            If i = 65 Then ' First line (A) not found
                Call LogInfo("No line A found - no open lines to process with FNL commands", "ProcessOpenStatusLines")
            Else ' Subsequent line not found
                Dim lastProcessedLineChar
                lastProcessedLineChar = Chr(i-1)
                Call LogInfo("No more open lines found after " & lastProcessedLineChar & " - FNL processing complete", "ProcessOpenStatusLines")
            End If
            ' System automatically returns to COMMAND prompt without manual ENTER
            Exit For ' Exit the For loop
        End If
        
        ' Check for "Line X is already finished" message
        If IsTextPresent("LINE " & lineLetterChar & " IS ALREADY FINISHED") Then
            Call LogEvent("comm", "high", "LINE ALREADY FINISHED detected", "ProcessOpenStatusLines", "Line " & lineLetterChar & " is already closed", "Skipping to next line")
            Call LogInfo("Line " & lineLetterChar & " is already finished, moving to next line", "ProcessOpenStatusLines")
            ' System automatically returns to COMMAND prompt without manual ENTER
        Else
            Call LogEvent("comm", "high", "FNL command succeeded", "ProcessOpenStatusLines", "Line " & lineLetterChar & " accepted FNL command", "Processing FNL prompts")
            ' If line exists, handle any prompts that appear after FNL command
            Call LogInfo("FNL command processed for line " & lineLetterChar, "ProcessOpenStatusLines")
            
            ' Small delay to allow screen updates
            Call WaitMs(1000)
            
            ' REFACTORED: Use ProcessPromptSequence instead of generic colon detection
            ' This handles specific FNL prompts explicitly rather than any text with ":"
            Call LogDebug("Processing remaining FNL prompts for line " & lineLetterChar & " using ProcessPromptSequence", "ProcessOpenStatusLines")
            
            Dim fnlPrompts
            Set fnlPrompts = CreateFnlPromptDictionary()
            Call ProcessPromptSequence(fnlPrompts)
        End If
        
        Call LogInfo("Completed FNL processing for line " & lineLetterChar, "ProcessOpenStatusLines")
    Next
    
    Call LogInfo("Completed OPEN status line processing with FNL commands", "ProcessOpenStatusLines")
    Call LogEvent("comm", "low", "All lines have been successfully closed", "ProcessOpenStatusLines", "", "")
End Sub

'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** Closeout_Default
' **DATE CREATED:** 2025-12-29
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' The standard/default closeout procedure. Contains the original closeout logic.
'-----------------------------------------------------------------------------------
Sub Closeout_Default()
    ' Process each line individually: R A -> FNL A -> R B -> FNL B, etc.
    Call ProcessLinesSequentially()

    ' Send the File (FC) command
    Call LogInfo("Sending file command after default processing", "Closeout_Default")
    WaitForPrompt "COMMAND:", "FC", True, g_PromptWait, ""
    If HandleCloseoutErrors() Then Exit Sub

    Call PerformFinalCloseout("Closeout_Default")
End Sub

Sub PerformFinalCloseout(callerName)
    Call LogInfo("Executing final closeout sequence with state machine", callerName)
    
    Dim closeoutPrompts
    Dim closeoutSequenceTimeoutMs, closeoutMaxNoPromptIterations, closeoutNoPromptRetryWaitMs
    Set closeoutPrompts = CreateCloseoutPromptDictionary()
    closeoutSequenceTimeoutMs = 120000
    closeoutMaxNoPromptIterations = 20
    closeoutNoPromptRetryWaitMs = 6000
    Call LogEvent("comm", "med", "Using extended closeout retry window", callerName, "timeoutMs=" & closeoutSequenceTimeoutMs & " maxNoPromptIterations=" & closeoutMaxNoPromptIterations, "noPromptRetryWaitMs=" & closeoutNoPromptRetryWaitMs)

    g_ProcessPromptSequenceTimeoutMsOverride = closeoutSequenceTimeoutMs
    g_ProcessPromptSequenceMaxNoPromptIterationsOverride = closeoutMaxNoPromptIterations
    g_ProcessPromptSequenceNoPromptRetryWaitMsOverride = closeoutNoPromptRetryWaitMs
    
    ' Process the entire closeout sequence using the state machine
    Call ProcessPromptSequence(closeoutPrompts)
    Call ResetProcessPromptSequenceOverrides()
    
    lastRoResult = "Successfully filed"
    Call LogInfo("RO filed successfully - ready for downstream review", callerName)
End Sub

' Handles the optional comeback prompt that occasionally appears during closeout.
Sub HandleOptionalComebackPrompt()
    ' Small delay to allow screen to update after sending operation code
    'Call WaitMs(5000)
    
    ' Capture screen state before checking for comeback prompt
    Call LogDebug("Checking for comeback prompt - screen snapshot: " & GetScreenSnapshot(24), "Closeout_Ro")

    ' Check for the exact comeback prompt (case-insensitive search)
    Dim comebackPrompt
    comebackPrompt = "Is this a comeback (Y/N)"

    Call LogDebug("Checking for comeback prompt: '" & comebackPrompt & "'", "HandleOptionalComebackPrompt")
    If WaitForTextSilent(comebackPrompt, 2000) Then
        Call LogInfo("Detected comeback prompt. Sending 'Y'.", "HandleOptionalComebackPrompt")
        Call FastKey("Y")
        Call FastKey("<NumpadEnter>")
        Call WaitMs(POST_PROMPT_WAIT_MS)

        ' Wait for the prompt to clear
        Dim clearStart, clearElapsed
        clearStart = Timer
        Do While IsTextPresent(comebackPrompt)
            Call WaitMs(100)
            clearElapsed = (Timer - clearStart) * 1000
            If clearElapsed > 2000 Then
                Call LogWarn("Comeback prompt did not clear within 2 seconds", "HandleOptionalComebackPrompt")
                Exit Do
            End If
        Loop
        Call LogInfo("Comeback prompt handled successfully", "HandleOptionalComebackPrompt")
    Else
        Call LogDebug("No comeback prompt detected within " & CStr(2000 / 1000) & " seconds", "HandleOptionalComebackPrompt")
        Call LogDebug("Comeback prompt not found - final screen snapshot: " & GetScreenSnapshot(24), "HandleOptionalComebackPrompt")
    End If
End Sub

' Helper: return the first matching trigger string or empty if none found.
'-----------------------------------------------------------
' When calling this function, it will check for multiple trigger strings.
' If any are found, it returns the first matching string.
' This determines whether the repair order should proceed to closeout.
' -----------------------------------------------------------

'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** FindTrigger
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' Searches the terminal screen for a set of predefined "trigger" strings. 
' These strings indicate that an open Repair Order is correct type for the closeout
' process to begin. It returns the first trigger that is found.
' There are many types of service orders; only those matching a trigger will be closed out
' as some orders may not require special closeout processing.
' 
' **RETURN VALUE:**
' (String) Returns the matching trigger text if found, otherwise returns an empty string.
'-----------------------------------------------------------------------------------
Function FindTrigger()
    Call LogEvent("comm", "high", "FindTrigger starting", "FindTrigger", "Scanning for closeout triggers", "")
    
    Dim i, candidate

    If Not IsArray(g_CloseoutTriggers) Then
        Call LogEvent("crit", "low", "Closeout triggers not initialized", "FindTrigger", "g_CloseoutTriggers is not an array", "InitializeConfig should load TriggerList")
        FindTrigger = ""
        Exit Function
    End If
    If UBound(g_CloseoutTriggers) < LBound(g_CloseoutTriggers) Then
        Call LogEvent("crit", "low", "Closeout triggers list empty at runtime", "FindTrigger", "g_CloseoutTriggers", "Check TriggerList configuration")
        FindTrigger = ""
        Exit Function
    End If
    
    For i = LBound(g_CloseoutTriggers) To UBound(g_CloseoutTriggers)
        candidate = g_CloseoutTriggers(i)
        Call LogEvent("comm", "max", "Checking trigger", "FindTrigger", "'" & candidate & "'", "")
        If IsTextPresent(candidate) Then
            Call LogEvent("comm", "med", "Trigger found", "FindTrigger", "'" & candidate & "'", "")
            FindTrigger = candidate
            Exit Function
        End If
    Next
    
    Call LogEvent("comm", "med", "No triggers found", "FindTrigger", "Checked " & (UBound(g_CloseoutTriggers) - LBound(g_CloseoutTriggers) + 1) & " triggers", "")
    FindTrigger = ""
End Function


'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** HandleCloseoutErrors
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' Checks the screen for known error messages that can occur during the closeout
' process. If an error is detected, it logs the error and executes a predefined
' sequence of corrective actions based on the error message.
' 
' 
' **RETURN VALUE:**
' (Boolean) Returns True if an error was detected and handled, False otherwise.
'-----------------------------------------------------------------------------------
Function HandleCloseoutErrors()
    Dim errorMap, key
    Set errorMap = GetCloseoutErrorMap()
    
    For Each key In errorMap.Keys
        If IsTextPresent(key) Then
            ' Try to extract the detailed message shown on the screen near the key
            Dim detailedMsg
            detailedMsg = ExtractMessageNear(key)

            If Len(Trim(CStr(currentRODisplay))) > 0 Then
                If Len(detailedMsg) > 0 Then
                    Call LogError("CLOSEOUT Error detected: " & key & " - " & detailedMsg, "HandleCloseoutErrors")
                Else
                    Call LogError("CLOSEOUT Error detected: " & key, "HandleCloseoutErrors")
                End If
            Else
                If Len(detailedMsg) > 0 Then
                    Call LogError("Error detected: " & key & " - " & detailedMsg, "HandleCloseoutErrors")
                Else
                    Call LogError("Error detected: " & key, "HandleCloseoutErrors")
                End If
            End If

            ExecuteActions errorMap.Item(key)
            ' Final result: mark the RO as left open for manual closing
            lastRoResult = "Left Open for manual closing"
            g_ShouldAbort = True
            ' Set the reason for the abort
            g_AbortReason = "Closeout Error: " & key
            If Len(detailedMsg) > 0 Then g_AbortReason = g_AbortReason & " (" & detailedMsg & ")"
            SafeMsg "Closeout error encountered: " & key & " - " & detailedMsg & vbCrLf & "Automation paused for manual recovery.", True, "Closeout Error"
            HandleCloseoutErrors = True
            Exit Function
        End If
    Next
    
    HandleCloseoutErrors = False
End Function


'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** GetCloseoutErrorMap
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' Defines and returns a dictionary that maps known error messages to a sequence
' of corrective actions. This provides a configurable way to handle different
' error conditions that may arise during the RO closeout process.
' 
' 
' **RETURN VALUE:**
' (Object) Returns a Scripting.Dictionary object where keys are error strings
' and values are arrays of action commands.
'-----------------------------------------------------------------------------------
Function GetCloseoutErrorMap()
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Action command formats (case-insensitive):
    '   "PressEnter"               -> press Enter key
    '   "PressKey:<key>"           -> PressKey with given key token (e.g. "<NumpadEnter>", "<Esc>")
    '   "Send:<text>"              -> EnterText (sends text then Enter)
    '   "SendNoEnter:<text>"       -> g_bzhao.SendKey text (no Enter appended)
    '   "Wait:<seconds>"           -> g_bzhao.Wait seconds
    '   "Log:<message>"            -> LogInfo("CLOSEOUT", message)
    ' 
    '   "Log:<message>"            -> LogResult("CLOSEOUT", message)
    ' 
    ' Customize or add entries below as you identify new conditions.
    
    dict.Add "ERROR", Array("PressEnter", "Send:E", "Wait:2")
    dict.Add "NOT AVAILABLE", Array("PressEnter", "Send:E", "Wait:2")
    dict.Add "INVALID", Array("PressEnter", "Send:E", "Wait:2")
    dict.Add "REQUEST CANNOT BE PROCESSED", Array("PressEnter", "Send:E", "Wait:2")
    dict.Add "INCOMPLETE SERVICE", Array("1", "Wait:9", "Send:E", "Wait:2")
    dict.Add "POSTING MESSAGES:", Array("PressEnter", "PressEnter", "PressEnter", "Send:E", "Wait:5")
    dict.Add "Is this a comeback", Array("PressKey:Y", "PressEnter", "Wait:3")

    ' Specific handler for:
    ' "NOT ALL LINES HAVE A COMPLETE STATUS...PRESS RETURN TO CONTINUE"
    ' Sequence: <Enter>, wait 2s, <Enter>, wait 2s, then send E (Exit) (Enter included by Send), then wait 2s.
    dict.Add "NOT ALL LINES HAVE A COMPLETE STATUS", Array( _
    "PressEnter", "Wait:2", _
    "PressEnter", "Wait:2", _
    "PressEnter", "Wait:2", _
    "Send:E", "Wait:2" _
    )
    
    ' Example of a different corrective sequence:
    ' dict.Add "SPECIAL POPUP", Array("PressEnter", "Wait:1", "PressKey:<Esc>", "Wait:1", "LogInfo:Handled SPECIAL POPUP")
    
    Set GetCloseoutErrorMap = dict
End Function


'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** ExtractMessageNear
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' When an error key is found on the screen, this function attempts to extract
' the full error message text surrounding the key. It reads the screen, cleans
' up the text (e.g., removes line breaks, separators), and returns a more
' detailed message for logging purposes.
' 
' 
' **PARAMETERS:**
' **key** (String): The error text that was found on the screen.
' 
' 
' **RETURN VALUE:**
' (String) Returns the cleaned-up, detailed error message.
'-----------------------------------------------------------------------------------
Function ExtractMessageNear(key)
    Dim screenContentBuffer, screenLength, keyPos, i, rowStart, rows, rowText, collected
    ' Read full screen (24 rows x 80 cols)
    screenLength = 24 * 80
    On Error Resume Next
    g_bzhao.ReadScreen screenContentBuffer, screenLength, 1, 1
    If Err.Number <> 0 Then
        Err.Clear
        ExtractMessageNear = ""
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    keyPos = InStr(1, screenContentBuffer, key, vbTextCompare)
    If keyPos = 0 Then
        ExtractMessageNear = ""
        Exit Function
    End If

    ' Capture a substring of the screen starting at the key position.
    ' This is more robust than relying on specific CRLF splits because
    ' the host may return different newline conventions or wrapped lines.
    Dim rest
    rest = Mid(screenContentBuffer, keyPos)

    ' Replace any CR/LF sequences with spaces so the returned text is a single line.
    rest = Replace(rest, vbCrLf, " ")
    rest = Replace(rest, vbCr, " ")
    rest = Replace(rest, vbLf, " ")

    ' Trim to a reasonable length to avoid giant log entries; keep most context.
    If Len(rest) > 1000 Then rest = Left(rest, 1000)


    collected = Trim(rest)

    ' If the key appears more than once near the start (e.g. "POSTING MESSAGES: - POSTING MESSAGES:"),
    ' trim everything up to the second appearance so we don't duplicate the heading.
    Dim firstKeyPos, nextKeyPos
    firstKeyPos = InStr(1, collected, key, vbTextCompare)
    If firstKeyPos > 0 Then
        nextKeyPos = InStr(firstKeyPos + Len(key), collected, key, vbTextCompare)
        If nextKeyPos > 0 And nextKeyPos <= 80 Then
            collected = Trim(Mid(collected, nextKeyPos + Len(key)))
        Else
            ' If it simply starts with the key, remove that leading occurrence
            If Len(collected) >= Len(key) Then
                If LCase(Left(collected, Len(key))) = LCase(key) Then
                    collected = Trim(Mid(collected, Len(key) + 1))
                End If
            End If
        End If
    End If

    ' Replace pipe characters (used as visual separators) with spaces
    collected = Replace(collected, "|", " ")

    ' Remove long runs of hyphens or boxed separators and lone plus signs
    On Error Resume Next
    Dim reg
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = "-{3,}"
    reg.Global = True
    collected = reg.Replace(collected, " ")
    reg.Pattern = "\+{1,}"
    collected = reg.Replace(collected, " ")
    Set reg = Nothing
    On Error GoTo 0

    ' Remove any 'PRESS RETURN' lines or known suffixes if accidentally included
    collected = Replace(collected, "PRESS RETURN TO CONTINUE", "", 1, -1, vbTextCompare)
    collected = Replace(collected, "PRESS RETURN", "", 1, -1, vbTextCompare)
    collected = Replace(collected, "...", "", 1, -1, vbTextCompare)

    ' Collapse multiple spaces
    Do While InStr(collected, "  ") > 0
        collected = Replace(collected, "  ", " ")
    Loop

    ExtractMessageNear = Trim(collected)
End Function


'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** ExecuteActions
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' Executes a series of corrective actions provided as an array of command strings. 
' It parses each string (e.g., "Send:E", "Wait:2") and performs the corresponding
' action, such as sending text, pressing a key, or waiting. This is used by the
' error handling system to recover from known error states.
' 
' 
' **PARAMETERS:**
' **actions** (Array): An array of strings, where each string is a command to execute.
'-----------------------------------------------------------------------------------
Sub ExecuteActions(actions)
    Dim idx, action, parts, cmd, arg
    For idx = LBound(actions) To UBound(actions)
        action = Trim(actions(idx))
        parts = Split(action, ":", 2)
        cmd = UCase(Trim(parts(0)))
        arg = ""
        If UBound(parts) >= 1 Then arg = parts(1)
        
        Select Case cmd
            Case "PRESSENTER"
            Call FastKey("<Enter>")
            Call WaitMs(100)
            Case "PRESSKEY"
            Call FastKey(Trim(arg))
            Case "SEND"
            ' This sends text and then Enter.
            Call FastText(arg)
            Call FastKey("<NumpadEnter>")
            Call WaitMs(1000)
            Case "SENDNOENTER"
            Call FastText(arg)
            Call WaitMs(100)
            Case "WAIT"
            If IsNumeric(arg) Then Call WaitMs(CInt(arg) * 1000) ' WaitMs expects milliseconds
            Case "LOG"
            If Len(Trim(CStr(currentRODisplay))) > 0 Then
                Call LogInfo(arg, "ExecuteActions")
            Else
                Call LogInfo(arg, "ExecuteActions")
            End If
            Case Else
            Call LogError("Unknown corrective action: " & action, "ExecuteActions")
        End Select
    Next
End Sub




'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** GetRepairOrderEnhanced
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' A specialized function to extract a Repair Order (RO) number from the screen
' after it has been created. It looks for the specific text "Created repair order "
' and parses the number that follows.
' 
' 
' **RETURN VALUE:**
' (String) Returns the extracted RO number, or an empty string if not found.
'-----------------------------------------------------------------------------------
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




'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** LogEntryWithRO
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' A simple logging routine that writes an entry to the log file, including
' a timestamp, a vehicle number (MVA), and an optional Repair Order (RO) number.
' 
' 
' **PARAMETERS:**
' **mva** (String): The vehicle number to log.
' **roNumber** (String): The associated RO number to log.
'-----------------------------------------------------------------------------------
Sub LogEntryWithRO(mva, roNumber)
    If Trim(mva) = "" Then Exit Sub

    Dim logFSO, logFile
    Set logFSO = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set logFile = logFSO.OpenTextFile(LOG_FILE_PATH, 8, True)
    If Err.Number <> 0 Then
        Err.Clear
        ' Try legacy path as fallback
        Set logFile = logFSO.OpenTextFile(LEGACY_LOG_PATH, 8, True)
        If Err.Number <> 0 Then
            Err.Clear
            ' If both attempts fail, exit gracefully
            Call LogError("Failed to open log file for MVA logging: " & LOG_FILE_PATH, "LogEntryWithRO")
            Exit Sub
        End If
    End If
    On Error GoTo 0

    If roNumber = "" Then
        logFile.WriteLine Now & " - MVA: " & mva
    Else
        logFile.WriteLine Now & " - MVA: " & mva & " - RO: " & roNumber
    End If

    logFile.Close
    Set logFile = Nothing
    Set logFSO = Nothing
End Sub




'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** UpdateDiagnosticLog
' **DATE CREATED:** 2025-11-06
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' If diagnostic logging is enabled, this captures the bottom 4 lines of the
' screen and adds it to a rolling 5-item queue. The entire queue is then
' written to a separate diagnostic log file, overwriting it each time. This
' provides a "breadcrumb trail" of the last 5 actions.
' 
' 
' **PARAMETERS:**
' **actionName** (String): A description of the action that was just performed.
'-----------------------------------------------------------------------------------
Sub UpdateDiagnosticLog(actionName)
    If Not g_EnableDiagnosticLogging Then Exit Sub

    On Error Resume Next

    Dim logFSO, logFile
    Set logFSO = CreateObject("Scripting.FileSystemObject")
    Set logFile = logFSO.OpenTextFile(DIAGNOSTIC_LOG_PATH, 8, True) ' Open for appending
    If Err.Number = 0 Then
        logFile.WriteLine Now & " [TEST] Diagnostic Log Test - Action: " & actionName
        logFile.Close
    Else
        Call LogError("Failed to write test message to diagnostic log: " & Err.Description, "UpdateDiagnosticLog")
        Err.Clear
    End If

    Set logFile = Nothing
    Set logFSO = Nothing
    Err.Clear
End Sub

'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** WaitForContinuePrompt
' **DATE CREATED:** 2025-12-23
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Waits until the COMMAND:(SEQ#/E/N/B/?) prompt appears, with a delay.
'-----------------------------------------------------------------------------------
Sub WaitForContinuePrompt()
    Dim promptText, timeoutMs, startTime, elapsedMs
    promptText = "COMMAND:(SEQ#/E/N/B/?)"
    timeoutMs = 10000 ' 10 seconds, adjust as needed
    startTime = Timer
    Do
        If IsTextPresent(promptText) Then
            Exit Sub
        End If
        Call WaitMs(120)
        elapsedMs = (Timer - startTime) * 1000
    Loop While elapsedMs < timeoutMs
    ' If not found, log and continue
    Call LogWarn("Timeout waiting for continue prompt: " & promptText, "WaitForContinuePrompt")
End Sub

Sub StartScript()
    ' === Validate all dependencies before proceeding ===
    MustHaveValidDependencies
    
    Call LogInfo("PostFinalCharges script bootstrap starting", "Bootstrap")
    ' === Include CommonLib.vbs (optional - script has built-in functions) ===
    Dim commonLibPath
    commonLibPath = GetConfigPath("PostFinalCharges", "CommonLib")
    If IncludeFile(commonLibPath) Then
        commonLibLoaded = True
        Call LogInfo("CommonLib.vbs loaded successfully", "Bootstrap")
    Else
        Call LogInfo("CommonLib.vbs not found - using built-in functions", "Bootstrap")
        commonLibLoaded = False
    End If
    ' === End Include CommonLib.vbs ===
    Call RunMainProcess
End Sub

StartScript

