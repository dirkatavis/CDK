Option Explicit


' Global script variables
Dim CSV_FILE_PATH, LOG_FILE_PATH
Dim fso, roNumber
Dim bzhao
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
Dim g_StartSequenceNumber, g_EndSequenceNumber
Dim g_DebugDelayFactor
Dim MainPromptLine
Dim LEGACY_CSV_PATH, LEGACY_LOG_PATH, LEGACY_DIAG_LOG_PATH, LEGACY_COMMONLIB_PATH
Dim g_CurrentCriticality, g_CurrentVerbosity
Dim g_SessionDateLogged
Dim g_LastSuccessfulLine
Dim g_NoPromptCount

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
Const LEGACY_BASE_PATH = "C:\Temp\Code\Scripts\VBScript\CDK\PostFinalCharges"
' Simplified timeout logic - no midnight handling needed for current debugging

' --- EARLY LOGGING: Force maximum logging for startup ---
g_CurrentCriticality = CRIT_COMMON ' Log all criticality levels
g_CurrentVerbosity = VERB_MAX ' Show maximum detail during startup
g_SessionDateLogged = False



LEGACY_CSV_PATH = LEGACY_BASE_PATH & "\CashoutRoList.csv"  '<== DEPRECATED... REMOVE AS PART ANY ANY COMMIT
LEGACY_LOG_PATH = LEGACY_BASE_PATH & "\PostFinalCharges.log"
LEGACY_DIAG_LOG_PATH = LEGACY_BASE_PATH & "\PostFinalCharges.screendump.log"
LEGACY_COMMONLIB_PATH = LEGACY_BASE_PATH & "\CommonLib.vbs"



' Bootstrap defaults so logging works before config initialization
g_BaseScriptPath = "C:\Temp\Code\Scripts\VBScript\CDK\PostFinalCharges"
CSV_FILE_PATH = ResolvePath("CashoutRoList.csv", LEGACY_CSV_PATH, True)
LOG_FILE_PATH = ResolvePath("PostFinalCharges.log", LEGACY_LOG_PATH, False)
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
End Class

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
Sub ProcessPromptSequence(prompts)
    Dim finished, promptKey, promptDetails, bestMatchKey, bestMatchLength
    Dim sequenceStartTime, sequenceElapsed
    finished = False
    sequenceStartTime = Now() ' Use actual date/time instead of Timer
    
    ' DIAGNOSTIC: Log Timer behavior at start
    Call LogEvent("comm", "high", "ProcessPromptSequence started", "ProcessPromptSequence", "Timer diagnostics", "sequenceStartTime=" & sequenceStartTime & " (Now() at start)")
    
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
        
        If sequenceElapsed > 30000 Then ' 30-second timeout
            Call LogEvent("crit", "low", "ProcessPromptSequence timed out after 30 seconds", "ProcessPromptSequence", "Automation stopped", "Now()=" & Now() & " sequenceStartTime=" & sequenceStartTime & " calculated=" & sequenceElapsed & "ms > 30000ms")
            SafeMsg "ProcessPromptSequence timed out after 30 seconds.\nAutomation stopped.", True, "Sequence Timeout"
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
        
        ' Use single line scanning instead of full screen scrape to avoid false positives
        ' Check key lines where prompts typically appear
        Dim lineToCheck, lineText, linesToCheck
        linesToCheck = Array(1, 2, 3, 4, 5, 20, 21, 22, 23, 24) ' Common prompt locations
        
        For Each lineToCheck In linesToCheck
            lineText = GetScreenLine(lineToCheck)
            If Len(lineText) > 0 Then
                ' Check each prompt key against this line
                For Each promptKey In prompts.Keys
                    Dim isRegex, re, regexError
                    isRegex = False
                    regexError = False
                    ' Heuristic: treat as regex if starts with ^ or contains ( or [ or .*
                    If Left(promptKey, 1) = "^" Or InStr(promptKey, "(") > 0 Or InStr(promptKey, "[") > 0 Or InStr(promptKey, ".*") > 0 Or InStr(promptKey, "\\d") > 0 Then
                        isRegex = True
                    End If
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
                    ' Only fall back to plain text if this was NOT a regex pattern
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

        ' --- If a prompt was found, handle it ---
        If bestMatchLength > 0 Then
            ' Reset the no-prompt counter since we found a prompt
            g_NoPromptCount = 0
            
            ' CRITICAL FIX: Reset timer for each individual prompt
            ' Each prompt gets its own 30-second timeout window
            sequenceStartTime = Now()
            Call LogEvent("comm", "med", "TIMER RESET for new prompt", "ProcessPromptSequence", "Individual prompt timeout starts now", "sequenceStartTime=" & sequenceStartTime)
            
            Set promptDetails = prompts.Item(bestMatchKey)
            Call LogEvent("comm", "med", "Matched prompt: '" & bestMatchKey & "'", "ProcessPromptSequence", "Found most specific match", "match length=" & bestMatchLength)
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
                    
                    ' Determine if bestMatchKey is a regex pattern
                    If Left(bestMatchKey, 1) = "^" Or InStr(bestMatchKey, "(") > 0 Or InStr(bestMatchKey, "[") > 0 Or InStr(bestMatchKey, ".*") > 0 Or InStr(bestMatchKey, "\\d") > 0 Then
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
                    Call LogEvent("comm", "med", "Default value detected in prompt", "ProcessPromptSequence", "Accepting by sending only key press", "")
                Else
                    Call LogEvent("comm", "high", "No valid default value detected", "ProcessPromptSequence", "Will send ResponseText", "")
                End If
            Else
                Call LogEvent("comm", "high", "Prompt has AcceptDefault=False", "ProcessPromptSequence", "Will always send ResponseText if provided", "")
            End If

            If promptDetails.ResponseText <> "" And Not shouldAcceptDefault Then
                Call FastText(promptDetails.ResponseText)
                Call LogEvent("comm", "med", "Sent ResponseText", "ProcessPromptSequence", "'" & promptDetails.ResponseText & "'", "")
            Else
                Call LogEvent("comm", "med", "No ResponseText to send", "ProcessPromptSequence", "Empty or accepting default", "")
            End If
            
            Call FastKey(promptDetails.KeyPress)
            
            ' Add extra logging for problematic prompts
            If InStr(bestMatchKey, "ADD A LABOR OPERATION") > 0 Then
                Call LogEvent("comm", "med", "Responded to ADD A LABOR OPERATION prompt", "ProcessPromptSequence", "Waiting for screen to stabilize", "")
                Call WaitMs(2000) ' Extra wait for this specific prompt
                Call LogEvent("comm", "high", "Screen after ADD A LABOR OPERATION response", "ProcessPromptSequence", "", GetScreenSnapshot(5))
                
                ' Check if we're back at COMMAND prompt
                If IsTextPresent("COMMAND:") Then
                    Call LogEvent("comm", "med", "Successfully returned to COMMAND prompt", "ProcessPromptSequence", "After ADD A LABOR OPERATION", "")
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
                Call LogEvent("comm", "med", "Responded to SOLD HOURS prompt", "ProcessPromptSequence", "Waiting for screen to stabilize", "")
                Call WaitMs(1500) ' Extra wait for this specific prompt
                Call LogEvent("comm", "high", "Screen after SOLD HOURS response", "ProcessPromptSequence", "", GetScreenSnapshot(5))
            End If

            If promptDetails.IsSuccess Then
                finished = True
                Call LogEvent("comm", "med", "Success prompt reached", "ProcessPromptSequence", bestMatchKey, "")
            End If

            ' TRACE: Log screen snapshot after key send
            Call LogScreenSnapshot("AfterKeySend")

            ' Wait for the prompt to clear before rescanning
            Dim clearStart, clearElapsed
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
                Call LogEvent("comm", "med", "Detected return to COMMAND prompt on MainPromptLine", "ProcessPromptSequence", "Line processing complete", "")
                finished = True
            Else
                ' Track consecutive "no prompt" iterations to prevent infinite loops
                g_NoPromptCount = g_NoPromptCount + 1
                
                If g_NoPromptCount > 20 Then ' Maximum 20 iterations of no prompt (5 seconds total)
                    Call LogEvent("crit", "low", "Too many consecutive iterations with no prompt detected", "ProcessPromptSequence", "Possible infinite loop - aborting", "noPromptCount=" & g_NoPromptCount & " line=" & MainPromptLine & " text='" & mainPromptText & "'")
                    Call SafeMsg("Automation appears stuck - too many iterations with no prompt detected." & vbCrLf & "Line " & MainPromptLine & ": '" & mainPromptText & "'" & vbCrLf & "Stopping automation.", True, "Infinite Loop Detection")
                    g_ShouldAbort = True
                    finished = True
                Else
                    Call LogEvent("min", "med", "No prompt found - waiting and retrying", "ProcessPromptSequence", "Attempt " & g_NoPromptCount & " of 20", "line=" & MainPromptLine & " text='" & mainPromptText & "'")
                    Call WaitMs(250)
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
    bzhao.ReadScreen screenContentBuffer, 80, lineNum, 1
    If Err.Number <> 0 Then
        GetScreenLine = ""
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    lineText = Trim(screenContentBuffer)
    GetScreenLine = lineText
End Function


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
        Dim isRegex, re
        isRegex = False
        If Left(key, 1) = "^" Or InStr(key, "(") > 0 Or InStr(key, "[") > 0 Or InStr(key, ".*") > 0 Or InStr(key, "\\d") > 0 Then
            isRegex = True
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
    Else
        SafeMsg "Unable to connect to BlueZone. Check that it’s open and logged in.", True, "Connection Error"
    End If

    ' Cleanup
    ' Guard object cleanup with IsObject to avoid 'Object required' when variables are Empty
    If IsObject(bzhao) Then
        On Error Resume Next
        bzhao.Disconnect
        Set bzhao = Nothing
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If
    If IsObject(fso) Then Set fso = Nothing
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
    ' Hardcoded path for the script's base directory, as per user request.
    ' This bypasses dynamic path resolution which can fail in hosted environments.
    g_BaseScriptPath = "C:\Temp\Code\Scripts\VBScript\CDK\PostFinalCharges"
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
    configPath = ResolvePath("config.ini", "", False)
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
    Set fso = CreateObject("Scripting.FileSystemObject")
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
        ' Include and use mock bzhao for testing
        Dim mockPath
        mockPath = ResolvePath("MockBzhao.vbs", "", True)
        If IncludeFile(mockPath) Then
            Set bzhao = New MockBzhao
            Call LogEvent("comm", "med", "Using MockBzhao for testing", "InitializeObjects", "", "")
            
            ' Setup initial test scenario
            bzhao.SetupTestScenario("basic_command_prompt")
        Else
            Call LogEvent("maj", "low", "Could not load MockBzhao.vbs for test mode", "InitializeObjects", "", "")
            g_IsTestMode = False
        End If
    End If
    
    If Not g_IsTestMode Then
        Set bzhao = CreateObject("BZWhll.WhllObj")
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
    ' Force g_BaseScriptPath to the known project root
    g_BaseScriptPath = "C:\Temp\Code\Scripts\VBScript\CDK\PostFinalCharges"
    
    ' --- Initialize Logging Configuration from INI file ---
    Dim criticalityValue, verbosityValue
    criticalityValue = LCase(GetIniSetting("Settings", "LogCriticality", "comm"))
    verbosityValue = LCase(GetIniSetting("Settings", "UserVerbosity", "med"))
    
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
    startSequenceNumberValue = GetIniSetting("Processing", "StartSequenceNumber", "")
    endSequenceNumberValue = GetIniSetting("Processing", "EndSequenceNumber", "")

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
    CSV_FILE_PATH = ResolvePath("CashoutRoList.csv", LEGACY_CSV_PATH, True)
    LOG_FILE_PATH = ResolvePath("PostFinalCharges.log", LEGACY_LOG_PATH, False)
    g_LongWait = 2000
    g_SendRetryCount = 2
    g_DelayBetweenTextAndEnterMs = 2000
    POST_PROMPT_WAIT_MS = Int(1000 * g_DebugDelayFactor)  ' Base 1000ms scaled by debug delay factor
    g_EnableDiagnosticLogging = False
    DIAGNOSTIC_LOG_PATH = ResolvePath("PostFinalCharges.screendump.log", LEGACY_DIAG_LOG_PATH, False)
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
' It uses the global bzhao object and attempts to connect to the default
' session. It logs the outcome of the connection attempt.
' 
' 
' **RETURN VALUE:**
' (Boolean) Returns True if the connection is successful, False otherwise.
'-----------------------------------------------------------------------------------
Function ConnectBlueZone()
    On Error Resume Next
    If bzhao Is Nothing Then
        Call LogEvent("crit", "low", "BlueZone object is not available", "ConnectBlueZone", "CreateObject failed", "")
        ConnectBlueZone = False
        Exit Function
    End If
    
    bzhao.Connect ""
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

    Dim roNumber
    Dim lineCount
    Dim sequenceLabel
    lineCount = 0
    
    ' In test mode, only process one RO
    If g_IsTestMode Then
        roNumber = 900
        Call LogROHeader(roNumber)
        sequenceLabel = "Sequence " & roNumber
        Call LogEvent("comm", "low", sequenceLabel & " - Processing", "ProcessRONumbers", "", "")
        
        lastRoResult = ""
        Call Main(roNumber)
        
        Call LogEvent("comm", "med", sequenceLabel & " - Result: " & lastRoResult, "ProcessRONumbers", "", "")
        Call LogEvent("comm", "med", "Test mode: Processed single RO " & roNumber, "ProcessRONumbers", "", "")
        Exit Sub
    End If

    For roNumber = g_StartSequenceNumber To g_EndSequenceNumber
        lineCount = lineCount + 1
        'WaitMs(2000)
        Call LogROHeader(roNumber)
        sequenceLabel = "Sequence " & roNumber
        Call LogEvent("comm", "low", sequenceLabel & " - Processing", "ProcessRONumbers", "", "")

        ' Start performance timing for this RO
        Dim roStartTime
        roStartTime = Now

        lastRoResult = ""
        Call Main(roNumber)

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
        Dim finalDisplay
        If Len(Trim(CStr(currentRODisplay))) > 0 Then
            finalDisplay = currentRODisplay
        Else
            finalDisplay = roNumber
        End If
        Dim finalMessage
        finalMessage = sequenceLabel
        If Len(Trim(CStr(finalDisplay))) > 0 And CStr(finalDisplay) <> CStr(roNumber) Then
            finalMessage = finalMessage & " (RO " & finalDisplay & ")"
        End If
        finalMessage = finalMessage & " - Result: " & lastRoResult
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
    
    ' Wait for RO detail screen to load before scraping RO number
    If Not WaitForScreenTransition("RO STATUS:", 5000, "RO detail screen") Then
        Call LogEvent("maj", "low", "RO detail screen did not load within timeout, attempting RO extraction anyway", "Main", "", "")
    End If
    
    ' Scrape the actual RO number from the screen (top of screen shows 'RO:  123456')
    Dim actualRO
    actualRO = GetROFromScreen()
    If Len(Trim(CStr(actualRO))) > 0 Then
        currentRODisplay = actualRO
    Else
        currentRODisplay = roNumber
        Call LogEvent("maj", "low", "RO not found on screen, using sequence: " & roNumber, "Main", "", "")
    End If
    
    If Len(Trim(CStr(currentRODisplay))) > 0 Then
        Call LogEvent("comm", "med", "Sent RO to BlueZone", "Main", "", "")
    Else
        ' No scraped RO available; log against the sequence number and note unknown RO
        Call LogEvent("comm", "med", roNumber & " - Sent RO to BlueZone", "Main", "RO: (unknown) - will use sequence number for checks", "")
    End If
    
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
    
    ' Otherwise, assume repair order is open — prefer the scraped RO for logging
    If Len(Trim(CStr(currentRODisplay))) > 0 Then
        Call LogEvent("comm", "med", "Repair Order Open", "Main", "", "")
    Else
        Call LogEvent("comm", "med", roNumber & " - Repair Order Open", "Main", "", "")
    End If
    
    ' Allow time for RO details to fully load before checking status
    'Call WaitMs(2000)
    
    ' After opening an RO, ensure it has the expected READY TO POST status.
    If Not IsStatusReady() Then
        ' Call LogCore("RO STATUS not READY TO POST - exiting (E) and moving to next", "Main") ' Redundant, removed
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
    trigger = FindTrigger()
    If trigger <> "" Then
        Call LogEvent("comm", "med", "Trigger found: " & trigger, "Main", "Proceeding to Closeout", "")
        Call Closeout_Ro(roStatusForDecision)
        ' Closeout_Ro should set lastRoResult appropriately
    Else
        ' If no trigger text found, but the scraped RO status is valid for closeout,
        ' proceed to closeout anyway (status supersedes trigger text).
        If IsValidCloseoutStatus(roStatusForDecision) Then
            Call LogEvent("comm", "med", "No closeout trigger text found", "Main", "RO STATUS is " & roStatusForDecision & " — proceeding to Closeout", "")
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
    If Not bzhao Is Nothing Then
        bzhao.MsgBox text
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit Sub
        Else
            Err.Clear
        End If
    End If

    Dim tmp
    tmp = MsgBox(text, IIf(isCritical, vbCritical, vbOKOnly), title)
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
    If bzhao Is Nothing Then
        Call LogEvent("maj", "low", "bzhao object is not available", "GetROFromScreen", "", "")
        GetROFromScreen = ""
        Exit Function
    End If

    Dim screenContentBuffer, screenLength, re, matches
    screenLength = 3 * 80 ' top three lines
    On Error Resume Next
    bzhao.ReadScreen screenContentBuffer, screenLength, 1, 1
    If Err.Number <> 0 Then
        Call LogEvent("maj", "med", "GetROFromScreen ReadScreen failed", "GetROFromScreen", Err.Description, "")
        Err.Clear
        GetROFromScreen = ""
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "RO:\s*(\d{4,})"
    re.IgnoreCase = True
    re.Global = False
    
    If re.Test(screenContentBuffer) Then
        Set matches = re.Execute(screenContentBuffer)
        GetROFromScreen = matches(0).SubMatches(0)
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
' Reads from config.ini [Processing] ValidCloseoutStatuses setting.
' Falls back to default statuses if not configured.
' 
' **RETURN VALUE:**
' (Array) Array of valid status strings for closeout processing
'-----------------------------------------------------------------------------------
Function GetValidCloseoutStatuses()
    Dim configStatuses, statusArray, i
    
    ' Read from config.ini with fallback to defaults
    configStatuses = GetIniSetting("Processing", "ValidCloseoutStatuses", "READY TO POST,PREASSIGNED,OPENED")
    
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
    bzhao.pause 1000 ' brief pause to ensure screen is stable
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
    If bzhao Is Nothing Then
        Call LogEvent("min", "med", "GetRepairOrderStatus: bzhao object not available", "GetRepairOrderStatus", "", "")
        GetRepairOrderStatus = ""
        Exit Function
    End If

    Dim buf, lengthToRead, lineNum, colNum
    lengthToRead = 30
    lineNum = 5
    colNum = 1
    bzhao.ReadScreen buf, lengthToRead, lineNum, colNum
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
    bzhao.ReadScreen screenContentBuffer, screenLength, 1, 1
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
    Call LogEvent("comm", "med", "Phase 1: Running FNL commands for all lines", "ProcessLineItems", "", "")
    For i = 65 To 90 ' ASCII for A to Z
        lineLetterChar = Chr(i)
        Call LogEvent("comm", "med", "Running FNL " & lineLetterChar & " command", "ProcessLineItems", "", "")
        
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
        
        ' Process any other prompts that appear (including technician assignment)
        Call ProcessPromptSequence(lineItemPrompts)
    Next

    Call LogEvent("comm", "low", "Phase 1 completed - All lines finalized", "ProcessLineItems", "", "")

    ' Phase 2: Run R commands and process prompts for all lines
    Call LogEvent("comm", "med", "Phase 2: Processing line prompts with R commands", "ProcessLineItems", "", "")
    For i = 65 To 90 ' ASCII for A to Z
        lineLetterChar = Chr(i)
        Call LogEvent("comm", "med", "Running R " & lineLetterChar & " command", "ProcessLineItems", "", "")
        
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
    
    ' Check if status-specific closeout is enabled
    Dim useStatusSpecific
    useStatusSpecific = GetIniSetting("Processing", "UseStatusSpecificCloseout", "true")
    
    If LCase(Trim(useStatusSpecific)) = "true" Then
        ' Route to status-specific closeout logic
        Select Case UCase(Trim(roStatus))
            Case "READY TO POST"
                Call Closeout_ReadyToPost()
            Case "PREASSIGNED"
                Call Closeout_Preassigned()
            Case "OPENED"
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
    
    ' Handle the "ALL LABOR POSTED (Y/N)?" prompt after F command
    Call LogInfo("Waiting for 'ALL LABOR POSTED (Y/N)?' prompt", "Closeout_ReadyToPost")
    WaitForPrompt "ALL LABOR POSTED (Y/N)?", "Y", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'MILEAGE OUT' prompt", "Closeout_ReadyToPost")
    WaitForPrompt "MILEAGE OUT", "", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'MILEAGE IN' prompt", "Closeout_ReadyToPost")
    WaitForPrompt "MILEAGE IN", "", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'O.K. TO CLOSE RO (Y/N)?' prompt", "Closeout_ReadyToPost")
    WaitForPrompt "O.K. TO CLOSE RO", "Y", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'INVOICE PRINTER' prompt", "Closeout_ReadyToPost")
    WaitForPrompt "INVOICE PRINTER", "2", True, g_PromptWait, ""
    
    lastRoResult = "Successfully filed"
    Call LogInfo("RO filed successfully - ready for downstream review", "Closeout_ReadyToPost")
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
    
    ' Handle the "ALL LABOR POSTED (Y/N)?" prompt after F command
    Call LogInfo("Waiting for 'ALL LABOR POSTED (Y/N)?' prompt", "Closeout_Preassigned")
    WaitForPrompt "ALL LABOR POSTED (Y/N)?", "Y", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'MILEAGE OUT' prompt", "Closeout_Preassigned")
    WaitForPrompt "MILEAGE OUT", "", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'MILEAGE IN' prompt", "Closeout_Preassigned")
    WaitForPrompt "MILEAGE IN", "", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'O.K. TO CLOSE RO (Y/N)?' prompt", "Closeout_Preassigned")
    WaitForPrompt "O.K. TO CLOSE RO", "Y", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'INVOICE PRINTER' prompt", "Closeout_Preassigned")
    WaitForPrompt "INVOICE PRINTER", "2", True, g_PromptWait, ""
    
    lastRoResult = "Successfully filed"
    Call LogInfo("RO filed successfully - ready for downstream review", "Closeout_Preassigned")
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
    
    ' Handle the "ALL LABOR POSTED (Y/N)?" prompt after F command
    Call LogInfo("Waiting for 'ALL LABOR POSTED (Y/N)?' prompt", "Closeout_Open")
    WaitForPrompt "ALL LABOR POSTED (Y/N)?", "Y", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'MILEAGE OUT' prompt", "Closeout_Open")
    WaitForPrompt "MILEAGE OUT", "", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'MILEAGE IN' prompt", "Closeout_Open")
    WaitForPrompt "MILEAGE IN", "", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'O.K. TO CLOSE RO (Y/N)?' prompt", "Closeout_Open")
    WaitForPrompt "O.K. TO CLOSE RO", "Y", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'INVOICE PRINTER' prompt", "Closeout_Open")
    WaitForPrompt "INVOICE PRINTER", "2", True, g_PromptWait, ""
    
    lastRoResult = "Successfully filed"
    Call LogInfo("RO filed successfully - ready for downstream review", "Closeout_Open")
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
        Call LogEvent("comm", "med", "Processing line " & lineLetterChar & " - Review then Close", "ProcessLinesSequentially", "", "")
        
        ' Step 1: Review the line with R command
        Call LogEvent("comm", "med", "Running R " & lineLetterChar & " command", "ProcessLinesSequentially", "", "")
        Call WaitForPrompt("COMMAND:", "R " & lineLetterChar, True, g_PromptWait, "")
        
        ' Brief wait to let the response appear
        Call WaitMs(500)
        
        ' Check if the line exists FIRST
        If IsTextPresent("LINE CODE " & lineLetterChar & " IS NOT ON FILE") Then
            Dim rScreenResponse
            rScreenResponse = GetScreenSnapshot(24)
            Call LogEvent("comm", "high", "R " & lineLetterChar & " command response", "ProcessLinesSequentially", rScreenResponse, "")
            If g_LastSuccessfulLine = "" Then
                Call LogEvent("comm", "low", "No line items found to process", "ProcessLinesSequentially", "Line " & lineLetterChar & " does not exist", "")
            Else
                Call LogEvent("comm", "low", "Finished processing lines. No more lines found after " & g_LastSuccessfulLine, "ProcessLinesSequentially", "Line " & lineLetterChar & " does not exist", "")
            End If
            Exit For ' Exit the For loop
        End If
        
        ' Wait for the line item screen to appear
        Dim expectedLineText, lineScreenLoaded
        expectedLineText = "LINE " & lineLetterChar & " STORY :"
        lineScreenLoaded = WaitForScreenTransition(expectedLineText, 3000, "line " & lineLetterChar & " screen")
        
        If Not lineScreenLoaded Then
            Call LogEvent("crit", "low", "CRITICAL: Line " & lineLetterChar & " screen failed to load", "ProcessLinesSequentially", "Cannot continue", "")
            Call FastKey("<Enter>")
            Exit For
        End If
        
        ' Process review prompts for this line
        Call LogDebug("Processing R " & lineLetterChar & " prompts", "ProcessLinesSequentially")
        Call ProcessPromptSequence(lineItemPrompts)
        
        ' Step 2: Close the line with FNL command  
        Call LogEvent("comm", "med", "Running FNL " & lineLetterChar & " command", "ProcessLinesSequentially", "", "")
        Call WaitForPrompt("COMMAND:", "FNL " & lineLetterChar, True, g_PromptWait, "")
        
        ' Check if the line is already finished
        If IsTextPresent("LINE " & lineLetterChar & " IS ALREADY FINISHED") Then
            Call LogEvent("comm", "high", "Line " & lineLetterChar & " already finished", "ProcessLinesSequentially", "Skipping FNL processing", "")
        Else
            ' Process FNL prompts for this line
            Call LogDebug("Processing FNL " & lineLetterChar & " prompts", "ProcessLinesSequentially")
            Call ProcessPromptSequence(fnlPrompts)
        End If
        
        ' Track successful line processing
        g_LastSuccessfulLine = lineLetterChar
        Call LogInfo("Completed sequential processing for line " & lineLetterChar, "ProcessLinesSequentially")
    Next
    
    Call LogEvent("comm", "low", "All lines processed sequentially (R->FNL per line)", "ProcessLinesSequentially", "", "")
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
    Call LogEvent("comm", "med", "Processing line prompts with R commands", "ProcessReviewLines", "", "")
    For i = 65 To 90 ' ASCII for A to Z
        lineLetterChar = Chr(i)
        Call LogEvent("comm", "med", "Running R " & lineLetterChar & " command", "ProcessReviewLines", "", "")
        
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
' **PROCEDURE NAME:** ProcessFinalCloseoutPrompts
' **DATE CREATED:** 2025-12-30
' **AUTHOR:** GitHub Copilot
' **STATUS:** DEPRECATED - No longer used (prompt handling moved to ProcessPromptSequence)
' 
' **FUNCTIONALITY:**
' Originally handled final prompts after F (File) command.
' Functionality replaced by ProcessPromptSequence for better reliability.
' Kept for reference/potential future use - can be removed in cleanup.
'-----------------------------------------------------------------------------------
Sub ProcessFinalCloseoutPrompts()
    Call LogInfo("Processing final closeout prompts", "ProcessFinalCloseoutPrompts")
    
    ' ALL LABOR POSTED
    Dim send_enter_key_all_labor_posted
    send_enter_key_all_labor_posted = True

    ' Add a 3 second delay before ALL LABOR POSTED prompt
    Call WaitMs(3000)

    If Not WaitForPrompt("ALL LABOR POSTED", "Y", send_enter_key_all_labor_posted, g_TimeoutMs, "") Then
        Call LogEvent("maj", "low", "Failed to get ALL LABOR POSTED prompt", "ProcessFinalCloseoutPrompts", "Aborting closeout", "")
        lastRoResult = "Failed - Could not confirm all labor posted"
        Exit Sub
    End If
    If HandleCloseoutErrors() Then Exit Sub
    
    ' MILEAGE OUT
    Dim mileageOutTimeout
    mileageOutTimeout = 6000
    WaitForPrompt "MILEAGE OUT", "", True, mileageOutTimeout, ""
    If HandleCloseoutErrors() Then Exit Sub

    ' NEW CORNER CASE: Current Mileage less than Previous Mileage
    ' If this prompt appears, send "Y" to confirm.
    If WaitForPrompt("Current Mileage less than Previous Mileage", "", False, 5000, "") Then
        Call LogInfo("Detected 'Current Mileage less than Previous Mileage' prompt. Sending 'Y'.", "ProcessFinalCloseoutPrompts")
        Dim send_enter_key_mileage_less
        send_enter_key_mileage_less = True
        Call WaitForPrompt("Current Mileage less than Previous Mileage", "Y", send_enter_key_mileage_less, g_DefaultWait, "")
        If HandleCloseoutErrors() Then Exit Sub
    End If

    ' MILEAGE IN
    Dim mileageInTimeout
    mileageInTimeout = 6000
    Dim send_enter_key_mileage_in
    send_enter_key_mileage_in = True
    WaitForPrompt "MILEAGE IN", "", send_enter_key_mileage_in, mileageInTimeout, ""
    If HandleCloseoutErrors() Then Exit Sub

    ' O.K. TO CLOSE RO
    Dim okToCloseTimeout
    okToCloseTimeout = 15000
    Dim send_enter_key_ok_to_close_ro
    send_enter_key_ok_to_close_ro = True
    If Not WaitForPrompt("O.K. TO CLOSE RO", "Y", send_enter_key_ok_to_close_ro, okToCloseTimeout, "") Then
        Call LogEvent("maj", "low", "Failed to get O.K. TO CLOSE RO prompt", "ProcessFinalCloseoutPrompts", "Aborting closeout", "")
        lastRoResult = "Failed - Could not confirm closeout"
        Exit Sub
    End If
    If HandleCloseoutErrors() Then Exit Sub

    ' Send to printer 2
    Dim send_enter_key_invoice_printer
    send_enter_key_invoice_printer = True
    Dim invoicePromptTimeout
    invoicePromptTimeout = 5000
    If Not WaitForPrompt("INVOICE PRINTER", "2", send_enter_key_invoice_printer, invoicePromptTimeout, "") Then
        Call LogError("Failed to get INVOICE PRINTER prompt - closeout may be incomplete", "ProcessFinalCloseoutPrompts")
        lastRoResult = "Failed - Could not send to printer"
        Exit Sub
    End If
    
    ' Use the state machine for the rest of the closeout prompts
    Dim closeoutPrompts
    Set closeoutPrompts = CreateCloseoutPromptDictionary()
    
    ' Wait for the continue prompt
    Dim continuePromptTimeout
    continuePromptTimeout = 10000
    Dim continuePromptDetected
    continuePromptDetected = WaitForTextSilent("COMMAND:(SEQ#/E/N/B/?)", continuePromptTimeout)
    If continuePromptDetected Then
        Call LogInfo("Detected continue prompt: COMMAND:(SEQ#/E/N/B/?)", "ProcessFinalCloseoutPrompts")
    Else
        Call LogWarn("Timeout waiting for continue prompt: COMMAND:(SEQ#/E/N/B/?)", "ProcessFinalCloseoutPrompts")
    End If
    Call ProcessPromptSequence(closeoutPrompts)

    ' Final error handling
    If HandleCloseoutErrors() Then Exit Sub

    lastRoResult = "Successfully closed"
    Call LogInfo("Final closeout prompts completed successfully", "ProcessFinalCloseoutPrompts")
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

    ' Handle the "ALL LABOR POSTED (Y/N)?" prompt after F command
    Call LogInfo("Waiting for 'ALL LABOR POSTED (Y/N)?' prompt", "Closeout_Default")
    WaitForPrompt "ALL LABOR POSTED (Y/N)?", "Y", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'MILEAGE OUT' prompt", "Closeout_Default")
    WaitForPrompt "MILEAGE OUT", "", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'MILEAGE IN' prompt", "Closeout_Default")
    WaitForPrompt "MILEAGE IN", "", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'O.K. TO CLOSE RO (Y/N)?' prompt", "Closeout_Default")
    WaitForPrompt "O.K. TO CLOSE RO", "Y", True, g_PromptWait, ""

    Call LogInfo("Waiting for 'INVOICE PRINTER' prompt", "Closeout_Default")
    WaitForPrompt "INVOICE PRINTER", "2", True, g_PromptWait, ""

    lastRoResult = "Successfully filed"
    Call LogInfo("RO filed successfully - ready for downstream review", "Closeout_Default")
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
    
    Dim triggers, i, candidate
    ' Add or remove entries in this array as needed.
    triggers = Array( _
    "CHECK AND ADJUST TIRE PRESSURE", _
    "REPLACE TIRE SENSOR", _
    "PM CHANGE OIL & FILTER", _
    "PMS PERFORMED ON LOT", _
    "LABOR POSTED", _
    "CHECK ENGINE LIGHT", _
    "TIRE ROTATION", _
    "MILEAGE IN" _
    )
    
    For i = LBound(triggers) To UBound(triggers)
        candidate = triggers(i)
        Call LogEvent("comm", "max", "Checking trigger", "FindTrigger", "'" & candidate & "'", "")
        If IsTextPresent(candidate) Then
            Call LogEvent("comm", "med", "Trigger found", "FindTrigger", "'" & candidate & "'", "")
            FindTrigger = candidate
            Exit Function
        End If
    Next
    
    Call LogEvent("comm", "med", "No triggers found", "FindTrigger", "Checked " & (UBound(triggers) - LBound(triggers) + 1) & " triggers", "")
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
    '   "SendNoEnter:<text>"       -> bzhao.SendKey text (no Enter appended)
    '   "Wait:<seconds>"           -> bzhao.Wait seconds
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
    bzhao.ReadScreen screenContentBuffer, screenLength, 1, 1
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
    
    bzhao.ReadScreen screenContent, screenLength, 1, 1
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
    Call LogInfo("PostFinalCharges script bootstrap starting", "Bootstrap")
    ' === Include CommonLib.vbs ===
    Dim commonLibPath
    commonLibPath = ResolvePath("CommonLib.vbs", LEGACY_COMMONLIB_PATH, True)
    If Not IncludeFile(commonLibPath) Then
        Call LogError("Could not include CommonLib.vbs. Script will terminate.", "Init")
        If IsObject(bzhao) Then bzhao.Disconnect
        Exit Sub
    Else
        commonLibLoaded = True
    End If
    ' === End Include CommonLib.vbs ===
    Call RunMainProcess
End Sub

StartScript