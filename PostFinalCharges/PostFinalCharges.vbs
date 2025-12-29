Option Explicit


' Global script variables
Dim CSV_FILE_PATH, LOG_FILE_PATH
'Dim fso, csvStream, currentLine, roNumber
Dim fso, roNumber
Dim bzhao
Dim lastRoResult
Dim currentRODisplay
Dim commonLibLoaded
'Dim WAIT_BETWEEN_RETRIES, RETRY_COUNT, TEXT_VERIFY_DELAY, TIMEOUT_MS, POLL_INTERVAL_MS, CLOSEOUT_WAIT, CLOSEOUT_LONG_WAIT
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
Dim MainPromptLine
Dim LEGACY_CSV_PATH, LEGACY_LOG_PATH, LEGACY_DIAG_LOG_PATH, LEGACY_COMMONLIB_PATH
' Current minimum logging level (configurable)
Dim g_CurrentLogLevel

MainPromptLine = 23



' Logging level constants (lower numbers = higher priority)
Const LOG_LEVEL_CORE = 0  ' Always logged
Const LOG_LEVEL_ERROR = 1
Const LOG_LEVEL_WARN = 2  
Const LOG_LEVEL_INFO = 3
Const LOG_LEVEL_DEBUG = 4
Const LOG_LEVEL_TRACE = 5
Const DEBUG_SCREEN_LINES = 3
Const g_DiagLogQueueSize = 5
Const LEGACY_BASE_PATH = "C:\Temp\Code\Scripts\VBScript\CDK\PostFinalCharges"

' --- EARLY LOGGING: Force TRACE level for startup logs ---
g_CurrentLogLevel = 5 ' LOG_LEVEL_TRACE (ensure all TRACE logs are written at startup)



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

' --- TRACE LOGGING: Script Startup and Path Resolution (after globals/bootstrap) ---
Call Log("TRACE", "Script entrypoint reached", "Startup")
Call Log("TRACE", "About to resolve g_BaseScriptPath", "Startup")
Call Log("TRACE", "g_BaseScriptPath set to: " & g_BaseScriptPath, "Startup")
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
    Public TriggerText
    Public ResponseText
    Public KeyPress
    Public IsSuccess
    Public AcceptDefault
End Class

' Helper to create a Prompt object and add it to the dictionary
Sub AddPromptToDict(dict, trigger, response, key, isSuccess)
    Dim p
    Set p = New Prompt
    p.TriggerText = trigger
    p.ResponseText = response
    p.KeyPress = key
    p.IsSuccess = isSuccess
    p.AcceptDefault = False
    dict.Add trigger, p
End Sub

' Extended helper to create a Prompt object with AcceptDefault support
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

' Creates and returns the prompt dictionary for the line item processing sequence.
Function CreateLineItemPromptDictionary()
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")
    ' Handle end-of-sequence error
    Call AddPromptToDict(dict, "SEQUENCE NUMBER \d+ DOES NOT EXIST", "", "", True)
    Call AddPromptToDict(dict, "OPERATION CODE FOR LINE", "I", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "COMMAND:\(SEQ#/E/N/B/\?\)", "", "", False)
    Call AddPromptToDict(dict, "COMMAND:", "", "", True)
    Call AddPromptToDict(dict, "This OpCode was performed in the last 270 days.", "", "", False)
    Call AddPromptToDict(dict, "LINE CODE X IS NOT ON FILE", "", "<Enter>", True)
    Call AddPromptToDict(dict, "LABOR TYPE FOR LINE", "", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "DESC:", "", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "Enter a technician number", "", "<F3>", False)
    Call AddPromptToDictEx(dict, "TECHNICIAN \(\d+\)", "99", "<NumpadEnter>", False, True)
    Call AddPromptToDictEx(dict, "TECHNICIAN?", "99", "<NumpadEnter>", False, True)
    Call AddPromptToDictEx(dict, "TECHNICIAN \([A-Za-z0-9]+\)\?", "99", "<NumpadEnter>", False, True)
    Call AddPromptToDictEx(dict, "ACTUAL HOURS \(\d+\)", "0", "<NumpadEnter>", False, True)
    Call AddPromptToDict(dict, "SOLD HOURS?", "0", "<NumpadEnter>", False)
    Call AddPromptToDictEx(dict, "SOLD HOURS \([0-9]+\)\?", "0", "<NumpadEnter>", False, True)
    Call AddPromptToDict(dict, "ADD A LABOR OPERATION \(N\)\?", "N", "<NumpadEnter>", True)
    Call AddPromptToDict(dict, "Is this a comeback \(Y/N\)\.\.\.", "Y", "<NumpadEnter>", False)
    Call AddPromptToDict(dict, "NOT ON FILE", "", "<Enter>", True)
    Call AddPromptToDict(dict, "NOT AVAILABLE", "", "<Enter>", True)
    Call AddPromptToDict(dict, "NOT ALL LINES HAVE A COMPLETE STATUS", "", "<Enter>", True)
    Call AddPromptToDict(dict, "PRESS RETURN TO CONTINUE", "", "<Enter>", False)
    Call AddPromptToDict(dict, "Press F3 to exit.", "", "<F3>", False)

    Set CreateLineItemPromptDictionary = dict
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

' Generic state machine to process a sequence of prompts from a given dictionary.
Sub ProcessPromptSequence(prompts)
    Dim finished, promptKey, promptDetails, bestMatchKey, bestMatchLength
    finished = False

    Do While Not finished
        ' TRACE: Log screen snapshot and main prompt line before each scan
        Call LogTrace("Screen snapshot before prompt scan:", "ProcessPromptSequence")
        Call LogScreenSnapshot("BeforePromptScan")
        Dim mainPromptText
        mainPromptText = GetScreenLine(MainPromptLine)
        Call LogTrace("MainPromptLine text: '" & mainPromptText & "'", "ProcessPromptSequence")
        If Len(mainPromptText) > 0 And Not IsPromptInConfig(mainPromptText, prompts) Then
            Call LogError("Unknown prompt on line " & MainPromptLine & ": '" & mainPromptText & "' - aborting script.", "ProcessPromptSequence")
            SafeMsg "Unknown prompt detected on line " & MainPromptLine & ": '" & mainPromptText & "\nAutomation stopped for manual review.", True, "Unknown Prompt Error"
            g_ShouldAbort = True
            Exit Sub
        End If

        ' --- Find the longest (most specific) matching prompt ---
        bestMatchKey = ""
        bestMatchLength = 0
        Dim screenSnapshot
        screenSnapshot = ""
        ' Read the whole screen buffer for regex matching
        On Error Resume Next
        Dim buf
        bzhao.ReadScreen buf, 1920, 1, 1 ' 24x80
        If Err.Number <> 0 Then
            buf = ""
            Err.Clear
        End If
        On Error GoTo 0
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
                    If re.Test(buf) Then
                        Call LogTrace("Regex PromptKey detected: '" & promptKey & "'", "ProcessPromptSequence")
                        If Len(promptKey) > bestMatchLength Then
                            bestMatchKey = promptKey
                            bestMatchLength = Len(promptKey)
                        End If
                    End If
                End If
                On Error GoTo 0
            End If
            ' If not regex or regex failed, fall back to plain text
            If Not isRegex Or regexError Then
                If IsTextPresent(promptKey) Then
                    Call LogTrace("PromptKey detected: '" & promptKey & "'", "ProcessPromptSequence")
                    If Len(promptKey) > bestMatchLength Then
                        bestMatchKey = promptKey
                        bestMatchLength = Len(promptKey)
                    End If
                End If
            End If
        Next

        ' --- If a prompt was found, handle it ---
        If bestMatchLength > 0 Then
            Set promptDetails = prompts.Item(bestMatchKey)
            Call LogInfo("Matched most specific prompt: '" & bestMatchKey & "'", "ProcessPromptSequence")
            Call LogTrace("Prompt details: ResponseText='" & promptDetails.ResponseText & "', KeyPress='" & promptDetails.KeyPress & "', IsSuccess=" & promptDetails.IsSuccess, "ProcessPromptSequence")

            ' Check if this prompt should accept default values and if one is present
            Dim shouldAcceptDefault
            shouldAcceptDefault = False
            If promptDetails.AcceptDefault Then
                shouldAcceptDefault = HasDefaultValueInPrompt(bestMatchKey, buf)
                If shouldAcceptDefault Then
                    Call LogInfo("Default value detected in prompt - accepting by sending only key press", "ProcessPromptSequence")
                End If
            End If

            If promptDetails.ResponseText <> "" And Not shouldAcceptDefault Then
                Call LogTrace("Sending ResponseText: '" & promptDetails.ResponseText & "'", "ProcessPromptSequence")
                Call FastText(promptDetails.ResponseText)
            End If
            Call LogTrace("Sending KeyPress: '" & promptDetails.KeyPress & "'", "ProcessPromptSequence")
            Call FastKey(promptDetails.KeyPress)

            ' TRACE: Log screen snapshot after key send
            Call LogScreenSnapshot("AfterKeySend")

            ' Wait for the prompt to clear before rescanning
            Dim clearStart, clearElapsed
            clearStart = Timer
            Do While IsTextPresent(bestMatchKey)
                Call LogTrace("Prompt '" & bestMatchKey & "' still present after key send.", "ProcessPromptSequence")
                Call WaitMs(500)
                clearElapsed = (Timer - clearStart) * 1000
                If clearElapsed < 0 Then clearElapsed = clearElapsed + 86400000 ' Handle midnight rollover
                If clearElapsed > 5000 Then ' 5-second timeout
                    Call LogWarn("Prompt '" & bestMatchKey & "' did not clear within 5 seconds.", "ProcessPromptSequence")
                    Exit Do
                End If
            Loop

            If promptDetails.IsSuccess Then
                finished = True
                Call LogInfo("Success prompt reached: " & bestMatchKey, "ProcessPromptSequence")
                Call LogTrace("Exiting ProcessPromptSequence on success.", "ProcessPromptSequence")
            End If
            ' The loop will now naturally restart and rescan for the next prompt
        Else
            ' No prompt found, wait a moment before trying again
            Call LogTrace("No prompt found in current scan.", "ProcessPromptSequence")
            Call WaitMs(250)
        End If
    Loop
End Sub

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
    Dim key
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
                If re.Test(promptText) Then
                    IsPromptInConfig = True
                    Exit Function
                End If
            End If
            Err.Clear
            On Error GoTo 0
        Else
            ' For non-regex patterns, check for exact match or substring match
            If StrComp(Trim(promptText), Trim(key), vbTextCompare) = 0 Or InStr(1, promptText, key, vbTextCompare) > 0 Then
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
    
    ' Look for patterns like TECHNICIAN(12345)? or ACTUAL HOURS (8) in the screen content
    ' Use regex to find the actual prompt text and check if it has a non-empty value in parentheses
    On Error Resume Next
    Dim re, matches, match, parenContent
    Set re = CreateObject("VBScript.RegExp")
    
    ' Common patterns that indicate a default value is present:
    ' TECHNICIAN(12345)? - has a value in parentheses
    ' ACTUAL HOURS (8) - has a value in parentheses
    ' SOLD HOURS (10)? - has a value in parentheses
    
    ' Extract the base pattern and look for it with actual values
    If InStr(promptPattern, "TECHNICIAN") > 0 Then
        re.Pattern = "TECHNICIAN\s*\(([A-Za-z0-9]+)\)"
    ElseIf InStr(promptPattern, "ACTUAL HOURS") > 0 Then
        re.Pattern = "ACTUAL HOURS\s*\(([0-9]+)\)"
    ElseIf InStr(promptPattern, "SOLD HOURS") > 0 Then
        re.Pattern = "SOLD HOURS\s*\(([0-9]+)\)"
    Else
        ' Generic pattern for any prompt with parentheses containing a value
        re.Pattern = "\([A-Za-z0-9]+\)"
    End If
    
    re.IgnoreCase = True
    re.Global = False
    
    If Err.Number = 0 Then
        Set matches = re.Execute(screenContent)
        If matches.Count > 0 Then
            Set match = matches(0)
            If match.SubMatches.Count > 0 Then
                parenContent = Trim(match.SubMatches(0))
                ' If there's content in parentheses and it's not empty or just question marks
                If Len(parenContent) > 0 And parenContent <> "?" And parenContent <> "" Then
                    HasDefaultValueInPrompt = True
                    Call LogTrace("Found default value in prompt: " & parenContent, "HasDefaultValueInPrompt")
                End If
            End If
        End If
    End If
    
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
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
    ' We log them as DEBUG, and errors are parsed and logged at the ERROR level.
    If Left(UCase(Trim(logMsg)), 6) = "ERROR:" Then
        Call Log("ERROR", "ADAPTER: " & logMsg, "CommonLib")
    Else
        Call Log("DEBUG", "ADAPTER: " & logMsg, "CommonLib")
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
        Call LogInfo("CommonLib loaded successfully in PostFinalCharges.vbs", "RunMainProcess")
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
        Call LogError("IncludeFile - File not found: " & filePath, "IncludeFile")
        IncludeFile = False
        Exit Function
    End If

    Set includeStream = fsoInclude.OpenTextFile(filePath, 1)
    fileContent = includeStream.ReadAll
    includeStream.Close
    Set includeStream = Nothing

    ExecuteGlobal fileContent

    If Err.Number <> 0 Then
        Call LogError("IncludeFile - Error executing file '" & filePath & "': " & Err.Description & " (" & Err.Number & ", " & Err.Source & ")", "IncludeFile")
        Err.Clear
        IncludeFile = False
        Exit Function
    End If
    On Error GoTo 0
    IncludeFile = True
End Function




Function ResolvePath(targetPath, defaultPath, mustExist)
    Dim fsoLocal, basePath, candidate, hasDefault, requireExists
    Set fsoLocal = CreateObject("Scripting.FileSystemObject")

    hasDefault = (Len(CStr(defaultPath)) > 0)
    requireExists = CBool(mustExist)

    If IsAbsolutePath(targetPath) Then
        candidate = targetPath
    Else
        basePath = GetBaseScriptPath()
        If Len(basePath) > 0 Then
            On Error Resume Next
            candidate = fsoLocal.BuildPath(basePath, targetPath)
            If Err.Number <> 0 Then
                candidate = ""
                Err.Clear
            End If
            On Error GoTo 0
        End If

        If Len(candidate) = 0 Then
            candidate = fsoLocal.GetAbsolutePathName(targetPath)
        End If
    End If

    If requireExists Then
        If Not PathExists(fsoLocal, candidate) Then
            If hasDefault Then
                ResolvePath = defaultPath
            Else
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
    Call LogDebug("Reading INI: " & configPath, "GetIniSetting")

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(configPath) Then
        Call LogDebug("INI file not found at: " & configPath, "GetIniSetting")
        GetIniSetting = defaultValue
        Exit Function
    End If
    
    Set file = fso.OpenTextFile(configPath, 1)
    If Err.Number <> 0 Then
        Call LogWarn("Could not open INI file: " & configPath, "GetIniSetting")
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
                Call LogDebug("Entered section [" & section & "]", "GetIniSetting")
            ElseIf inSection Then
                ' We have passed the relevant section, so we can stop.
                Call LogDebug("Exited section [" & section & "]", "GetIniSetting")
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
                Call LogDebug("Found key '" & key & "' with value '" & result & "'", "GetIniSetting")
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
            Call LogInfo("Using MockBzhao for testing", "InitializeObjects")
            
            ' Setup initial test scenario
            bzhao.SetupTestScenario("basic_command_prompt")
        Else
            Call LogError("Could not load MockBzhao.vbs for test mode", "InitializeObjects")
            g_IsTestMode = False
        End If
    End If
    
    If Not g_IsTestMode Then
        Set bzhao = CreateObject("BZWhll.WhllObj")
        If Err.Number <> 0 Then
            Call LogError("Failed to create BZWhll.WhllObj: " & Err.Description, "InitializeObjects")
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
    ' --- Initialize Log Level from INI file FIRST ---
    Dim logLevelValue
    logLevelValue = UCase(GetIniSetting("Settings", "LogLevel", "WARN"))
    g_CurrentLogLevel = LOG_LEVEL_WARN ' Default
    
    Select Case logLevelValue
        Case "ERROR": g_CurrentLogLevel = LOG_LEVEL_ERROR
        Case "WARN": g_CurrentLogLevel = LOG_LEVEL_WARN
        Case "INFO": g_CurrentLogLevel = LOG_LEVEL_INFO
        Case "DEBUG": g_CurrentLogLevel = LOG_LEVEL_DEBUG
        Case "TRACE": g_CurrentLogLevel = LOG_LEVEL_TRACE
    End Select

    ' --- Now load other settings ---
    g_DefaultWait = GetIniSetting("Settings", "DefaultWait", 1000)
    g_PromptWait = GetIniSetting("Settings", "PromptWait", 5000)
    
    Dim startSequenceNumberValue, endSequenceNumberValue
    startSequenceNumberValue = GetIniSetting("Processing", "StartSequenceNumber", "")
    endSequenceNumberValue = GetIniSetting("Processing", "EndSequenceNumber", "")

    If startSequenceNumberValue = "" Or endSequenceNumberValue = "" Then
        Call LogError("Critical config missing: 'StartSequenceNumber' and/or 'EndSequenceNumber' not found in config.ini. Aborting run.", "InitializeConfig")
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
    'POLL_INTERVAL_MS = 150
    POST_PROMPT_WAIT_MS = 150
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
    ' This function now only handles overrides via environment variables.
    ' Default log level is set in InitializeConfig from the INI file.
    
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
        g_CurrentLogLevel = LOG_LEVEL_TRACE ' Override to most verbose
    End If
    
    ' --- PFC_LOG_LEVEL override ---
    ' Allows specific log level setting, overriding both INI and PFC_DEBUG.
    Dim logLevelEnvValue
    logLevelEnvValue = UCase(Trim(shell.Environment("PROCESS")("PFC_LOG_LEVEL")))
    If Len(logLevelEnvValue) = 0 Then logLevelEnvValue = UCase(Trim(shell.Environment("USER")("PFC_LOG_LEVEL")))
    If Len(logLevelEnvValue) = 0 Then logLevelEnvValue = UCase(Trim(shell.Environment("SYSTEM")("PFC_LOG_LEVEL")))
    
    If Len(logLevelEnvValue) > 0 Then
        Select Case logLevelEnvValue
            Case "ERROR": g_CurrentLogLevel = LOG_LEVEL_ERROR
            Case "WARN": g_CurrentLogLevel = LOG_LEVEL_WARN
            Case "INFO": g_CurrentLogLevel = LOG_LEVEL_INFO
            Case "DEBUG": g_CurrentLogLevel = LOG_LEVEL_DEBUG
            Case "TRACE": g_CurrentLogLevel = LOG_LEVEL_TRACE
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
        Call LogError("BlueZone object is not available (CreateObject failed).", "ConnectBlueZone")
        ConnectBlueZone = False
        Exit Function
    End If
    
    bzhao.Connect ""
    If Err.Number <> 0 Then
        Call LogError("BlueZone connection failed: " & Err.Description, "ConnectBlueZone")
        Err.Clear
        ConnectBlueZone = False
    Else
        Call LogInfo("Connected to BlueZone.", "ConnectBlueZone")
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
        Call LogInfo(sequenceLabel & " - Processing", "ProcessRONumbers")
        
        lastRoResult = ""
        Call Main(roNumber)
        
        Call LogInfo(sequenceLabel & " - Result: " & lastRoResult, "ProcessRONumbers")
        Call LogInfo("Test mode: Processed single RO " & roNumber, "ProcessRONumbers")
        Exit Sub
    End If

    For roNumber = g_StartSequenceNumber To g_EndSequenceNumber
        lineCount = lineCount + 1
        'WaitMs(2000)
        Call LogROHeader(roNumber)
        sequenceLabel = "Sequence " & roNumber
        Call LogInfo(sequenceLabel & " - Processing", "ProcessRONumbers")

        ' Start performance timing for this RO
        Dim roStartTime
        roStartTime = Now

        lastRoResult = ""
        Call Main(roNumber)

        ' Check for end-of-sequence error
        If IsTextPresent("SEQUENCE NUMBER " & roNumber & " DOES NOT EXIST") Then
            Call LogError("End of sequence detected: SEQUENCE NUMBER " & roNumber & " DOES NOT EXIST. Stopping script.", "ProcessRONumbers")
            Exit Sub
        End If

        ' Calculate and log performance timing for successful closures only
        If InStr(1, LCase(lastRoResult), "successfully closed") > 0 Then
            Dim roEndTime, roDuration
            roEndTime = Now
            roDuration = DateDiff("s", roStartTime, roEndTime)
            Call LogInfo(sequenceLabel & " - E2E Duration: " & roDuration & " seconds", "ProcessRONumbers")
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
            Call LogError(errorLabel & " - " & lastRoResult, "ProcessRONumbers")
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
        Call LogCore(finalMessage, "ProcessRONumbers")
        ' Always write the scraped RO status to the core log for troubleshooting
        Dim statusForLog
        statusForLog = Trim(CStr(g_LastScrapedStatus))
        If Len(statusForLog) = 0 Then statusForLog = "(none)"
        Call LogCore("RO STATUS FOUND: " & statusForLog, "ProcessRONumbers")

        If (lineCount Mod 10) = 0 Then
            Call LogInfo("Processed " & lineCount & " ROs...", "ProcessRONumbers")
        End If

        If g_ShouldAbort Then
            Call LogError("Aborting sequence processing. Reason: " & g_AbortReason, "ProcessRONumbers")
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
            Call LogError("Failed to open log file: " & LOG_FILE_PATH, "LogROHeader")
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
    ' Scrape the actual RO number from the screen (top of screen shows 'RO:  123456')
    Dim actualRO
    actualRO = GetROFromScreen()
    If Len(Trim(CStr(actualRO))) > 0 Then
        currentRODisplay = actualRO
    Else
        currentRODisplay = roNumber
    End If
    
    If Len(Trim(CStr(currentRODisplay))) > 0 Then
        Call LogInfo("Sent RO to BlueZone", "Main")
    Else
        ' No scraped RO available; log against the sequence number and note unknown RO
        Call LogInfo(roNumber & " - Sent RO to BlueZone - RO: (unknown) - will use sequence number for checks", "Main")
    End If
    
    ' Check for "closed" response
    If IsTextPresent("Repair Order " & currentRODisplay & " is closed.") Then
        Call LogInfo("Repair Order Closed", "Main")
        lastRoResult = "Closed"
        Exit Sub
    End If
    
    ' Check for "NOT ON FILE" response
    If IsTextPresent("NOT ON FILE") Then
        Call LogInfo("Not On File", "Main")
        lastRoResult = "Not On File"
        Exit Sub
    End If
    
    ' Otherwise, assume repair order is open — prefer the scraped RO for logging
    If Len(Trim(CStr(currentRODisplay))) > 0 Then
        Call LogInfo("Repair Order Open", "Main")
    Else
        Call LogInfo(roNumber & " - Repair Order Open", "Main")
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
        Call LogInfo("RO STATUS: READY TO POST", "Main")
    End If
    
    ' Snapshot the scraped status now to avoid timing races, then detect triggers.
    Dim trigger, roStatusForDecision
    roStatusForDecision = Trim(CStr(g_LastScrapedStatus))
    Call LogDebug("Pre-trigger check - scraped status: '" & roStatusForDecision & "'", "Main")
    trigger = FindTrigger()
    If trigger <> "" Then
        Call LogInfo("Trigger found: " & trigger & " - Proceeding to Closeout", "Main")
        Call Closeout_Ro()
        ' Closeout_Ro should set lastRoResult appropriately
    Else
        ' If no trigger text found, but the scraped RO status is READY TO POST,
        ' proceed to closeout anyway (status supersedes trigger text).
        If StrComp(roStatusForDecision, "READY TO POST", vbTextCompare) = 0 Then
            Call LogInfo("No closeout trigger text found, but RO STATUS is READY TO POST — proceeding to Closeout", "Main")
            Call Closeout_Ro()
        Else
            Call LogInfo("No Closeout Text Found - Skipping Closeout", "Main")
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
' **FUNCTION NAME:** Log
' **DATE CREATED:** 2025-11-19
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Unified logging interface with level-based filtering.
' Automatically filters messages based on current log level.
' 
' **PARAMETERS:**
' level (String): Log level - "CORE", "ERROR", "WARN", "INFO", "DEBUG", "TRACE"
' message (String): The message to log
' source (String): The source component/function
'-----------------------------------------------------------------------------------
Sub Log(level, message, source)
    Dim levelValue
    
    ' Convert level string to numeric value
    Select Case UCase(level)
        Case "CORE": levelValue = LOG_LEVEL_CORE
        Case "ERROR": levelValue = LOG_LEVEL_ERROR
        Case "WARN": levelValue = LOG_LEVEL_WARN
        Case "INFO": levelValue = LOG_LEVEL_INFO
        Case "DEBUG": levelValue = LOG_LEVEL_DEBUG
        Case "TRACE": levelValue = LOG_LEVEL_TRACE
        Case Else: levelValue = LOG_LEVEL_INFO ' Default to INFO for unknown levels
    End Select
    
    ' Only log if the message's level is less than or equal to the current configured level
    If levelValue <= g_CurrentLogLevel Then
        Call WriteLogEntry(level, message, source)
    End If
End Sub

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** LogError
' **DATE CREATED:** 2025-11-19
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Convenience function for ERROR level logging.
'-----------------------------------------------------------------------------------
Sub LogError(message, source)
    Call Log("ERROR", message, source)
End Sub

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** LogWarn
' **DATE CREATED:** 2025-11-19
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Convenience function for WARN level logging.
'-----------------------------------------------------------------------------------
Sub LogWarn(message, source)
    Call Log("WARN", message, source)
End Sub

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** LogInfo
' **DATE CREATED:** 2025-11-19
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Convenience function for INFO level logging.
'-----------------------------------------------------------------------------------
Sub LogInfo(message, source)
    Call Log("INFO", message, source)
End Sub

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** LogDebug
' **DATE CREATED:** 2025-11-19
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Convenience function for DEBUG level logging.
'-----------------------------------------------------------------------------------
Sub LogDebug(message, source)
    Call Log("DEBUG", message, source)
End Sub

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** LogTrace
' **DATE CREATED:** 2025-11-19
' **AUTHOR:** GitHub Copilot
' 
' **FUNCTIONALITY:**
' Convenience function for TRACE level logging.
'-----------------------------------------------------------------------------------
Sub LogTrace(message, source)
    Call Log("TRACE", message, source)
End Sub

'-----------------------------------------------------------------------------------
' **FUNCTION NAME:** LogCore
' **DATE CREATED:** 2025-11-28
' **AUTHOR:** Gemini
' 
' **FUNCTIONALITY:**
' Logs essential messages by using the CORE log level, which is always output.
'-----------------------------------------------------------------------------------
Sub LogCore(message, source)
    Call Log("CORE", message, source)
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
    Call Log(level, message, source)
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
    Call Log(level, message, source)
End Sub

Sub WriteLogEntry(level, message, source)
    Dim logFSO, logFile, logLine, logFolder

    logLine = Now & " [" & level & "] [" & source & "] " & message

    On Error Resume Next
    Set logFSO = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    'Dim logFolder
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
        Call LogError(text, "SafeMsg")
    Else
        Call LogInfo(text, "SafeMsg")
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
        Call LogError("bzhao object is not available", "GetROFromScreen")
        GetROFromScreen = ""
        Exit Function
    End If

    Dim screenContentBuffer, screenLength, re, matches
    screenLength = 3 * 80 ' top three lines
    On Error Resume Next
    bzhao.ReadScreen screenContentBuffer, screenLength, 1, 1
    If Err.Number <> 0 Then
        Call LogError("GetROFromScreen ReadScreen failed: " & Err.Description, "GetROFromScreen")
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
    End If
End Function


'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** IsStatusReady
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' Checks if the current screen indicates that the Repair Order status is
' "READY TO POST". It performs an exact, case-insensitive search for this
' specific string to ensure the RO is in the correct state to proceed.
' 
' 
' **RETURN VALUE:**
' (Boolean) Returns True if the status is "READY TO POST", False otherwise.
'-----------------------------------------------------------------------------------
Function IsStatusReady()
    ' Use GetRepairOrderStatus() to scrape the exact RO status from the screen
    ' Caller may choose to add waits before calling if needed
    bzhao.pause 1000 ' brief pause to ensure screen is stable
    Dim roStatus
    roStatus = GetRepairOrderStatus()
    ' Return True only when the scraped status exactly equals "READY TO POST"
    IsStatusReady = (StrComp(Trim(CStr(roStatus)), "READY TO POST", vbTextCompare) = 0)
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
        Call LogWarn("GetRepairOrderStatus: bzhao object not available", "GetRepairOrderStatus")
        GetRepairOrderStatus = ""
        Exit Function
    End If

    Dim buf, lengthToRead, lineNum, colNum
    lengthToRead = 30
    lineNum = 5
    colNum = 1
    bzhao.ReadScreen buf, lengthToRead, lineNum, colNum
    If Err.Number <> 0 Then
        Call LogWarn("GetRepairOrderStatus: ReadScreen failed: " & Err.Description, "GetRepairOrderStatus")
        Err.Clear
        GetRepairOrderStatus = ""
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Call LogDebug("GetRepairOrderStatus - raw buffer: '" & Replace(buf, vbCrLf, " ") & "'", "GetRepairOrderStatus")

    Dim prefix, pos, raw
    prefix = "RO STATUS: "
    pos = InStr(1, buf, prefix, vbTextCompare)
    If pos = 0 Then
        ' Not found in this slice
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
    Call LogDebug("GetRepairOrderStatus - parsed status: '" & parsedStatus & "'", "GetRepairOrderStatus")
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
        If elapsedMs < 0 Then elapsedMs = elapsedMs + 86400000 ' Handle midnight rollover
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
        Call LogDebug("LogScreenSnapshot failed to read screen: " & Err.Description, "LogScreenSnapshot")
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
    Call LogDebug("ScreenSnapshot(" & name & "): " & snippet, "LogScreenSnapshot")
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

    For i = 65 To 90 ' ASCII for A to Z
        lineLetterChar = Chr(i)
        ' Wait for the COMMAND prompt and then enter "R" + the current line letter.
        Call WaitForPrompt("COMMAND:", "R " & lineLetterChar, True, g_PromptWait, "")
        ' Check if the line exists. If not, we are done with line processing.
        If IsTextPresent("LINE CODE " & lineLetterChar & " IS NOT ON FILE") Then
            Call LogInfo("Finished processing line items. No more lines found.", "ProcessLineItems")
            ' Press Enter to clear the "NOT ON FILE" message from the screen.
            Call FastKey("<Enter>")
            Exit For ' Exit the For loop.
        End If
        ' Use the new state machine method for all prompt handling
        Call LogInfo("Processing line item " & lineLetterChar & " using ProcessSingleLine_Dynamic", "ProcessLineItems")
        Call LogDetailed("INFO", "Processing line item " & lineLetterChar & " using ProcessPromptSequence", "ProcessLineItems")

        Call ProcessPromptSequence(lineItemPrompts)
    Next
End Sub

'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** Closeout_Ro
' **DATE CREATED:** 2025-11-04
' **AUTHOR:** Dirk Steele
' 
' 
' **FUNCTIONALITY:**
' Automates the sequence of steps required to close out a Repair Order (RO).
' This involves first processing each line item, and then navigating through 
' final prompts to approve and close the RO.
'-----------------------------------------------------------------------------------
Sub Closeout_Ro()
    ' Process all line items before proceeding to final closeout.
    Call ProcessLineItems

    ' Send the Final Closeout (FC) command
    WaitForPrompt "COMMAND:", "FC", True, g_PromptWait, ""
    If HandleCloseoutErrors() Then Exit Sub

    ' ALL LABOR POSTED
    Dim send_enter_key_all_labor_posted
    send_enter_key_all_labor_posted = True

    ' Add a 3 second delay before ALL LABOR POSTED prompt
    Call WaitMs(3000)

    
    ' DEBUG: Show a message box for diagnostic purposes
    ' If Not bzhao Is Nothing Then
    '     bzhao.MsgBox "DEBUG: At ALL LABOR POSTED prompt in Closeout_Ro"
    ' End If
    If Not WaitForPrompt("ALL LABOR POSTED", "Y", send_enter_key_all_labor_posted, g_TimeoutMs, "") Then
        Call LogError("Failed to get ALL LABOR POSTED prompt - aborting closeout", "Closeout_Ro")
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
    If WaitForPrompt("Current Mileage less than Previous Mileage", "", False, 5000, "") Then ' Increased timeout for optional prompt
        Call LogInfo("Detected 'Current Mileage less than Previous Mileage' prompt. Sending 'Y'.", "Closeout_Ro")
        Dim send_enter_key_mileage_less
        send_enter_key_mileage_less = True
            Call WaitForPrompt("Current Mileage less than Previous Mileage", "Y", send_enter_key_mileage_less, g_DefaultWait, "")
        ' After sending Y, another error might appear, so we should check for errors again.
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
        Call LogError("Failed to get O.K. TO CLOSE RO prompt - aborting closeout", "Closeout_Ro")
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
        Call LogError("Failed to get INVOICE PRINTER prompt - closeout may be incomplete", "Closeout_Ro")
        lastRoResult = "Failed - Could not send to printer"
        Exit Sub
    End If
    ' Use the state machine for the rest of the closeout prompts
    Dim closeoutPrompts
    Set closeoutPrompts = CreateCloseoutPromptDictionary()
    ' Add a 5 second delay before processing the final closeout prompts
    ' Wait for the continue prompt using WaitForTextSilent directly.
    Dim continuePromptTimeout
    continuePromptTimeout = 10000 ' 10 seconds, adjust as needed

    Dim continuePromptDetected
    continuePromptDetected = WaitForTextSilent("COMMAND:(SEQ#/E/N/B/?)", continuePromptTimeout)
    If continuePromptDetected Then
        Call LogInfo("Detected continue prompt: COMMAND:(SEQ#/E/N/B/?)", "Closeout_Ro")
    Else
        Call LogWarn("Timeout waiting for continue prompt: COMMAND:(SEQ#/E/N/B/?)", "Closeout_Ro")
    End If
    Call ProcessPromptSequence(closeoutPrompts)

    ' Give the terminal a short moment to surface any follow-up messages before scanning for errors.
    'Call WaitMs(2000)
    If HandleCloseoutErrors() Then Exit Sub

    lastRoResult = "Successfully closed"
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
            If clearElapsed < 0 Then clearElapsed = clearElapsed + 86400000
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
        If IsTextPresent(candidate) Then
            FindTrigger = candidate
            Exit Function
        End If
    Next
    
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
        If elapsedMs < 0 Then elapsedMs = elapsedMs + 86400000 ' Handle midnight rollover
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