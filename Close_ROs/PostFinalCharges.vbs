Option Explicit
'=============================================
' TODO:
' - Remove hardcoded paths
' - Centralize wait/retry values
' - Validate all user input
' - Improve error reporting
' - Remove WScript.Quit usage
' - Add config file support
' - Check bzhao before use
' - Option for log file path
' - Document all constants
'=============================================

' Replace generic declarations with clearer names and defer creation to initializer.
Dim CSV_FILE_PATH
Dim LOG_FILE_PATH
Dim fso, csvStream, currentLine, roNumber
Dim bzhao
Dim lastRoResult
Dim currentRODisplay

Dim g_DefaultWait, g_LongWait, g_SendRetryCount

' -- Prompt Detection Constants --
Const POLL_INTERVAL = 100   ' Check every 100ms (10 times per second)
Const POST_ENTRY_WAIT = 200  ' Minimal wait after entry
Const PRE_KEY_WAIT = 150     ' Pause before sending special keys
Const POST_KEY_WAIT = 350    ' Pause after sending special keys
Const PROMPT_TIMEOUT_MS = 10000 ' Default prompt timeout
Const DelayTimeAfterPromptDetection = 500 ' Delay after prompt detection before sending input


'-----------------------------------------------------------
' Define file paths and connect to BlueZone
'-----------------------------------------------------------

'------------------------------
' Initialization and main flow
'------------------------------
Call InitializeObjects()
If ConnectBlueZone() Then
    ProcessCSV()
Else
    SafeMsg "Unable to connect to BlueZone. Check that it’s open and logged in.", True, "Connection Error"
End If
Call FlushLogBuffer() ' Flush any remaining log messages before script ends

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

' Initialize required objects (FileSystemObject and BlueZone instance)
Sub InitializeObjects()
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set bzhao = CreateObject("BZWhll.WhllObj")
    If Err.Number <> 0 Then
        Call LogResultWithSource("ERROR", "Failed to create BZWhll.WhllObj: " & Err.Description, "InitializeObjects")
        Err.Clear
    End If
    Call InitializeConfig
End Sub

Sub InitializeConfig()
    ' Hardcode file paths
    CSV_FILE_PATH = "C:\Temp\Code\Scripts\VBScript\CDK\Close_ROs\Close_ROs_Pt1.csv"
    LOG_FILE_PATH = "C:\Temp\Code\Scripts\VBScript\CDK\Close_ROs\PostFinalCharges.log"

    ' Default wait times (can be moved to constants if not configurable)
    g_DefaultWait = 1000
    g_LongWait = 2000
    g_SendRetryCount = 2
End Sub

Function ReadIniValue(filePath, section, key, defaultValue)
    Dim file, line, inSection, value
    inSection = False
    value = ""

    Set file = fso.OpenTextFile(filePath, 1)

    Do While Not file.AtEndOfStream
        line = Trim(file.ReadLine)

        If Left(line, 1) = "[" And Right(line, 1) = "]" Then
            If LCase(Trim(Mid(line, 2, Len(line) - 2))) = LCase(section) Then
                inSection = True
            Else
                inSection = False
            End If
        ElseIf inSection And InStr(line, "=") > 0 Then
            Dim parts
            parts = Split(line, "=", 2)
            If LCase(Trim(parts(0))) = LCase(key) Then
                value = Trim(parts(1))
                Exit Do
            End If
        End If
    Loop

    file.Close

    If value = "" Then
        ReadIniValue = defaultValue
    Else
        ReadIniValue = value
    End If
End Function

' DetermineDebugMode: if a file named Cashout_ROs.debug exists next to the script, enable DEBUG_LOGGING



'=============================================
' PROMPT DETECTION SUBROUTINES
'=============================================

'--------------------------------------------------------------------
' Subroutine: WaitForPrompt - Requires 4 parameters: promptText, valueToEnter, sendEnter, timeoutMs
'--------------------------------------------------------------------
Sub WaitForPrompt(promptText, valueToEnter, sendEnter, timeoutMs)
    
    Call LogResultWithSource("INFO", "WaitForPrompt called - Looking for: [" & promptText & "] Value: [" & valueToEnter & "] SendEnter: " & sendEnter & " Timeout: " & timeoutMs & "ms", "WaitForPrompt")
    Dim startTime, currentTime, elapsedMs, promptFound
    
    startTime = Timer 
    promptFound = False
    
    Do
        ' Check for the prompt text first
        If IsTextPresent(promptText) Then
            Call LogResultWithSource("INFO", "Detected prompt: " & promptText, "WaitForPrompt")
            promptFound = True
            Exit Do
        End If
        
        ' Wait a bit before checking again
        Call WaitMs(POLL_INTERVAL)
        
        ' Calculate elapsed time and check timeout
        currentTime = Timer
        If currentTime < startTime Then currentTime = currentTime + 86400
        elapsedMs = (currentTime - startTime) * 1000
        
        ' Exit if timeout reached
        If elapsedMs >= timeoutMs Then
            Call LogResultWithSource("ERROR", "Timeout waiting for prompt: " & promptText, "WaitForPrompt")
            Exit Do
        End If
    Loop
    
    ' Only send input if prompt was actually found
    If promptFound Then
        
        ' Check if the value is a special key command
        bzhao.Pause DelayTimeAfterPromptDetection
        If InStr(1, valueToEnter, "<") > 0 And InStr(1, valueToEnter, ">") > 0 Then
            Call LogResultWithSource("INFO", "Sending key command: " & valueToEnter, "WaitForPrompt")
            Call FastKey(valueToEnter)
        Else
            Call FastText(valueToEnter)
        End If
        
        If sendEnter Then
            Call FastKey("<NumpadEnter>")
        End If
        
        Call WaitMs(POST_ENTRY_WAIT)
    Else
        Call LogResultWithSource("ERROR", "Prompt not found - skipping input for prompt " & promptText, "WaitForPrompt")
    End If
End Sub


'--------------------------------------------------------------------
' Subroutine: FastText - Minimal delay text entry
'--------------------------------------------------------------------
Sub FastText(text)
    Call LogResultWithSource("INFO", "Sending text: " & text, "FastText")
    bzhao.SendKey text
    Call WaitMs(100)
End Sub

'--------------------------------------------------------------------
' Subroutine: FastKey - Minimal delay key press
'--------------------------------------------------------------------
Sub FastKey(key)
    Call LogResultWithSource("INFO", "Sending key command: " & key, "FastKey")
    ' Pause briefly before sending a special key to avoid injecting escape sequences into active fields
    Call WaitMs(PRE_KEY_WAIT)
    bzhao.SendKey key

End Sub

'--------------------------------------------------------------------
' Subroutine: WaitMs - Optimized waiting
'--------------------------------------------------------------------
Sub WaitMs(ms)
    If ms <= 0 Then Exit Sub
    
    Dim startTime, endTime, waitSeconds
    waitSeconds = ms / 1000
    startTime = Timer
    endTime = startTime + waitSeconds

    If endTime > 86400 Then
        endTime = endTime - 86400
        Do While Timer >= startTime Or Timer < endTime
        Loop
    Else
        Do While Timer < endTime
        Loop
    End If
End Sub

'=============================================
' END PROMPT DETECTION SUBROUTINES
'=============================================

' Attempt BlueZone connection with scoped error handling.

Function ConnectBlueZone()
    On Error Resume Next
    If bzhao Is Nothing Then
        Call LogResultWithSource("ERROR", "BlueZone object is not available (CreateObject failed).", "ConnectBlueZone")
        ConnectBlueZone = False
        Exit Function
    End If
    
    bzhao.Connect ""
    If Err.Number <> 0 Then
        Call LogResultWithSource("ERROR", "BlueZone connection failed: " & Err.Description, "ConnectBlueZone")
        Err.Clear
        ConnectBlueZone = False
    Else
        Call LogResultWithSource("INFO", "Connected to BlueZone.", "ConnectBlueZone")
        ConnectBlueZone = True
    End If
    On Error GoTo 0
End Function



' Read/process CSV; uses clearer variable names and closes file handles.
Sub ProcessCSV()
    If Not fso.FileExists(CSV_FILE_PATH) Then
        Call LogResultWithSource("ERROR", "CSV file not found: " & CSV_FILE_PATH, "ProcessCSV")
        SafeMsg "Error: The file '" & CSV_FILE_PATH & "' was not found.", True, "File Not Found"
        Exit Sub
    End If
    
    Set csvStream = fso.OpenTextFile(CSV_FILE_PATH, 1)
    If Err.Number <> 0 Then
        Call LogResultWithSource("ERROR", "Failed to open CSV file: " & Err.Description, "ProcessCSV")
        Err.Clear
        Exit Sub
    End If
    If Not csvStream.AtEndOfStream Then
        csvStream.ReadLine  ' Skip header row if present
    End If
    
    Dim lineCount
    lineCount = 0
    
    Do While Not csvStream.AtEndOfStream
        currentLine = csvStream.ReadLine
        roNumber = Trim(currentLine)
        lineCount = lineCount + 1
        
        Call LogROHeader(roNumber)
        Call LogResultWithSource("INFO", roNumber & " - Processing Sequence: " & roNumber, "ProcessCSV")
        
        lastRoResult = ""
        Call Main(roNumber)
        If Err.Number <> 0 Then
            lastRoResult = "Error in Main: " & Err.Description
            ' Prefer the scraped RO for error/result logging when available
            Dim displayId
            If Len(Trim(CStr(currentRODisplay))) > 0 Then
                displayId = currentRODisplay
            Else
                displayId = roNumber
            End If
            Call LogResultWithSource("ERROR", displayId & " - " & lastRoResult, "ProcessCSV")
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
        Call LogResultWithSource("INFO", finalDisplay & " - Result: " & lastRoResult, "ProcessCSV")
    

    Loop
    Call FlushLogBuffer() ' Final flush after all ROs are processed
End Sub

' Writes a header to the log for readability
Sub LogROHeader(ro)
    Dim logFSO, logFile, sep
    sep = "==================="
    Set logFSO = CreateObject("Scripting.FileSystemObject")
    Set logFile = logFSO.OpenTextFile(LOG_FILE_PATH, 8, True)
    logFile.WriteLine sep
    logFile.WriteLine "Sequence: " & CStr(ro)
    logFile.WriteLine sep
    logFile.Close
    Set logFile = Nothing
    Set logFSO = Nothing
End Sub

'-----------------------------------------------------------
' Main subroutine to check and process each RO number
' (renamed parameter to roNumber for clarity)
'-----------------------------------------------------------
Sub Main(roNumber)
    '==== INPUT POINT 1: BEFORE ENTERING RO Number ====
    ' NEED TO IDENTIFY: What prompt appears when CDK is ready for RO Number?
    Call WaitForPrompt("COMMAND: (SEQ#", roNumber, True, g_LongWait)
    ' Scrape the actual RO number from the screen (top of screen shows 'RO:  123456')
    Dim actualRO
    actualRO = GetROFromScreen()
    If Len(Trim(CStr(actualRO))) > 0 Then
        currentRODisplay = actualRO
    Else
        currentRODisplay = roNumber
    End If
    
    If Len(Trim(CStr(currentRODisplay))) > 0 Then
        Call LogResultWithSource("INFO", "Sent RO to BlueZone", "Main")
    Else
        ' No scraped RO available; log against the sequence number and note unknown RO
        Call LogResultWithSource("INFO", roNumber & " - Sent RO to BlueZone - RO: (unknown) - will use sequence number for checks", "Main")
    End If
    
    ' Check for "closed" response
    If IsTextPresent("Repair Order " & currentRODisplay & " is closed.") Then
        Call LogResultWithSource("INFO", "Repair Order Closed", "Main")
        lastRoResult = "Closed"
        Exit Sub
    End If
    
    ' Check for "NOT ON FILE" response
    If IsTextPresent("NOT ON FILE") Then
        Call LogResultWithSource("INFO", "Not On File", "Main")
        lastRoResult = "Not On File"
        Exit Sub
    End If
    
    ' Otherwise, assume repair order is open — prefer the scraped RO for logging
    If Len(Trim(CStr(currentRODisplay))) > 0 Then
        Call LogResultWithSource("INFO", "Repair Order Open", "Main")
    Else
        Call LogResultWithSource("INFO", roNumber & " - Repair Order Open", "Main")
    End If
    
    ' After opening an RO, ensure it has the expected READY TO POST status.
    If Not IsStatusReady() Then
        Call LogResultWithSource("INFO", "RO STATUS not READY TO POST - exiting (E) and moving to next", "Main")
        Call WaitForPrompt("", "E", True, 2000)
        lastRoResult = "Skipped - Status not ready"
        Exit Sub
    Else
        Call LogResultWithSource("INFO", "RO STATUS: READY TO POST", "Main")
    End If
    
    ' Define closeout triggers locally for explicit checking
    Dim closeoutTriggers
    closeoutTriggers = Array( _
    "CHECK AND ADJUST TIRE PRESSURE", _
    "REPLACE TIRE SENSOR", _
    "PM CHANGE OIL & FILTER", _
    "PMS PERFORMED ON LOT" _
    )

    ' Use explicit checks for closeout triggers
    Dim triggerFound
    triggerFound = False
    Dim currentTrigger

    For Each currentTrigger In closeoutTriggers
        If IsTextPresent(currentTrigger) Then
            Call LogResultWithSource("INFO", "Trigger found: " & currentTrigger & " - Proceeding to Closeout", "Main")
            triggerFound = True
            Exit For
        End If
    Next

    If triggerFound Then
        Call Closeout_Ro()
        ' Closeout_Ro should set lastRoResult appropriately
    Else
        Call LogResultWithSource("INFO", "No Closeout Text Found - Skipping Closeout", "Main")
        Call WaitForPrompt("", "E", True, 2000) ' Send 'E' to exit if no trigger found
        lastRoResult = "Skipped - No closeout text found"
    End If
End Sub

' Improved logging function — accepts either (level, message)
' or (key, message) (back-compat with existing calls)
'-----------------------------------------------------------
Sub LogResult(p1, p2)
    Dim level, message, fso, logFile, logLine

    If UCase(CStr(p1)) = "INFO" Or UCase(CStr(p1)) = "ERROR" Or UCase(CStr(p1)) = "CLOSEOUT" Or UCase(CStr(p1)) = "DEBUG" Then
        level = p1
        message = p2
    Else
        level = "INFO"
        message = CStr(p1) & " - " & CStr(p2)
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set logFile = fso.OpenTextFile(LOG_FILE_PATH, 8, True)
    If Err.Number = 0 Then
        logLine = Now & " [" & level & "] " & message
        logFile.WriteLine logLine
        logFile.Close
    Else
        Err.Clear
    End If
    Set logFile = Nothing
    Set fso = Nothing
End Sub

Sub LogResultWithSource(level, message, source)
    Dim fso, logFile, logLine
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set logFile = fso.OpenTextFile(LOG_FILE_PATH, 8, True) ' 8 = ForAppending, True = Create if not exists
    If Err.Number = 0 Then
        logLine = Now & " [" & level & "] [" & source & "] " & message
        logFile.WriteLine logLine
        logFile.Close
    Else
        bzhao.MsgBox "ERROR: Failed to write to log file: " & Err.Description & " (Error #" & Err.Number & ")", vbCritical, "Logging Error"
        Err.Clear
    End If
    Set logFile = Nothing
    Set fso = Nothing
End Sub



' SafeMsg: host-safe message display for interactive runs. For embedded hosts
' (no MsgBox/WScript), it will only log the message instead of popping UI.
Sub SafeMsg(text, isCritical, title)
    If Len(Trim(CStr(title))) = 0 Then title = ""
    If isCritical Then
        Call LogResultWithSource("ERROR", text, "SafeMsg")
    Else
        Call LogResultWithSource("INFO", text, "SafeMsg")
    End If

    Call FlushLogBuffer() ' Flush logs before showing MsgBox

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



'-----------------------------------------------------------
' Helper functions and subroutines
'-----------------------------------------------------------
Function IsTextPresent(textToFind)
    Call LogResultWithSource("DEBUG", "Entering IsTextPresent for: '" & textToFind & "'", "IsTextPresent")
    Dim screenContentBuffer 
    Dim screenLength
    screenLength = 24 * 80 ' Read full screen
    bzhao.ReadScreen screenContentBuffer, screenLength, 1, 1
    
    IsTextPresent = (InStr(1, screenContentBuffer, textToFind, vbTextCompare) > 0)
    
    Call LogResultWithSource("DEBUG", "Screen content checked for: '" & textToFind & "'. Result: " & IsTextPresent, "IsTextPresent")
    
    If Not IsTextPresent Then ' If text not found, log the full screen content
        Call LogResultWithSource("DEBUG", "Full screen content when '" & textToFind & "' was NOT found:\n" & String(80, "=") & "\n" & screenContentBuffer & "\n" & String(80, "="), "IsTextPresent")
    End If
End Function

' GetROFromScreen: reads screen and extracts 'RO:  123456' pattern
Function GetROFromScreen()
    If bzhao Is Nothing Then
        Call LogResultWithSource("ERROR", "bzhao object is not available", "GetROFromScreen")
        GetROFromScreen = ""
        Exit Function
    End If

    Dim screenContentBuffer, screenLength, re, matches
    screenLength = 3 * 80 ' top three lines
    On Error Resume Next
    bzhao.ReadScreen screenContentBuffer, screenLength, 1, 1
    If Err.Number <> 0 Then
        Call LogResultWithSource("ERROR", "GetROFromScreen ReadScreen failed: " & Err.Description, "GetROFromScreen")
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

' IsStatusReady: centralizes status matching (exact + tolerant variants)
Function IsStatusReady()
    ' Only accept the exact expected string to avoid false positives
    If IsTextPresent("RO STATUS: READY TO POST") Then
        IsStatusReady = True
    Else
        IsStatusReady = False
    End If
End Function




'-----------------------------------------------------------
' Closeout_Ro script subroutines
' (replace EnterText(...) calls with EnterTextAndWait(..., 1))
'-----------------------------------------------------------
Sub Closeout_Ro()
    '*******************************************************
    ' Final Closeout Steps
    '*******************************************************
    Call WaitForPrompt("COMMAND", "FC", True, g_DefaultWait)
    If HandleCloseoutErrors() Then Exit Sub
    
    ' Have all hours been entered
    Call WaitForPrompt("ALL LABOR POSTED", "Y", True, g_DefaultWait)
    If HandleCloseoutErrors() Then Exit Sub

    ' Confirming the next screen
    Call WaitForPrompt("VERIFY", "", True, g_DefaultWait)
    If HandleCloseoutErrors() Then Exit Sub
    
    ' OUT MILEAGE
    Call WaitForPrompt("MILEAGE OUT", "", True, g_DefaultWait)
    If HandleCloseoutErrors() Then Exit Sub
    
    ' IN MILEAGE
    Call WaitForPrompt("MILEAGE IN", "", True, g_DefaultWait)
    If HandleCloseoutErrors() Then Exit Sub
    
    ' OK TO CLOSE THE RO?
    Call WaitForPrompt("O.K. TO CLOSE RO", "Y", True, g_DefaultWait)
    If HandleCloseoutErrors() Then Exit Sub
    
    ' SEND TO PRINTER 2
    Call WaitForPrompt("INVOICE PRINTER", "2", True, g_DefaultWait)
    ' Record successful close for the final result summary; avoid immediate duplicate log line
    lastRoResult = "Successfully closed"
End Sub

' Helper: return the first matching trigger string or empty if none found.


' Replace the previous fixed-array error handler with a dictionary-driven handler.
Function HandleCloseoutErrors()
    Dim errorMap, key
    Set errorMap = GetCloseoutErrorMap()
    
    Dim fullScreenContent, screenLength
    screenLength = 24 * 80
    On Error Resume Next
    bzhao.ReadScreen fullScreenContent, screenLength, 1, 1
    If Err.Number <> 0 Then
        Call LogResultWithSource("ERROR", "ReadScreen failed in HandleCloseoutErrors: " & Err.Description, "HandleCloseoutErrors")
        Err.Clear
        HandleCloseoutErrors = False
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    For Each key In errorMap.Keys
        If InStr(1, fullScreenContent, key, vbTextCompare) > 0 Then
            ' Try to extract the detailed message shown on the screen near the key
            Dim detailedMsg
            detailedMsg = ExtractMessageNear(key, fullScreenContent)

            If Len(Trim(CStr(currentRODisplay))) > 0 Then
                If Len(detailedMsg) > 0 Then
                    Call LogResultWithSource(currentRODisplay, "CLOSEOUT Error detected: " & key & " - " & detailedMsg, "HandleCloseoutErrors")
                Else
                    Call LogResultWithSource(currentRODisplay, "CLOSEOUT Error detected: " & key, "HandleCloseoutErrors")
                End If
            Else
                If Len(detailedMsg) > 0 Then
                    Call LogResultWithSource("CLOSEOUT", "Error detected: " & key & " - " & detailedMsg, "HandleCloseoutErrors")
                Else
                    Call LogResultWithSource("CLOSEOUT", "Error detected: " & key, "HandleCloseoutErrors")
                End If
            End If

            ExecuteActions errorMap.Item(key)
            ' Final result: mark the RO as left open for manual closing
            lastRoResult = "Left Open for manual closing"
            HandleCloseoutErrors = True
            Exit Function
        End If
    Next
    
    HandleCloseoutErrors = False
End Function

' Return a dictionary mapping error text -> array of corrective action strings.
' Edit this function to add new error keys and their associated action sequences.
Function GetCloseoutErrorMap()
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Action command formats (case-insensitive):
    '   "PressEnter"               -> press Enter key
    '   "PressKey:<key>"           -> PressKey with given key token (e.g. "<NumpadEnter>", "<Esc>")
    '   "Send:<text>"              -> EnterText (sends text then Enter)
    '   "SendNoEnter:<text>"       -> bzhao.SendKey text (no Enter appended)
    '   "Wait:<seconds>"           -> bzhao.Wait seconds
    '   "Log:<message>"            -> LogResult("CLOSEOUT", message)
    '
    ' Customize or add entries below as you identify new conditions.
    
    dict.Add "ERROR", Array("PressEnter", "Send:E", "Wait:2")
    dict.Add "NOT AVAILABLE", Array("PressEnter", "Send:E", "Wait:2")
    dict.Add "INVALID", Array("PressEnter", "Send:E", "Wait:2")
    dict.Add "REQUEST CANNOT BE PROCESSED", Array("PressEnter", "Send:E", "Wait:2")
    dict.Add "INCOMPLETE SERVICE", Array("1", "Wait:9", "Send:E", "Wait:2")
    dict.Add "POSTING MESSAGES:", Array("PressEnter", "PressEnter", "PressEnter", "Send:E", "Wait:2")
    
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
    ' dict.Add "SPECIAL POPUP", Array("PressEnter", "Wait:1", "PressKey:<Esc>", "Wait:1", "Log:Handled SPECIAL POPUP")
    
    Set GetCloseoutErrorMap = dict
End Function

Function ExtractMessageNear(key)
    Dim screenContentBuffer, screenLength, keyPos, i, rowStart, rows, rowText, collected

    ' Read full screen (24 rows x 80 cols)
    screenLength = 24 * 80
    On Error Resume Next
    bzhao.ReadScreen screenContentBuffer, screenLength, 1, 1
    If Err.Number <> 0 Then
        Call LogResultWithSource("ERROR", "ReadScreen failed in ExtractMessageNear: " & Err.Description, "ExtractMessageNear")
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
    ' the host may really return different newline conventions or wrapped lines.
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
            PressKey "<Enter>"
            bzhao.Wait 1
            Case "PRESSKEY"
            PressKey Trim(arg)
            Case "SEND"
            ' Use WaitForPrompt (sends text then Enter)
            Call WaitForPrompt("", arg, True, 1000)
            Case "SENDNOENTER"
            bzhao.SendKey arg
            bzhao.Wait 1
            Case "WAIT"
            If IsNumeric(arg) Then bzhao.Wait CInt(arg)
            Case "LOG"
            If Len(Trim(CStr(currentRODisplay))) > 0 Then
                Call LogResultWithSource(currentRODisplay, arg, "ExecuteActions")
            Else
                Call LogResultWithSource("CLOSEOUT", arg, "ExecuteActions")
            End If
            Case Else
            Call LogResultWithSource("CLOSEOUT", "Unknown corrective action: " & action, "ExecuteActions")
        End Select
    Next
End Sub

' Ensure BlueZone is cleanly disconnected and object released
' cleanup handled earlier; no-op here to avoid duplicate Quit
