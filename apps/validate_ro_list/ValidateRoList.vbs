Option Explicit

' ======================================================================
' ValidateRoList.vbs
' Reads a CSV of RO numbers and checks each RO in BlueZone.
' Writes results to utilities\ValidateRoList_Results.txt in format: RO,STATUS
' Expected statuses: "NOT ON FILE" or "(PFC) POST FINAL CHARGES"
' ======================================================================

' --- Bootstrap ---
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_sh: Set g_sh = CreateObject("WScript.Shell")
Dim g_root: g_root = g_sh.Environment("USER")("CDK_BASE")
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\PathHelper.vbs")).ReadAll

Dim mockMode: mockMode = False
Dim envVal: envVal = g_sh.Environment("PROCESS")("MOCK_VALIDATE_RO")
If envVal = "" Then envVal = g_sh.Environment("USER")("MOCK_VALIDATE_RO")
If LCase(envVal) = "1" Or LCase(envVal) = "true" Then mockMode = True

' --- Use GetConfigPath for required files (fail-fast) ---
Dim inputFile: inputFile = g_sh.Environment("PROCESS")("MOCK_INPUT_FILE")
If mockMode And inputFile <> "" Then
    ' Use environment override
Else
    inputFile = GetConfigPath("ValidateRoList", "InputFile")
End If

' Validate inputFile unless we're in mock mode with MOCK_SCREEN_MAPS
Dim mockMapsEnvCheck: mockMapsEnvCheck = g_sh.Environment("PROCESS")("MOCK_SCREEN_MAPS")
If mockMapsEnvCheck = "" Then mockMapsEnvCheck = g_sh.Environment("USER")("MOCK_SCREEN_MAPS")
If Not (mockMode And mockMapsEnvCheck <> "") Then
    If inputFile = "" Then
        Err.Raise 53, "ValidateRoList", "Missing config.ini entry: [ValidateRoList] InputFile"
    End If
    If Not g_fso.FileExists(inputFile) Then
        Err.Raise 53, "ValidateRoList", "Input file not found: " & inputFile
    End If
End If

Dim toolsOutDir: toolsOutDir = GetConfigPath("ValidateRoList", "OutDir")
If toolsOutDir = "" Then
    Err.Raise 53, "ValidateRoList", "Missing config.ini entry: [ValidateRoList] OutDir"
End If
If Not g_fso.FolderExists(toolsOutDir) Then
    Err.Raise 53, "ValidateRoList", "OutDir path does not exist: " & toolsOutDir
End If

' Use explicit output file from config (Mandatory - Fail Fast)
Dim outputFile: outputFile = g_sh.Environment("PROCESS")("MOCK_OUTPUT_FILE")
If mockMode And outputFile <> "" Then
    ' Use environment override
Else
    outputFile = GetConfigPath("ValidateRoList", "OutputFile")
End If

If outputFile = "" Then
    Err.Raise 53, "ValidateRoList", "Missing config.ini entry: [ValidateRoList] OutputFile"
End If

Dim outputBaseName: outputBaseName = "ValidateRoList"
If outputFile <> "" Then outputBaseName = g_fso.GetBaseName(outputFile)

' --- Logging initialization (must be available before WaitForOneOf uses it) ---
' Logging level: 1=ERROR,2=INFO,3=DEBUG. Can override with env VALIDATERO_DEBUG
Dim DEBUG_LEVEL: DEBUG_LEVEL = 1
Dim envDbg: envDbg = g_sh.Environment("PROCESS")("VALIDATERO_DEBUG")
If envDbg = "" Then envDbg = g_sh.Environment("USER")("VALIDATERO_DEBUG")
If IsNumeric(envDbg) Then DEBUG_LEVEL = CInt(envDbg)

' Log file placed in the configured OutDir next to results
Dim logFile: logFile = g_fso.BuildPath(toolsOutDir, outputBaseName & "_log.txt")
Dim logTS: Set logTS = Nothing
On Error Resume Next
Set logTS = g_fso.OpenTextFile(logFile, 8, True)
If Err.Number <> 0 Then
    Err.Clear
    Set logTS = g_fso.CreateTextFile(logFile, True)
End If
On Error GoTo 0
logTS.WriteLine "ValidateRoList log started: " & Now & " | Init"

Sub LogResult(logType, message)
    Dim typeLevel
    Select Case UCase(logType)
        Case "ERROR": typeLevel = 1
        Case "INFO": typeLevel = 2
        Case "DEBUG": typeLevel = 3
        Case Else: typeLevel = 2
    End Select

    If typeLevel <= DEBUG_LEVEL Then
        On Error Resume Next
        logTS.WriteLine Now & " [" & logType & "] " & message
        On Error GoTo 0
    End If
End Sub

' Helper to extract leading row number from lines like "02 | ..."
Function ExtractRowNumber(line)
    Dim p, prefix, k, ch, numStr
    p = InStr(line, "|")
    If p > 0 Then
        prefix = Left(line, p - 1)
    Else
        prefix = line
    End If
    prefix = Trim(prefix)
    numStr = ""
    For k = 1 To Len(prefix)
        ch = Mid(prefix, k, 1)
        If ch >= "0" And ch <= "9" Then
            numStr = numStr & ch
        End If
    Next
    If numStr <> "" Then
        ExtractRowNumber = CInt(numStr)
    Else
        ExtractRowNumber = -1
    End If
End Function

 ' NOTE: Input and Out paths are provided by config via GetConfigPath() above.


' --- Screen map discovery (optional) ---
Dim MainPromptLine
Dim screenMapPath
screenMapPath = GetConfigPath("ValidateRoList", "ScreenMap")
' Allow overriding the screen map via env var for mock runs
' Check both singular and plural env names; prefer explicit MOCK_SCREEN_MAP if set,
' otherwise accept the first entry of MOCK_SCREEN_MAPS
Dim mockScreenMap: mockScreenMap = g_sh.Environment("PROCESS")("MOCK_SCREEN_MAP")
If mockScreenMap = "" Then mockScreenMap = g_sh.Environment("USER")("MOCK_SCREEN_MAP")
If mockScreenMap = "" Then
    Dim tmpMaps: tmpMaps = g_sh.Environment("PROCESS")("MOCK_SCREEN_MAPS")
    If tmpMaps = "" Then tmpMaps = g_sh.Environment("USER")("MOCK_SCREEN_MAPS")
    If tmpMaps <> "" Then
        Dim firstMap: firstMap = Split(tmpMaps, ";")(0)
        mockScreenMap = Trim(firstMap)
    End If
End If
If mockScreenMap <> "" Then screenMapPath = mockScreenMap

' COORDINATE-FREE SCREEN READING
' We read the entire 24x80 screen (1920 chars) in a single operation
' to maximize terminal throughput and ensure robustness.
MainPromptLine = 23

' --- Prepare BlueZone object ---
' Support a quick mock mode that takes one-or-more screen-map files and
' produces a single _out file with one line per map. Use env var
' MOCK_SCREEN_MAPS with semicolon-separated paths (absolute or repo-relative).
Dim mockMapsEnv: mockMapsEnv = g_sh.Environment("PROCESS")("MOCK_SCREEN_MAPS")
If mockMapsEnv = "" Then mockMapsEnv = g_sh.Environment("USER")("MOCK_SCREEN_MAPS")
If mockMode And mockMapsEnv <> "" Then
    ' Use environment override for output if provided, else use default in toolsOutDir
    Dim mockOut: mockOut = g_sh.Environment("PROCESS")("MOCK_OUTPUT_FILE")
    If mockOut = "" Then mockOut = g_fso.BuildPath(toolsOutDir, "ValidateRoList_mock_out.txt")
    
    Dim mockLog: mockLog = g_fso.BuildPath(toolsOutDir, "ValidateRoList_mock_log.txt")
    Dim mockOutTS: Set mockOutTS = g_fso.CreateTextFile(mockOut, True)
    Dim mockLogTS: Set mockLogTS = g_fso.CreateTextFile(mockLog, True)
    mockLogTS.WriteLine "Mock map run started: " & Now & " | maps=" & mockMapsEnv
    Dim mapsArr: mapsArr = Split(mockMapsEnv, ";")
    Dim mi
    For mi = 0 To UBound(mapsArr)
        Dim mapPath: mapPath = Trim(mapsArr(mi))
        If mapPath <> "" Then
            If Not g_fso.FileExists(mapPath) Then
                ' Do NOT fallback to hardcoded repo paths here; require maps to be
                ' provided via config or absolute paths. Record missing and continue.
            End If

            If Not g_fso.FileExists(mapPath) Then
                mockLogTS.WriteLine Now & " | SKIP missing map: " & mapPath
                mockOutTS.WriteLine g_fso.GetBaseName(mapPath) & ",MISSING"
            Else
                Dim simTS: Set simTS = g_fso.OpenTextFile(mapPath, 1, False)
                Dim simBuf: simBuf = ""
                Do Until simTS.AtEndOfStream
                    simBuf = simBuf & " " & simTS.ReadLine
                Loop
                simTS.Close
                simBuf = UCase(simBuf)
                Dim status
                If InStr(simBuf, "RO:") > 0 Then
                    status = "Open"
                ElseIf InStr(simBuf, "NOT ON FILE") > 0 Then
                    status = "NOT ON FILE"
                ElseIf InStr(simBuf, "IS CLOSED") > 0 Then
                    status = "ALREADY CLOSED"
                Else
                    status = "UNKNOWN"
                End If
                mockLogTS.WriteLine Now & " | MAP=" & mapPath & " -> " & status
                mockOutTS.WriteLine g_fso.GetFileName(mapPath) & "," & status
            End If
        End If
    Next
    mockOutTS.Close
    mockLogTS.WriteLine "Mock map run finished: " & Now & " | out=" & mockOut
    mockLogTS.Close
    Dim conclMsg: conclMsg = "Mock map run complete. Results: " & mockOut
    On Error Resume Next
    WScript.Echo conclMsg
    If Err.Number <> 0 Then
        Err.Clear
        Dim echoTS: Set echoTS = g_fso.OpenTextFile(mockLog, 8, True)
        echoTS.WriteLine Now & " | " & conclMsg
        echoTS.Close
    End If
    On Error GoTo 0
    On Error Resume Next
    WScript.Quit 0
    If Err.Number <> 0 Then Err.Raise 9999, "ValidateRoList", "Completed (non-WSH host)"
End If

If Not mockMode Then
    Dim g_bzhao: Set g_bzhao = CreateObject("BZWhll.WhllObj")
    On Error Resume Next
    g_bzhao.Connect ""
    If Err.Number <> 0 Then
        MsgBox "ERROR: Failed to connect to BlueZone terminal session. " & Err.Description, vbCritical, "ValidateRoList"
        On Error Resume Next
        WScript.Quit 1
        If Err.Number <> 0 Then Err.Raise 9999, "ValidateRoList", "BlueZone connect failed"
    End If
    On Error GoTo 0
End If

' PROMPT-DRIVEN UI SYNCHRONIZATION
' High-speed polling (250ms) using full-screen buffer reads.
Function WaitForOneOf(targetsCSV, timeoutMs)
    Dim targets: targets = Split(targetsCSV, "|")
    Dim elapsed: elapsed = 0
    Dim screenBuffer, i, pollCount: pollCount = 0

    Do
        If mockMode Then
            ' In mock mode, use the provided screen map file as a simulated terminal snapshot
            If screenMapPath <> "" And g_fso.FileExists(screenMapPath) Then
                Dim simTS: Set simTS = g_fso.OpenTextFile(screenMapPath, 1, False)
                Dim simBuf: simBuf = ""
                Do Until simTS.AtEndOfStream
                    simBuf = simBuf & " " & simTS.ReadLine
                Loop
                simTS.Close
                simBuf = UCase(simBuf)
                ' Check for any target in the simulated buffer
                For i = 0 To UBound(targets)
                    Dim tm: tm = Trim(targets(i))
                    If InStr(simBuf, UCase(tm)) > 0 Then
                        WaitForOneOf = tm
                        Exit Function
                    End If
                Next
                WaitForOneOf = "__TIMEOUT__"
                Exit Function
            End If
        End If

        pollCount = pollCount + 1
        g_bzhao.Pause 250
        elapsed = elapsed + 250
        
        ' Read the entire 24x80 screen in a single COM call (fastest method)
        g_bzhao.ReadScreen screenBuffer, 1920, 1, 1
        screenBuffer = UCase(screenBuffer)

        For i = 0 To UBound(targets)
            Dim t: t = targets(i)
            If InStr(screenBuffer, UCase(t)) > 0 Then
                WaitForOneOf = t
                LogResult "DEBUG", "MATCH -> [" & t & "] | Elapsed=" & elapsed & "ms"
                Exit Function
            End If
        Next

        If elapsed >= timeoutMs Then
            LogResult "ERROR", "TIMEOUT after " & elapsed & "ms"
            WaitForOneOf = "__TIMEOUT__"
            Exit Function
        End If
    Loop
End Function

' --- Process input file ---
Dim inTS: Set inTS = g_fso.OpenTextFile(inputFile, 1, False)
Dim outTS: Set outTS = g_fso.CreateTextFile(outputFile, True)

Dim ln, roVal, roStatus, foundResult
Dim timeoutMs: timeoutMs = 3000 ' 3 seconds timeout for responsiveness
' Primary UI cues identified from screen maps (No leading/trailing spaces around pipe)
Dim targetsCSV: targetsCSV = "RO:|NOT ON FILE|IS CLOSED"

Do Until inTS.AtEndOfStream
    ln = inTS.ReadLine
    roVal = Trim(ln)
    If roVal <> "" Then
        ' Direct terminal entry: includes requested 500ms delays for UI stability
        If Not mockMode Then
            g_bzhao.Pause 700 ' Delay before typing RO
            g_bzhao.SendKey roVal
            g_bzhao.Pause 500 ' Delay before Enter
            g_bzhao.SendKey "<NumpadEnter>"	    
            g_bzhao.Pause 500 ' Delay before Enter
        End If

        ' Wait for the UI cue to appear (250ms polling)
        foundResult = WaitForOneOf(targetsCSV, timeoutMs)

        If foundResult = "__TIMEOUT__" Then
            roStatus = "TIMEOUT"
        Else
            ' Map specific triggers to user-facing statuses using InStr for robustness
            If InStr(foundResult, "RO:") > 0 Then
                roStatus = "Open"
            ElseIf InStr(foundResult, "NOT ON FILE") > 0 Then
                roStatus = "NOT ON FILE"
            ElseIf InStr(foundResult, "IS CLOSED") > 0 Then
                roStatus = "ALREADY CLOSED"
            Else
                roStatus = "UNKNOWN"
            End If
        End If

        outTS.WriteLine roVal & "," & roStatus

        ' Return to main COMMAND: prompt if valid, or sync at prompt if invalid
        If roStatus = "Open" Then
            If Not mockMode Then
                g_bzhao.SendKey "E<NumpadEnter>"
                On Error Resume Next
                WaitForOneOf "COMMAND:", timeoutMs
                On Error GoTo 0
            End If
        Else
            ' If invalid, ensure we sync with the current prompt before looping
            If Not mockMode Then WaitForOneOf "?|COMMAND:", timeoutMs
        End If
    End If
Loop

inTS.Close
outTS.Close
On Error Resume Next
logTS.WriteLine "ValidateRoList finished: " & Now & " | Results=" & outputFile
logTS.Close

Dim finalMsg: finalMsg = "ValidateRoList complete. Results: " & outputFile
On Error Resume Next
WScript.Echo finalMsg
If Err.Number <> 0 Then
    Err.Clear
    Dim finalLogTS: Set finalLogTS = g_fso.OpenTextFile(g_fso.BuildPath(toolsOutDir, "ValidateRoList_final_log.txt"), 8, True)
    finalLogTS.WriteLine Now & " | " & finalMsg
    finalLogTS.Close
End If
On Error GoTo 0
On Error Resume Next
WScript.Quit 0
If Err.Number <> 0 Then Err.Raise 9999, "ValidateRoList", "Completed (non-WSH host)"
