Option Explicit

' ======================================================================
' ValidateRoList.vbs
' Reads a CSV of RO numbers and checks each RO in BlueZone.
' Writes results to utilities\ValidateRoList_Results.txt in format: RO,STATUS
' Expected statuses: "NOT ON FILE" or "(PFC) POST FINAL CHARGES"
' ======================================================================

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim sh: Set sh = CreateObject("WScript.Shell")

' --- Determine repo root and paths ---
Dim repoRoot
' Prefer CDK_BASE environment variable when available and valid
Dim envBase: envBase = sh.ExpandEnvironmentStrings("%CDK_BASE%")
If envBase = "%CDK_BASE%" Then envBase = sh.Environment("USER")("CDK_BASE")
If envBase <> "" And fso.FolderExists(envBase) And fso.FileExists(fso.BuildPath(envBase, ".cdkroot")) Then
    repoRoot = envBase
Else
    ' Fall back to walking up from current working directory
    repoRoot = FindRepoRoot(fso.GetAbsolutePathName("."))
End If

If repoRoot = "" Then
    MsgBox "ERROR: Could not determine repository root. Ensure CDK is extracted with .cdkroot present or set CDK_BASE.", vbCritical, "ValidateRoList"
    On Error Resume Next
    WScript.Quit 1
    If Err.Number <> 0 Then Err.Raise 9999, "ValidateRoList", "Terminating: repo root not found"
End If

 ' HostCompat not loaded here; use guarded WScript.Quit fallbacks instead

Function FindRepoRoot(startDir)
    Dim current: current = startDir
    Dim i: i = 0
    Do While i < 20
        If fso.FileExists(fso.BuildPath(current, ".cdkroot")) Then
            FindRepoRoot = current
            Exit Function
        End If
        Dim parent: parent = fso.GetParentFolderName(current)
        If parent = "" Or parent = current Then Exit Do
        current = parent
        i = i + 1
    Loop
    FindRepoRoot = ""
End Function

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

 ' Prefer an input file in sensible locations; check several fallbacks
Dim cwd: cwd = fso.GetAbsolutePathName(".")
Dim candidateA: candidateA = fso.BuildPath(cwd, "ValidateRoList_IN.csv")
Dim candidateB: candidateB = fso.BuildPath(cwd, "tools\ValidateRoList\ValidateRoList_IN.csv")
Dim candidateC: candidateC = fso.BuildPath(repoRoot, "tools\ValidateRoList\ValidateRoList_IN.csv")
Dim defaultInputPath: defaultInputPath = fso.BuildPath(repoRoot, "utilities\ValidateRoList_IN.csv")
Dim inputFile
If fso.FileExists(candidateA) Then
    inputFile = candidateA
ElseIf fso.FileExists(candidateB) Then
    inputFile = candidateB
ElseIf fso.FileExists(candidateC) Then
    inputFile = candidateC
Else
    inputFile = defaultInputPath
End If

Dim inputFolder: inputFolder = fso.GetParentFolderName(inputFile)
Dim inputBase: inputBase = fso.GetBaseName(inputFile)
' If input filename ends with _IN, strip that before appending _out
Dim baseRoot: baseRoot = inputBase
If Len(baseRoot) > 3 Then
    If LCase(Right(baseRoot, 3)) = "_in" Then
        baseRoot = Left(baseRoot, Len(baseRoot) - 3)
    End If
End If
Dim outputFile
' Ensure the results _out file is placed under tools\ValidateRoList next to the script
Dim toolsOutDir: toolsOutDir = fso.BuildPath(repoRoot, "tools\ValidateRoList")
If Not fso.FolderExists(toolsOutDir) Then toolsOutDir = inputFolder
outputFile = fso.BuildPath(toolsOutDir, baseRoot & "_out.txt")

If Not fso.FileExists(inputFile) Then
    MsgBox "ERROR: Input file not found: " & inputFile, vbCritical, "ValidateRoList"
    On Error Resume Next
    WScript.Quit 1
    If Err.Number <> 0 Then Err.Raise 9999, "ValidateRoList", "Missing input file"
End If

' --- Screen map discovery (optional) ---
Dim MainPromptLine
Dim screenMapPath
screenMapPath = ""
Dim cand1: cand1 = fso.BuildPath(repoRoot, "utilities\ro_screen_map.txt")
Dim cand2: cand2 = fso.BuildPath(cwd, "ro_screen_map.txt")
Dim cand3: cand3 = fso.BuildPath(repoRoot, "tools\ValidateRoList\ro_screen_map.txt")
 ' Allow overriding the screen map via env var for mock runs
Dim mockScreenMap: mockScreenMap = sh.Environment("PROCESS")("MOCK_SCREEN_MAP")
If mockScreenMap = "" Then mockScreenMap = sh.Environment("USER")("MOCK_SCREEN_MAP")
If mockScreenMap <> "" And fso.FileExists(mockScreenMap) Then
    screenMapPath = mockScreenMap
End If
If fso.FileExists(cand1) Then
    screenMapPath = cand1
ElseIf fso.FileExists(cand2) Then
    screenMapPath = cand2
ElseIf fso.FileExists(cand3) Then
    screenMapPath = cand3
End If

If screenMapPath <> "" Then
    On Error Resume Next
    Dim mapTS: Set mapTS = fso.OpenTextFile(screenMapPath, 1, False)
    If Err.Number = 0 Then
        
        Dim mapLines()
        ReDim mapLines(0)
        Dim idx: idx = 0
        Do Until mapTS.AtEndOfStream
            Dim ml: ml = mapTS.ReadLine
            ReDim Preserve mapLines(idx)
            mapLines(idx) = ml
            idx = idx + 1
        Loop
        mapTS.Close

        ' Try to discover a sensible read-start line from the map file.
        ' Prefer the line containing the R.O. prompt (e.g. "R.O. NUMBER")
        ' and also look for where "NOT ON FILE" appears so we can read that
        ' area if it's above the prompt.
        Dim foundROLine: foundROLine = -1
        Dim foundNotOnFileLine: foundNotOnFileLine = -1
        Dim i, lineText
        For i = 0 To UBound(mapLines)
            lineText = UCase(mapLines(i))
            If InStr(lineText, "R.O. NUMBER") > 0 Or InStr(lineText, "RO NUMBER") > 0 Then
                foundROLine = ExtractRowNumber(mapLines(i))
            End If
            If InStr(lineText, "NOT ON FILE") > 0 Then
                foundNotOnFileLine = ExtractRowNumber(mapLines(i))
            End If
            If foundROLine > 0 And foundNotOnFileLine > 0 Then Exit For
        Next

        If foundROLine > 0 Then
            MainPromptLine = foundROLine
        Else
            ' Fallback: try to find a generic COMMAND: prompt, else default to 23
            MainPromptLine = 23
            For i = 0 To UBound(mapLines)
                If InStr(UCase(mapLines(i)), "COMMAND:") > 0 Then
                    MainPromptLine = i + 1
                    Exit For
                End If
            Next
        End If

        ' Choose the earliest relevant line to start reading from so we capture
        ' statuses that may appear above the prompt (e.g. "NOT ON FILE").
        Dim ReadStartLine
        If foundNotOnFileLine > 0 Then
            If foundNotOnFileLine < MainPromptLine Then
                ReadStartLine = foundNotOnFileLine
            Else
                ReadStartLine = MainPromptLine
            End If
        Else
            ReadStartLine = MainPromptLine
        End If
    Else
        MainPromptLine = 23
        ReadStartLine = 23
    End If
    On Error GoTo 0
Else
    MainPromptLine = 23
    ReadStartLine = 23
End If

' --- Prepare BlueZone object ---
Dim mockMode: mockMode = False
Dim envVal: envVal = sh.Environment("PROCESS")("MOCK_VALIDATE_RO")
If envVal = "" Then envVal = sh.Environment("USER")("MOCK_VALIDATE_RO")
If LCase(envVal) = "1" Or LCase(envVal) = "true" Then mockMode = True

' Support a quick mock mode that takes one-or-more screen-map files and
' produces a single _out file with one line per map. Use env var
' MOCK_SCREEN_MAPS with semicolon-separated paths (absolute or repo-relative).
Dim mockMapsEnv: mockMapsEnv = sh.Environment("PROCESS")("MOCK_SCREEN_MAPS")
If mockMapsEnv = "" Then mockMapsEnv = sh.Environment("USER")("MOCK_SCREEN_MAPS")
If mockMode And mockMapsEnv <> "" Then
    toolsOutDir = fso.BuildPath(repoRoot, "tools\ValidateRoList")
    If Not fso.FolderExists(toolsOutDir) Then toolsOutDir = fso.GetAbsolutePathName(".")
    Dim mockOut: mockOut = fso.BuildPath(toolsOutDir, "ValidateRoList_mock_out.txt")
    Dim mockLog: mockLog = fso.BuildPath(toolsOutDir, "ValidateRoList_mock_log.txt")
    Dim mockOutTS: Set mockOutTS = fso.CreateTextFile(mockOut, True)
    Dim mockLogTS: Set mockLogTS = fso.CreateTextFile(mockLog, True)
    mockLogTS.WriteLine "Mock map run started: " & Now & " | maps=" & mockMapsEnv
    Dim mapsArr: mapsArr = Split(mockMapsEnv, ";")
    Dim mi
    For mi = 0 To UBound(mapsArr)
        Dim mapPath: mapPath = Trim(mapsArr(mi))
        If mapPath <> "" Then
            If Not fso.FileExists(mapPath) Then
                Dim alt: alt = fso.BuildPath(repoRoot, mapPath)
                If fso.FileExists(alt) Then mapPath = alt
            End If

            If Not fso.FileExists(mapPath) Then
                mockLogTS.WriteLine Now & " | SKIP missing map: " & mapPath
                mockOutTS.WriteLine fso.GetBaseName(mapPath) & ",MISSING"
            Else
                Dim simTS: Set simTS = fso.OpenTextFile(mapPath, 1, False)
                Dim simBuf: simBuf = ""
                Do Until simTS.AtEndOfStream
                    simBuf = simBuf & " " & simTS.ReadLine
                Loop
                simTS.Close
                simBuf = UCase(simBuf)
                Dim status
                If InStr(simBuf, "(PFC) POST FINAL CHARGES") > 0 Or InStr(simBuf, "PFC") > 0 Then
                    status = "(PFC) POST FINAL CHARGES"
                ElseIf InStr(simBuf, "NOT ON FILE") > 0 Then
                    status = "NOT ON FILE"
                Else
                    status = "UNKNOWN"
                End If
                mockLogTS.WriteLine Now & " | MAP=" & mapPath & " -> " & status
                mockOutTS.WriteLine fso.GetFileName(mapPath) & "," & status
            End If
        End If
    Next
    mockOutTS.Close
    mockLogTS.WriteLine "Mock map run finished: " & Now & " | out=" & mockOut
    mockLogTS.Close
    MsgBox "Mock map run complete. Results: " & mockOut, vbInformation, "ValidateRoList"
    On Error Resume Next
    WScript.Quit 0
    If Err.Number <> 0 Then Err.Raise 9999, "ValidateRoList", "Completed (non-WSH host)"
End If

If Not mockMode Then
    Dim bzhao: Set bzhao = CreateObject("BZWhll.WhllObj")
    On Error Resume Next
    bzhao.Connect ""
    If Err.Number <> 0 Then
        MsgBox "ERROR: Failed to connect to BlueZone terminal session. " & Err.Description, vbCritical, "ValidateRoList"
        On Error Resume Next
        WScript.Quit 1
        If Err.Number <> 0 Then Err.Raise 9999, "ValidateRoList", "BlueZone connect failed"
    End If
    On Error GoTo 0
End If

' --- Helper subs (copied pattern used across repo) ---
Sub PressKey(key)
    If mockMode Then Exit Sub
    bzhao.SendKey key
    bzhao.Pause 100
End Sub

Sub EnterTextAndWait(text)
    If mockMode Then Exit Sub
    bzhao.SendKey text
    bzhao.Pause 100
    Call PressKey("<NumpadEnter>")
    bzhao.Pause 500
End Sub

' Wait for any of the target texts to appear at bottom rows (23 or 24)
Function WaitForOneOf(targetsCSV, timeoutMs)
    Dim targets: targets = Split(targetsCSV, "|")
    Dim elapsed: elapsed = 0
    Dim col: col = 1
    ' Increase capture width; some terminal adapters present wider buffers
    Dim screenLength: screenLength = 160
    Dim screenBuffer, i
    Dim numReadLines: numReadLines = 6
    Dim j

    Do
        If mockMode Then
            ' In mock mode, use the provided screen map file as a simulated terminal snapshot
            If screenMapPath <> "" And fso.FileExists(screenMapPath) Then
                Dim simTS: Set simTS = fso.OpenTextFile(screenMapPath, 1, False)
                Dim simBuf: simBuf = ""
                Do Until simTS.AtEndOfStream
                    simBuf = simBuf & " " & simTS.ReadLine
                Loop
                simTS.Close
                simBuf = UCase(simBuf)
                On Error Resume Next
                logTS.WriteLine Now & " | MOCK-SIM snapshot read from " & screenMapPath
                On Error GoTo 0
                ' Check for any target in the simulated buffer
                For i = 0 To UBound(targets)
                    Dim tm: tm = Trim(targets(i))
                    If InStr(simBuf, UCase(tm)) > 0 Then
                        WaitForOneOf = tm
                        On Error Resume Next
                        logTS.WriteLine Now & " | MOCK-SIM MATCH -> " & tm
                        On Error GoTo 0
                        Exit Function
                    End If
                Next
                ' nothing matched in simulated snapshot -> timeout
                WaitForOneOf = "__TIMEOUT__"
                Exit Function
            Else
                On Error Resume Next
                logTS.WriteLine Now & " | MOCK mode - no screen map available"
                On Error GoTo 0
                WaitForOneOf = "__MOCK__"
                Exit Function
            End If
        End If

        bzhao.Pause 500
        elapsed = elapsed + 500
        screenBuffer = ""
        ' Read starting one line above ReadStartLine when possible to capture
        ' messages that appear above the prompt (e.g. status lines).
        Dim readBase
        ' Try to include up to two lines above the prompt to capture status text
        readBase = ReadStartLine
        If ReadStartLine > 2 Then
            readBase = ReadStartLine - 2
        ElseIf ReadStartLine > 1 Then
            readBase = 1
        End If
        On Error Resume Next
        logTS.WriteLine Now & " | READBASE=" & readBase & " | numReadLines=" & numReadLines
        On Error GoTo 0
        For j = 0 To numReadLines - 1
            Dim tmpBuf
            bzhao.ReadScreen tmpBuf, screenLength, readBase + j, col
            screenBuffer = screenBuffer & " " & tmpBuf
            ' Log the raw line we just read so we can see exact screen content
            On Error Resume Next
            logTS.WriteLine Now & " | LINE " & (readBase + j) & " | " & tmpBuf
            On Error GoTo 0
        Next
        screenBuffer = UCase(screenBuffer)

        ' Log a trimmed snapshot for debugging
        On Error Resume Next
        Dim snap
        snap = Left(Replace(screenBuffer, vbCrLf, " "), 2000)
        logTS.WriteLine Now & " | Elapsed=" & elapsed & "ms | Snapshot=" & snap
        On Error GoTo 0

        For i = 0 To UBound(targets)
            Dim t: t = Trim(targets(i))
            If InStr(screenBuffer, UCase(t)) > 0 Then
                WaitForOneOf = t
                On Error Resume Next
                logTS.WriteLine Now & " | MATCH -> " & t & " | Elapsed=" & elapsed & "ms"
                On Error GoTo 0
                Exit Function
            End If
        Next

        If elapsed >= timeoutMs Then
            On Error Resume Next
            logTS.WriteLine Now & " | TIMEOUT after " & elapsed & "ms"
            On Error GoTo 0
                ' Before giving up, try one expanded scan from the top of the screen
                On Error Resume Next
                logTS.WriteLine Now & " | Performing expanded retry scan from top of screen"
                On Error GoTo 0
                Dim retryBase, retryLines, kbuf
                retryBase = 1
                retryLines = MainPromptLine
                If retryLines < numReadLines Then retryLines = numReadLines
                If retryLines > 24 Then retryLines = 24
                kbuf = ""
                For j = 0 To retryLines - 1
                    Dim tbuf
                    bzhao.ReadScreen tbuf, screenLength, retryBase + j, col
                    kbuf = kbuf & " " & tbuf
                Next
                kbuf = UCase(kbuf)
                For i = 0 To UBound(targets)
                    Dim tr: tr = Trim(targets(i))
                    If InStr(kbuf, UCase(tr)) > 0 Then
                        WaitForOneOf = tr
                        On Error Resume Next
                        logTS.WriteLine Now & " | EXPANDED MATCH -> " & tr
                        On Error GoTo 0
                        Exit Function
                    End If
                Next
                WaitForOneOf = "__TIMEOUT__"
                Exit Function
        End If
    Loop
End Function

' --- Process input file ---
Dim inTS: Set inTS = fso.OpenTextFile(inputFile, 1, False)
Dim outTS: Set outTS = fso.CreateTextFile(outputFile, True)

' Debug log for scraping activity
Dim logFile: logFile = fso.BuildPath(inputFolder, baseRoot & "_log.txt")
Dim logTS: Set logTS = fso.CreateTextFile(logFile, True)
logTS.WriteLine "ValidateRoList log started: " & Now & " | ReadStartLine=" & ReadStartLine & " MainPromptLine=" & MainPromptLine

Dim ln, roVal, roStatus, foundResult
Dim timeoutMs: timeoutMs = 10000 ' 10 seconds per your choice
' Only these statuses are considered valid results for scraping.
' Order matters: check for PFC/post-final first to avoid partial matches.
' Include a short token `PFC` to catch variants like missing parentheses or line breaks.
Dim targets: targets = "(PFC) POST FINAL CHARGES|PFC|NOT ON FILE"

Do Until inTS.AtEndOfStream
    ln = inTS.ReadLine
    roVal = Trim(ln)
    If roVal <> "" Then
        ' send RO to COMMAND prompt
        EnterTextAndWait roVal

        ' wait for one of the two expected results
        foundResult = WaitForOneOf(targets, timeoutMs)

        If foundResult = "__TIMEOUT__" Then
            roStatus = "TIMEOUT"
        ElseIf foundResult = "__MOCK__" Then
            roStatus = "__MOCK__"
        Else
            ' Only accept the two canonical statuses; map loosely if needed.
            If InStr(UCase(foundResult), "PFC") > 0 Or UCase(Trim(foundResult)) = "(PFC) POST FINAL CHARGES" Then
                roStatus = "(PFC) POST FINAL CHARGES"
            ElseIf UCase(Trim(foundResult)) = "NOT ON FILE" Then
                roStatus = "NOT ON FILE"
            Else
                roStatus = "UNKNOWN"
            End If
        End If

        outTS.WriteLine roVal & "," & roStatus

        ' If we found a valid RO (post-final charges present), exit back to main screen
        If roStatus = "(PFC) POST FINAL CHARGES" Then
            EnterTextAndWait "E"
            ' Wait for main command prompt to ensure we're back at the start
            On Error Resume Next
            WaitForOneOf "COMMAND:", timeoutMs
            On Error GoTo 0
        End If
    End If
Loop

inTS.Close
outTS.Close
On Error Resume Next
logTS.WriteLine "ValidateRoList finished: " & Now & " | Results=" & outputFile
logTS.Close

MsgBox "ValidateRoList complete. Results: " & outputFile, vbInformation, "ValidateRoList"
On Error Resume Next
WScript.Quit 0
If Err.Number <> 0 Then Err.Raise 9999, "ValidateRoList", "Completed (non-WSH host)"
