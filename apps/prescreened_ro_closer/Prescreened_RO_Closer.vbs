'=====================================================================================
' Prescreened RO Closer
' Part of the CDK DMS Automation Suite
'
' Purpose: Close ROs from a pre-screened input list with no gate logic.
'          The input CSV is assumed to contain only ROs that are safe to close.
'          All business rule evaluation happens externally (e.g. spreadsheet filter).
'=====================================================================================

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

' Load host compatibility helpers
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\HostCompat.vbs")).ReadAll

' --- Configuration ---
Dim MAIN_PROMPT: MAIN_PROMPT = "R.O. NUMBER"
Dim LOG_FILE_PATH: LOG_FILE_PATH = GetConfigPath("Prescreened_RO_Closer", "Log")
Dim INPUT_CSV_PATH: INPUT_CSV_PATH = GetConfigPath("Prescreened_RO_Closer", "InputCSV")
Dim DEBUG_LEVEL: DEBUG_LEVEL = 2 ' 1=Error, 2=Info

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

Dim STABILITY_PAUSE: STABILITY_PAUSE = GetConfigSetting("Prescreened_RO_Closer", "StabilityPause", 1000)
Dim LOOP_PAUSE: LOOP_PAUSE = GetConfigSetting("Prescreened_RO_Closer", "LoopPause", 500)
Dim REVIEW_PAUSE: REVIEW_PAUSE = GetConfigSetting("Prescreened_RO_Closer", "ReviewPause", 250)
Dim BLACKLIST_TERMS: BLACKLIST_TERMS = GetConfigSetting("Prescreened_RO_Closer", "blacklist_terms", "")

' --- CDK Objects ---
Dim bzhao: Set bzhao = CreateObject("BZWhll.WhllObj")

' --- Main ---
Sub RunAutomation()
    Dim successfulCount: successfulCount = 0
    Dim skippedCount: skippedCount = 0

    If Not g_fso.FileExists(INPUT_CSV_PATH) Then
        LogResult "ERROR", "Input CSV not found: " & INPUT_CSV_PATH
        MsgBox "Error: Input CSV not found at: " & INPUT_CSV_PATH, vbCritical, "File Not Found"
        Exit Sub
    End If

    LogResult "INFO", "Starting Prescreened RO Closer. Input: " & INPUT_CSV_PATH

    On Error Resume Next
    bzhao.Connect ""
    If Err.Number <> 0 Then
        LogResult "ERROR", "Failed to connect to BlueZone: " & Err.Description
        MsgBox "Failed to connect to BlueZone terminal session.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ProcessRoList successfulCount, skippedCount

    On Error Resume Next
    If Not bzhao Is Nothing Then bzhao.Disconnect
    On Error GoTo 0

    LogResult "INFO", "Automation complete. Closed: " & successfulCount & " | Skipped: " & skippedCount
    MsgBox "Prescreened RO Closer Finished." & vbCrLf & "Successful Closures: " & successfulCount & vbCrLf & "Skipped: " & skippedCount, vbInformation
End Sub

' --- Process List ---
Sub ProcessRoList(ByRef successfulCount, ByRef skippedCount)
    Dim ts, strLine, roFromCsv, currentRo

    On Error Resume Next
    Set ts = g_fso.OpenTextFile(INPUT_CSV_PATH, 1)
    If Err.Number <> 0 Then
        LogResult "ERROR", "Failed to open input CSV: " & Err.Description
        MsgBox "Failed to open input CSV: " & INPUT_CSV_PATH, vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    Do While Not ts.AtEndOfStream
        strLine = Trim(ts.ReadLine)
        If strLine <> "" Then
            roFromCsv = Split(strLine, ",")(0)
            currentRo = Trim(roFromCsv)

            If Len(currentRo) = 6 And IsNumeric(currentRo) Then
                LogResult "INFO", String(50, "=")
                LogResult "INFO", "Processing RO: " & currentRo

                WaitForText MAIN_PROMPT
                EnterTextWithStability currentRo

                If IsRoProcessable(currentRo) Then
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
                Else
                    skippedCount = skippedCount + 1
                End If

                ReturnToMainPrompt
            ElseIf Len(currentRo) > 0 Then
                LogResult "INFO", "Skipping invalid format row: '" & currentRo & "'"
            End If
        End If
    Loop

    If Not ts Is Nothing Then
        ts.Close
        Set ts = Nothing
    End If
End Sub

' --- RO State Check ---
Function IsRoProcessable(roNumber)
    Dim screenContent
    ' Wait for the RO screen to fully load before reading state.
    ' Any of these indicate the screen is ready — COMMAND: is the normal loaded state.
    WaitForText "COMMAND:|NOT ON FILE|ALREADY CLOSED|is closed|ENTER SEQUENCE NUMBER|VARIABLE HAS NOT BEEN ASSIGNED"
    bzhao.Pause STABILITY_PAUSE
    bzhao.ReadScreen screenContent, 1920, 1, 1

    Dim matchedBlacklistTerm: matchedBlacklistTerm = GetMatchedBlacklistTerm(BLACKLIST_TERMS, screenContent)
    If matchedBlacklistTerm <> "" Then
        LogResult "INFO", "RO " & roNumber & " | Blacklisted ('" & matchedBlacklistTerm & "'). Skipping."
        IsRoProcessable = False
        Exit Function
    End If

    If InStr(1, screenContent, "NOT ON FILE", vbTextCompare) > 0 Then
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
        LogResult "INFO", "RO " & roNumber & " prompted for Sequence Number. Skipping."
        IsRoProcessable = False
        Exit Function
    End If

    IsRoProcessable = True
End Function

' --- Review Sequence ---
Function DiscoverLineLetters()
    Dim i, capturedLetter, screenContentBuffer, readLength
    Dim foundLetters, foundCount
    Dim startReadRow, startReadColumn, emptyRowCount, nextColChar
    Dim tempLetters(25)
    foundCount = 0
    emptyRowCount = 0

    For startReadRow = 10 To 22
        startReadColumn = 1
        readLength = 1

        On Error Resume Next
        bzhao.ReadScreen screenContentBuffer, readLength, startReadRow, startReadColumn
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        On Error GoTo 0

        capturedLetter = Trim(screenContentBuffer)
        If Len(capturedLetter) = 1 Then
            If Asc(UCase(capturedLetter)) >= Asc("A") And Asc(UCase(capturedLetter)) <= Asc("Z") Then
                nextColChar = ""
                On Error Resume Next
                bzhao.ReadScreen nextColChar, 1, startReadRow, startReadColumn + 1
                If Err.Number <> 0 Then
                    Err.Clear
                    nextColChar = ""
                End If
                On Error GoTo 0

                If Len(nextColChar) > 0 And Asc(nextColChar) = 32 Then
                    tempLetters(foundCount) = UCase(capturedLetter)
                    foundCount = foundCount + 1
                    emptyRowCount = 0
                Else
                    emptyRowCount = emptyRowCount + 1
                End If
            Else
                emptyRowCount = emptyRowCount + 1
            End If
        Else
            emptyRowCount = emptyRowCount + 1
        End If

        If emptyRowCount >= 3 Then Exit For
    Next

    If foundCount = 0 Then
        DiscoverLineLetters = Array()
        Exit Function
    End If

    ReDim foundLetters(foundCount - 1)
    For i = 0 To foundCount - 1
        foundLetters(i) = tempLetters(i)
    Next

    DiscoverLineLetters = foundLetters
End Function

Function ProcessRoReview()
    Dim lineLetters, i, startIndex
    lineLetters = DiscoverLineLetters()

    If UBound(lineLetters) = -1 Then
        LogResult "INFO", "No service lines detected for review. Skipping RO."
        ProcessRoReview = False
        Exit Function
    End If

    startIndex = 0
    For i = 0 To UBound(lineLetters)
        If lineLetters(i) = "A" Then
            startIndex = i
            Exit For
        End If
    Next

    For i = startIndex To UBound(lineLetters)
        LogResult "INFO", "Reviewing discovered Line " & lineLetters(i)
        WaitForText "COMMAND:"
        EnterTextWithStability "R " & lineLetters(i)

        If Not HandleReviewPrompts(lineLetters(i)) Then
            ProcessRoReview = False
            Exit Function
        End If
    Next

    ProcessRoReview = True
End Function

Function HandleReviewPrompts(lineLetter)
    Dim screenContent, startTime, elapsed, regEx
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = False

    startTime = Timer

    Do
        bzhao.Pause REVIEW_PAUSE
        bzhao.ReadScreen screenContent, 1920, 1, 1

        If InStr(1, screenContent, "COMMAND:", vbTextCompare) > 0 Then
            HandleReviewPrompts = True
            Exit Function
        End If

        If TestPrompt(regEx, screenContent, "LABOR TYPE") Then
            EnterReviewPrompt ""
        ElseIf TestPrompt(regEx, screenContent, "OP CODE.*\([A-Za-z0-9]+\)\?|OPERATION CODE.*\([A-Za-z0-9]+\)\?") Then
            EnterReviewPrompt ""
        ElseIf TestPrompt(regEx, screenContent, "OP CODE.*\?|OPERATION CODE.*\?") Then
            EnterReviewPrompt "I"
        ElseIf TestPrompt(regEx, screenContent, "DESC:") Then
            EnterReviewPrompt ""
        ElseIf TestPrompt(regEx, screenContent, "TECHNICIAN.*\([A-Za-z0-9]+\)\?") Then
            EnterReviewPrompt ""
        ElseIf TestPrompt(regEx, screenContent, "TECHNICIAN.*\?") Then
            LogResult "WARN", "TECHNICIAN prompt has no default for Line " & lineLetter & " — sending 99"
            EnterReviewPrompt "99"
        ElseIf TestPrompt(regEx, screenContent, "ACTUAL HOURS") Then
            EnterReviewPrompt ""
        ElseIf TestPrompt(regEx, screenContent, "SOLD HOURS") Then
            EnterReviewPrompt ""
        ElseIf TestPrompt(regEx, screenContent, "ADD A LABOR OPERATION") Then
            EnterReviewPrompt ""
        End If

        elapsed = Timer - startTime
        If elapsed > 45 Then
            LogResult "ERROR", "Timeout in HandleReviewPrompts for Line " & lineLetter
            HandleReviewPrompts = False
            Exit Function
        End If
    Loop
End Function

Sub EnterReviewPrompt(text)
    If text <> "" Then bzhao.SendKey CStr(text)
    bzhao.Pause 50
    bzhao.SendKey "<NumpadEnter>"
    bzhao.Pause REVIEW_PAUSE
End Sub

Function TestPrompt(regEx, text, pattern)
    regEx.Pattern = pattern
    TestPrompt = regEx.Test(text)
End Function

' --- Close Sequence ---
Function CloseRoFinal()
    Dim mileage, screenContent, startTime, elapsed, pos
    Dim stage: stage = 1

    WaitForText "COMMAND:"
    EnterTextWithStability "FC"

    WaitForText "ALL LABOR POSTED"
    EnterTextWithStability "Y"

    mileage = ""
    bzhao.ReadScreen screenContent, 480, 1, 1
    pos = InStr(1, screenContent, "MILEAGE:", vbTextCompare)
    If pos > 0 Then
        mileage = Trim(Mid(screenContent, pos + 8, 10))
        If InStr(mileage, " ") > 0 Then mileage = Left(mileage, InStr(mileage, " ") - 1)
        LogResult "INFO", "Extracted mileage from screen: " & mileage
    End If

    If mileage = "" Then
        LogResult "INFO", "WARNING: Could not extract mileage. Using '0' as fallback."
        mileage = "0"
    End If

    startTime = Timer
    Do
        bzhao.Pause LOOP_PAUSE
        bzhao.ReadScreen screenContent, 1920, 1, 1
        screenContent = UCase(screenContent)

        If InStr(screenContent, "COMMAND:") > 0 Or InStr(screenContent, UCase(MAIN_PROMPT)) > 0 Then
            LogResult "INFO", "Close flow returned to command/main prompt at Stage " & stage & ". Treating as successful close."
            CloseRoFinal = True
            Exit Function
        ElseIf InStr(screenContent, "IS THIS A COMEBACK") > 0 Then
            LogResult "INFO", "Detected comeback prompt. Sending Y."
            EnterTextWithStability "Y"
            startTime = Timer
        ElseIf stage = 1 And (InStr(screenContent, "MILES OUT") > 0 Or InStr(screenContent, "MILEAGE OUT") > 0) Then
            EnterTextWithStability mileage
            stage = 2
            startTime = Timer
        ElseIf stage = 2 And (InStr(screenContent, "MILES IN") > 0 Or InStr(screenContent, "MILEAGE IN") > 0) Then
            EnterTextWithStability mileage
            stage = 3
            startTime = Timer
        ElseIf stage <= 3 And InStr(screenContent, "O.K. TO CLOSE RO") > 0 Then
            EnterTextWithStability "Y"
            stage = 4
            startTime = Timer
        ElseIf InStr(screenContent, "INVOICE PRINTER") > 0 Then
            EnterTextWithStability "2"
            CloseRoFinal = True
            Exit Function
        End If

        elapsed = Timer - startTime
        If elapsed > 120 Then
            LogResult "ERROR", "Timeout during close sequence at Stage " & stage
            CloseRoFinal = False
            Exit Function
        End If
    Loop
End Function

' --- Navigation Helpers ---
Sub ReturnToMainPrompt()
    Dim screenContent, i, targets, j, isFound, waitStep
    targets = Split(MAIN_PROMPT, "|")

    For waitStep = 1 To 10
        bzhao.Pause LOOP_PAUSE
        bzhao.ReadScreen screenContent, 1920, 1, 1

        For j = 0 To UBound(targets)
            If InStr(1, screenContent, targets(j), vbTextCompare) > 0 Then
                LogResult "INFO", "Confirmed at main prompt: " & targets(j)
                Exit Sub
            End If
        Next
    Next

    For i = 1 To 3
        LogResult "INFO", "ReturnToMainPrompt: Still not at target. Attempting recovery 'E' (" & i & "/3)..."
        bzhao.SendKey "E"
        bzhao.SendKey "<NumpadEnter>"

        For waitStep = 1 To 4
            bzhao.Pause LOOP_PAUSE
            bzhao.ReadScreen screenContent, 1920, 1, 1

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
    Dim elapsed, screenContent, targets, found, i
    targets = Split(targetText, "|")
    elapsed = 0

    Do
        bzhao.Pause LOOP_PAUSE
        elapsed = elapsed + LOOP_PAUSE

        bzhao.ReadScreen screenContent, 1920, 1, 1

        found = False
        For i = 0 To UBound(targets)
            If InStr(1, screenContent, targets(i), vbTextCompare) > 0 Then
                found = True
                Exit For
            End If
        Next

        If found Then Exit Sub

        If elapsed >= 60000 Then
            TerminateScript "Critical timeout waiting for: " & targetText
            Exit Do
        End If
    Loop
End Sub

Sub EnterTextWithStability(text)
    LogResult "INFO", "Input State: Sending text '" & text & "' to terminal."
    bzhao.SendKey CStr(text)
    bzhao.Pause 150
    bzhao.SendKey "<NumpadEnter>"
    bzhao.Pause STABILITY_PAUSE
End Sub

' --- Logging ---
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

Function GetMatchedBlacklistTerm(blacklistTermsCsv, screenContent)
    Dim terms, i, term

    If Trim(blacklistTermsCsv) = "" Then
        GetMatchedBlacklistTerm = ""
        Exit Function
    End If

    terms = Split(blacklistTermsCsv, ",")

    For i = 0 To UBound(terms)
        term = Trim(terms(i))
        If term <> "" Then
            If InStr(1, screenContent, term, vbTextCompare) > 0 Then
                GetMatchedBlacklistTerm = term
                Exit Function
            End If
        End If
    Next

    GetMatchedBlacklistTerm = ""
End Function

Sub TerminateScript(reason)
    LogResult "ERROR", "TERMINATING SCRIPT: " & reason
    On Error Resume Next
    If Not bzhao Is Nothing Then
        bzhao.Disconnect
        bzhao.StopScript
    End If
    On Error GoTo 0
End Sub

' --- Start ---
RunAutomation
