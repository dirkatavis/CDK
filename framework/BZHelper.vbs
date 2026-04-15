'=====================================================================================
' BZHelper.vbs
' Part of the CDK DMS Automation Suite ??? framework\BZHelper.vbs
'
' Purpose: Authoritative shared library for BlueZone terminal automation.
'          Provides connection management, screen reading, text detection,
'          keystroke sending, and prompt waiting.
'
' Usage:
'   Load after PathHelper.vbs via ExecuteGlobal:
'     ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\BZHelper.vbs")).ReadAll
'
'   The calling script must DECLARE g_bzhao before loading this file (required by
'   Option Explicit). It must be SET before any BZHelper function is first called,
'   but does not need to be set at load time ??? loading with g_bzhao = Nothing is safe.
'
'   Typical bootstrap (g_bzhao set immediately):
'     Dim g_bzhao: Set g_bzhao = CreateObject("BZWhll.WhllObj")
'     ExecuteGlobal g_fso.OpenTextFile(...BZHelper.vbs).ReadAll
'
'   Deferred assignment (e.g. test-mode scripts that conditionally use MockBzhao):
'     Dim g_bzhao
'     ExecuteGlobal g_fso.OpenTextFile(...BZHelper.vbs).ReadAll
'     ' ... later, before any BZHelper function is called:
'     Set g_bzhao = CreateObject("BZWhll.WhllObj")  ' or Set g_bzhao = New MockBzhao
'
'   This library does NOT instantiate g_bzhao. Each script owns its own connection
'   object so that multiple scripts can run independently without session conflicts.
'
' Load guard: safe to ExecuteGlobal multiple times.
'=====================================================================================

If Not IsObject(g_BZHelper_Loaded) Then
Set g_BZHelper_Loaded = CreateObject("Scripting.Dictionary")

'-------------------------------------------------------------------------------------
' BZH_Log ??? Internal logging shim.
' Calls LogResult(level, message) if the calling script has defined it.
' Silently no-ops if LogResult is not available, so BZHelper works in any script.
'-------------------------------------------------------------------------------------
Sub BZH_Log(level, message)
    On Error Resume Next
    LogResult level, message
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------------
' ConnectBZ ??? Connect g_bzhao to the active BlueZone session.
' Returns True on success, False on failure.
'-------------------------------------------------------------------------------------
Function ConnectBZ()
    On Error Resume Next

    If Not IsObject(g_bzhao) Then
        BZH_Log "ERROR", "BZHelper.ConnectBZ: g_bzhao is not initialised. Declare and Set it before loading BZHelper."
        ConnectBZ = False
        Exit Function
    End If

    g_bzhao.Connect ""
    If Err.Number <> 0 Then
        BZH_Log "ERROR", "BZHelper.ConnectBZ: Connection failed ??? " & Err.Description
        Err.Clear
        ConnectBZ = False
    Else
        BZH_Log "INFO", "BZHelper.ConnectBZ: Connected to BlueZone session."
        ConnectBZ = True
    End If
    On Error GoTo 0
End Function

'-------------------------------------------------------------------------------------
' DisconnectBZ ??? Cleanly disconnect and release g_bzhao.
'-------------------------------------------------------------------------------------
Sub DisconnectBZ()
    On Error Resume Next
    If IsObject(g_bzhao) Then
        g_bzhao.Disconnect
        Set g_bzhao = Nothing
        BZH_Log "INFO", "BZHelper.DisconnectBZ: Disconnected from BlueZone session."
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------------
' BZReadScreen ??? Read a block of characters from the terminal screen.
' Parameters:
'   length  ??? number of characters to read (max 1920 for full 24x80 screen)
'   row     ??? starting row (1-based)
'   col     ??? starting column (1-based)
' Returns the screen content as a string.
'-------------------------------------------------------------------------------------
Function BZReadScreen(length, row, col)
    Dim buf
    buf = ""
    On Error Resume Next
    g_bzhao.ReadScreen buf, length, row, col
    If Err.Number <> 0 Then
        BZH_Log "ERROR", "BZHelper.BZReadScreen: ReadScreen failed at row " & row & ", col " & col & " ??? " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    BZReadScreen = buf
End Function

'-------------------------------------------------------------------------------------
' IsTextPresent ??? Search the full terminal screen (24 rows x 80 cols) for text.
' Pipe-delimited multi-target: "PROMPT A|PROMPT B" returns True if either matches.
' Search is case-insensitive.
' Returns True if any target is found, False otherwise.
'-------------------------------------------------------------------------------------
Function IsTextPresent(searchText)
    Dim targets, i, target, lineNum, lineContent

    IsTextPresent = False
    If Len(Trim(searchText)) = 0 Then Exit Function

    targets = Split(searchText, "|")

    On Error Resume Next
    For lineNum = 1 To 24
        lineContent = ""
        g_bzhao.ReadScreen lineContent, 80, lineNum, 1
        If Err.Number <> 0 Then
            Err.Clear
        Else
            For i = 0 To UBound(targets)
                target = Trim(targets(i))
                If Len(target) > 0 Then
                    If InStr(1, lineContent, target, vbTextCompare) > 0 Then
                        IsTextPresent = True
                        Exit Function
                    End If
                End If
            Next
        End If
    Next
    On Error GoTo 0
End Function

'-------------------------------------------------------------------------------------
' BZSendKey ??? Send a keystroke or text string to the terminal.
' Handles both special keys (e.g. "<NumpadEnter>") and plain text.
' Returns True on success, False on error.
'-------------------------------------------------------------------------------------
Function BZSendKey(keyValue)
    On Error Resume Next
    BZSendKey = False

    If Len(keyValue) = 0 Then
        BZH_Log "WARN", "BZHelper.BZSendKey: Empty key value ??? nothing sent."
        On Error GoTo 0
        Exit Function
    End If

    g_bzhao.SendKey keyValue
    If Err.Number <> 0 Then
        BZH_Log "ERROR", "BZHelper.BZSendKey: Failed to send '" & keyValue & "' ??? " & Err.Description
        Err.Clear
    Else
        BZSendKey = True
    End If
    On Error GoTo 0
End Function

'-------------------------------------------------------------------------------------
' WaitMs ??? Busy-wait for a number of milliseconds.
' Uses Timer-based loop; handles midnight rollover (Timer resets to 0 at midnight).
'-------------------------------------------------------------------------------------
Sub WaitMs(milliseconds)
    If milliseconds <= 0 Then Exit Sub
    Dim startTime, endTime
    startTime = Timer
    endTime = startTime + (milliseconds / 1000)
    Do While Timer < endTime
        ' Midnight rollover: Timer resets to 0 at midnight
        If Timer < startTime Then Exit Do
    Loop
End Sub

'-------------------------------------------------------------------------------------
' WaitForPrompt ??? Wait for a prompt to appear on screen, optionally send input.
'
' Parameters:
'   promptText  ??? text to wait for (pipe-delimited for multi-target: "A|B")
'   inputValue  ??? text or key to send once prompt is detected (pass "" to skip)
'                 Special keys detected by presence of "<" and ">" (e.g. "<NumpadEnter>")
'   sendEnter   ??? Boolean; if True, sends <NumpadEnter> after inputValue
'   timeoutMs   ??? milliseconds to wait before giving up (0 = use default 5000ms)
'   description ??? optional label used in log messages (pass "" if not needed)
'
' Returns True if the prompt was found, False if timeout elapsed.
'
' Canonical version. Authoritative source: framework\BZHelper.vbs
' Derived from: PostFinalCharges.vbs (structure, logging, error handling)
'               Open_RO.vbs (midnight rollover guard, <> key detection, NumpadEnter)
'-------------------------------------------------------------------------------------
Function WaitForPrompt(promptText, inputValue, sendEnter, timeoutMs, description)
    Dim found, waitStart, waitElapsed, label

    If timeoutMs <= 0 Then timeoutMs = 5000

    If Len(Trim(description)) > 0 Then label = description Else label = promptText
    found = False
    waitStart = Timer

    BZH_Log "INFO", "BZHelper.WaitForPrompt: Waiting for '" & label & "' (timeout " & timeoutMs & "ms)"

    Do
        If IsTextPresent(promptText) Then
            found = True
            BZH_Log "INFO", "BZHelper.WaitForPrompt: Found '" & label & "'"

            ' Send input if provided
            If Len(inputValue) > 0 Then
                On Error Resume Next
                ' Detect special key sequences (e.g. <NumpadEnter>, <Enter>)
                If InStr(inputValue, "<") > 0 And InStr(inputValue, ">") > 0 Then
                    g_bzhao.SendKey inputValue
                Else
                    g_bzhao.SendKey inputValue
                End If
                If Err.Number <> 0 Then
                    BZH_Log "ERROR", "BZHelper.WaitForPrompt: Failed to send input '" & inputValue & "' ??? " & Err.Description
                    Err.Clear
                Else
                    BZH_Log "INFO", "BZHelper.WaitForPrompt: Sent '" & inputValue & "'"
                End If
                On Error GoTo 0
                WaitMs 100
            End If

            ' Send Enter if requested
            If sendEnter Then
                On Error Resume Next
                g_bzhao.SendKey "<NumpadEnter>"
                If Err.Number <> 0 Then
                    BZH_Log "ERROR", "BZHelper.WaitForPrompt: Failed to send Enter ??? " & Err.Description
                    Err.Clear
                Else
                    BZH_Log "INFO", "BZHelper.WaitForPrompt: Enter sent."
                End If
                On Error GoTo 0
                WaitMs 100
            End If

            Exit Do
        End If

        WaitMs 50

        ' Elapsed calculation with midnight rollover guard
        waitElapsed = Timer - waitStart
        If waitElapsed < 0 Then waitElapsed = waitElapsed + 86400
        waitElapsed = waitElapsed * 1000

        If waitElapsed > timeoutMs Then
            BZH_Log "WARN", "BZHelper.WaitForPrompt: Timeout after " & timeoutMs & "ms waiting for '" & label & "'"
            Exit Do
        End If
    Loop

    WaitForPrompt = found
End Function

'-------------------------------------------------------------------------------------
' WaitForAnyOf ??? Wait for any one of several pipe-delimited targets to appear.
'
' Parameters:
'   targets    ??? pipe-delimited list of strings to search for (e.g. "CAMP|PASTEUR")
'   timeoutMs  ??? milliseconds to wait before giving up (0 = use default 5000ms)
'
' Returns True if any target is found, False if timeout elapsed.
'
' Uses IsTextPresent internally ??? search is case-insensitive, row-by-row.
' Canonical version. Authoritative source: framework\BZHelper.vbs
'-------------------------------------------------------------------------------------
Function WaitForAnyOf(targets, timeoutMs)
    Dim waitStart, waitElapsed

    If timeoutMs <= 0 Then timeoutMs = 5000

    BZH_Log "INFO", "BZHelper.WaitForAnyOf: Waiting for '" & targets & "' (timeout " & timeoutMs & "ms)"

    waitStart = Timer
    Do
        If IsTextPresent(targets) Then
            BZH_Log "INFO", "BZHelper.WaitForAnyOf: Found match in '" & targets & "'"
            WaitForAnyOf = True
            Exit Function
        End If

        WaitMs 500

        waitElapsed = Timer - waitStart
        If waitElapsed < 0 Then waitElapsed = waitElapsed + 86400
        waitElapsed = waitElapsed * 1000

        If waitElapsed > timeoutMs Then
            BZH_Log "WARN", "BZHelper.WaitForAnyOf: Timeout after " & timeoutMs & "ms waiting for '" & targets & "'"
            WaitForAnyOf = False
            Exit Function
        End If
    Loop
End Function

'-------------------------------------------------------------------------------------
' BZH_RecoverFromVehidError ??? Shared recovery for "VEHID not on file" errors.
'
' Called when the terminal shows "PRESS RETURN TO CONTINUE" after a VEHID lookup
' failure. Navigates back to the PFC function menu and selects the caller-specified
' option, leaving the terminal in a stable state for the caller to continue.
'
' Parameters:
'   employeeNumber   ??? Employee ID string (e.g. "18351"), read from config
'   nameConfirmText  ??? Pipe-delimited name fragment(s) to wait for on the
'                      name confirmation screen (e.g. "CAMP|PASTEUR"), read from config
'   menuOption       ??? Option to select at the PFC ENTER OPTION menu:
'                        "1" = return to main RO screen (Maintenance_RO_Closer)
'                        "2" = return to PFC sequence prompt (PFC_Scrapper)
'
' Returns True on success, False if any step times out.
'-------------------------------------------------------------------------------------
Function BZH_RecoverFromVehidError(employeeNumber, nameConfirmText, menuOption)
    BZH_RecoverFromVehidError = False

    BZH_Log "INFO", "BZHelper.BZH_RecoverFromVehidError: Step 1 - dismissing VEHID error."
    g_bzhao.SendKey "<Enter>"
    If Not WaitForPrompt("FUNCTION CODE", "", False, 5000, "FUNCTION CODE after VEHID dismiss") Then
        BZH_Log "ERROR", "BZHelper.BZH_RecoverFromVehidError: Step 1 failed - FUNCTION CODE not found."
        Exit Function
    End If

    BZH_Log "INFO", "BZHelper.BZH_RecoverFromVehidError: Step 2 - entering PFC."
    g_bzhao.SendKey "PFC"
    g_bzhao.Pause 100
    g_bzhao.SendKey "<NumpadEnter>"
    If Not WaitForPrompt("EMPLOYEE NUMBER", "", False, 10000, "EMPLOYEE NUMBER prompt") Then
        BZH_Log "ERROR", "BZHelper.BZH_RecoverFromVehidError: Step 2 failed - EMPLOYEE NUMBER not found."
        Exit Function
    End If

    BZH_Log "INFO", "BZHelper.BZH_RecoverFromVehidError: Step 3 - entering employee number."
    g_bzhao.SendKey employeeNumber
    g_bzhao.Pause 100
    g_bzhao.SendKey "<NumpadEnter>"
    If Not WaitForAnyOf(nameConfirmText, 10000) Then
        BZH_Log "ERROR", "BZHelper.BZH_RecoverFromVehidError: Step 3 failed - name confirmation not found."
        Exit Function
    End If

    BZH_Log "INFO", "BZHelper.BZH_RecoverFromVehidError: Step 4 - confirming employee name."
    g_bzhao.SendKey "<NumpadEnter>"
    If Not WaitForPrompt("ENTER OPTION", "", False, 10000, "ENTER OPTION menu") Then
        BZH_Log "ERROR", "BZHelper.BZH_RecoverFromVehidError: Step 4 failed - ENTER OPTION menu not found."
        Exit Function
    End If

    BZH_Log "INFO", "BZHelper.BZH_RecoverFromVehidError: Step 5 - selecting option " & menuOption & "."
    g_bzhao.SendKey menuOption
    g_bzhao.Pause 100
    g_bzhao.SendKey "<NumpadEnter>"

    BZH_Log "INFO", "BZHelper.BZH_RecoverFromVehidError: Recovery complete."
    BZH_RecoverFromVehidError = True
End Function

'-------------------------------------------------------------------------------------
' BZH_GetMatchedBlacklistTerm ??? Scan all RO service line pages for a blacklist term.
'
' Parameters:
'   blacklistTermsCsv ??? comma-separated list of terms to search for (case-insensitive)
'   pauseMs           ??? milliseconds to pause between page advances (configurable per script)
'
' Pages through multi-screen ROs using the CDK pagination pattern:
'   "(MORE ON NEXT SCREEN)" on row 22 ??? advance with "N" + NumpadEnter
'   "(END OF DISPLAY)" on row 22      ??? last page reached
' Returns to page 1 via "B" + NumpadEnter after scanning.
'
' Returns the first matched term, or empty string if no match found.
'-------------------------------------------------------------------------------------
Function BZH_GetMatchedBlacklistTerm(blacklistTermsCsv, pauseMs)
    Dim terms, i, term, lineNum, lineContent
    Dim pagesAdvanced, matchedTerm, pageIndicator
    Dim foundMatch, doneScanning, p

    BZH_GetMatchedBlacklistTerm = ""

    blacklistTermsCsv = Trim(CStr(blacklistTermsCsv))
    If Len(blacklistTermsCsv) = 0 Then Exit Function
    If pauseMs <= 0 Then pauseMs = 500

    terms = Split(blacklistTermsCsv, ",")
    pagesAdvanced = 0
    matchedTerm = ""
    doneScanning = False

    Do While Not doneScanning

        ' Scan all 24 rows of current page
        foundMatch = False
        On Error Resume Next
        For lineNum = 1 To 24
            lineContent = ""
            g_bzhao.ReadScreen lineContent, 80, lineNum, 1
            If Err.Number <> 0 Then
                Err.Clear
            Else
                For i = 0 To UBound(terms)
                    term = Trim(terms(i))
                    If Len(term) > 0 Then
                        If InStr(1, lineContent, term, vbTextCompare) > 0 Then
                            matchedTerm = term
                            foundMatch = True
                            Exit For
                        End If
                    End If
                Next
            End If
            If foundMatch Then Exit For
        Next
        On Error GoTo 0

        If foundMatch Then
            BZH_Log "INFO", "BZHelper.BZH_GetMatchedBlacklistTerm: Matched '" & matchedTerm & "' on page " & (pagesAdvanced + 1) & "."
            doneScanning = True
        Else
            ' Check row 22 for pagination indicator
            pageIndicator = ""
            On Error Resume Next
            g_bzhao.ReadScreen pageIndicator, 80, 22, 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0

            If InStr(1, pageIndicator, "(END OF DISPLAY)", vbTextCompare) > 0 Then
                doneScanning = True
            ElseIf InStr(1, pageIndicator, "(MORE ON NEXT SCREEN)", vbTextCompare) > 0 Then
                BZH_Log "INFO", "BZHelper.BZH_GetMatchedBlacklistTerm: Advancing to page " & (pagesAdvanced + 2) & "."
                On Error Resume Next
                g_bzhao.SendKey "N"
                g_bzhao.SendKey "<NumpadEnter>"
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                g_bzhao.Pause pauseMs
                pagesAdvanced = pagesAdvanced + 1
            Else
                BZH_Log "INFO", "BZHelper.BZH_GetMatchedBlacklistTerm: No pagination indicator on row 22 ??? treating as end of display."
                doneScanning = True
            End If
        End If

    Loop

    ' Return to page 1 if we paged forward
    If pagesAdvanced > 0 Then
        BZH_Log "INFO", "BZHelper.BZH_GetMatchedBlacklistTerm: Returning to page 1 (" & pagesAdvanced & " page(s))."
        For p = 1 To pagesAdvanced
            On Error Resume Next
            g_bzhao.SendKey "B"
            g_bzhao.SendKey "<NumpadEnter>"
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            g_bzhao.Pause pauseMs
        Next
    End If

    BZH_GetMatchedBlacklistTerm = matchedTerm
End Function

End If ' g_BZHelper_Loaded load guard

