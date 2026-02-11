'=====================================================================================
' Maintenance RO Auto-Closer
' Part of the CDK DMS Automation Suite
'
' Strategic Context: Legacy system scheduled for retirement in 3-6 months.
' Purpose: Automate closing of specific Maintenance ROs with exact footprint match.
'=====================================================================================

Option Explicit

' --- Execution Parameters ---
Dim START_RO: START_RO = 872080 ' Edit this number as needed
Dim TARGET_COUNT: TARGET_COUNT = 500
Dim MAIN_PROMPT: MAIN_PROMPT = "R.O. NUMBER" ' Reduced to substring for better matching
Dim LOG_FILE_PATH: LOG_FILE_PATH = "C:\Temp_alt\CDK\Close_ROs\Maintenance_RO_Closer.log"
Dim DEBUG_LEVEL: DEBUG_LEVEL = 2 ' 1=Error, 2=Info

' --- Picky Match Configuration ---
Dim MATCH_LINES: MATCH_LINES = Array(_
    Array(9, 4, "PM CHANGE OIL & FILTER"), _
    Array(11, 4, "CHECK AND ADJUST TIRE PRESSURE"), _
    Array(13, 4, "MEASURE AND DOCUMENT TIRE TREAD DEPTH") _
)
Dim EXCLUSION_ROW: EXCLUSION_ROW = 15
Dim EXCLUSION_COL: EXCLUSION_COL = 4

' --- CDK Objects ---
Dim bzhao: Set bzhao = CreateObject("BZWhll.WhllObj")

' --- Main Loop ---
Sub RunAutomation()
    Dim currentRo, successfulCount, i
    currentRo = START_RO
    successfulCount = 0

    LogResult "INFO", "Starting Maintenance RO Auto-Closer at RO: " & START_RO
    
    bzhao.Connect ""
    
    For i = 1 To TARGET_COUNT
        LogResult "INFO", "Processing RO: " & currentRo & " (" & i & "/" & TARGET_COUNT & ")"
        
        ' Ensure we are at the main prompt (Checks Row 11 as confirmed)
        WaitForText MAIN_PROMPT
        
        ' Enter RO Number
        EnterTextWithStability currentRo
        
        ' Check for errors or closed status
        If IsRoProcessable(currentRo) Then
            ' Check "Picky" Match Logic
            If CheckPickyMatch() Then
                LogResult "INFO", "Match found for RO: " & currentRo & ". Proceeding to review."
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
                LogResult "INFO", "RO: " & currentRo & " does not match footprint. Skipping."
            End If
        End If

        ' Always send "E" to return to main prompt
        ReturnToMainPrompt()
        
        currentRo = currentRo + 1
    Next

    LogResult "INFO", "Automation complete. Total successful closures: " & successfulCount
    MsgBox "Maintenance RO Auto-Closer Finished." & vbCrLf & "Successful Closures: " & successfulCount, vbInformation
    
    bzhao.Disconnect
End Sub

' --- Helper Subroutines & Functions ---

Function IsRoProcessable(roNumber)
    Dim screenContent
    bzhao.Pause 2000
    bzhao.ReadScreen screenContent, 80, 2, 1
    
    If InStr(screenContent, "NOT ON FILE") > 0 Then
        LogResult "INFO", "RO " & roNumber & " NOT ON FILE. Skipping."
        IsRoProcessable = False
        Exit Function
    ElseIf InStr(screenContent, "is closed") > 0 Or InStr(screenContent, "ALREADY CLOSED") > 0 Then
        LogResult "INFO", "RO " & roNumber & " ALREADY CLOSED. Skipping."
        IsRoProcessable = False
        Exit Function
    End If
    
    IsRoProcessable = True
End Function

Function CheckPickyMatch()
    Dim row, col, expectedText, screenContent, i, anchorRow
    
    ' Phase 1: Hunt for the anchor (Line "A" in Column 1)
    anchorRow = 0
    For i = 8 To 15 ' Search expected range where LC A usually lives
        bzhao.ReadScreen screenContent, 1, i, 1
        If UCase(Trim(screenContent)) = "A" Then
            anchorRow = i
            Exit For
        End If
    Next

    If anchorRow = 0 Then
        LogResult "INFO", "Footprint mismatch: Line 'A' not detected in Col 1 (checked rows 8-15)."
        CheckPickyMatch = False
        Exit Function
    Else
        LogResult "INFO", "Line 'A' detected at Row " & anchorRow
    End If

    ' Phase 2: Verify descriptions relative to anchor (A=Anchor, B=A+2, C=A+4)
    Dim checkLayout: checkLayout = Array(_
        Array(anchorRow, 4, MATCH_LINES(0)(2)), _
        Array(anchorRow + 2, 4, MATCH_LINES(1)(2)), _
        Array(anchorRow + 4, 4, MATCH_LINES(2)(2)) _
    )
    
    For i = 0 To UBound(checkLayout)
        row = checkLayout(i)(0)
        col = checkLayout(i)(1)
        expectedText = checkLayout(i)(2)
        
        bzhao.ReadScreen screenContent, 40, row, col
        
        If InStr(1, screenContent, expectedText, vbTextCompare) = 0 Then
            LogResult "INFO", "Mismatch at Row " & row & ". Expected: '" & expectedText & "' | Found: '" & Trim(screenContent) & "'"
            CheckPickyMatch = False
            Exit Function
        End If
    Next
    
    ' Phase 3: Check Exclusion (Line D at Anchor + 6 should be empty)
    Dim exclusionRowFinal: exclusionRowFinal = anchorRow + 6
    bzhao.ReadScreen screenContent, 20, exclusionRowFinal, 4 
    If Trim(screenContent) <> "" Then
        LogResult "INFO", "Exclusion match failed: Row " & exclusionRowFinal & " is not empty ('" & Trim(screenContent) & "'). Skipping."
        CheckPickyMatch = False
        Exit Function
    End If
    
    CheckPickyMatch = True
End Function

Function ProcessRoReview()
    Dim lineLetters, i
    lineLetters = Array("A", "B", "C")
    
    For i = 0 To UBound(lineLetters)
        LogResult "INFO", "Reviewing Line " & lineLetters(i)
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
    Dim screenContent, startTime, elapsed
    startTime = Timer
    
    Do
        bzhao.Pause 1000
        ' Read the entire screen to handle prompts that might appear mid-screen
        bzhao.ReadScreen screenContent, 1920, 1, 1
        screenContent = UCase(screenContent)
        
        ' Exit condition: Back to COMMAND prompt
        If InStr(screenContent, "COMMAND:") > 0 Then
            HandleReviewPrompts = True
            Exit Function
        End If
        
        ' Match Prompts
        If InStr(screenContent, "LABOR TYPE") > 0 Or InStr(screenContent, "LTYPE") > 0 Then
            EnterTextWithStability ""
        ElseIf InStr(screenContent, "OP CODE") > 0 Or InStr(screenContent, "OPERATION CODE") > 0 Then
            EnterTextWithStability ""
        ElseIf InStr(screenContent, "DESC:") > 0 Then
            EnterTextWithStability ""
        ElseIf InStr(screenContent, "TECHNICIAN") > 0 Then
            EnterTextWithStability "99"
        ElseIf InStr(screenContent, "ACTUAL HOURS") > 0 Then
            EnterTextWithStability ""
        ElseIf InStr(screenContent, "SOLD HOURS") > 0 Then
            EnterTextWithStability ""
        ElseIf InStr(screenContent, "ADD A LABOR OPERATION") > 0 Then
            EnterTextWithStability "" ' Defaults to "N"
        End If
        
        elapsed = Timer - startTime
        If elapsed > 45 Then ' Increased timeout for slow terminal moves
            LogResult "ERROR", "Timeout in HandleReviewPrompts for Line " & lineLetter
            HandleReviewPrompts = False
            Exit Function
        End If
    Loop
End Function

Function CloseRoFinal()
    Dim mileage, screenContent, startTime, elapsed
    Dim lastActionTime: lastActionTime = Timer
    
    ' Phase II: The Closing
    WaitForText "COMMAND:"
    EnterTextWithStability "FC"
    
    ' ALL LABOR POSTED
    WaitForText "ALL LABOR POSTED"
    EnterTextWithStability "Y"
    
    ' MILEAGE / MILES OUT
    ' Read from Row 2, Col 47 (Mileage Header)
    bzhao.ReadScreen mileage, 10, 2, 47
    mileage = Trim(mileage)
    LogResult "INFO", "Using mileage from header: " & mileage
    
    ' Define sequence-based state tracking to avoid double-tapping
    Dim stage: stage = 1 ' 1=MilesOut, 2=MilesIn, 3=OkToClose, 4=Printer
    
    startTime = Timer
    Do
        bzhao.Pause 1000
        bzhao.ReadScreen screenContent, 1920, 1, 1
        screenContent = UCase(screenContent)
        
        ' Stage 1: MILEAGE OUT
        If stage = 1 And (InStr(screenContent, "MILES OUT") > 0 Or InStr(screenContent, "MILEAGE OUT") > 0) Then
            EnterTextWithStability mileage
            stage = 2
            startTime = Timer ' Reset timer for next expected prompt due to 5-10s delay
        
        ' Stage 2: MILEAGE IN
        ElseIf stage = 2 And (InStr(screenContent, "MILES IN") > 0 Or InStr(screenContent, "MILEAGE IN") > 0) Then
            EnterTextWithStability mileage
            stage = 3
            startTime = Timer
            
        ' Stage 3: OK TO CLOSE (Sometimes Miles In doesn't appear, or OK appears immediately)
        ElseIf stage >= 2 And stage <= 3 And InStr(screenContent, "O.K. TO CLOSE RO") > 0 Then
            EnterTextWithStability "Y"
            stage = 4
            startTime = Timer
            
        ' Stage 4: INVOICE PRINTER
        ElseIf stage >= 3 And InStr(screenContent, "INVOICE PRINTER") > 0 Then
            EnterTextWithStability "2"
            CloseRoFinal = True
            Exit Function
        End If
        
        elapsed = Timer - startTime
        If elapsed > 120 Then ' Give closing 2 minutes for slow UI/Printer logic
            LogResult "ERROR", "Timeout during Phase II Closing sequence at Stage " & stage
            CloseRoFinal = False
            Exit Function
        End If
    Loop
End Function

Sub ReturnToMainPrompt()
    Dim screenContent, i
    ' Try sending "E" a couple of times to get back to the RO number prompt
    For i = 1 To 3
        ' Read the entire screen (1920 chars) to find the main prompt anywhere (e.g., Row 11)
        bzhao.ReadScreen screenContent, 1920, 1, 1
        If InStr(screenContent, MAIN_PROMPT) > 0 Then Exit Sub
        
        bzhao.SendKey "E"
        bzhao.SendKey "<NumpadEnter>"
        bzhao.Pause 1000
    Next
End Sub

Sub WaitForText(targetText)
    Dim elapsed, screenContent, targets, found, i, isMainPrompt
    targets = Split(targetText, "|")
    elapsed = 0
    isMainPrompt = (InStr(UCase(targetText), UCase(MAIN_PROMPT)) > 0)
    
    Do
        bzhao.Pause 500
        elapsed = elapsed + 500
        
        ' Read the entire screen (24 rows * 80 cols) to be robust
        bzhao.ReadScreen screenContent, 1920, 1, 1
        screenContent = UCase(screenContent)
        
        found = False
        For i = 0 To UBound(targets)
            If InStr(screenContent, UCase(targets(i))) > 0 Then
                found = True
                Exit For
            End If
        Next
        
        If found Then Exit Sub
        
        ' Blind Entry Fallback: If we are looking for the main RO prompt and 5 seconds have passed,
        ' we will "blindly" assume we are there (or the prompt is hidden/scrolled) and try to proceed.
        If isMainPrompt And elapsed >= 5000 Then
            LogResult "INFO", "Prompt '" & targetText & "' not detected after 5s. Proceeding blindly."
            Exit Sub
        End If

        If elapsed >= 15000 Then
            LogResult "ERROR", "Timeout waiting for: " & targetText
            bzhao.StopScript
        End If
    Loop
End Sub

Sub EnterTextWithStability(text)
    bzhao.SendKey CStr(text)
    bzhao.Pause 150
    bzhao.SendKey "<NumpadEnter>"
    bzhao.Pause 2000 ' PRD Stability requirement: 2-second pause after every command
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

' Execute
RunAutomation
