'=====================================================================================
' Maintenance RO Auto-Closer
' Part of the CDK DMS Automation Suite
'
' Strategic Context: Legacy system scheduled for retirement in 3-6 months.
' Purpose: Automate closing of specific Maintenance ROs with exact footprint match.
'=====================================================================================

Option Explicit

' --- Execution Parameters ---
Dim MAIN_PROMPT: MAIN_PROMPT = "R.O. NUMBER|SEQUENCE NUMBER|ENTER SEQUENCE NUMBER" ' Accept both as valid input states
Dim LOG_FILE_PATH: LOG_FILE_PATH = "C:\Temp_alt\CDK\Maintenance_RO_Closer\Maintenance_RO_Closer.log"
Dim CRITERIA_FILE: CRITERIA_FILE = "C:\Temp_alt\CDK\Maintenance_RO_Closer\PM_Match_Criteria.txt"
Dim DEBUG_LEVEL: DEBUG_LEVEL = 2 ' 1=Error, 2=Info
Dim RO_LIST_PATH: RO_LIST_PATH = "C:\Temp_alt\CDK\Maintenance_RO_Closer\RO_List.csv"

' --- Picky Match State ---
Dim CriteriaA, CriteriaB, CriteriaC

' --- CDK Objects ---
Dim bzhao: Set bzhao = CreateObject("BZWhll.WhllObj")

' --- Main Loop ---
Sub RunAutomation()
    Dim currentRo, successfulCount, fso, ts, strLine, roFromCsv
    successfulCount = 0

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(RO_LIST_PATH) Then
        MsgBox "Error: RO List file not found at: " & RO_LIST_PATH, vbCritical, "File Not Found"
        LogResult "ERROR", "RO List file not found: " & RO_LIST_PATH
        Exit Sub
    End If

    LogResult "INFO", "Starting Maintenance RO Auto-Closer using list: " & RO_LIST_PATH
    
    ' Load Matching Criteria from External File
    LoadMatchCriteria()
    
    bzhao.Connect ""
    
    Set ts = fso.OpenTextFile(RO_LIST_PATH, 1) ' 1 = ForReading
    
    ' Skip header row
    If Not ts.AtEndOfStream Then ts.ReadLine
    
    Do While Not ts.AtEndOfStream
        strLine = Trim(ts.ReadLine)
        If strLine <> "" Then
            ' Handle potential CSV splitting (take first column)
            roFromCsv = Split(strLine, ",")(0)
            currentRo = Trim(roFromCsv)
            
            ' Validate 6-digit RO
            If Len(currentRo) = 6 And IsNumeric(currentRo) Then
                LogResult "INFO", "Processing RO: " & currentRo
                
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
            ElseIf Len(currentRo) > 0 Then
                LogResult "INFO", "Skipping invalid RO format: " & currentRo & " (Must be 6 digits)"
            End If
        End If
    Loop
    
    ts.Close

    LogResult "INFO", "Automation complete. Total successful closures: " & successfulCount
    MsgBox "Maintenance RO Auto-Closer Finished." & vbCrLf & "Successful Closures: " & successfulCount, vbInformation
    
    bzhao.Disconnect
End Sub

' --- Helper Subroutines & Functions ---

Function IsRoProcessable(roNumber)
    Dim screenContent
    bzhao.Pause 2000
    ' Read screen starting from Row 2 down to Row 6 to catch status (Row 5) and RO info
    ' We also read more to catch system errors (Pick/BASIC errors)
    bzhao.ReadScreen screenContent, 1920, 1, 1 
    
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
        ' This is actually a valid prompt now, but we skip it here to let the main loop handle it
        LogResult "INFO", "RO " & roNumber & " prompted for Sequence Number. Treating as valid prompt."
        IsRoProcessable = False
        Exit Function
    ElseIf InStr(1, screenContent, "READY TO POST", vbTextCompare) = 0 Then
        LogResult "INFO", "RO " & roNumber & " status is NOT 'READY TO POST'. Found instead: " & GetStatusSnip(screenContent)
        IsRoProcessable = False
        Exit Function
    End If
    
    IsRoProcessable = True
End Function

Function GetStatusSnip(screenContent)
    ' Helper to grab a small snip of where the status usually is for logging
    Dim pos: pos = InStr(1, screenContent, "STATUS:", vbTextCompare)
    If pos > 0 Then
        GetStatusSnip = "'" & Trim(Mid(screenContent, pos, 30)) & "'"
    Else
        GetStatusSnip = "(Status line not found in read buffer)"
    End If
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

    ' Phase 2: Verify descriptions relative to anchor using Variant Matching
    ' Layout remains: Line A (anchor), B (anchor+2), C (anchor+4)
    Dim checkItems: checkItems = Array(_
        Array(anchorRow, CriteriaA, "A"), _
        Array(anchorRow + 2, CriteriaB, "B"), _
        Array(anchorRow + 4, CriteriaC, "C") _
    )
    
    For i = 0 To UBound(checkItems)
        row = checkItems(i)(0)
        bzhao.ReadScreen screenContent, 50, row, 4 ' Read 50 chars for comparison
        
        If Not MatchesAnyVariant(screenContent, checkItems(i)(1)) Then
            LogResult "INFO", "Mismatch at Row " & row & " (Line " & checkItems(i)(2) & "). Found: '" & Trim(screenContent) & "'"
            CheckPickyMatch = False
            Exit Function
        End If
    Next
    
    ' Phase 3: Exclusion Check - Skip if Line D exists
    ' We check Column 1 at anchor + 6 for a "D"
    bzhao.ReadScreen screenContent, 1, anchorRow + 6, 1
    If UCase(Trim(screenContent)) = "D" Then
        LogResult "INFO", "Exclusion match failed: Line 'D' detected at Row " & (anchorRow + 6) & ". Too many service lines. Skipping."
        CheckPickyMatch = False
        Exit Function
    End If
    
    CheckPickyMatch = True
End Function

Sub LoadMatchCriteria()
    Dim fso, txtFile, line, parts, key, variants, i, commentPos, vVal
    Dim cleanArr, count
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(CRITERIA_FILE) Then
        LogResult "ERROR", "CRITICAL ERROR: Configuration file missing: " & CRITERIA_FILE
        MsgBox "Missing mandatory config file: " & CRITERIA_FILE, vbCritical
        bzhao.StopScript
    End If
    
    Set txtFile = fso.OpenTextFile(CRITERIA_FILE, 1)
    Do Until txtFile.AtEndOfStream
        line = txtFile.ReadLine
        
        ' 1. Strip comments
        commentPos = InStr(line, "#")
        If commentPos > 0 Then line = Left(line, commentPos - 1)
        line = Trim(line)
        
        ' 2. Process valid key=value lines
        If line <> "" And InStr(line, "=") > 0 Then
            parts = Split(line, "=")
            key = UCase(Trim(parts(0)))
            variants = Split(parts(1), "|")
            
            ' Filter variants
            count = 0
            cleanArr = Array() ' Initialize as empty array
            For i = 0 To UBound(variants)
                vVal = Trim(variants(i))
                If vVal <> "" Then
                    ReDim Preserve cleanArr(count)
                    cleanArr(count) = vVal
                    count = count + 1
                End If
            Next
            
            If count > 0 Then
                Select Case key
                    Case "A": CriteriaA = cleanArr
                    Case "B": CriteriaB = cleanArr
                    Case "C": CriteriaC = cleanArr
                End Select
            End If
        End If
    Loop
    txtFile.Close
    
    ' Validate we have all required lines
    If Not IsArray(CriteriaA) Or Not IsArray(CriteriaB) Or Not IsArray(CriteriaC) Then
        LogResult "ERROR", "CRITICAL ERROR: Config file incomplete or corrupted (Lines A, B, and C required)."
        bzhao.StopScript
    End If
End Sub

Function MatchesAnyVariant(screenStr, variantsArray)
    Dim i, sText, vText
    MatchesAnyVariant = False
    
    ' Truncate screen string to 50 chars for comparison as requested
    sText = Left(Trim(screenStr), 50)
    
    For i = 0 To UBound(variantsArray)
        vText = Left(variantsArray(i), 50)
        ' Use InStr(..., 1) = 1 for "StartsWith" logic (case-insensitive)
        If InStr(1, sText, vText, vbTextCompare) = 1 Then
            MatchesAnyVariant = True
            Exit Function
        End If
    Next
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
    Dim screenContent, i, promptPos, targets, j, isFound
    targets = Split(MAIN_PROMPT, "|")
    
    ' SAFETY NET: We only send keys IF we are truly lost.
    ' If the prompt is visible anywhere (even shifted by an error), we stay put.
    For i = 1 To 5
        bzhao.Pause 500 ' Reduced to 500ms for production speed
        bzhao.ReadScreen screenContent, 1920, 1, 1
        
        ' Check all possible prompts
        promptPos = 0
        isFound = False
        For j = 0 To UBound(targets)
            promptPos = InStr(1, screenContent, targets(j), vbTextCompare)
            If promptPos > 0 Then
                isFound = True
                Exit For
            End If
        Next
        
        ' If prompt is anywhere from Row 2 down to the bottom, we are HAPPY.
        ' Row 2 starts at Pos 81.
        If isFound And promptPos > 80 Then
            LogResult "INFO", "Prompt found and satisfied (Pos: " & promptPos & ")."
            Exit Sub
        End If
        
        ' RECOVERY: Only reached if prompt is missing or in Row 1 (Header)
        LogResult "INFO", "Self-Correction Required (Attempt " & i & ")..."
        bzhao.SendKey "E" ' Exit (Standard Recovery)
        bzhao.SendKey "<NumpadEnter>"
        
        bzhao.Pause 1500 ' Reduced for production
    Next
End Sub

Sub WaitForText(targetText)
    Dim elapsed, screenContent, targets, found, i, isMainPrompt, promptPos, isFoundAnywhere
    targets = Split(targetText, "|")
    elapsed = 0
    isMainPrompt = (InStr(1, targetText, MAIN_PROMPT, vbTextCompare) > 0)
    
    Do
        bzhao.Pause 500
        elapsed = elapsed + 500
        
        bzhao.ReadScreen screenContent, 1920, 1, 1
        
        found = False
        isFoundAnywhere = False
        For i = 0 To UBound(targets)
            promptPos = InStr(1, screenContent, targets(i), vbTextCompare)
            If promptPos > 0 Then
                isFoundAnywhere = True
                
                ' Check if it's in a valid prompt position (Row 4+: pos 241+)
                If isMainPrompt Then
                    If promptPos > 240 Then found = True
                Else
                    found = True ' For other technical prompts, any location is fine
                End If
                
                If found Then Exit For
            End If
        Next
        
        If found Then Exit Sub
        
        ' Recovery logic: ALWAYS send 'E' to exit if we are stuck.
        ' We previously tried sending Caret (^) but user reports 'E' is more reliable.
        If isMainPrompt And elapsed >= 5000 Then
            If elapsed Mod 5000 = 0 Then 
                LogResult "INFO", "Target text not found at bottom. Sending Exit (E)."
                bzhao.SendKey "E"
                bzhao.SendKey "<NumpadEnter>"
                bzhao.Pause 1000
            End If
        End If

        If elapsed >= 45000 Then 
            LogResult "ERROR", "Critical Timeout waiting for: " & targetText
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
