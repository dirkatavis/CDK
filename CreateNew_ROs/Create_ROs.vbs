'====================================================================
' Script: Efficient VehicleDataAutomation.vbs
' Focus: Maximum speed with smart prompt detection - FULLY STANDARDIZED
'====================================================================

Option Explicit

' Configuration Constants
Const CSV_FILE_PATH = "C:\Temp\Code\Scripts\VBScript\CDK\CreateNew_ROs\create_RO.csv"
Const POLL_INTERVAL = 500   ' Check every 500ms (optimized from 1000ms for faster response)
Const POST_ENTRY_WAIT = 100  ' Minimal wait after entry (optimized from 200ms)
Const PRE_KEY_WAIT = 100     ' Pause before sending special keys (optimized from 150ms)
Const POST_KEY_WAIT = 200    ' Pause after sending special keys (optimized from 350ms)
Const PROMPT_TIMEOUT_MS = 5000 ' Default prompt timeout

Dim fso, ts, strLine, arrValues, i, MVA, Mileage
Dim Bzhao
On Error Resume Next
Set Bzhao = CreateObject("BZWhll.WhllObj")
If Err.Number <> 0 Then
    ' Will report at connect time; clear for now
    Err.Clear
End If
On Error GoTo 0

Dim SCRIPT_FOLDER, SLOW_MARKER_PATH, LOG_FILE_PATH
' Hardcode paths to eliminate variables
SCRIPT_FOLDER = "C:\Temp\bluezone_backup\Scripts\archive"
SLOW_MARKER_PATH = "C:\Temp\Code\Scripts\VBScript\CDK\CreateNew_ROs\Create_RO.debug"
LOG_FILE_PATH = "C:\Temp\Code\Scripts\VBScript\CDK\CreateNew_ROs\VehicleData.log"

' Test logging immediately to verify log file creation
LOG "Script started - Log file path: " & LOG_FILE_PATH

Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(CSV_FILE_PATH) Then
    LOG "CSV file found: " & CSV_FILE_PATH
    bzhao.Connect ""
    LOG "Connected to BlueZone"
    Set ts = fso.OpenTextFile(CSV_FILE_PATH, 1)
    ts.ReadLine   ' Skip header row
    LOG "Processing CSV data..."

    Do While Not ts.AtEndOfStream
        strLine = ts.ReadLine
        arrValues = Split(strLine, ",")
        
        If UBound(arrValues) >= 1 Then
            MVA = Trim(arrValues(0))
            Mileage = Trim(arrValues(1))
            Call Main(MVA, Mileage)
        End If
    Loop

    ts.Close
    Set ts = Nothing
End If

Set fso = Nothing
bzhao.Disconnect

'--------------------------------------------------------------------
' Subroutine: Main - Fully Standardized with WaitForPrompt
'--------------------------------------------------------------------
Sub Main(mva, mileage)
    
    '==== INPUT POINT 1: BEFORE ENTERING MVA ====
    ' NEED TO IDENTIFY: What prompt appears when CDK is ready for Vehicle ID?
    ' CURRENT: Using "Vehid....." - NEEDS VERIFICATION
    Call WaitForPrompt("Vehid.....", mva, true, PROMPT_TIMEOUT_MS)
    ' bzhao.Pause 1000


    ' Skip if no matching vehicle - check but don't enter anything
    If IsTextPresent("No matching") Then Exit Sub



    Call WaitForPrompt("ENTER SEQUENCE NUMBER", "1", true, 1000)


    '==== INPUT POINT 2: BEFORE ENTERING COMMAND SELECTION ====
    ' NEED TO IDENTIFY: What menu/prompt shows before selecting command?
    ' CURRENT: Looking for "Command?" - NEEDS VERIFICATION
    Call WaitForPrompt("Command?", "<NumpadEnter>", False, PROMPT_TIMEOUT_MS)
    


    '==== INPUT POINT 3: BEFORE CONFIRMING DISPLAY ====
    ' NEED TO IDENTIFY: What text appears before "Display them now" response?
    ' CURRENT: Using "Display them now" - NEEDS VERIFICATION
    Call WaitForPrompt("Display them now", "<NumpadEnter>", False, 1000)

    '==== INPUT POINT 4: BEFORE ENTERING MILEAGE ====
    ' NEED TO IDENTIFY: What prompt shows when mileage field is ready?
    ' CURRENT: Using "Miles In...:" - NEEDS VERIFICATION
    Call WaitForPrompt("Miles In", mileage, true, PROMPT_TIMEOUT_MS)
    

    '==== INPUT POINT 5: BEFORE ENTERING MILEAGE VALIDATION ====
    ' NEED TO IDENTIFY: What prompt asks for Y/N on mileage validation?
    ' CURRENT: Using "greater than" - NEEDS VERIFICATION
    Call WaitForPrompt("greater than", "Y", true, 1000)

    '==== INPUT POINT 6: BEFORE ENTERING TAG ====
    ' NEED TO IDENTIFY: What field label appears for tag entry?
    ' CURRENT: Using "Tag......" - NEEDS VERIFICATION
    Call WaitForPrompt("Tag......", mva, true, PROMPT_TIMEOUT_MS)
    ' bzhao.Pause 1000

    '==== INPUT POINT 7: BEFORE ENTERING VENDOR ====
    ' NEED TO IDENTIFY: What prompt shows for vendor field?
    ' CURRENT: Using "PMVEND" - NEEDS VERIFICATION
    Call WaitForPrompt("Quick Codes", "PMVEND", True, PROMPT_TIMEOUT_MS)
    ' bzhao.Pause 1000

    '==== INPUT POINT 8: BEFORE F3 KEY ====
    ' NEED TO IDENTIFY: What screen/text indicates ready for F3?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    Call WaitForPrompt("Quick Code Description", "<F3>", False, PROMPT_TIMEOUT_MS)
    ' bzhao.Pause 1000

    '==== INPUT POINT 9: BEFORE F8 KEY ====
    ' NEED to IDENTIFY: What screen/text indicates ready for F8?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    Call WaitForPrompt("Quick Codes", "<F8>", False, PROMPT_TIMEOUT_MS)
    ' bzhao.Pause 1000
    
    '==== INPUT POINT 10: BEFORE ENTERING "99" ====
    ' NEED TO IDENTIFY: What prompt shows for "99" entry?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    Call WaitForPrompt("Tech", "99", False, PROMPT_TIMEOUT_MS)
    ' bzhao.Pause 1000
    
    '==== INPUT POINT 11: BEFORE SECOND F3 ====
    ' NEED TO IDENTIFY: What indicates ready for second F3?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    Call WaitForPrompt("Tech", "<F3>", False, PROMPT_TIMEOUT_MS)
    ' bzhao.Pause 1000
    
    '==== INPUT POINT 12: BEFORE THIRD F3 ====
    ' NEED TO IDENTIFY: What indicates ready for third F3?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    Call WaitForPrompt("Quick Codes", "<F3>", False, PROMPT_TIMEOUT_MS)
    ' bzhao.Pause 1000    
    
    '==== INPUT POINT 13: BEFORE FIRST ENTER KEY ====
    ' NEED TO IDENTIFY: What text shows system is ready for Enter?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    Call WaitForPrompt("Choose an option", "<NumpadEnter>", False, PROMPT_TIMEOUT_MS)
    ' bzhao.Pause 1000
    
    '==== INPUT POINT 14: BEFORE SECOND ENTER KEY ====
    ' NEED TO IDENTIFY: What prompt appears before second Enter?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    Call WaitForPrompt("MILEAGE OUT", "<NumpadEnter>", False, 10000)
    ' bzhao.Pause 1000
    
    '==== INPUT POINT 15: BEFORE THIRD ENTER KEY ====
    ' NEED TO IDENTIFY: What prompt appears before third Enter?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION

    Call WaitForPrompt("MILEAGE IN", "<NumpadEnter>", False, 10000)
    ' bzhao.Pause 1000
    
    '==== INPUT POINT 16: BEFORE ENTERING FINAL "N" ====
    ' NEED TO IDENTIFY: What question/prompt is asking for N response?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    Call WaitForPrompt("O.K. TO CLOSE RO", "N", true, PROMPT_TIMEOUT_MS)
    ' bzhao.Pause 1000

    ' Scrape and log
    Dim roNumber
    roNumber = GetRepairOrderEnhanced()
    Call LogEntryWithRO(mva, roNumber)
    


    '==== INPUT POINT 17: BEFORE FINAL F3 ====
    ' NEED TO IDENTIFY: What indicates ready for final F3?
    ' CURRENT: No verification - NEEDS PROMPT DETECTION
    Call WaitForPrompt("Created repair order", "<F3>", False, PROMPT_TIMEOUT_MS)
End Sub



'--------------------------------------------------------------------
' Subroutine: WaitForPrompt - Requires 4 parameters: promptText, valueToEnter, sendEnter, timeoutMs
'--------------------------------------------------------------------
Sub WaitForPrompt(promptText, valueToEnter, sendEnter, timeoutMs)
    
    LOG "WaitForPrompt called - Looking for: [" & promptText & "] Value: [" & valueToEnter & "] SendEnter: " & sendEnter & " Timeout: " & timeoutMs & "ms"
    Dim startTime, currentTime, elapsedMs, promptFound
    
    startTime = Timer
    promptFound = False
    
    Do
        ' Check for the prompt text first
        If IsTextPresent(promptText) Then
            LOG "Detected prompt: " & promptText
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
            LOG "Timeout waiting for prompt: " & promptText
            Exit Do
        End If
        LOG elapsedMs & "ms elapsed waiting for prompt: "
    Loop
    
    ' Only send input if prompt was actually found
    If promptFound Then
        ' Apply slow mode delay if enabled
        If IsSlowModeEnabled() Then Call WaitMs(1000)
        
        ' Check if the value is a special key command (removed redundant bzhao.Pause)
        If InStr(1, valueToEnter, "<") > 0 And InStr(1, valueToEnter, ">") > 0 Then
            LOG "Sending key command: " & valueToEnter
            Call FastKey(valueToEnter)
        Else
            Call FastText(valueToEnter)
        End If
        
        If sendEnter Then
            Call FastKey("<NumpadEnter>")
        End If
        
        Call WaitMs(POST_ENTRY_WAIT)
    Else
        LOG "Prompt not found - skipping input"
    End If
End Sub


'--------------------------------------------------------------------
' Subroutine: FastText - Minimal delay text entry
'--------------------------------------------------------------------
Sub FastText(text)
    LOG "Sending text: " & text
    bzhao.SendKey text
    If IsSlowModeEnabled() Then
        Call WaitMs(1000)
    Else
        Call WaitMs(100)
    End If
End Sub

'--------------------------------------------------------------------
' Subroutine: FastKey - Minimal delay key press
'--------------------------------------------------------------------
Sub FastKey(key)
    LOG "Sending key command: " & key
    ' Pause briefly before sending a special key to avoid injecting escape sequences into active fields
    If IsSlowModeEnabled() Then
        Call WaitMs(1000)
    Else
        Call WaitMs(PRE_KEY_WAIT)
    End If
    bzhao.SendKey key
    ' Allow the host some time to process the special key and transition screens
    If IsSlowModeEnabled() Then
        Call WaitMs(1000)
    Else
        Call WaitMs(POST_KEY_WAIT)
    End If
End Sub

'--------------------------------------------------------------------
' Function: IsSlowModeEnabled
' Returns True if the debug marker file (Create_RO.debug) is present next to the script
'--------------------------------------------------------------------
Function IsSlowModeEnabled()
    On Error Resume Next
    Dim f
    Set f = CreateObject("Scripting.FileSystemObject")
    IsSlowModeEnabled = f.FileExists(SLOW_MARKER_PATH)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

'--------------------------------------------------------------------
' Function: GetRepairOrderEnhanced - Quick repair order extraction
'--------------------------------------------------------------------
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

'--------------------------------------------------------------------
' Function: IsTextPresent - Fast screen reading
'--------------------------------------------------------------------
Function IsTextPresent(textToFind)
    Dim screenContentBuffer 
    Dim screenLength
    screenLength = 24 * 80 
    bzhao.ReadScreen screenContentBuffer, screenLength, 1, 1
    IsTextPresent = (InStr(1, screenContentBuffer, textToFind, vbTextCompare) > 0)
    LOG "Screen content checked for: " & textToFind
    LOG "Text presence result: " & IsTextPresent
End Function

'--------------------------------------------------------------------
' Subroutine: LogEntryWithRO - Simple logging
'--------------------------------------------------------------------
Sub LogEntryWithRO(mva, roNumber)
    If Trim(mva) = "" Then Exit Sub

    Dim logFSO, logFile
    Set logFSO = CreateObject("Scripting.FileSystemObject")
    Set logFile = logFSO.OpenTextFile(LOG_FILE_PATH, 8, True)

    If roNumber = "" Then
        logFile.WriteLine Now & " - MVA: " & mva
    Else
        logFile.WriteLine Now & " - MVA: " & mva & " - RO: " & roNumber
    End If

    logFile.Close
    Set logFile = Nothing
    Set logFSO = Nothing
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

'--------------------------------------------------------------------
' Subroutine: LOG - lightweight logger used by this archived script
'--------------------------------------------------------------------
Sub LOG(msg)
    On Error Resume Next
    Dim lfs, lfile, errorNum, errorDesc
    Set lfs = CreateObject("Scripting.FileSystemObject")
    
    ' Try to create log entry
    Set lfile = lfs.OpenTextFile(LOG_FILE_PATH, 8, True)
    errorNum = Err.Number
    errorDesc = Err.Description
    
    If errorNum = 0 Then
        lfile.WriteLine Now & " - " & CStr(msg)
        lfile.Close
    Else
        ' If main log fails, try creating a fallback log with error info
        Dim fallbackPath
        fallbackPath = "C:\Temp\LOG_ERROR.txt"
        Set lfile = lfs.OpenTextFile(fallbackPath, 8, True)
        If Err.Number = 0 Then
            lfile.WriteLine Now & " - LOG ERROR: " & errorNum & " - " & errorDesc
            lfile.WriteLine Now & " - Failed LOG_FILE_PATH: " & LOG_FILE_PATH
            lfile.WriteLine Now & " - Original message: " & CStr(msg)
            lfile.Close
        End If
    End If
    
    Set lfile = Nothing
    Set lfs = Nothing
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub