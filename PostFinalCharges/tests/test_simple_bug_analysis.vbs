Option Explicit

' Simplified test to focus on the core issue with OPERATION CODE pattern
' This will help us identify exactly where the production bug occurs

' Include the main script functions we need
Dim scriptPath, fso
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\..\PostFinalCharges.vbs"

' Read and extract just what we need
Dim scriptFile, scriptContent
Set scriptFile = fso.OpenTextFile(scriptPath, 1)
scriptContent = scriptFile.ReadAll()
scriptFile.Close()

' Extract HasDefaultValueInPrompt function
Dim functionStart, functionEnd, functionCode
functionStart = InStr(scriptContent, "Function HasDefaultValueInPrompt(")
functionEnd = InStr(functionStart, scriptContent, "End Function")

If functionStart > 0 And functionEnd > 0 Then
    functionCode = Mid(scriptContent, functionStart, functionEnd - functionStart + 12)
    ' Add stub for LogTrace
    functionCode = "Sub LogTrace(msg, context)" & vbCrLf & "End Sub" & vbCrLf & vbCrLf & functionCode
    ExecuteGlobal functionCode
    
    WScript.Echo "FOCUSED Production Bug Analysis"
    WScript.Echo "==============================="
    WScript.Echo ""
    
    Dim testCount, passCount
    testCount = 0
    passCount = 0
    
    ' Test the exact production scenario with detailed analysis
    testCount = testCount + 1
    WScript.Echo "Test " & testCount & ": Production scenario analysis"
    WScript.Echo "Screen content: 'OPERATION CODE FOR LINE A, L1 (I)?'"
    WScript.Echo "Expected behavior: Accept default 'I' (press Enter only)"
    WScript.Echo "Actual behavior: Send 'I' + press Enter"
    WScript.Echo ""
    
    ' Check what HasDefaultValueInPrompt returns
    Dim screenContent, pattern, result
    screenContent = "OPERATION CODE FOR LINE A, L1 (I)?"
    pattern = "OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?"
    result = HasDefaultValueInPrompt(pattern, screenContent)
    
    WScript.Echo "HasDefaultValueInPrompt('" & pattern & "', '" & screenContent & "') = " & result
    
    If result = True Then
        WScript.Echo "FUNCTION RESULT: TRUE (should accept default)"
        WScript.Echo "PRODUCTION BEHAVIOR: Sends ResponseText 'I'"
        WScript.Echo "CONCLUSION: Function is correct, but production ignores the result!"
        WScript.Echo ""
        WScript.Echo "LIKELY CAUSES:"
        WScript.Echo "1. Pattern matching is selecting wrong dictionary entry"
        WScript.Echo "2. AcceptDefault flag is False for the matched pattern"
        WScript.Echo "3. Logic error in ProcessPromptSequence"
        passCount = passCount + 1
    Else
        WScript.Echo "FUNCTION RESULT: FALSE (should send ResponseText)"
        WScript.Echo "PRODUCTION BEHAVIOR: Sends ResponseText 'I'"
        WScript.Echo "CONCLUSION: Function has a bug - not detecting default value!"
    End If
    WScript.Echo ""
    
    ' Let's check what's in the current PostFinalCharges.vbs file for the OPERATION CODE pattern
    testCount = testCount + 1
    WScript.Echo "Test " & testCount & ": Check current OPERATION CODE configuration"
    
    ' Look for the AddPromptToDictEx call for OPERATION CODE
    Dim operationCodeStart, operationCodeLine, operationCodeEnd
    operationCodeStart = InStr(scriptContent, "OPERATION CODE FOR LINE")
    
    If operationCodeStart > 0 Then
        ' Find the line containing this
        Dim contentBeforeMatch, lineBreaks, lineNumber
        contentBeforeMatch = Left(scriptContent, operationCodeStart)
        lineBreaks = Len(contentBeforeMatch) - Len(Replace(contentBeforeMatch, vbCrLf, ""))
        lineNumber = (lineBreaks \ 2) + 1  ' Rough line number
        
        ' Extract the line
        operationCodeEnd = InStr(operationCodeStart, scriptContent, vbCrLf)
        If operationCodeEnd > 0 Then
            operationCodeLine = Mid(scriptContent, operationCodeStart, operationCodeEnd - operationCodeStart)
        Else
            operationCodeLine = Mid(scriptContent, operationCodeStart, 200) ' Get some context
        End If
        
        WScript.Echo "Found OPERATION CODE configuration around line " & lineNumber & ":"
        WScript.Echo "  " & Trim(operationCodeLine)
        WScript.Echo ""
        
        ' Check if AcceptDefault is True
        If InStr(operationCodeLine, "True") > 0 Then
            WScript.Echo "AcceptDefault appears to be: TRUE"
            passCount = passCount + 1
        Else
            WScript.Echo "AcceptDefault appears to be: FALSE or missing!"
            WScript.Echo "THIS COULD BE THE BUG!"
        End If
    Else
        WScript.Echo "Could not find OPERATION CODE configuration in script!"
    End If
    WScript.Echo ""
    
    ' Test: What if we had different patterns?
    testCount = testCount + 1
    WScript.Echo "Test " & testCount & ": Test pattern variations"
    
    Dim patterns, testPattern, i
    patterns = Array( _
        "OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?", _
        "OPERATION CODE FOR LINE.*\([A-Za-z0-9]+\)\?", _
        "OPERATION CODE FOR LINE A, L1 \(I\)\?" _
    )
    
    For i = 0 To UBound(patterns)
        testPattern = patterns(i)
        result = HasDefaultValueInPrompt(testPattern, screenContent)
        WScript.Echo "  Pattern " & (i+1) & ": " & testPattern
        WScript.Echo "    Result: " & result
    Next
    passCount = passCount + 1
    WScript.Echo ""
    
    WScript.Echo "========================================"
    WScript.Echo "Analysis Summary: " & passCount & "/" & testCount & " areas checked"
    WScript.Echo ""
    WScript.Echo "NEXT STEPS TO FIX THE BUG:"
    WScript.Echo "1. Check the exact prompt dictionary entry for OPERATION CODE"
    WScript.Echo "2. Verify AcceptDefault=True is set correctly"
    WScript.Echo "3. Check for multiple conflicting patterns"
    WScript.Echo "4. Add more detailed logging to ProcessPromptSequence"
    
Else
    WScript.Echo "ERROR: Could not extract HasDefaultValueInPrompt function"
    WScript.Quit 1
End If