Option Explicit

' Test to reproduce the EXACT production issue with OPERATION CODE FOR LINE prompts
' Issue: Script sends "I" instead of accepting the default (I) value
' Based on production log from 1/2/2026 9:48:51 AM

' Include the main script to get access to functions
Dim scriptPath, fso
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\..\PostFinalCharges.vbs"

' Read and extract necessary functions
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
    
    WScript.Echo "PRODUCTION ISSUE Reproduction Test"
    WScript.Echo "===================================="
    WScript.Echo "Reproducing exact production failure from 1/2/2026 9:48:51 AM"
    WScript.Echo "Issue: Script sends 'I' instead of accepting default (I)"
    WScript.Echo ""
    
    Dim testNum, passCount, totalTests
    testNum = 0
    passCount = 0
    totalTests = 0
    
    ' Current pattern from production logs
    Dim pattern, screenContent
    pattern = "OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?"
    
    ' Test 1: EXACT production scenario that's failing
    testNum = testNum + 1
    totalTests = totalTests + 1
    screenContent = "OPERATION CODE FOR LINE A, L1 (I)?"
    Dim result1, expected1
    expected1 = True  ' Should accept default "I" and return True
    result1 = HasDefaultValueInPrompt(pattern, screenContent)
    WScript.Echo "Test " & testNum & ": EXACT production failure scenario"
    WScript.Echo "  Pattern: " & pattern
    WScript.Echo "  Screen Content: '" & screenContent & "'"
    WScript.Echo "  Expected: " & expected1 & " (should accept default 'I')"
    WScript.Echo "  Actual: " & result1
    If result1 = expected1 Then
        WScript.Echo "  RESULT: PASS"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL - This is the EXACT production bug!"
        WScript.Echo "  Impact: System sends 'I' explicitly instead of accepting default"
    End If
    WScript.Echo ""
    
    ' Test 2: Verify the regex pattern can match the content
    testNum = testNum + 1
    totalTests = totalTests + 1
    WScript.Echo "Test " & testNum & ": Verify pattern matching capability"
    Dim re, matches
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.IgnoreCase = True
    re.Global = False
    Set matches = re.Execute(screenContent)
    WScript.Echo "  Pattern: " & pattern
    WScript.Echo "  Content: " & screenContent
    WScript.Echo "  Matches Found: " & matches.Count
    If matches.Count > 0 Then
        WScript.Echo "  First Match: '" & matches(0).Value & "'"
        WScript.Echo "  RESULT: PASS - Pattern can match the content"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL - Pattern cannot match content at all!"
    End If
    WScript.Echo ""
    
    ' Test 3: Check what HasDefaultValueInPrompt internal regex finds
    testNum = testNum + 1
    totalTests = totalTests + 1
    WScript.Echo "Test " & testNum & ": Internal regex analysis"
    Dim internalPattern
    internalPattern = ".*\(([A-Za-z0-9]+)\)"  ' The pattern used inside HasDefaultValueInPrompt
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = internalPattern
    re.IgnoreCase = True
    re.Global = False
    Set matches = re.Execute(screenContent)
    WScript.Echo "  Internal Pattern: " & internalPattern
    WScript.Echo "  Content: " & screenContent
    WScript.Echo "  Matches Found: " & matches.Count
    If matches.Count > 0 Then
        WScript.Echo "  First Match: '" & matches(0).Value & "'"
        If matches(0).SubMatches.Count > 0 Then
            WScript.Echo "  Captured Default: '" & matches(0).SubMatches(0) & "'"
        End If
        WScript.Echo "  RESULT: PASS - Internal regex can find default value"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL - Internal regex cannot find default!"
    End If
    WScript.Echo ""
    
    ' Test 4: Step-by-step debugging of HasDefaultValueInPrompt logic
    testNum = testNum + 1
    totalTests = totalTests + 1
    WScript.Echo "Test " & testNum & ": Step-by-step HasDefaultValueInPrompt debugging"
    WScript.Echo "  This will manually trace the function's logic..."
    
    ' Recreate the HasDefaultValueInPrompt logic step by step
    On Error Resume Next
    Dim debugRe, debugMatches, debugMatch, debugParenContent
    Set debugRe = CreateObject("VBScript.RegExp")
    debugRe.Pattern = ".*\(([A-Za-z0-9]+)\)"
    debugRe.IgnoreCase = True
    debugRe.Global = False
    
    WScript.Echo "  Step 1: Creating regex object - " & IIf(Err.Number = 0, "SUCCESS", "ERROR: " & Err.Description)
    
    If Err.Number = 0 Then
        Set debugMatches = debugRe.Execute(screenContent)
        WScript.Echo "  Step 2: Executing regex - " & IIf(Err.Number = 0, "SUCCESS", "ERROR: " & Err.Description)
        
        If Err.Number = 0 And debugMatches.Count > 0 Then
            Set debugMatch = debugMatches(0)
            WScript.Echo "  Step 3: Found " & debugMatches.Count & " matches"
            
            If debugMatch.SubMatches.Count > 0 Then
                debugParenContent = debugMatch.SubMatches(0)
                WScript.Echo "  Step 4: Captured content: '" & debugParenContent & "'"
                
                If Len(debugParenContent) > 0 Then
                    WScript.Echo "  Step 5: Content length > 0: TRUE"
                    WScript.Echo "  RESULT: PASS - Should return True but function returns False!"
                    passCount = passCount + 1
                Else
                    WScript.Echo "  Step 5: Content length = 0: FALSE"
                    WScript.Echo "  RESULT: FAIL - Empty content when should have 'I'"
                End If
            Else
                WScript.Echo "  Step 4: No submatches found"
                WScript.Echo "  RESULT: FAIL - Regex matched but no capture group"
            End If
        Else
            WScript.Echo "  Step 3: No matches found or error occurred"
            WScript.Echo "  RESULT: FAIL - Regex should match production content"
        End If
    Else
        WScript.Echo "  RESULT: FAIL - Cannot create regex object"
    End If
    
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    WScript.Echo ""
    
    ' Summary
    WScript.Echo "========================================"
    WScript.Echo "Test Summary: " & passCount & "/" & totalTests & " tests passed"
    
    If passCount = totalTests Then
        WScript.Echo "UNEXPECTED: All tests passed - this suggests the function should work correctly!"
        WScript.Echo "There may be a different issue in production (environment, timing, etc.)"
    Else
        WScript.Echo "REPRODUCTION SUCCESS: " & (totalTests - passCount) & " tests failed!"
        WScript.Echo "This confirms the production bug can be reproduced in testing."
        WScript.Echo ""
        WScript.Echo "NEXT STEPS:"
        WScript.Echo "1. Fix the HasDefaultValueInPrompt function"
        WScript.Echo "2. Run this test again to verify the fix"
        WScript.Echo "3. Update the production script"
    End If
    
    If passCount < totalTests Then
        WScript.Quit 1  ' Signal test failure
    Else
        WScript.Quit 0  ' Signal test success
    End If
Else
    WScript.Echo "ERROR: Could not extract HasDefaultValueInPrompt function"
    WScript.Quit 1
End If

' Helper function for inline conditionals
Function IIf(condition, trueValue, falseValue)
    If condition Then
        IIf = trueValue
    Else
        IIf = falseValue
    End If
End Function