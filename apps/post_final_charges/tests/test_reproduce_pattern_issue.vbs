Option Explicit

' Test to reproduce pattern matching issues with OPERATION CODE FOR LINE prompts
' This test validates the current implementation and helps identify any bugs

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
    
    WScript.Echo "OPERATION CODE Pattern Reproduction Test"
    WScript.Echo "========================================"
    WScript.Echo "Testing current pattern: OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?"
    WScript.Echo ""
    
    Dim testNum, passCount, totalTests
    testNum = 0
    passCount = 0
    totalTests = 0
    
    ' Current pattern from the code
    Dim pattern
    pattern = "OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?"
    
    ' Test 1: Normal case with default value "I"
    testNum = testNum + 1
    totalTests = totalTests + 1
    Dim result1, expected1
    expected1 = True
    result1 = HasDefaultValueInPrompt(pattern, "OPERATION CODE FOR LINE A, L1 (I)?")
    WScript.Echo "Test " & testNum & ": Standard case with default 'I'"
    WScript.Echo "  Content: OPERATION CODE FOR LINE A, L1 (I)?"
    WScript.Echo "  Expected: " & expected1 & " (should accept default)"
    WScript.Echo "  Actual: " & result1
    If result1 = expected1 Then
        WScript.Echo "  RESULT: PASS"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL - This is the bug we need to fix!"
    End If
    WScript.Echo ""
    
    ' Test 2: Empty parentheses case
    testNum = testNum + 1
    totalTests = totalTests + 1
    Dim result2, expected2
    expected2 = False
    result2 = HasDefaultValueInPrompt(pattern, "OPERATION CODE FOR LINE A, L1 ()?")
    WScript.Echo "Test " & testNum & ": Empty parentheses case"
    WScript.Echo "  Content: OPERATION CODE FOR LINE A, L1 ()?"
    WScript.Echo "  Expected: " & expected2 & " (should send 'I')"
    WScript.Echo "  Actual: " & result2
    If result2 = expected2 Then
        WScript.Echo "  RESULT: PASS"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL - This is the bug we need to fix!"
    End If
    WScript.Echo ""
    
    ' Test 3: Different default value
    testNum = testNum + 1
    totalTests = totalTests + 1
    Dim result3, expected3
    expected3 = True
    result3 = HasDefaultValueInPrompt(pattern, "OPERATION CODE FOR LINE B, L2 (X)?")
    WScript.Echo "Test " & testNum & ": Different default value 'X'"
    WScript.Echo "  Content: OPERATION CODE FOR LINE B, L2 (X)?"
    WScript.Echo "  Expected: " & expected3 & " (should accept default)"
    WScript.Echo "  Actual: " & result3
    If result3 = expected3 Then
        WScript.Echo "  RESULT: PASS"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL - This is the bug we need to fix!"
    End If
    WScript.Echo ""
    
    ' Test 4: Edge case - single space in parentheses
    testNum = testNum + 1
    totalTests = totalTests + 1
    Dim result4, expected4
    expected4 = False
    result4 = HasDefaultValueInPrompt(pattern, "OPERATION CODE FOR LINE A, L1 ( )?")
    WScript.Echo "Test " & testNum & ": Single space in parentheses"
    WScript.Echo "  Content: OPERATION CODE FOR LINE A, L1 ( )?"
    WScript.Echo "  Expected: " & expected4 & " (should send 'I', space is not alphanumeric)"
    WScript.Echo "  Actual: " & result4
    If result4 = expected4 Then
        WScript.Echo "  RESULT: PASS"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL - This is the bug we need to fix!"
    End If
    WScript.Echo ""
    
    ' Test 5: Edge case - special characters in parentheses
    testNum = testNum + 1
    totalTests = totalTests + 1
    Dim result5, expected5
    expected5 = False
    result5 = HasDefaultValueInPrompt(pattern, "OPERATION CODE FOR LINE A, L1 (-)?")
    WScript.Echo "Test " & testNum & ": Special character in parentheses"
    WScript.Echo "  Content: OPERATION CODE FOR LINE A, L1 (-)?"
    WScript.Echo "  Expected: " & expected5 & " (should send 'I', dash is not alphanumeric)"
    WScript.Echo "  Actual: " & result5
    If result5 = expected5 Then
        WScript.Echo "  RESULT: PASS"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL - This is the bug we need to fix!"
    End If
    WScript.Echo ""
    
    ' Test 6: Alphanumeric default
    testNum = testNum + 1
    totalTests = totalTests + 1
    Dim result6, expected6
    expected6 = True
    result6 = HasDefaultValueInPrompt(pattern, "OPERATION CODE FOR LINE C, L3 (I2)?")
    WScript.Echo "Test " & testNum & ": Alphanumeric default 'I2'"
    WScript.Echo "  Content: OPERATION CODE FOR LINE C, L3 (I2)?"
    WScript.Echo "  Expected: " & expected6 & " (should accept default)"
    WScript.Echo "  Actual: " & result6
    If result6 = expected6 Then
        WScript.Echo "  RESULT: PASS"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL - This is the bug we need to fix!"
    End If
    WScript.Echo ""
    
    ' Test 7: Pattern Matching Check - Does the pattern actually match?
    testNum = testNum + 1
    totalTests = totalTests + 1
    WScript.Echo "Test " & testNum & ": Pattern Matching Verification"
    Dim re, matches, testContent
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.IgnoreCase = True
    re.Global = False
    
    testContent = "OPERATION CODE FOR LINE A, L1 (I)?"
    Set matches = re.Execute(testContent)
    WScript.Echo "  Testing if pattern matches: " & testContent
    WScript.Echo "  Pattern: " & pattern
    If matches.Count > 0 Then
        WScript.Echo "  RESULT: PASS - Pattern matches successfully"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL - Pattern does NOT match! This is a critical bug!"
    End If
    WScript.Echo ""
    
    ' Test 8: Pattern Matching Check - Empty parentheses
    testNum = testNum + 1
    totalTests = totalTests + 1
    WScript.Echo "Test " & testNum & ": Pattern Matching Empty Parentheses"
    testContent = "OPERATION CODE FOR LINE A, L1 ()?"
    Set matches = re.Execute(testContent)
    WScript.Echo "  Testing if pattern matches: " & testContent
    WScript.Echo "  Pattern: " & pattern
    If matches.Count > 0 Then
        WScript.Echo "  RESULT: PASS - Pattern matches empty parentheses"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL - Pattern should match empty parentheses too!"
    End If
    WScript.Echo ""
    
    ' Summary
    WScript.Echo "========================================"
    WScript.Echo "Test Summary: " & passCount & "/" & totalTests & " tests passed"
    If passCount = totalTests Then
        WScript.Echo "SUCCESS: All tests passed! Current implementation is working correctly."
        WScript.Quit 0
    Else
        WScript.Echo "FAILURE: " & (totalTests - passCount) & " tests failed!"
        WScript.Echo "These failures indicate bugs that need to be fixed."
        WScript.Quit 1
    End If
Else
    WScript.Echo "ERROR: Could not extract HasDefaultValueInPrompt function"
    WScript.Quit 1
End If