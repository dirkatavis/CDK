Option Explicit

' Test the single OPERATION CODE FOR LINE pattern with both default and empty scenarios
' This verifies that the new unified pattern works correctly

' Include the main script to get access to HasDefaultValueInPrompt function
Dim scriptPath, fso
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\..\PostFinalCharges.vbs"

' Read the script file and extract just the HasDefaultValueInPrompt function
Dim scriptFile, scriptContent, functionStart, functionEnd
Set scriptFile = fso.OpenTextFile(scriptPath, 1)
scriptContent = scriptFile.ReadAll()
scriptFile.Close()

' Find and extract the HasDefaultValueInPrompt function
functionStart = InStr(scriptContent, "Function HasDefaultValueInPrompt(")
functionEnd = InStr(functionStart, scriptContent, "End Function")

If functionStart > 0 And functionEnd > 0 Then
    Dim functionCode
    functionCode = Mid(scriptContent, functionStart, functionEnd - functionStart + 12) ' 12 = length of "End Function"
    
    ' Add the LogTrace stub function
    functionCode = "Sub LogTrace(msg, context)" & vbCrLf & "    ' Stub for testing" & vbCrLf & "End Sub" & vbCrLf & vbCrLf & functionCode
    
    ' Execute the function code
    ExecuteGlobal functionCode
    
    WScript.Echo "OPERATION CODE Single Pattern Test"
    WScript.Echo "=================================="
    WScript.Echo ""
    
    Dim testCount, passCount
    testCount = 0
    passCount = 0
    
    ' Test Case 1: OPERATION CODE with default value "I"
    testCount = testCount + 1
    Dim result1, expected1
    expected1 = True
    result1 = HasDefaultValueInPrompt("OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?", "OPERATION CODE FOR LINE A, L1 (I)?")
    WScript.Echo "Test " & testCount & ": OPERATION CODE with default value (I)"
    WScript.Echo "  Pattern: OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?"
    WScript.Echo "  Content: OPERATION CODE FOR LINE A, L1 (I)?"
    WScript.Echo "  Expected: " & expected1
    WScript.Echo "  Actual: " & result1
    If result1 = expected1 Then
        WScript.Echo "  RESULT: PASS"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL"
    End If
    WScript.Echo ""
    
    ' Test Case 2: OPERATION CODE with empty parentheses
    testCount = testCount + 1
    Dim result2, expected2
    expected2 = False
    result2 = HasDefaultValueInPrompt("OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?", "OPERATION CODE FOR LINE A, L1 ()?")
    WScript.Echo "Test " & testCount & ": OPERATION CODE with empty parentheses"
    WScript.Echo "  Pattern: OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?"
    WScript.Echo "  Content: OPERATION CODE FOR LINE A, L1 ()?"
    WScript.Echo "  Expected: " & expected2
    WScript.Echo "  Actual: " & result2
    If result2 = expected2 Then
        WScript.Echo "  RESULT: PASS"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL"
    End If
    WScript.Echo ""
    
    ' Test Case 3: OPERATION CODE with different default value "X"
    testCount = testCount + 1
    Dim result3, expected3
    expected3 = True
    result3 = HasDefaultValueInPrompt("OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?", "OPERATION CODE FOR LINE B, L2 (X)?")
    WScript.Echo "Test " & testCount & ": OPERATION CODE with default value (X)"
    WScript.Echo "  Pattern: OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?"
    WScript.Echo "  Content: OPERATION CODE FOR LINE B, L2 (X)?"
    WScript.Echo "  Expected: " & expected3
    WScript.Echo "  Actual: " & result3
    If result3 = expected3 Then
        WScript.Echo "  RESULT: PASS"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL"
    End If
    WScript.Echo ""
    
    ' Test Case 4: OPERATION CODE with alphanumeric default value "I2"
    testCount = testCount + 1
    Dim result4, expected4
    expected4 = True
    result4 = HasDefaultValueInPrompt("OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?", "OPERATION CODE FOR LINE C, L3 (I2)?")
    WScript.Echo "Test " & testCount & ": OPERATION CODE with alphanumeric default (I2)"
    WScript.Echo "  Pattern: OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?"
    WScript.Echo "  Content: OPERATION CODE FOR LINE C, L3 (I2)?"
    WScript.Echo "  Expected: " & expected4
    WScript.Echo "  Actual: " & result4
    If result4 = expected4 Then
        WScript.Echo "  RESULT: PASS"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL"
    End If
    WScript.Echo ""
    
    WScript.Echo "========================================"
    WScript.Echo "Test Summary: " & passCount & "/" & testCount & " tests passed"
    If passCount = testCount Then
        WScript.Echo "SUCCESS: All tests passed!"
        WScript.Quit 0
    Else
        WScript.Echo "FAILURE: Some tests failed!"
        WScript.Quit 1
    End If
Else
    WScript.Echo "ERROR: Could not extract HasDefaultValueInPrompt function from main script"
    WScript.Quit 1
End If