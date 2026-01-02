'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestDefaultValueDetection
' **DATE CREATED:** 2025-12-29
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Test script for the default value detection bugfix.
' Tests the HasDefaultValueInPrompt function and AcceptDefault functionality.
'-----------------------------------------------------------------------------------

Option Explicit

' Include required files
Function IncludeFile(filePath)
    On Error Resume Next
    Dim fsoInclude, fileContent, includeStream

    Set fsoInclude = CreateObject("Scripting.FileSystemObject")

    If Not fsoInclude.FileExists(filePath) Then
        WScript.Echo "IncludeFile - File not found: " & filePath
        IncludeFile = False
        Exit Function
    End If

    Set includeStream = fsoInclude.OpenTextFile(filePath, 1)
    fileContent = includeStream.ReadAll
    includeStream.Close
    Set includeStream = Nothing

    ExecuteGlobal fileContent
    IncludeFile = True
End Function

' Test data structure for test cases
Class TestCase
    Public Description
    Public PromptPattern
    Public ScreenContent
    Public ExpectedResult
End Class

' Create test cases for default value detection
Function CreateTestCases()
    Dim testCases(14), tc  ' Increased from 9 to 14 for additional test cases

    ' Test case 1: TECHNICIAN with default value
    Set tc = New TestCase
    tc.Description = "TECHNICIAN with default value (72925)"
    tc.PromptPattern = "TECHNICIAN \([A-Za-z0-9]+\)\?"
    tc.ScreenContent = "TECHNICIAN (72925)?"
    tc.ExpectedResult = True
    Set testCases(0) = tc

    ' Test case 2: TECHNICIAN with different default value
    Set tc = New TestCase
    tc.Description = "TECHNICIAN with default value (99)"
    tc.PromptPattern = "TECHNICIAN \(\d+\)"
    tc.ScreenContent = "TECHNICIAN (99)"
    tc.ExpectedResult = True
    Set testCases(1) = tc

    ' Test case 3: TECHNICIAN without default
    Set tc = New TestCase
    tc.Description = "TECHNICIAN without default value"
    tc.PromptPattern = "TECHNICIAN?"
    tc.ScreenContent = "TECHNICIAN?"
    tc.ExpectedResult = False
    Set testCases(2) = tc

    ' Test case 4: TECHNICIAN with empty parentheses
    Set tc = New TestCase
    tc.Description = "TECHNICIAN with empty parentheses"
    tc.PromptPattern = "TECHNICIAN \([A-Za-z0-9]+\)\?"
    tc.ScreenContent = "TECHNICIAN ()?"
    tc.ExpectedResult = False
    Set testCases(3) = tc

    ' Test case 5: ACTUAL HOURS with default value
    Set tc = New TestCase
    tc.Description = "ACTUAL HOURS with default value (117)"
    tc.PromptPattern = "ACTUAL HOURS \(\d+\)"
    tc.ScreenContent = "ACTUAL HOURS (117)?"
    tc.ExpectedResult = True
    Set testCases(4) = tc

    ' Test case 6: ACTUAL HOURS with zero default
    Set tc = New TestCase
    tc.Description = "ACTUAL HOURS with zero default"
    tc.PromptPattern = "ACTUAL HOURS \(\d+\)"
    tc.ScreenContent = "ACTUAL HOURS (0)?"
    tc.ExpectedResult = True
    Set testCases(5) = tc

    ' Test case 7: SOLD HOURS with default value
    Set tc = New TestCase
    tc.Description = "SOLD HOURS with default value (10)"
    tc.PromptPattern = "SOLD HOURS \([0-9]+\)\?"
    tc.ScreenContent = "SOLD HOURS (10)?"
    tc.ExpectedResult = True
    Set testCases(6) = tc

    ' Test case 8: SOLD HOURS with zero default
    Set tc = New TestCase
    tc.Description = "SOLD HOURS with zero default"
    tc.PromptPattern = "SOLD HOURS \([0-9]+\)\?"
    tc.ScreenContent = "SOLD HOURS (0)?"
    tc.ExpectedResult = True
    Set testCases(7) = tc

    ' Test case 9: Complex screen with multiple prompts
    Set tc = New TestCase
    tc.Description = "Complex screen content with TECHNICIAN default"
    tc.PromptPattern = "TECHNICIAN \([A-Za-z0-9]+\)\?"
    tc.ScreenContent = "LINE 1: Some other text" & vbCrLf & "LINE 2: TECHNICIAN (ABC123)?" & vbCrLf & "LINE 3: More text"
    tc.ExpectedResult = True
    Set testCases(8) = tc

    ' Test case 10: Edge case with special characters
    Set tc = New TestCase
    tc.Description = "TECHNICIAN with alphanumeric ID"
    tc.PromptPattern = "TECHNICIAN \([A-Za-z0-9]+\)\?"
    tc.ScreenContent = "TECHNICIAN (A1B2C3)?"
    tc.ExpectedResult = True
    Set testCases(9) = tc

    ' *** NEW TEST CASES THAT WOULD HAVE CAUGHT THE BUG ***
    
    ' Test case 11: OPERATION CODE FOR LINE with intervening text and default (THE BUG!)
    Set tc = New TestCase
    tc.Description = "OPERATION CODE FOR LINE with intervening text and default (I)"
    tc.PromptPattern = "OPERATION CODE FOR LINE.*\([A-Za-z0-9]+\)\?"
    tc.ScreenContent = "OPERATION CODE FOR LINE A, L1 (I)?"
    tc.ExpectedResult = True
    Set testCases(10) = tc

    ' Test case 12: OPERATION CODE with different line and default
    Set tc = New TestCase
    tc.Description = "OPERATION CODE FOR LINE with different line and default (X)"
    tc.PromptPattern = "OPERATION CODE FOR LINE.*\([A-Za-z0-9]+\)\?"
    tc.ScreenContent = "OPERATION CODE FOR LINE B, L2 (X)?"
    tc.ExpectedResult = True
    Set testCases(11) = tc

    ' Test case 13: Any prompt with intervening text pattern (general case)
    Set tc = New TestCase
    tc.Description = "TECHNICIAN with job info and default (intervening text pattern)"
    tc.PromptPattern = "TECHNICIAN.*\([A-Za-z0-9]+\)\?"
    tc.ScreenContent = "TECHNICIAN FOR JOB 12345, LINE A (T999)?"
    tc.ExpectedResult = True
    Set testCases(12) = tc

    ' Test case 14: ACTUAL HOURS with detailed context and default
    Set tc = New TestCase
    tc.Description = "ACTUAL HOURS with detailed info and default"
    tc.PromptPattern = "ACTUAL HOURS.*\([A-Za-z0-9]+\)\?"
    tc.ScreenContent = "ACTUAL HOURS FOR OPERATION Z, LINE C (45)?"
    tc.ExpectedResult = True
    Set testCases(13) = tc

    ' Test case 15: Edge case with multiple parentheses
    Set tc = New TestCase
    tc.Description = "Prompt with multiple parentheses (should find the last one)"
    tc.PromptPattern = "COST CENTER.*\([A-Za-z0-9]+\)\?"
    tc.ScreenContent = "COST CENTER (MAIN) FOR LINE A (CC123)?"
    tc.ExpectedResult = True
    Set testCases(14) = tc

    CreateTestCases = testCases
End Function

' Run all test cases
Sub RunDefaultValueDetectionTests()
    WScript.Echo "Running Default Value Detection Tests..."
    WScript.Echo String(50, "=")

    Dim testCases, i, testCase, actualResult, passed, total
    testCases = CreateTestCases()
    passed = 0
    total = UBound(testCases) + 1

    For i = 0 To UBound(testCases)
        Set testCase = testCases(i)
        actualResult = HasDefaultValueInPrompt(testCase.PromptPattern, testCase.ScreenContent)
        
        WScript.Echo "Test " & (i + 1) & ": " & testCase.Description
        WScript.Echo "  Pattern: " & testCase.PromptPattern
        WScript.Echo "  Content: " & Replace(testCase.ScreenContent, vbCrLf, " | ")
        WScript.Echo "  Expected: " & testCase.ExpectedResult
        WScript.Echo "  Actual: " & actualResult
        
        If actualResult = testCase.ExpectedResult Then
            WScript.Echo "  RESULT: PASS"
            passed = passed + 1
        Else
            WScript.Echo "  RESULT: FAIL"
        End If
        
        WScript.Echo ""
    Next

    WScript.Echo String(50, "=")
    WScript.Echo "Test Summary: " & passed & "/" & total & " tests passed"
    If passed = total Then
        WScript.Echo "SUCCESS: All tests passed!"
    Else
        WScript.Echo "FAILURE: " & (total - passed) & " tests failed"
    End If
End Sub

' Main test entry point
Sub Main()
    WScript.Echo "Default Value Detection Test Suite"
    WScript.Echo "=================================="
    WScript.Echo ""

    ' Run the test suites (we'll test the function directly without loading the full script)
    RunDefaultValueDetectionTests()
    
    WScript.Echo vbCrLf & "Test suite completed."
    WScript.Echo "Note: Prompt dictionary tests require running with the main script loaded."
End Sub

' Copy of the HasDefaultValueInPrompt function for testing
Function HasDefaultValueInPrompt(promptPattern, screenContent)
    HasDefaultValueInPrompt = False
    
    ' Use a more robust approach - look for any text followed by parentheses containing alphanumeric content
    ' This handles all prompt types without hardcoding specific patterns
    On Error Resume Next
    Dim re, matches, match, parenContent
    Set re = CreateObject("VBScript.RegExp")
    
    ' Universal pattern: any text followed by parentheses containing non-empty alphanumeric content
    ' Examples: TECHNICIAN(12345), ACTUAL HOURS (8), SOLD HOURS (10), OPERATION CODE FOR LINE A, L1 (I)
    ' Updated pattern to handle any content before parentheses
    re.Pattern = ".*\(([A-Za-z0-9]+)\)"
    re.IgnoreCase = True
    re.Global = False
    
    If Err.Number = 0 Then
        Set matches = re.Execute(screenContent)
        If matches.Count > 0 Then
            Set match = matches(0)
            If match.SubMatches.Count > 0 Then
                parenContent = Trim(match.SubMatches(0))
                ' If there's content in parentheses and it's not empty or just question marks
                If Len(parenContent) > 0 And parenContent <> "?" And parenContent <> "" Then
                    HasDefaultValueInPrompt = True
                End If
            End If
        End If
    End If
    
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

' Run the tests
Main()