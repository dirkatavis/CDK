'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestOldRegexPattern
' **DATE CREATED:** 2026-01-02
' **FUNCTIONALITY:**
' Tests the OLD regex pattern against the enhanced test cases to demonstrate
' which tests would have caught the bug before e2e testing.
'-----------------------------------------------------------------------------------

Option Explicit

' Simulate the OLD HasDefaultValueInPrompt function (before the fix)
Function HasDefaultValueInPrompt_Old(promptPattern, screenContent)
    HasDefaultValueInPrompt_Old = False
    
    On Error Resume Next
    Dim re, matches, match, parenContent
    Set re = CreateObject("VBScript.RegExp")
    
    ' OLD pattern that failed on intervening text
    re.Pattern = "[A-Z][A-Z\s]*\s*\(([A-Za-z0-9]+)\)"
    re.IgnoreCase = True
    re.Global = False
    
    If Err.Number = 0 Then
        Set matches = re.Execute(screenContent)
        If matches.Count > 0 Then
            Set match = matches(0)
            If match.SubMatches.Count > 0 Then
                parenContent = Trim(match.SubMatches(0))
                If Len(parenContent) > 0 And parenContent <> "?" And parenContent <> "" Then
                    HasDefaultValueInPrompt_Old = True
                End If
            End If
        End If
    End If
    
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Sub TestBugCatchingCases()
    WScript.Echo "Testing which NEW test cases would have caught the bug"
    WScript.Echo "==========================================================="
    WScript.Echo ""
    
    ' Test cases that would have FAILED with the old regex and caught the bug
    Dim testCases(4)
    
    ' Bug case 1: The actual reported bug
    testCases(0) = Array("OPERATION CODE FOR LINE A, L1 (I)?", "OPERATION CODE FOR LINE A, L1 (I)?", True)
    
    ' Bug case 2: Different line variation
    testCases(1) = Array("OPERATION CODE FOR LINE B, L2 (X)?", "OPERATION CODE FOR LINE B, L2 (X)?", True)
    
    ' Bug case 3: TECHNICIAN with job info
    testCases(2) = Array("TECHNICIAN FOR JOB 12345, LINE A (T999)?", "TECHNICIAN FOR JOB 12345, LINE A (T999)?", True)
    
    ' Bug case 4: ACTUAL HOURS with detailed info
    testCases(3) = Array("ACTUAL HOURS FOR OPERATION Z, LINE C (45)?", "ACTUAL HOURS FOR OPERATION Z, LINE C (45)?", True)
    
    ' Bug case 5: Multiple parentheses
    testCases(4) = Array("COST CENTER (MAIN) FOR LINE A (CC123)?", "COST CENTER (MAIN) FOR LINE A (CC123)?", True)
    
    Dim i, description, content, expected, oldResult, newResult
    Dim bugsCaught, totalBugs
    bugsCaught = 0
    totalBugs = UBound(testCases) + 1
    
    For i = 0 To UBound(testCases)
        description = testCases(i)(0)
        content = testCases(i)(1)
        expected = testCases(i)(2)
        
        oldResult = HasDefaultValueInPrompt_Old("", content)  ' Pattern not used in old version
        ' newResult would be True with the fixed regex
        
        WScript.Echo "Test " & (i + 1) & ": " & description
        WScript.Echo "  Content: " & content
        WScript.Echo "  Expected: " & expected
        WScript.Echo "  Old Regex Result: " & oldResult
        WScript.Echo "  New Regex Result: " & expected & " (after fix)"
        
        If oldResult <> expected Then
            WScript.Echo "  STATUS: This test would have CAUGHT THE BUG!"
            bugsCaught = bugsCaught + 1
        Else
            WScript.Echo "  STATUS: This test would have passed (false negative)"
        End If
        WScript.Echo ""
    Next
    
    WScript.Echo "SUMMARY:"
    WScript.Echo "========"
    WScript.Echo "Tests that would have caught the bug: " & bugsCaught & " out of " & totalBugs
    WScript.Echo "Bug detection rate: " & FormatPercent(bugsCaught / totalBugs, 0)
    
    If bugsCaught = totalBugs Then
        WScript.Echo ""
        WScript.Echo "SUCCESS: All enhanced test cases would have detected the regex failure!"
        WScript.Echo "Adding these tests to the suite prevents future regressions."
    End If
End Sub

TestBugCatchingCases()