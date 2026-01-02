Option Explicit

' Test to reproduce the EXACT bug in ProcessPromptSequence
' Issue: InStr() is used to find regex patterns in screen text, which always fails

WScript.Echo "EXACT BUG Reproduction Test"
WScript.Echo "==========================="
WScript.Echo "Bug: InStr() used to find regex pattern in screen text"
WScript.Echo ""

Dim testCount, passCount
testCount = 0
passCount = 0

' Test 1: Reproduce the exact InStr() bug
testCount = testCount + 1
WScript.Echo "Test " & testCount & ": Reproduce InStr() regex search bug"

Dim bestMatchKey, screenContent, lineText, matchedLineContent
bestMatchKey = "OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?"
screenContent = "OPERATION CODE FOR LINE A, L1 (I)?"
lineText = screenContent

WScript.Echo "  Regex pattern (bestMatchKey): '" & bestMatchKey & "'"
WScript.Echo "  Screen content (lineText): '" & lineText & "'"
WScript.Echo "  InStr() search: InStr(1, lineText, bestMatchKey, vbTextCompare)"

Dim instrResult
instrResult = InStr(1, lineText, bestMatchKey, vbTextCompare)
WScript.Echo "  InStr() result: " & instrResult

If instrResult > 0 Then
    WScript.Echo "  RESULT: FAIL - InStr() should NOT find regex pattern in plain text!"
    matchedLineContent = lineText
Else
    WScript.Echo "  RESULT: PASS - This reproduces the production bug!"
    WScript.Echo "  Impact: matchedLineContent remains empty"
    matchedLineContent = ""
    passCount = passCount + 1
End If
WScript.Echo ""

' Test 2: Show what HasDefaultValueInPrompt gets called with
testCount = testCount + 1
WScript.Echo "Test " & testCount & ": Show impact on HasDefaultValueInPrompt"

' Include HasDefaultValueInPrompt function
Dim scriptPath, fso, scriptFile, scriptContent
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\..\PostFinalCharges.vbs"
Set scriptFile = fso.OpenTextFile(scriptPath, 1)
scriptContent = scriptFile.ReadAll()
scriptFile.Close()

Dim functionStart, functionEnd, functionCode
functionStart = InStr(scriptContent, "Function HasDefaultValueInPrompt(")
functionEnd = InStr(functionStart, scriptContent, "End Function")

If functionStart > 0 And functionEnd > 0 Then
    functionCode = Mid(scriptContent, functionStart, functionEnd - functionStart + 12)
    functionCode = "Sub LogTrace(msg, context)" & vbCrLf & "End Sub" & vbCrLf & vbCrLf & functionCode
    ExecuteGlobal functionCode
    
    Dim buggyResult, correctResult
    buggyResult = HasDefaultValueInPrompt(bestMatchKey, matchedLineContent)
    correctResult = HasDefaultValueInPrompt(bestMatchKey, screenContent)
    
    WScript.Echo "  Current production behavior:"
    WScript.Echo "    HasDefaultValueInPrompt(pattern, '" & matchedLineContent & "') = " & buggyResult
    WScript.Echo "  Correct behavior should be:"
    WScript.Echo "    HasDefaultValueInPrompt(pattern, '" & screenContent & "') = " & correctResult
    
    If buggyResult = False And correctResult = True Then
        WScript.Echo "  RESULT: PASS - This exactly reproduces the production bug!"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL - Bug reproduction incomplete"
    End If
End If
WScript.Echo ""

' Test 3: Show the correct solution
testCount = testCount + 1
WScript.Echo "Test " & testCount & ": Demonstrate the correct solution"

WScript.Echo "  Problem: Using InStr() to find regex patterns fails"
WScript.Echo "  Solution: Use regex matching to find the line containing the matched pattern"
WScript.Echo ""

' Show correct approach using regex
Dim re, matches
Set re = CreateObject("VBScript.RegExp")
re.Pattern = bestMatchKey
re.IgnoreCase = True
re.Global = False

Set matches = re.Execute(screenContent)
If matches.Count > 0 Then
    WScript.Echo "  Correct approach: regex.Execute(screenContent).Count = " & matches.Count
    WScript.Echo "  This would properly find the line containing the pattern"
    WScript.Echo "  Then HasDefaultValueInPrompt would get the correct content"
    WScript.Echo "  RESULT: PASS - Solution identified"
    passCount = passCount + 1
Else
    WScript.Echo "  RESULT: FAIL - Regex should match the screen content"
End If
WScript.Echo ""

' Summary
WScript.Echo "========================================"
WScript.Echo "Test Summary: " & passCount & "/" & testCount & " tests passed"
WScript.Echo ""
WScript.Echo "BUG IDENTIFIED:"
WScript.Echo "In ProcessPromptSequence lines ~423, the code uses:"
WScript.Echo "  If InStr(1, lineText, bestMatchKey, vbTextCompare) > 0 Then"
WScript.Echo ""
WScript.Echo "This fails because:"
WScript.Echo "- bestMatchKey contains regex patterns like '.*\([A-Za-z0-9]*\)\?'"
WScript.Echo "- lineText contains plain text like 'OPERATION CODE FOR LINE A, L1 (I)?'"
WScript.Echo "- InStr() does literal string search, not regex matching"
WScript.Echo ""
WScript.Echo "SOLUTION:"
WScript.Echo "Replace the InStr() logic with proper regex matching"
WScript.Echo "to find which line contains the pattern match"

If passCount = testCount Then
    WScript.Echo ""
    WScript.Echo "GOOD: Bug reproduction test shows the logic should work correctly"
    WScript.Echo "If this test passes but production still fails, check the actual implementation"
    WScript.Quit 0  ' Exit normally - this confirms understanding of the problem
Else
    WScript.Echo ""
    WScript.Echo "ISSUE: Some aspects of the bug could not be reproduced"
    WScript.Quit 1
End If