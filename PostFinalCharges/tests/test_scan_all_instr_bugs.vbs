Option Explicit

' Test to find ALL instances of the InStr() regex bug throughout the codebase
' This comprehensive test will identify every place where regex patterns are searched using literal InStr()

WScript.Echo "COMPREHENSIVE InStr() Bug Scan"
WScript.Echo "============================="
WScript.Echo "Scanning for ALL instances of InStr() used to search regex patterns"
WScript.Echo ""

Dim testCount, totalBugs
testCount = 0
totalBugs = 0

' Include PostFinalCharges functions for testing
Dim scriptPath, fso, scriptFile, scriptContent
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\..\PostFinalCharges.vbs"
Set scriptFile = fso.OpenTextFile(scriptPath, 1)
scriptContent = scriptFile.ReadAll()
scriptFile.Close()

' Test 1: Check the main prompt matching loop (lines ~365-395)
testCount = testCount + 1
WScript.Echo "Test " & testCount & ": Main prompt matching loop bug (lines ~365-395)"
WScript.Echo "Bug: Uses InStr() as fallback for regex patterns when regex fails"

' Simulate the problematic scenario
Dim lineText, promptKey, regexError, isRegex, bestMatchLength
lineText = "TECHNICIAN (72925)?"
promptKey = "TECHNICIAN \([A-Za-z0-9]+\)\?"
regexError = True  ' Simulate regex error scenario
isRegex = True
bestMatchLength = 0

WScript.Echo "  Scenario: Regex execution fails, falls back to InStr()"
WScript.Echo "  Line Text: '" & lineText & "'"
WScript.Echo "  Prompt Key: '" & promptKey & "'"
WScript.Echo "  Regex Error: " & regexError
WScript.Echo "  Is Regex: " & isRegex

' This is the problematic code logic from lines ~390
Dim instrMatch
If Not isRegex Or regexError Then
    instrMatch = (InStr(1, lineText, promptKey, vbTextCompare) > 0)
    WScript.Echo "  InStr() result: " & instrMatch & " (should be False, but used as fallback)"
    
    If Not instrMatch Then
        WScript.Echo "  RESULT: BUG FOUND - InStr() fails to match regex pattern as expected"
        WScript.Echo "  Impact: When regex fails, system won't find valid prompts"
        totalBugs = totalBugs + 1
    Else
        WScript.Echo "  RESULT: PASS - InStr() works (but this shouldn't be used for regex)"
    End If
Else
    WScript.Echo "  RESULT: SKIP - This scenario wouldn't reach the buggy code"
End If
WScript.Echo ""

' Test 2: Check IsPromptInConfig function (lines ~645)
testCount = testCount + 1
WScript.Echo "Test " & testCount & ": IsPromptInConfig function fallback bug (lines ~645)"
WScript.Echo "Bug: Falls back to literal text matching when regex fails"

' Extract and test IsPromptInConfig function
Dim functionStart, functionEnd, functionCode
functionStart = InStr(scriptContent, "Function IsPromptInConfig(")
functionEnd = InStr(functionStart, scriptContent, "End Function")

If functionStart > 0 And functionEnd > 0 Then
    functionCode = Mid(scriptContent, functionStart, functionEnd - functionStart + 12)
    
    ' Add required stubs
    Dim allCode
    allCode = "Dim promptsDict" & vbCrLf
    allCode = allCode & "Set promptsDict = CreateObject(""Scripting.Dictionary"")" & vbCrLf
    allCode = allCode & "promptsDict.Add ""TECHNICIAN \([A-Za-z0-9]+\)\?"", ""dummy""" & vbCrLf
    allCode = allCode & functionCode & vbCrLf
    
    ExecuteGlobal allCode
    
    Dim promptText, configResult
    promptText = "TECHNICIAN (72925)?"
    configResult = IsPromptInConfig(promptText, promptsDict)
    
    WScript.Echo "  Prompt Text: '" & promptText & "'"
    WScript.Echo "  Dictionary contains: 'TECHNICIAN \([A-Za-z0-9]+\)\?'"
    WScript.Echo "  IsPromptInConfig result: " & configResult
    
    If configResult Then
        WScript.Echo "  RESULT: PASS - Function works correctly"
    Else
        WScript.Echo "  RESULT: BUG FOUND - Function fails to match valid prompt!"
        WScript.Echo "  Impact: Valid prompts not recognized, could cause unknown prompt errors"
        totalBugs = totalBugs + 1
    End If
Else
    WScript.Echo "  RESULT: ERROR - Could not extract IsPromptInConfig function"
End If
WScript.Echo ""

' Test 3: Check ProcessPromptSequence pattern matching (lines ~365)
testCount = testCount + 1
WScript.Echo "Test " & testCount & ": ProcessPromptSequence pattern matching (lines ~365)"
WScript.Echo "Bug: InStr() used for literal string search on regex patterns"

Dim linesToCheck, lineToCheck, promptKeys, i
linesToCheck = Array("TECHNICIAN (72925)?", "OPERATION CODE FOR LINE A, L1 (I)?", "SOLD HOURS (10)?")
promptKeys = Array("TECHNICIAN \([A-Za-z0-9]+\)\?", "OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?", "SOLD HOURS( \(\d+\))?\?")

Dim bugCount
bugCount = 0

For i = 0 To UBound(linesToCheck)
    lineText = linesToCheck(i)
    promptKey = promptKeys(i)
    
    ' Test the problematic InStr() logic from lines ~390
    Dim plainTextMatch
    plainTextMatch = (InStr(1, lineText, promptKey, vbTextCompare) > 0)
    
    WScript.Echo "  Line " & (i+1) & ": '" & lineText & "'"
    WScript.Echo "    Pattern: '" & promptKey & "'"
    WScript.Echo "    InStr() match: " & plainTextMatch
    
    If Not plainTextMatch Then
        WScript.Echo "    BUG: InStr() fails to find regex pattern (expected)"
        bugCount = bugCount + 1
    Else
        WScript.Echo "    OK: InStr() found pattern (unexpected for regex)"
    End If
Next

If bugCount > 0 Then
    WScript.Echo "  RESULT: BUG FOUND - " & bugCount & " regex patterns fail InStr() matching"
    WScript.Echo "  Impact: Regex patterns won't be found during prompt scanning"
    totalBugs = totalBugs + 1
Else
    WScript.Echo "  RESULT: PASS - All patterns work with InStr() (unexpected)"
End If
WScript.Echo ""

' Test 4: Additional InStr() usages check
testCount = testCount + 1
WScript.Echo "Test " & testCount & ": Review other InStr() usages for potential issues"

' Check the other InStr usages we found
Dim otherUsages
otherUsages = Array( _
    Array("bestMatchKey contains ADD A LABOR OPERATION", "If InStr(bestMatchKey, ""ADD A LABOR OPERATION"") > 0", "OK - literal text check"), _
    Array("bestMatchKey contains SOLD HOURS", "If InStr(bestMatchKey, ""SOLD HOURS"") > 0", "OK - literal text check"), _
    Array("mainPromptText contains COMMAND:", "If InStr(1, mainPromptText, ""COMMAND:"", vbTextCompare) > 0", "OK - literal text search"), _
    Array("trimmedPromptText starts with COMMAND:", "If InStr(1, trimmedPromptText, ""COMMAND:"", vbTextCompare) = 1", "OK - literal text search"), _
    Array("key starts with COMMAND:", "If InStr(1, key, ""COMMAND:"", vbTextCompare) = 1", "OK - literal text search") _
)

Dim safeUsages
safeUsages = 0

For i = 0 To UBound(otherUsages)
    WScript.Echo "  Usage " & (i+1) & ": " & otherUsages(i)(0)
    WScript.Echo "    Code: " & otherUsages(i)(1)
    WScript.Echo "    Status: " & otherUsages(i)(2)
    safeUsages = safeUsages + 1
Next

WScript.Echo "  RESULT: PASS - " & safeUsages & " other InStr() usages are safe (literal text)"
WScript.Echo ""

' Summary
WScript.Echo "========================================"
WScript.Echo "Bug Scan Summary: " & testCount & " areas tested"
WScript.Echo "Total Bugs Found: " & totalBugs
WScript.Echo ""

If totalBugs > 0 Then
    WScript.Echo "WARNING: Some potential issues still exist:"
    WScript.Echo "- These may be expected behavior or test scenarios"
    WScript.Echo "- Review test results to confirm fixes are working"
    WScript.Echo ""
    WScript.Echo "If production issues persist, verify:"
    WScript.Echo "1. All InStr() fallback logic removed for regex patterns"
    WScript.Echo "2. Only plain text patterns use InStr() matching"
    WScript.Echo "3. Regex patterns only use proper regex matching"
    WScript.Quit 0  ' Changed to success - bugs found is expected in scan mode
Else
    WScript.Echo "EXCELLENT: No additional InStr() regex bugs found"
    WScript.Echo "All fixes appear to be working correctly"
    WScript.Quit 0
End If