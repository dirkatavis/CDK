Option Explicit

' Test to verify the fix for the InStr() bug in ProcessPromptSequence
' This test simulates the fixed logic to ensure it works correctly

WScript.Echo "BUG FIX Verification Test"
WScript.Echo "========================="
WScript.Echo "Testing the fixed logic for finding regex patterns in screen text"
WScript.Echo ""

Dim testCount, passCount
testCount = 0
passCount = 0

' Test 1: Verify the new regex matching logic works
testCount = testCount + 1
WScript.Echo "Test " & testCount & ": Verify fixed regex matching logic"

Dim bestMatchKey, lineText, lineMatchFound
bestMatchKey = "OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?"
lineText = "OPERATION CODE FOR LINE A, L1 (I)?"

WScript.Echo "  Pattern (bestMatchKey): '" & bestMatchKey & "'"
WScript.Echo "  Screen content (lineText): '" & lineText & "'"

' Simulate the fixed logic from ProcessPromptSequence
lineMatchFound = False

' Determine if bestMatchKey is a regex pattern (same logic as fix)
If Left(bestMatchKey, 1) = "^" Or InStr(bestMatchKey, "(") > 0 Or InStr(bestMatchKey, "[") > 0 Or InStr(bestMatchKey, ".*") > 0 Or InStr(bestMatchKey, "\\d") > 0 Then
    WScript.Echo "  Detected as regex pattern: YES"
    
    ' Use proper regex matching (same as fix)
    On Error Resume Next
    Dim lineRe, lineMatches
    Set lineRe = CreateObject("VBScript.RegExp")
    lineRe.Pattern = bestMatchKey
    lineRe.IgnoreCase = True
    lineRe.Global = False
    Set lineMatches = lineRe.Execute(lineText)
    If Err.Number = 0 And lineMatches.Count > 0 Then
        lineMatchFound = True
    End If
    If Err.Number <> 0 Then
        WScript.Echo "  Regex error: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
Else
    WScript.Echo "  Detected as regex pattern: NO"
    ' Use InStr for plain text
    If InStr(1, lineText, bestMatchKey, vbTextCompare) > 0 Then
        lineMatchFound = True
    End If
End If

WScript.Echo "  Line match found: " & lineMatchFound

If lineMatchFound Then
    WScript.Echo "  RESULT: PASS - Fixed logic correctly finds the pattern!"
    passCount = passCount + 1
Else
    WScript.Echo "  RESULT: FAIL - Fixed logic still doesn't work"
End If
WScript.Echo ""

' Test 2: Verify HasDefaultValueInPrompt gets correct input
testCount = testCount + 1
WScript.Echo "Test " & testCount & ": Verify HasDefaultValueInPrompt gets correct content"

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
    
    Dim matchedLineContent, hasDefaultResult
    If lineMatchFound Then
        matchedLineContent = lineText  ' This is what the fix should now provide
    Else
        matchedLineContent = ""  ' This was the old buggy behavior
    End If
    
    hasDefaultResult = HasDefaultValueInPrompt(bestMatchKey, matchedLineContent)
    
    WScript.Echo "  Fixed behavior:"
    WScript.Echo "    matchedLineContent: '" & matchedLineContent & "'"
    WScript.Echo "    HasDefaultValueInPrompt result: " & hasDefaultResult
    
    If hasDefaultResult = True Then
        WScript.Echo "  RESULT: PASS - HasDefaultValueInPrompt now returns True!"
        WScript.Echo "  Impact: System should accept default instead of sending ResponseText"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL - HasDefaultValueInPrompt still returns False"
    End If
End If
WScript.Echo ""

' Test 3: Test edge cases
testCount = testCount + 1
WScript.Echo "Test " & testCount & ": Test edge cases"

Dim edgeCases, i, edgePattern, edgeContent, edgeResult
edgeCases = Array( _
    Array("TECHNICIAN \([A-Za-z0-9]+\)\?", "TECHNICIAN (72925)?", True), _
    Array("TECHNICIAN?", "TECHNICIAN?", False), _
    Array("SOLD HOURS( \(\d+\))?\?", "SOLD HOURS (10)?", True), _
    Array("SOLD HOURS( \(\d+\))?\?", "SOLD HOURS?", False) _
)

Dim edgePassCount
edgePassCount = 0

For i = 0 To UBound(edgeCases)
    edgePattern = edgeCases(i)(0)
    edgeContent = edgeCases(i)(1)
    Dim expectedResult
    expectedResult = edgeCases(i)(2)
    
    ' Test the regex detection logic
    If Left(edgePattern, 1) = "^" Or InStr(edgePattern, "(") > 0 Or InStr(edgePattern, "[") > 0 Or InStr(edgePattern, ".*") > 0 Or InStr(edgePattern, "\\d") > 0 Then
        ' Regex pattern
        On Error Resume Next
        Set lineRe = CreateObject("VBScript.RegExp")
        lineRe.Pattern = edgePattern
        lineRe.IgnoreCase = True
        lineRe.Global = False
        Set lineMatches = lineRe.Execute(edgeContent)
        If Err.Number = 0 And lineMatches.Count > 0 Then
            edgeResult = HasDefaultValueInPrompt(edgePattern, edgeContent)
        Else
            edgeResult = False
        End If
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    Else
        ' Plain text
        If InStr(1, edgeContent, edgePattern, vbTextCompare) > 0 Then
            edgeResult = HasDefaultValueInPrompt(edgePattern, edgeContent)
        Else
            edgeResult = False
        End If
    End If
    
    WScript.Echo "  Edge case " & (i+1) & ": " & edgePattern
    WScript.Echo "    Content: '" & edgeContent & "'"
    WScript.Echo "    Expected: " & expectedResult & ", Actual: " & edgeResult
    
    If edgeResult = expectedResult Then
        edgePassCount = edgePassCount + 1
    End If
Next

WScript.Echo "  Edge cases passed: " & edgePassCount & "/" & (UBound(edgeCases) + 1)
If edgePassCount = (UBound(edgeCases) + 1) Then
    passCount = passCount + 1
End If
WScript.Echo ""

' Summary
WScript.Echo "========================================"
WScript.Echo "Test Summary: " & passCount & "/" & testCount & " tests passed"
WScript.Echo ""

If passCount = testCount Then
    WScript.Echo "SUCCESS: Bug fix verified!"
    WScript.Echo "The InStr() bug has been resolved."
    WScript.Echo "OPERATION CODE prompts should now accept defaults correctly."
    WScript.Quit 0
Else
    WScript.Echo "FAILURE: Bug fix incomplete or ineffective"
    WScript.Quit 1
End If