Option Explicit

' Test to reproduce the COMPLETE production flow for OPERATION CODE FOR LINE prompts
' This simulates the full ProcessPromptSequence logic to find where the bug occurs

' Include the main script to get access to functions
Dim scriptPath, fso
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\..\PostFinalCharges.vbs"

' Read the script and extract functions we need
Dim scriptFile, scriptContent
Set scriptFile = fso.OpenTextFile(scriptPath, 1)
scriptContent = scriptFile.ReadAll()
scriptFile.Close()

' Extract multiple functions we need for testing
Dim functionsToExtract, functionName, functionStart, functionEnd, allFunctions
functionsToExtract = Array("HasDefaultValueInPrompt", "CreateLineItemPromptDictionary")
allFunctions = ""

' Add stubs for missing functions first
allFunctions = allFunctions & "Sub LogTrace(msg, context)" & vbCrLf & "    WScript.Echo ""[TRACE] "" & msg" & vbCrLf & "End Sub" & vbCrLf & vbCrLf
allFunctions = allFunctions & "Sub LogDebug(msg, context)" & vbCrLf & "    WScript.Echo ""[DEBUG] "" & msg" & vbCrLf & "End Sub" & vbCrLf & vbCrLf
allFunctions = allFunctions & "Sub LogInfo(msg, context)" & vbCrLf & "    WScript.Echo ""[INFO] "" & msg" & vbCrLf & "End Sub" & vbCrLf & vbCrLf

' Add required classes
Dim classStart, classEnd, promptClassCode
classStart = InStr(scriptContent, "Class Prompt")
classEnd = InStr(classStart, scriptContent, "End Class")
If classStart > 0 And classEnd > 0 Then
    promptClassCode = Mid(scriptContent, classStart, classEnd - classStart + 9) ' 9 = length of "End Class"
    allFunctions = allFunctions & promptClassCode & vbCrLf & vbCrLf
End If

' Add AddPromptToDictEx function
Dim addPromptStart, addPromptEnd, addPromptCode
addPromptStart = InStr(scriptContent, "Sub AddPromptToDictEx(")
addPromptEnd = InStr(addPromptStart, scriptContent, "End Sub")
If addPromptStart > 0 And addPromptEnd > 0 Then
    addPromptCode = Mid(scriptContent, addPromptStart, addPromptEnd - addPromptStart + 7) ' 7 = length of "End Sub"
    allFunctions = allFunctions & addPromptCode & vbCrLf & vbCrLf
End If

' Add AddPromptToDict function
addPromptStart = InStr(scriptContent, "Sub AddPromptToDict(")
addPromptEnd = InStr(addPromptStart, scriptContent, "End Sub")
If addPromptStart > 0 And addPromptEnd > 0 Then
    addPromptCode = Mid(scriptContent, addPromptStart, addPromptEnd - addPromptStart + 7) ' 7 = length of "End Sub"
    allFunctions = allFunctions & addPromptCode & vbCrLf & vbCrLf
End If

' Extract each function
For Each functionName In functionsToExtract
    functionStart = InStr(scriptContent, "Function " & functionName & "(")
    functionEnd = InStr(functionStart, scriptContent, "End Function")
    
    If functionStart > 0 And functionEnd > 0 Then
        Dim functionCode
        functionCode = Mid(scriptContent, functionStart, functionEnd - functionStart + 12) ' 12 = length of "End Function"
        allFunctions = allFunctions & functionCode & vbCrLf & vbCrLf
    End If
Next

' Execute all the extracted code
ExecuteGlobal allFunctions

WScript.Echo "COMPLETE PRODUCTION FLOW Reproduction Test"
WScript.Echo "==========================================="
WScript.Echo "Simulating the full prompt processing logic"
WScript.Echo ""

Dim testNum, passCount, totalTests
testNum = 0
passCount = 0
totalTests = 0

' Test 1: Create the prompt dictionary and check the OPERATION CODE entry
testNum = testNum + 1
totalTests = totalTests + 1
WScript.Echo "Test " & testNum & ": Check prompt dictionary configuration"

Dim promptDict
Set promptDict = CreateLineItemPromptDictionary()

' Find the OPERATION CODE pattern in the dictionary
Dim foundPattern, foundPrompt, operationPattern
foundPattern = ""
operationPattern = "OPERATION CODE FOR LINE.*\([A-Za-z0-9]*\)\?"

For Each key In promptDict.Keys
    If InStr(key, "OPERATION CODE") > 0 Then
        foundPattern = key
        Set foundPrompt = promptDict.Item(key)
        Exit For
    End If
Next

WScript.Echo "  Looking for pattern containing 'OPERATION CODE'"
WScript.Echo "  Found pattern: '" & foundPattern & "'"
If foundPattern <> "" Then
    WScript.Echo "  Pattern matches expected: " & (foundPattern = operationPattern)
    WScript.Echo "  ResponseText: '" & foundPrompt.ResponseText & "'"
    WScript.Echo "  AcceptDefault: " & foundPrompt.AcceptDefault
    WScript.Echo "  KeyPress: '" & foundPrompt.KeyPress & "'"
    
    If foundPrompt.AcceptDefault = True Then
        WScript.Echo "  RESULT: PASS - Prompt configured for AcceptDefault=True"
        passCount = passCount + 1
    Else
        WScript.Echo "  RESULT: FAIL - AcceptDefault should be True!"
    End If
Else
    WScript.Echo "  RESULT: FAIL - No OPERATION CODE pattern found in dictionary!"
End If
WScript.Echo ""

' Test 2: Simulate the exact ProcessPromptSequence logic
testNum = testNum + 1
totalTests = totalTests + 1
WScript.Echo "Test " & testNum & ": Simulate ProcessPromptSequence logic"

Dim screenContent, pattern, shouldAcceptDefault, promptDetails
screenContent = "OPERATION CODE FOR LINE A, L1 (I)?"
pattern = operationPattern

WScript.Echo "  Screen Content: '" & screenContent & "'"
WScript.Echo "  Matched Pattern: '" & pattern & "'"

If foundPattern <> "" Then
    Set promptDetails = promptDict.Item(foundPattern)
    WScript.Echo "  Prompt AcceptDefault: " & promptDetails.AcceptDefault
    
    If promptDetails.AcceptDefault Then
        shouldAcceptDefault = HasDefaultValueInPrompt(pattern, screenContent)
        WScript.Echo "  HasDefaultValueInPrompt result: " & shouldAcceptDefault
        
        If shouldAcceptDefault Then
            WScript.Echo "  Expected behavior: Send ONLY KeyPress ('" & promptDetails.KeyPress & "')"
            WScript.Echo "  Should NOT send ResponseText: '" & promptDetails.ResponseText & "'"
            WScript.Echo "  RESULT: PASS - Logic should accept default"
            passCount = passCount + 1
        Else
            WScript.Echo "  Expected behavior: Send ResponseText ('" & promptDetails.ResponseText & "') + KeyPress"
            WScript.Echo "  RESULT: FAIL - HasDefaultValueInPrompt should return True!"
        End If
    Else
        WScript.Echo "  Expected behavior: Always send ResponseText (AcceptDefault=False)"
        WScript.Echo "  RESULT: FAIL - AcceptDefault should be True for this prompt!"
    End If
Else
    WScript.Echo "  RESULT: FAIL - Cannot test without prompt configuration"
End If
WScript.Echo ""

' Test 3: Check if there are multiple conflicting patterns
testNum = testNum + 1
totalTests = totalTests + 1
WScript.Echo "Test " & testNum & ": Check for conflicting OPERATION CODE patterns"

Dim conflictCount, conflictingPatterns
conflictCount = 0
conflictingPatterns = ""

For Each key In promptDict.Keys
    If InStr(UCase(key), "OPERATION CODE") > 0 Then
        conflictCount = conflictCount + 1
        conflictingPatterns = conflictingPatterns & "'" & key & "' "
    End If
Next

WScript.Echo "  Found " & conflictCount & " OPERATION CODE patterns:"
WScript.Echo "  Patterns: " & conflictingPatterns
If conflictCount = 1 Then
    WScript.Echo "  RESULT: PASS - Only one pattern, no conflicts"
    passCount = passCount + 1
Else
    WScript.Echo "  RESULT: FAIL - Multiple patterns could cause conflicts!"
    WScript.Echo "  Issue: Pattern matching might select wrong entry"
End If
WScript.Echo ""

' Test 4: Test the regex pattern matching priority
testNum = testNum + 1
totalTests = totalTests + 1
WScript.Echo "Test " & testNum & ": Test pattern matching and priority"

Dim longestMatch, longestLength, currentLength
longestMatch = ""
longestLength = 0

' Simulate the "longest match" logic from ProcessPromptSequence
For Each key In promptDict.Keys
    ' Check if this pattern matches our screen content
    Dim re, matches
    Set re = CreateObject("VBScript.RegExp")
    
    ' Determine if this is a regex pattern (simplified heuristic)
    If InStr(key, "(") > 0 Or InStr(key, ".*") > 0 Or InStr(key, "\\d") > 0 Then
        re.Pattern = key
        re.IgnoreCase = True
        re.Global = False
        
        On Error Resume Next
        Set matches = re.Execute(screenContent)
        If Err.Number = 0 And matches.Count > 0 Then
            currentLength = Len(matches(0).Value)
            If currentLength > longestLength Then
                longestLength = currentLength
                longestMatch = key
            End If
        End If
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    Else
        ' Plain text matching
        If InStr(1, screenContent, key, vbTextCompare) > 0 Then
            currentLength = Len(key)
            If currentLength > longestLength Then
                longestLength = currentLength
                longestMatch = key
            End If
        End If
    End If
Next

WScript.Echo "  Screen content: '" & screenContent & "'"
WScript.Echo "  Longest matching pattern: '" & longestMatch & "'"
WScript.Echo "  Match length: " & longestLength

If longestMatch = operationPattern Then
    WScript.Echo "  RESULT: PASS - Correct pattern selected"
    passCount = passCount + 1
Else
    WScript.Echo "  RESULT: FAIL - Wrong pattern selected!"
    WScript.Echo "  Expected: '" & operationPattern & "'"
    WScript.Echo "  This could explain why AcceptDefault logic isn't working!"
End If
WScript.Echo ""

' Summary
WScript.Echo "========================================"
WScript.Echo "Test Summary: " & passCount & "/" & totalTests & " tests passed"

If passCount = totalTests Then
    WScript.Echo "CONFIGURATION OK: All prompt processing logic is correct"
    WScript.Echo "The issue may be in:"
    WScript.Echo "- Timing/screen reading in production"
    WScript.Echo "- Different screen content than expected"
    WScript.Echo "- MockBzhao vs production BlueZone differences"
Else
    WScript.Echo "ISSUE FOUND: " & (totalTests - passCount) & " configuration problems detected!"
    WScript.Echo "These issues explain the production failure."
End If

If passCount < totalTests Then
    WScript.Quit 1  ' Signal test failure - we found the bug!
Else
    WScript.Quit 0  ' Signal test success - need to investigate further
End If