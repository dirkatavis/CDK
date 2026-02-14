'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestOperationCodeDefault
' **DATE CREATED:** 2026-01-02
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Tests the specific OPERATION CODE FOR LINE default value detection issue
' reported by the user where "OPERATION CODE FOR LINE A, L1 (I)?" should 
' accept the default "I" instead of sending a new "I".
'-----------------------------------------------------------------------------------

Option Explicit

' Include the HasDefaultValueInPrompt function
' This is a standalone test that doesn't require the full script loading
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

Sub Main()
    WScript.Echo "OPERATION CODE FOR LINE Default Value Test"
    WScript.Echo "========================================"
    WScript.Echo ""
    
    ' Test the specific case reported by the user
    Dim testCase, pattern, content, expected, actual, result
    
    ' Test Case 1: OPERATION CODE FOR LINE with default (I)
    WScript.Echo "Test 1: OPERATION CODE FOR LINE with default value (I)"
    pattern = "OPERATION CODE FOR LINE.*\([A-Za-z0-9]+\)\?"
    content = "OPERATION CODE FOR LINE A, L1 (I)?"
    expected = True
    actual = HasDefaultValueInPrompt(pattern, content)
    
    WScript.Echo "  Pattern: " & pattern
    WScript.Echo "  Content: " & content
    WScript.Echo "  Expected: " & expected
    WScript.Echo "  Actual: " & actual
    
    If actual = expected Then
        result = "PASS"
    Else
        result = "FAIL"
    End If
    WScript.Echo "  RESULT: " & result
    WScript.Echo ""
    
    ' Test Case 2: OPERATION CODE FOR LINE without default
    WScript.Echo "Test 2: OPERATION CODE FOR LINE without default value"
    pattern = "OPERATION CODE FOR LINE"
    content = "OPERATION CODE FOR LINE"
    expected = False
    actual = HasDefaultValueInPrompt(pattern, content)
    
    WScript.Echo "  Pattern: " & pattern
    WScript.Echo "  Content: " & content
    WScript.Echo "  Expected: " & expected
    WScript.Echo "  Actual: " & actual
    
    If actual = expected Then
        result = "PASS"
    Else
        result = "FAIL"
    End If
    WScript.Echo "  RESULT: " & result
    WScript.Echo ""
    
    ' Test Case 3: More complex OPERATION CODE format
    WScript.Echo "Test 3: OPERATION CODE FOR LINE with line B, L2 (X)"
    pattern = "OPERATION CODE FOR LINE.*\([A-Za-z0-9]+\)\?"
    content = "OPERATION CODE FOR LINE B, L2 (X)?"
    expected = True
    actual = HasDefaultValueInPrompt(pattern, content)
    
    WScript.Echo "  Pattern: " & pattern
    WScript.Echo "  Content: " & content
    WScript.Echo "  Expected: " & expected
    WScript.Echo "  Actual: " & actual
    
    If actual = expected Then
        result = "PASS"
    Else
        result = "FAIL"
    End If
    WScript.Echo "  RESULT: " & result
    WScript.Echo ""
    
    WScript.Echo "Expected Behavior:"
    WScript.Echo "- When prompt shows 'OPERATION CODE FOR LINE A, L1 (I)?' -> Should send ENTER only (accept default I)"
    WScript.Echo "- When prompt shows 'OPERATION CODE FOR LINE' -> Should send 'I' + ENTER (provide value)"
    WScript.Echo ""
    WScript.Echo "Test completed."
End Sub

Main()