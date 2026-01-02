'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestEmptyDefaultValues
' **DATE CREATED:** 2026-01-02
' **FUNCTIONALITY:**
' Tests empty/null default values in parentheses to verify they're handled correctly
'-----------------------------------------------------------------------------------

Option Explicit

' Test the regex pattern against empty/null values
Function HasDefaultValueInPrompt_Test(promptPattern, screenContent)
    HasDefaultValueInPrompt_Test = False
    
    On Error Resume Next
    Dim re, matches, match, parenContent
    Set re = CreateObject("VBScript.RegExp")
    
    ' Current pattern: any text followed by parentheses containing non-empty alphanumeric content
    re.Pattern = ".*\(([A-Za-z0-9]+)\)"
    re.IgnoreCase = True
    re.Global = False
    
    WScript.Echo "Testing content: '" & screenContent & "'"
    WScript.Echo "Using pattern: '" & re.Pattern & "'"
    
    If Err.Number = 0 Then
        Set matches = re.Execute(screenContent)
        WScript.Echo "Match count: " & matches.Count
        
        If matches.Count > 0 Then
            Set match = matches(0)
            WScript.Echo "Match value: '" & match.Value & "'"
            WScript.Echo "SubMatches count: " & match.SubMatches.Count
            
            If match.SubMatches.Count > 0 Then
                parenContent = Trim(match.SubMatches(0))
                WScript.Echo "Parentheses content: '" & parenContent & "'"
                
                If Len(parenContent) > 0 And parenContent <> "?" And parenContent <> "" Then
                    HasDefaultValueInPrompt_Test = True
                    WScript.Echo "Result: VALID DEFAULT FOUND"
                Else
                    WScript.Echo "Result: INVALID/EMPTY DEFAULT"
                End If
            Else
                WScript.Echo "Result: NO SUBMATCHES"
            End If
        Else
            WScript.Echo "Result: NO MATCHES"
        End If
    Else
        WScript.Echo "Regex error: " & Err.Description
    End If
    
    WScript.Echo "Final result: " & HasDefaultValueInPrompt_Test
    WScript.Echo ""
    
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Sub TestEmptyValues()
    WScript.Echo "Testing Empty/Null Default Value Scenarios"
    WScript.Echo "=========================================="
    WScript.Echo ""
    
    ' Test various empty/null scenarios
    Dim result
    
    WScript.Echo "Test 1: Normal default value"
    result = HasDefaultValueInPrompt_Test("", "OPERATION CODE FOR LINE A, L1 (I)?")
    
    WScript.Echo "Test 2: Empty parentheses"
    result = HasDefaultValueInPrompt_Test("", "OPERATION CODE FOR LINE A, L1 ()?")
    
    WScript.Echo "Test 3: Parentheses with just question mark"
    result = HasDefaultValueInPrompt_Test("", "OPERATION CODE FOR LINE A, L1 (?)?")
    
    WScript.Echo "Test 4: Parentheses with spaces"
    result = HasDefaultValueInPrompt_Test("", "OPERATION CODE FOR LINE A, L1 (   )?")
    
    WScript.Echo "Test 5: Parentheses with null/special chars"
    result = HasDefaultValueInPrompt_Test("", "OPERATION CODE FOR LINE A, L1 (null)?")
    
    WScript.Echo "Test 6: No parentheses at all"
    result = HasDefaultValueInPrompt_Test("", "OPERATION CODE FOR LINE")
    
End Sub

TestEmptyValues()