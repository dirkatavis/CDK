'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestBlacklistTerms
' **DATE CREATED:** 2026-03-12
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Comprehensive test suite for the RO blacklist filtering feature.
' Tests the CheckBlacklistedTerms function logic in isolation using mock screen data.
' Verifies:
' - Empty blacklist behavior (no filtering)
' - Single and multiple blacklist terms
' - Case-insensitive matching
' - Real-world RO screen scenarios
'-----------------------------------------------------------------------------------

Option Explicit

' Simplified CheckBlacklistedTerms for testing (copied from PostFinalCharges.vbs)
Function CheckBlacklistedTerms(blacklistArray, screenContent)
    Dim screenLength, i, matchedTerm
    
    ' Empty blacklist = no matches
    If UBound(blacklistArray) < 0 Then
        CheckBlacklistedTerms = ""
        Exit Function
    End If
    
    ' Convert screen to uppercase for case-insensitive matching
    screenContent = UCase(screenContent)
    
    ' Check each blacklist term
    For i = LBound(blacklistArray) To UBound(blacklistArray)
        matchedTerm = Trim(UCase(blacklistArray(i)))
        If Len(matchedTerm) > 0 And InStr(screenContent, matchedTerm) > 0 Then
            CheckBlacklistedTerms = matchedTerm
            Exit Function
        End If
    Next
    
    ' No match found
    CheckBlacklistedTerms = ""
End Function

' Test scenario 1: Empty blacklist (no matches expected)
Sub TestEmptyBlacklist()
    WScript.Echo "TEST 1: Empty blacklist"
    
    Dim blacklistArray, result, screenContent
    ReDim blacklistArray(-1) ' Empty array
    screenContent = "REPAIR ORDER #875084 DETAIL"
    
    result = CheckBlacklistedTerms(blacklistArray, screenContent)
    
    If result = "" Then
        WScript.Echo "  [PASS] Empty blacklist returns empty string"
    Else
        WScript.Echo "  [FAIL] Expected empty string, got '" & result & "'"
    End If
End Sub

' Test scenario 2: Single term matching "VEND TO DEALER"
Sub TestSingleTermMatch()
    WScript.Echo "TEST 2: Single term match - VEND TO DEALER"
    
    Dim blacklistArray(0), result, screenContent
    blacklistArray(0) = "VEND TO DEALER"
    screenContent = "REPAIR ORDER #875084 DETAIL" & vbCrLf & "B  VEND TO DEALER                        C92"
    
    result = CheckBlacklistedTerms(blacklistArray, screenContent)
    
    If result = "VEND TO DEALER" Then
        WScript.Echo "  [PASS] Matched VEND TO DEALER"
    Else
        WScript.Echo "  [FAIL] Expected 'VEND TO DEALER', got '" & result & "'"
    End If
End Sub

' Test scenario 3: Case-insensitive matching
Sub TestCaseInsensitive()
    WScript.Echo "TEST 3: Case-insensitive matching"
    
    Dim blacklistArray(0), result, screenContent
    blacklistArray(0) = "vend to dealer" ' lowercase in config
    screenContent = "B  VEND TO DEALER                        C92"
    
    result = CheckBlacklistedTerms(blacklistArray, screenContent)
    
    If result = "VEND TO DEALER" Then
        WScript.Echo "  [PASS] Case-insensitive match successful (returns uppercase)"
    Else
        WScript.Echo "  [FAIL] Expected 'VEND TO DEALER', got '" & result & "'"
    End If
End Sub

' Test scenario 4: Multiple blacklist terms, one matches
Sub TestMultipleTermsOneMatch()
    WScript.Echo "TEST 4: Multiple blacklist terms, one matches"
    
    Dim blacklistArray(2), result, screenContent
    blacklistArray(0) = "HOLD"
    blacklistArray(1) = "PENDING"
    blacklistArray(2) = "VEND TO DEALER"
    screenContent = "REPAIR ORDER #875084 DETAIL" & vbCrLf & "B  VEND TO DEALER                        C92"
    
    result = CheckBlacklistedTerms(blacklistArray, screenContent)
    
    If result = "VEND TO DEALER" Then
        WScript.Echo "  [PASS] Found correct term among multiple options"
    Else
        WScript.Echo "  [FAIL] Expected 'VEND TO DEALER', got '" & result & "'"
    End If
End Sub

' Test scenario 5: Multiple blacklist terms, no match
Sub TestMultipleTermsNoMatch()
    WScript.Echo "TEST 5: Multiple blacklist terms, no match"
    
    Dim blacklistArray(1), result, screenContent
    blacklistArray(0) = "HOLD"
    blacklistArray(1) = "PENDING"
    screenContent = "REPAIR ORDER #875084 DETAIL" & vbCrLf & "RO STATUS: READY TO POST"
    
    result = CheckBlacklistedTerms(blacklistArray, screenContent)
    
    If result = "" Then
        WScript.Echo "  [PASS] No match found as expected"
    Else
        WScript.Echo "  [FAIL] Expected empty string, got '" & result & "'"
    End If
End Sub

' Test scenario 6: Blacklist term appears as substring
Sub TestSubstringMatch()
    WScript.Echo "TEST 6: Substring match (VEND appears in VEND TO DEALER)"
    
    Dim blacklistArray(0), result, screenContent
    blacklistArray(0) = "VEND"
    screenContent = "B  VEND TO DEALER                        C92"
    
    result = CheckBlacklistedTerms(blacklistArray, screenContent)
    
    If result = "VEND" Then
        WScript.Echo "  [PASS] Substring match successful"
    Else
        WScript.Echo "  [FAIL] Expected 'VEND', got '" & result & "'"
    End If
End Sub

' Test scenario 7: Whitespace handling
Sub TestWhitespaceHandling()
    WScript.Echo "TEST 7: Whitespace handling in blacklist terms"
    
    Dim blacklistArray(0), result, screenContent
    blacklistArray(0) = "  VEND TO DEALER  " ' Terms with extra whitespace
    screenContent = "B  VEND TO DEALER                        C92"
    
    result = CheckBlacklistedTerms(blacklistArray, screenContent)
    
    ' Should trim and match
    If InStr(result, "VEND TO DEALER") > 0 Or InStr(result, "VEND") > 0 Then
        WScript.Echo "  [PASS] Whitespace handled, term matched"
    Else
        WScript.Echo "  [FAIL] Expected match, got '" & result & "'"
    End If
End Sub

' Test scenario 8: First matching term is returned
Sub TestFirstMatchReturn()
    WScript.Echo "TEST 8: First matching term returned when multiple match"
    
    Dim blacklistArray(2), result, screenContent
    blacklistArray(0) = "VEND"          ' Will match first in order
    blacklistArray(1) = "DEALER"        ' Also matches but comes second
    blacklistArray(2) = "TO"            ' Also matches but comes last
    screenContent = "B  VEND TO DEALER                        C92"
    
    result = CheckBlacklistedTerms(blacklistArray, screenContent)
    
    If result = "VEND" Then
        WScript.Echo "  [PASS] First matching term returned"
    Else
        WScript.Echo "  [FAIL] Expected 'VEND' (first match), got '" & result & "'"
    End If
End Sub

' Test scenario 9: Real-world RO screen content
Sub TestRealWorldScreen()
    WScript.Echo "TEST 9: Real-world RO screen with multiple fields"
    
    Dim blacklistArray(0), result, screenContent
    blacklistArray(0) = "VEND TO DEALER"
    
    ' Realistic example RO screen
    screenContent = "ONSITE PLUS SERVICE            (PFC) POST FINAL CHARGES            12MAR26 06:34" & vbCrLf & _
                    "RO: 875084     TAG: T5596556 SA: 18351   25 TOYOTA TOYOTA VIN: 4T1DAACK3SU553265" & vbCrLf & _
                    "NAME:                   PMT: CASH      MILEAGE:   33759     OPENED DATE: 05MAR26" & vbCrLf & _
                    "RO STATUS: READY TO POST          PROMISED: 05MAR26 17:00   OPENED TIME:   06:58" & vbCrLf & _
                    "REPAIR ORDER #875084 DETAIL" & vbCrLf & _
                    "LC DESCRIPTION                           TECH... LTYPE    ACT   SOLD    SALE AMT" & vbCrLf & _
                    "A  SECOND ROW SEAT BELT                  C93" & vbCrLf & _
                    "B  VEND TO DEALER                        C92"
    
    result = CheckBlacklistedTerms(blacklistArray, screenContent)
    
    If result = "VEND TO DEALER" Then
        WScript.Echo "  [PASS] Real-world screen match successful"
    Else
        WScript.Echo "  [FAIL] Expected 'VEND TO DEALER', got '" & result & "'"
    End If
End Sub

' Main test runner
Sub Main()
    WScript.Echo "PostFinalCharges Blacklist Terms - Unit Test Suite"
    WScript.Echo "=================================================="
    WScript.Echo ""
    
    ' Run all tests
    TestEmptyBlacklist
    WScript.Echo ""
    
    TestSingleTermMatch
    WScript.Echo ""
    
    TestCaseInsensitive
    WScript.Echo ""
    
    TestMultipleTermsOneMatch
    WScript.Echo ""
    
    TestMultipleTermsNoMatch
    WScript.Echo ""
    
    TestSubstringMatch
    WScript.Echo ""
    
    TestWhitespaceHandling
    WScript.Echo ""
    
    TestFirstMatchReturn
    WScript.Echo ""
    
    TestRealWorldScreen
    WScript.Echo ""
    
    WScript.Echo "=================================================="
    WScript.Echo "Test Suite Complete"
End Sub

' Run tests
Main()
