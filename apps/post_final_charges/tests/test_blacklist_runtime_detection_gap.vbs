'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestBlacklistRuntimeDetectionGap
' **DATE CREATED:** 2026-03-16
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Exploratory test that mirrors current runtime blacklist detection behavior
' (line-by-line InStr matching) to identify likely production miss scenarios.
'-----------------------------------------------------------------------------------

Option Explicit

Dim failures
failures = 0

Function IsTextPresentInLines(searchText, lines)
    Dim i, lineContent
    IsTextPresentInLines = False

    If Len(searchText) = 0 Then Exit Function

    For i = LBound(lines) To UBound(lines)
        lineContent = Trim(CStr(lines(i)))
        If InStr(1, lineContent, searchText, vbTextCompare) > 0 Then
            IsTextPresentInLines = True
            Exit Function
        End If
    Next
End Function

Function GetMatchedBlacklistTermFromLines(blacklistTermsCsv, lines)
    Dim terms, i, term
    GetMatchedBlacklistTermFromLines = ""

    blacklistTermsCsv = Trim(CStr(blacklistTermsCsv))
    If Len(blacklistTermsCsv) = 0 Then Exit Function

    terms = Split(blacklistTermsCsv, ",")
    For i = LBound(terms) To UBound(terms)
        term = Trim(CStr(terms(i)))
        If Len(term) > 0 Then
            If IsTextPresentInLines(term, lines) Then
                GetMatchedBlacklistTermFromLines = term
                Exit Function
            End If
        End If
    Next
End Function

Sub AssertEqual(label, actual, expected)
    If StrComp(CStr(actual), CStr(expected), vbTextCompare) = 0 Then
        WScript.Echo "[PASS] " & label & " -> '" & actual & "'"
    Else
        WScript.Echo "[FAIL] " & label & " -> expected '" & expected & "', got '" & actual & "'"
        failures = failures + 1
    End If
End Sub

Sub Main()
    Dim blacklist
    blacklist = "VEND TO DEALER"

    WScript.Echo "Blacklist Runtime Detection Gap Test"
    WScript.Echo "===================================="

    ' Case 1: Exact same-line phrase (should match)
    Dim linesExact
    linesExact = Array( _
        "REPAIR ORDER #875268 DETAIL", _
        "A  Electric Vehicle Sound Levels", _
        "B  VEND TO DEALER                        C92" _
    )
    AssertEqual "Exact phrase on one line", GetMatchedBlacklistTermFromLines(blacklist, linesExact), "VEND TO DEALER"

    ' Case 2: Double-space variation (likely visual match, exact-string miss)
    Dim linesDoubleSpace
    linesDoubleSpace = Array( _
        "REPAIR ORDER #875268 DETAIL", _
        "B  VEND  TO DEALER                       C92" _
    )
    AssertEqual "Double-space variation", GetMatchedBlacklistTermFromLines(blacklist, linesDoubleSpace), ""

    ' Case 3: Line-wrap split (visible to user, not contiguous in one line)
    Dim linesWrapped
    linesWrapped = Array( _
        "REPAIR ORDER #875268 DETAIL", _
        "B  VEND TO", _
        "   DEALER                               C92" _
    )
    AssertEqual "Line-wrap split across rows", GetMatchedBlacklistTermFromLines(blacklist, linesWrapped), ""

    WScript.Echo ""
    If failures = 0 Then
        WScript.Echo "SUCCESS: Runtime-equivalent matcher behavior reproduced expected gaps."
        WScript.Quit 0
    Else
        WScript.Echo "FAILED: " & failures & " check(s) failed."
        WScript.Quit 1
    End If
End Sub

Main()
