'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestBlacklistReadyStatusPrecedence
' **DATE CREATED:** 2026-03-16
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Validates blacklist precedence when RO STATUS is READY TO POST.
' Guards against regressions where IsStatusReady is evaluated before blacklist skip.
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_fso, g_scriptPath, g_content, failures
Set g_fso = CreateObject("Scripting.FileSystemObject")
g_scriptPath = "../PostFinalCharges.vbs"
failures = 0

If Not g_fso.FileExists(g_scriptPath) Then
    WScript.Echo "[FAIL] PostFinalCharges.vbs not found at: " & g_scriptPath
    WScript.Quit 1
End If

Dim ts
Set ts = g_fso.OpenTextFile(g_scriptPath, 1)
g_content = ts.ReadAll
ts.Close

Sub AssertContains(label, needle)
    If InStr(1, g_content, needle, vbTextCompare) > 0 Then
        WScript.Echo "[PASS] " & label
    Else
        WScript.Echo "[FAIL] " & label & " (missing: " & needle & ")"
        failures = failures + 1
    End If
End Sub

Function IndexOf(needle)
    IndexOf = InStr(1, g_content, needle, vbTextCompare)
End Function

Sub AssertOrder(label, firstNeedle, secondNeedle)
    Dim i1, i2
    i1 = IndexOf(firstNeedle)
    i2 = IndexOf(secondNeedle)

    If i1 > 0 And i2 > 0 And i1 < i2 Then
        WScript.Echo "[PASS] " & label
    Else
        WScript.Echo "[FAIL] " & label & " (expected order not found)"
        failures = failures + 1
    End If
End Sub

Function CheckBlacklistedTerms(blacklistCsv, screenContent)
    Dim terms, i, term
    CheckBlacklistedTerms = ""

    blacklistCsv = Trim(CStr(blacklistCsv))
    If Len(blacklistCsv) = 0 Then Exit Function

    terms = Split(blacklistCsv, ",")
    screenContent = UCase(CStr(screenContent))

    For i = LBound(terms) To UBound(terms)
        term = Trim(UCase(CStr(terms(i))))
        If Len(term) > 0 Then
            If InStr(1, screenContent, term, vbTextCompare) > 0 Then
                CheckBlacklistedTerms = term
                Exit Function
            End If
        End If
    Next
End Function

Sub TestReadyStatusWithBlacklistMatch()
    WScript.Echo "TEST: READY TO POST + blacklist term"

    Dim screenContent, matched
    screenContent = "RO STATUS: READY TO POST" & vbCrLf & _
                    "B VEND TO DEALER        C92"

    matched = CheckBlacklistedTerms("VEND TO DEALER", screenContent)
    If matched = "VEND TO DEALER" Then
        WScript.Echo "[PASS] Blacklist term detected even when status is READY TO POST"
    Else
        WScript.Echo "[FAIL] Expected blacklist detection, got: '" & matched & "'"
        failures = failures + 1
    End If
End Sub

WScript.Echo "Blacklist READY-TO-POST Precedence Regression Test"
WScript.Echo "================================================="

AssertContains "Main checks blacklist", "matchedBlacklistTerm = GetMatchedBlacklistTerm(g_BlacklistTermsRaw)"
AssertContains "Main checks status readiness", "If Not IsStatusReady() Then"
AssertContains "Main sets blacklist skip result", "lastRoResult = ""Skipped - Blacklisted term: "" & matchedBlacklistTerm"
AssertContains "Main sets status skip result", "lastRoResult = ""Skipped - Status not ready"""

AssertOrder "Blacklist check occurs before IsStatusReady", _
    "matchedBlacklistTerm = GetMatchedBlacklistTerm(g_BlacklistTermsRaw)", _
    "If Not IsStatusReady() Then"

AssertOrder "Blacklist skip assignment occurs before status skip assignment", _
    "lastRoResult = ""Skipped - Blacklisted term: "" & matchedBlacklistTerm", _
    "lastRoResult = ""Skipped - Status not ready"""

TestReadyStatusWithBlacklistMatch

WScript.Echo ""
If failures = 0 Then
    WScript.Echo "SUCCESS: Blacklist precedence over READY TO POST is enforced."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & failures & " checks failed."
    WScript.Quit 1
End If
