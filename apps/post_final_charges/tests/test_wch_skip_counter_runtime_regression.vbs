'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestPartsChargedGateRuntimeRegression
' **DATE CREATED:** 2026-04-09
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Regression guard for exception-aware parts gate and WCH skip gates in
' PostFinalCharges. Verifies no-parts bypass wiring, config-driven exception
' codes, and pagination command sequence expectations.
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_fso, g_scriptPath, g_content
Set g_fso = CreateObject("Scripting.FileSystemObject")
g_scriptPath = "../PostFinalCharges.vbs"

If Not g_fso.FileExists(g_scriptPath) Then
    WScript.Echo "[FAIL] PostFinalCharges.vbs not found at: " & g_scriptPath
    WScript.Quit 1
End If

Dim ts
Set ts = g_fso.OpenTextFile(g_scriptPath, 1)
g_content = ts.ReadAll
ts.Close

Dim failures
failures = 0

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

WScript.Echo "Parts-Charged Gate Runtime Regression Test"
WScript.Echo "==========================================="

' Exception-aware parts gate functions are present
AssertContains "EvaluatePartsChargedGate function is declared", "Function EvaluatePartsChargedGate(ByRef skipReason)"
AssertContains "Exception tech-code helper exists", "Function IsCdkLaborOnlyExceptionTech(techCode)"
AssertContains "Description exception helper exists", "Function IsCdkLaborOnlyExceptionDesc(descText)"
AssertContains "Parts gate scans P-line indicator", "Mid(buf, 6, 1) = ""P"""
AssertContains "Parts gate reads SALE AMT column", "Mid(buf, 70, 11)"
AssertContains "Parts gate supports labor-only bypass", "Labor-only exception matched - bypassing no-parts skip"
AssertContains "Parts gate checks line description exceptions", "hasDescException = IsCdkLaborOnlyExceptionDesc(lineDesc)"
AssertContains "Parts gate sets explicit offending-code result", "Skipped - No parts charged: "

' Guard is wired into Closeout_Ro before status routing
AssertContains "Closeout_Ro calls EvaluatePartsChargedGate", "If Not EvaluatePartsChargedGate(noPartsSkipReason) Then"
AssertContains "Closeout_Ro writes dynamic skip reason", "lastRoResult = noPartsSkipReason"

' Guard fires before FC/F commands (guard appears before Closeout_ReadyToPost)
AssertOrder "Parts guard precedes READY TO POST closeout", _
    "If Not EvaluatePartsChargedGate(noPartsSkipReason) Then", "Call Closeout_ReadyToPost()"

' Exception list is config-driven (default is I for internal labor)
AssertContains "Config reader loads labor-only exceptions", "GetIniSetting(""PostFinalCharges"", ""CDKLaborOnlyLTypeExceptions"", ""I"")"
AssertContains "Exception list is normalized to uppercase", "g_arrCDKExceptions(ei) = UCase(Trim(g_arrCDKExceptions(ei)))"
AssertContains "Config reader loads labor-only description exceptions", "GetIniSetting(""PostFinalCharges"", ""CDKLaborOnlyDescriptionExceptions"", ""check and adjust"")"
AssertContains "Description exceptions normalized lowercase", "g_arrCDKDescriptionExceptions(di) = LCase(Trim(g_arrCDKDescriptionExceptions(di)))"

' Configurable LTYPE block gate replaces WCH-specific gate
AssertContains "SkipLaborLTypes config key is read", "GetIniSetting(""PostFinalCharges"", ""SkipLaborLTypes"", ""WCH,WV,WF"")"
AssertContains "Blocked LTYPE array is declared", "Dim g_arrSkipLaborLTypes"
AssertContains "Blocked LTYPE helper is declared", "Function HasBlockedLTypeOnAnyPage()"
AssertContains "Main gate calls blocked LTYPE helper", "blockedLType = HasBlockedLTypeOnAnyPage()"
AssertContains "Blocked LTYPE skip result includes LTYPE code", "lastRoResult = ""Skipped - Blocked LTYPE: "" & blockedLType"
AssertContains "Blocked LTYPE summary line is present", "Skips - Blocked LTYPE:"
AssertContains "LTYPE gate reads col 50 for LTYPE", "Mid(buf, 50, 6)"
AssertContains "LTYPE gate checks L-row indicator", "Mid(buf, 4, 1) = ""L"""
AssertContains "LTYPE gate guards against empty lTypeCode", "If Len(lTypeCode) > 0 And IsArray(g_arrSkipLaborLTypes)"
AssertContains "LTYPE gate filters empty entries in loop", "If Len(g_arrSkipLaborLTypes(i)) > 0 And lTypeCode"
AssertContains "LTYPE gate pagination uses next-screen command", "g_bzhao.SendKey ""N"""
AssertContains "LTYPE gate pagination uses ENTER command", "g_bzhao.SendKey ""<NumpadEnter>"""
AssertContains "LTYPE gate pagination waits after page advance", "g_bzhao.Pause 500"
AssertContains "InitializeConfig filters empty SkipLaborLTypes entries", "g_arrSkipLaborLTypes = Array()"

WScript.Echo ""
If failures = 0 Then
    WScript.Echo "SUCCESS: Parts-charged gate runtime wiring is correct."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & failures & " regression checks failed."
    WScript.Quit 1
End If
