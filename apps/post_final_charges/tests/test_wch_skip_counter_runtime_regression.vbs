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

' Exception list is config-driven
AssertContains "Config reader loads labor-only exceptions", "GetIniSetting(""PostFinalCharges"", ""CDKLaborOnlyLTypeExceptions"", ""WCH,WT,WF"")"
AssertContains "Exception list is normalized to uppercase", "g_arrCDKExceptions(ei) = UCase(Trim(g_arrCDKExceptions(ei)))"
AssertContains "Config reader loads labor-only description exceptions", "GetIniSetting(""PostFinalCharges"", ""CDKLaborOnlyDescriptionExceptions"", ""check and adjust"")"
AssertContains "Description exceptions normalized lowercase", "g_arrCDKDescriptionExceptions(di) = LCase(Trim(g_arrCDKDescriptionExceptions(di)))"

' WCH gate is enabled/disabled by config and uses paginated detection
AssertContains "WCH feature flag exists", "Dim g_SkipWchEnabled"
AssertContains "WCH pagination helper is declared", "Function HasWchOnAnyDetailPage()"
AssertContains "WCH gate uses pagination helper", "If g_SkipWchEnabled And HasWchOnAnyDetailPage() Then"
AssertContains "WCH skip result label is preserved", "lastRoResult = ""Skipped - WCH labor type"""
AssertContains "WCH summary line is present", "Skips - Warranty (WCH):"
AssertContains "WCH pagination uses next-screen command", "g_bzhao.SendKey ""N"""
AssertContains "WCH pagination uses ENTER command", "g_bzhao.SendKey ""<NumpadEnter>"""
AssertContains "WCH pagination waits after page advance", "g_bzhao.Pause 500"

WScript.Echo ""
If failures = 0 Then
    WScript.Echo "SUCCESS: Parts-charged gate runtime wiring is correct."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & failures & " regression checks failed."
    WScript.Quit 1
End If
