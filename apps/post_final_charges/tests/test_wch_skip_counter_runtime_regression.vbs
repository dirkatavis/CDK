'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestPartsChargedGateRuntimeRegression
' **DATE CREATED:** 2026-04-09
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Regression guard for the exception-aware parts gate and warranty review flow
' in PostFinalCharges. Verifies no-parts bypass wiring, config-driven exception
' codes, blacklist gate wiring, and warranty review flow wiring.
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

Sub AssertAbsent(label, needle)
    If InStr(1, g_content, needle, vbTextCompare) = 0 Then
        WScript.Echo "[PASS] " & label
    Else
        WScript.Echo "[FAIL] " & label & " (should be absent: " & needle & ")"
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

' Labor-only gate and parts-charged gate functions are present
AssertContains "EvaluateLaborOnlyGate function is declared", "Function EvaluateLaborOnlyGate(ByRef skipReason)"
AssertContains "EvaluatePartsChargedGate function is declared", "Function EvaluatePartsChargedGate(ByRef skipReason)"
AssertContains "Description exception helper exists", "Function IsCdkLaborOnlyExceptionDesc(descText)"
AssertContains "Labor-only gate tracks line header description", "currentLineHeaderDesc"
AssertContains "Labor-only gate scans P-line indicator", "Mid(buf, 6, 1) = ""P"""
AssertContains "Labor-only gate reads SALE AMT column", "Mid(buf, 70, 11)"
AssertContains "Labor-only gate checks description exceptions", "IsCdkLaborOnlyExceptionDesc("
AssertContains "Labor-only gate sets skip reason with descriptions", "Skipped - Labor line requires parts:"
AssertAbsent "Ltype exception helper is removed", "Function IsCdkLaborOnlyExceptionTech("
AssertAbsent "Ltype exception global is removed", "g_arrCDKExceptions"

' Labor-only gate is wired into main flow before FindTrigger
AssertContains "Main flow calls EvaluateLaborOnlyGate", "If Not EvaluateLaborOnlyGate(laborOnlySkipReason) Then"
AssertContains "Labor-only gate skip increments no-parts counter", "g_SkipNoPartsChargedCount = g_SkipNoPartsChargedCount + 1"
AssertOrder "Labor-only gate precedes FindTrigger", _
    "If Not EvaluateLaborOnlyGate(laborOnlySkipReason) Then", "trigger = FindTrigger()"

' Parts-charged gate is wired into Closeout_Ro as late safety net
AssertContains "Closeout_Ro calls EvaluatePartsChargedGate", "If Not EvaluatePartsChargedGate(noPartsSkipReason) Then"
AssertContains "Closeout_Ro writes dynamic skip reason", "lastRoResult = noPartsSkipReason"
AssertOrder "Parts charged guard precedes READY TO POST closeout", _
    "If Not EvaluatePartsChargedGate(noPartsSkipReason) Then", "Call Closeout_ReadyToPost()"

' Exception list is config-driven (description only)
AssertContains "Config reader loads labor-only description exceptions", "GetIniSetting(""PostFinalCharges"", ""CDKLaborOnlyDescriptionExceptions"", ""check and adjust"")"
AssertContains "Description exceptions normalized lowercase", "g_arrCDKDescriptionExceptions(di) = LCase(Trim(g_arrCDKDescriptionExceptions(di)))"
AssertAbsent "Ltype exception config key is removed", "CDKLaborOnlyLTypeExceptions"

' Blacklist gate is retained (general-purpose full-screen scan)
AssertContains "Blacklist raw terms global is declared", "Dim g_BlacklistTermsRaw"
AssertContains "Blacklist helper is called in main flow", "BZH_GetMatchedBlacklistTerm("

' Retired WCH gate is gone
AssertAbsent "WCH-specific skip gate is removed", "g_SkipWchEnabled"
AssertAbsent "HasWchOnAnyDetailPage is removed", "Function HasWchOnAnyDetailPage()"
AssertAbsent "FCA field-filling handler is removed", "Sub HandleFcaDialog("
AssertAbsent "FCA prompt dictionary builder is removed", "Function CreateFcaPromptDictionary("
AssertAbsent "IsWchLine hardcoded check is removed", "Function IsWchLine("
AssertAbsent "ExtractPartNumberForFca is removed", "Function ExtractPartNumberForFca()"

' Warranty review flow is present
AssertContains "IsWarrantyLine function is declared", "Function IsWarrantyLine(lineLetterChar)"
AssertContains "IsWarrantyLine checks config-driven array", "g_arrWarrantyLTypes"
AssertContains "HandleWarrantyClaimsDialog sub is declared", "Sub HandleWarrantyClaimsDialog(maxPolls)"
AssertContains "Dialog detects LABOR OP: prompt", "InStr(1, buf, ""LABOR OP:"", vbTextCompare)"
AssertContains "Dialog detects COMMAND: prompt", "InStr(1, buf, ""COMMAND:"", vbTextCompare)"
AssertContains "Dialog sends blank Enter for LABOR OP: state", "WaitForPrompt(""LABOR OP:"", """", True"
AssertContains "Dialog sends period to skip fields in COMMAND: state", "FastText(""."")"
AssertContains "Dialog sends E to exit in COMMAND: state", "FastText(""E"")"
AssertContains "WarrantyLTypes config key is read", "GetIniSetting(""PostFinalCharges"", ""WarrantyLTypes"", ""WCH,WV,WF"")"
AssertContains "WarrantyCauseText config key is read", "GetIniSetting(""PostFinalCharges"", ""WarrantyCauseText"", ""Device failure"")"
AssertContains "WarrantyDialogStepDelayMs config key is read", "GetIniSetting(""PostFinalCharges"", ""WarrantyDialogStepDelayMs"", ""2000"")"
AssertContains "WarrantyDialogSignatures config key is read", "GetIniSetting(""PostFinalCharges"", ""WarrantyDialogSignatures"","
AssertContains "g_WarrantyCauseText global is declared", "Dim g_WarrantyCauseText"
AssertContains "g_WarrantyDialogStepDelayMs global is declared", "Dim g_WarrantyDialogStepDelayMs"
AssertContains "g_WarrantyDialogSignatureTexts global is declared", "Dim g_WarrantyDialogSignatureTexts()"
AssertContains "g_WarrantyDialogSignatureTypes global is declared", "Dim g_WarrantyDialogSignatureTypes()"
AssertContains "CAUSE L prefix detection is present", "CAUSE L"
AssertContains "CAUSE L loop uses inner poll to avoid premature exit", "For causePoll = 1 To 6"
AssertContains "DetectWarrantyDialog function is declared", "Function DetectWarrantyDialog(maxPolls)"
AssertContains "HandleWarrantyClaimsDialog accepts maxPolls", "Sub HandleWarrantyClaimsDialog(maxPolls)"
AssertContains "HandleFcaClaimsDialog sub is declared", "Sub HandleFcaClaimsDialog()"
AssertContains "HandleVwWarrantyDialog sub is declared", "Sub HandleVwWarrantyDialog()"
AssertContains "IsWarrantyLine is called before FNL in ProcessLinesSequentially", "lineIsWarranty = IsWarrantyLine(lineLetterChar)"
AssertContains "warrantyPolls computed from lineIsWarranty", "warrantyPolls = 20"
AssertContains "HandleWarrantyClaimsDialog always called with poll count", "Call HandleWarrantyClaimsDialog(warrantyPolls)"
AssertContains "fnlPrompts has ADD A LABOR OPER safety net", "ADD A LABOR OPER"

' Warranty dialog handler fires after R prompts, not after FNL
AssertOrder "HandleWarrantyClaimsDialog fires after R review prompts", _
    "Call ProcessPromptSequence(lineItemPrompts)", "Call HandleWarrantyClaimsDialog(warrantyPolls)"
AssertOrder "HandleWarrantyClaimsDialog fires after R review prompts (not FNL)", _
    "lineIsWarranty = IsWarrantyLine(lineLetterChar)", "Call HandleWarrantyClaimsDialog(warrantyPolls)"

WScript.Echo ""
If failures = 0 Then
    WScript.Echo "SUCCESS: Parts-charged gate runtime wiring is correct."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & failures & " regression checks failed."
    WScript.Quit 1
End If
