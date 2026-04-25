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
AssertContains "Labor-only gate checks for P-row following L-row", "Mid(buf, 6, 1) = ""P"""
AssertContains "Labor-only gate reads SALE AMT column", "Mid(buf, 70, 11)"
AssertContains "Labor-only gate checks description exceptions", "IsCdkLaborOnlyExceptionDesc("
AssertContains "Labor-only gate sets no-parts skip reason", "Skipped - No parts charged: lrow=["
AssertContains "Labor-only gate sets unsupported warranty skip reason", "Skipped - Unsupported warranty ltype: ["
AssertAbsent "Unsupported warranty skip reason does not include lrow detail", "Unsupported warranty ltype: [WF] lrow=["
AssertContains "Labor-only gate uses pending pattern for P-row lookahead", "pendingLRowDesc"
AssertAbsent "Parts-order keyword function is removed", "Function DescMatchesPartsKeyword("
AssertAbsent "Parts-order scan function is removed", "Function GetPartsNeededLaborDesc("
AssertAbsent "Parts-order keywords global is removed", "g_PartsOrderKeywords"
AssertAbsent "Parts-order negators global is removed", "g_PartsOrderNegators"
AssertAbsent "Parts-order needed counter is removed", "g_SkipPartsOrderNeededCount"
AssertAbsent "Ltype exception helper is removed", "Function IsCdkLaborOnlyExceptionTech("
AssertAbsent "Ltype exception global is removed", "g_arrCDKExceptions"

' Labor-only gate is wired into main flow before FindTrigger
AssertContains "Main flow calls EvaluateLaborOnlyGate", "If Not EvaluateLaborOnlyGate(laborOnlySkipReason) Then"
AssertContains "Labor-only gate result routes through TrackPrimaryOutcomeCounters", "lastRoResult = laborOnlySkipReason"
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
AssertContains "WarrantyLTypes config key is read", "GetIniSetting(""PostFinalCharges"", ""WarrantyLTypes"", ""WCH,WF,W"")"
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
AssertContains "HandleWWarrantyDialog sub is declared", "Sub HandleWWarrantyDialog()"
AssertContains "HandleFordWarrantyDialog sub is declared", "Sub HandleFordWarrantyDialog()"
AssertContains "Ford dispatcher wired in HandleWarrantyClaimsDialog", "ElseIf dialogType = ""FORD"" Then"
AssertContains "g_FordWarrantyCauseText global is declared", "Dim g_FordWarrantyCauseText"
AssertContains "FordWarrantyCauseText config key is read", "GetIniSetting(""PostFinalCharges"", ""FordWarrantyCauseText"", ""Defective Part"")"
AssertContains "Ford dialog license state TODO is present", "TODO: license state is per-vehicle"
AssertContains "IsWarrantyLine is called before FNL in ProcessLinesSequentially", "lineIsWarranty = IsWarrantyLine(lineLetterChar)"
AssertContains "warrantyPolls computed from lineIsWarranty", "warrantyPolls = 20"
AssertContains "HandleWarrantyClaimsDialog always called with poll count", "Call HandleWarrantyClaimsDialog(warrantyPolls)"
AssertContains "fnlPrompts has ADD A LABOR OPER safety net", "ADD A LABOR OPER"

' Warranty dialog handler fires after R prompts, not after FNL
AssertOrder "HandleWarrantyClaimsDialog fires after R review prompts", _
    "Call ProcessPromptSequence(lineItemPrompts)", "Call HandleWarrantyClaimsDialog(warrantyPolls)"
AssertOrder "HandleWarrantyClaimsDialog fires after R review prompts (not FNL)", _
    "lineIsWarranty = IsWarrantyLine(lineLetterChar)", "Call HandleWarrantyClaimsDialog(warrantyPolls)"

' Per-line tech code routing in ProcessLinesSequentially
AssertContains "GetLineTechCode helper is declared", "Function GetLineTechCode(lineLetterChar)"
AssertContains "ProcessLinesSequentially reads per-line tech code", "lineTechCode = GetLineTechCode(lineLetterChar)"
AssertContains "C93 branch skips FNL and R", "lineTechCode = ""C93"""
AssertContains "C92 branch skips FNL only", "lineTechCode = ""C92"""
AssertContains "skipFnlForLine flag is set for C92", "skipFnlForLine = (lineTechCode = ""C92"")"
AssertOrder "C93 check precedes C92 check", _
    "lineTechCode = ""C93""", "lineTechCode = ""C92"""

' VTD labor gate wiring
AssertContains "ContainsWholeWordVtd helper is declared", "Function ContainsWholeWordVtd(text)"
AssertContains "EvaluateVtdLaborGate function is declared", "Function EvaluateVtdLaborGate(ByRef skipReason)"
AssertContains "VTD gate checks ltype I", "lTypeCode = ""I"""
AssertContains "VTD gate calls whole-word check", "ContainsWholeWordVtd(lRowDesc)"
AssertContains "VTD gate sets skip reason prefix", "Skipped - VTD labor line:"
AssertContains "g_SkipVtdLaborCount global is declared", "Dim g_SkipVtdLaborCount"
AssertContains "g_SkipVtdLaborCount reset in ProcessRONumbers", "g_SkipVtdLaborCount = 0"
AssertContains "Main calls EvaluateVtdLaborGate", "If Not EvaluateVtdLaborGate(vtdSkipReason) Then"
AssertContains "Main increments g_SkipVtdLaborCount on VTD gate failure", "g_SkipVtdLaborCount = g_SkipVtdLaborCount + 1"
AssertOrder "VTD gate precedes labor-only gate in Main", _
    "If Not EvaluateVtdLaborGate(vtdSkipReason) Then", "If Not EvaluateLaborOnlyGate(laborOnlySkipReason) Then"
AssertContains "BuildSessionSummary includes VTD count in miscTotal", "g_SkipVtdLaborCount + _"
AssertContains "BuildSessionSummary shows VTD detail line", "Skipped - VTD labor line: "" & g_SkipVtdLaborCount"
AssertContains "IsResultRepresentedInSummary handles VTD skip prefix", """SKIPPED - VTD LABOR LINE"""

' Ford dialog — config-driven license state
AssertContains "g_FordWarrantyLicenseState global is declared", "Dim g_FordWarrantyLicenseState"
AssertContains "FordWarrantyLicenseState config key is read", "GetIniSetting(""PostFinalCharges"", ""FordWarrantyLicenseState"", ""GA"")"
AssertContains "VLS step reads field content before sending state code", "vlsFieldText = Trim(Mid(vlsBuf, vlsLabelPos + 22, 5))"
AssertContains "VLS step uses config value when field is blank", "g_FordWarrantyLicenseState"
AssertContains "GetLineTechCode pagination limitation is documented", "Scans only the currently visible page"

WScript.Echo ""
If failures = 0 Then
    WScript.Echo "SUCCESS: Parts-charged gate runtime wiring is correct."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & failures & " regression checks failed."
    WScript.Quit 1
End If
