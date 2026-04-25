'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestSkipLaborOnlyGate
' **DATE CREATED:** 2026-04-23
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Regression tests for the SkipLaborOnlyGate config flag behavior at the
' EvaluateLaborOnlyGate call site in Main:
'   1) SkipLaborOnlyGate=true + no-parts failure      -> proceed (bypass)
'   2) SkipLaborOnlyGate=true + unsupported W* ltype  -> skip (always enforced)
'   3) SkipLaborOnlyGate=false + no-parts failure     -> skip (gate active)
'   4) SkipLaborOnlyGate=false + unsupported W* ltype -> skip (gate active)
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_Pass, g_Fail
g_Pass = 0
g_Fail = 0

' ---- Assert helpers ----
Sub AssertTrue(ByVal label, ByVal value)
    If value Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (expected True)"
    End If
End Sub

Sub AssertFalse(ByVal label, ByVal value)
    If Not value Then
        g_Pass = g_Pass + 1
        WScript.Echo "[PASS] " & label
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "[FAIL] " & label & " (expected False)"
    End If
End Sub

'-----------------------------------------------------------------------------------
' Extracted call-site decision logic from Main in PostFinalCharges.vbs.
' Returns True if the RO should be skipped, False if it should proceed.
'-----------------------------------------------------------------------------------
Function ShouldSkipForLaborOnly(ByVal skipReason, ByVal skipLaborOnlyGate)
    Dim isWarrantyFailure
    isWarrantyFailure = (InStr(1, skipReason, "Unsupported warranty ltype", vbTextCompare) > 0)
    If skipLaborOnlyGate And Not isWarrantyFailure Then
        ShouldSkipForLaborOnly = False
    Else
        ShouldSkipForLaborOnly = True
    End If
End Function

' ============================
' Tests
' ============================
WScript.Echo "SkipLaborOnlyGate — Call-Site Decision Tests"
WScript.Echo "============================================="

Const NO_PARTS_REASON = "Skipped - No parts charged: lrow=[VCC VERIFY CONNECTED CAR DEVICE IS] header=[VERIFY CC DEVICE OPERATION]"
Const WARRANTY_REASON_WM = "Skipped - Unsupported warranty ltype: [WM]"
Const WARRANTY_REASON_WT = "Skipped - Unsupported warranty ltype: [WT]"

' --- Test 1: SkipLaborOnlyGate=true + no-parts failure -> proceed ---
AssertFalse "Gate=true + no-parts failure -> should NOT skip (bypass)", _
    ShouldSkipForLaborOnly(NO_PARTS_REASON, True)

' --- Test 2: SkipLaborOnlyGate=true + unsupported WM ltype -> skip ---
AssertTrue "Gate=true + WM warranty ltype -> should still skip", _
    ShouldSkipForLaborOnly(WARRANTY_REASON_WM, True)

' --- Test 3: SkipLaborOnlyGate=true + unsupported WT ltype -> skip ---
AssertTrue "Gate=true + WT warranty ltype -> should still skip", _
    ShouldSkipForLaborOnly(WARRANTY_REASON_WT, True)

' --- Test 4: SkipLaborOnlyGate=false + no-parts failure -> skip ---
AssertTrue "Gate=false + no-parts failure -> should skip (gate active)", _
    ShouldSkipForLaborOnly(NO_PARTS_REASON, False)

' --- Test 5: SkipLaborOnlyGate=false + unsupported warranty ltype -> skip ---
AssertTrue "Gate=false + WM warranty ltype -> should skip (gate active)", _
    ShouldSkipForLaborOnly(WARRANTY_REASON_WM, False)

WScript.Echo ""
If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All " & g_Pass & " SkipLaborOnlyGate tests passed."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & g_Fail & " test(s) failed."
    WScript.Quit 1
End If
