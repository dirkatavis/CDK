'-----------------------------------------------------------------------------------
' test_bzhelper.vbs
' Smoke test for framework\BZHelper.vbs.
' Loads BZHelper against a mock terminal object and exercises every exported
' function once, verifying no runtime errors occur.
'
' Catches: undefined function calls (e.g. IIf), missing VBScript intrinsics,
'          broken logic in exported functions.
'
' Usage: cscript.exe //nologo tests\infrastructure\test_bzhelper.vbs
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_sh:  Set g_sh  = CreateObject("WScript.Shell")
Dim g_root: g_root = g_sh.Environment("USER")("CDK_BASE")

Dim failures: failures = 0

WScript.Echo "BZHelper Smoke Test"
WScript.Echo "==================="

' ---------------------------------------------------------------------------
' 1. Load AdvancedMock and wire up g_bzhao
' ---------------------------------------------------------------------------
On Error Resume Next
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\AdvancedMock.vbs")).ReadAll
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] AdvancedMock.vbs load: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

Dim g_bzhao: Set g_bzhao = New AdvancedMock
g_bzhao.Connect ""
g_bzhao.SetBuffer "COMMAND:" & String((24 * 80) - 8, " ")

' ---------------------------------------------------------------------------
' 2. Load BZHelper
' ---------------------------------------------------------------------------
On Error Resume Next
Err.Clear
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\BZHelper.vbs")).ReadAll
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] BZHelper.vbs load: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo "[PASS] BZHelper.vbs load"

' ---------------------------------------------------------------------------
' 3. WaitMs
' ---------------------------------------------------------------------------
On Error Resume Next
Err.Clear
WaitMs 1
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] WaitMs: " & Err.Description
    failures = failures + 1
Else
    WScript.Echo "[PASS] WaitMs"
End If
On Error GoTo 0

' ---------------------------------------------------------------------------
' 4. IsTextPresent — single term, should find COMMAND: in mock buffer
' ---------------------------------------------------------------------------
On Error Resume Next
Err.Clear
Dim result: result = IsTextPresent("COMMAND:")
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] IsTextPresent (single): " & Err.Description
    failures = failures + 1
ElseIf Not result Then
    WScript.Echo "[FAIL] IsTextPresent (single): expected True, got False"
    failures = failures + 1
Else
    WScript.Echo "[PASS] IsTextPresent (single term)"
End If
On Error GoTo 0

' ---------------------------------------------------------------------------
' 5. IsTextPresent — pipe-delimited, second term matches
' ---------------------------------------------------------------------------
On Error Resume Next
Err.Clear
result = IsTextPresent("NOTFOUND|COMMAND:")
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] IsTextPresent (pipe): " & Err.Description
    failures = failures + 1
ElseIf Not result Then
    WScript.Echo "[FAIL] IsTextPresent (pipe): expected True, got False"
    failures = failures + 1
Else
    WScript.Echo "[PASS] IsTextPresent (pipe-delimited)"
End If
On Error GoTo 0

' ---------------------------------------------------------------------------
' 6. IsTextPresent — no match, should return False
' ---------------------------------------------------------------------------
On Error Resume Next
Err.Clear
result = IsTextPresent("NOTFOUND")
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] IsTextPresent (no match): " & Err.Description
    failures = failures + 1
ElseIf result Then
    WScript.Echo "[FAIL] IsTextPresent (no match): expected False, got True"
    failures = failures + 1
Else
    WScript.Echo "[PASS] IsTextPresent (no match)"
End If
On Error GoTo 0

' ---------------------------------------------------------------------------
' 7. BZSendKey
' ---------------------------------------------------------------------------
On Error Resume Next
Err.Clear
BZSendKey "X"
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] BZSendKey: " & Err.Description
    failures = failures + 1
Else
    WScript.Echo "[PASS] BZSendKey"
End If
On Error GoTo 0

' ---------------------------------------------------------------------------
' 8. BZReadScreen
' ---------------------------------------------------------------------------
On Error Resume Next
Err.Clear
Dim buf: buf = BZReadScreen(8, 1, 1)
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] BZReadScreen: " & Err.Description
    failures = failures + 1
Else
    WScript.Echo "[PASS] BZReadScreen"
End If
On Error GoTo 0

' ---------------------------------------------------------------------------
' 9. WaitForPrompt — with description (exercises the IIf-equivalent line)
' ---------------------------------------------------------------------------
On Error Resume Next
Err.Clear
result = WaitForPrompt("COMMAND:", "", False, 1000, "smoke test")
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] WaitForPrompt (with description): " & Err.Description
    failures = failures + 1
ElseIf Not result Then
    WScript.Echo "[FAIL] WaitForPrompt (with description): prompt not found in mock buffer"
    failures = failures + 1
Else
    WScript.Echo "[PASS] WaitForPrompt (with description)"
End If
On Error GoTo 0

' ---------------------------------------------------------------------------
' 10. WaitForPrompt — empty description (exercises the else branch)
' ---------------------------------------------------------------------------
On Error Resume Next
Err.Clear
result = WaitForPrompt("COMMAND:", "", False, 1000, "")
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] WaitForPrompt (empty description): " & Err.Description
    failures = failures + 1
ElseIf Not result Then
    WScript.Echo "[FAIL] WaitForPrompt (empty description): prompt not found in mock buffer"
    failures = failures + 1
Else
    WScript.Echo "[PASS] WaitForPrompt (empty description)"
End If
On Error GoTo 0

' ---------------------------------------------------------------------------
' 11. WaitForAnyOf — pipe-delimited, second target matches
' ---------------------------------------------------------------------------
On Error Resume Next
Err.Clear
result = WaitForAnyOf("NOTFOUND|COMMAND:", 1000)
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] WaitForAnyOf: " & Err.Description
    failures = failures + 1
ElseIf Not result Then
    WScript.Echo "[FAIL] WaitForAnyOf: expected True, got False"
    failures = failures + 1
Else
    WScript.Echo "[PASS] WaitForAnyOf"
End If
On Error GoTo 0

' ---------------------------------------------------------------------------
' Summary
' ---------------------------------------------------------------------------
WScript.Echo ""
If failures = 0 Then
    WScript.Echo "SUCCESS: All BZHelper smoke tests passed."
    WScript.Quit 0
Else
    WScript.Echo "FAILED: " & failures & " smoke test(s) failed."
    WScript.Quit 1
End If
