'=====================================================================================
' test_mock_catch_bug.vbs - Demonstrating "Partial Load" race condition detection
'=====================================================================================

Option Explicit

' Include AdvancedMock
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim scriptDir: scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
Dim repoRoot: repoRoot = fso.GetParentFolderName(fso.GetParentFolderName(scriptDir))
Dim mockPath: mockPath = fso.BuildPath(repoRoot, "framework\AdvancedMock.vbs")
ExecuteGlobal fso.OpenTextFile(mockPath).ReadAll

' "The Brittle Script" Logic
' This simulates a script that assumes if it can see Row 1, the whole screen is there.
Function IsReadyForCommand(bzhao)
    Dim row1, row23
    bzhao.ReadScreen row1, 10, 1, 1
    If Trim(row1) <> "" Then
        ' Script assumes if Row 1 is here, Row 23 must be too!
        bzhao.ReadScreen row23, 8, 23, 1
        IsReadyForCommand = (Trim(row23) = "COMMAND:")
    Else
        IsReadyForCommand = False
    End If
End Function

' --- Test Execution ---
Dim mock: Set mock = New AdvancedMock
mock.Connect "A"

' Scenario: Full Screen Loaded
WScript.Echo "Testing with FULL screen load..."
Dim buffer: buffer = String(24 * 80, " ")
buffer = "HEADER" & Mid(buffer, 7)
buffer = Left(buffer, (22*80)) & "COMMAND:" & Mid(buffer, (22*80)+9)
mock.SetBuffer buffer
mock.SetPartialLoad 24 ' Full load

If IsReadyForCommand(mock) Then
    WScript.Echo "[PASS] Script correctly detected COMMAND: on full load."
Else
    WScript.Echo "[FAIL] Script failed to detect COMMAND: on full load."
End If

WScript.Echo ""

' Scenario: Partial Screen Load (Race Condition)
WScript.Echo "Testing with PARTIAL screen load (simulate network lag)..."
mock.SetPartialLoad 10 ' Only top 10 rows arrived

If IsReadyForCommand(mock) Then
    WScript.Echo "[FAIL] Script falsely claimed it was ready (Race Condition Not Caught!)"
Else
    ' The script should realize it's NOT ready because row 23 is empty
    WScript.Echo "[SUCCESS] Mock caught the race condition! Script correctly reported NOT ready."
End If
