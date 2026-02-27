'=====================================================================================
' TestAdvancedMock.vbs - Validation for the AdvancedMock Framework
'=====================================================================================

Option Explicit

' Include AdvancedMock
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim scriptDir: scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
Dim repoRoot: repoRoot = fso.GetParentFolderName(fso.GetParentFolderName(scriptDir))
Dim mockPath: mockPath = fso.BuildPath(repoRoot, "framework\AdvancedMock.vbs")
ExecuteGlobal fso.OpenTextFile(mockPath).ReadAll

Dim g_TestsPassed: g_TestsPassed = 0
Dim g_TestsFailed: g_TestsFailed = 0

Sub Assert(condition, description)
    If condition Then
        WScript.Echo "[PASS] " & description
        g_TestsPassed = g_TestsPassed + 1
    Else
        WScript.Echo "[FAIL] " & description
        g_TestsFailed = g_TestsFailed + 1
    End If
End Sub

' ------------------------------------------------------------------------------
' Test Case 1: Basic Read/Write
' ------------------------------------------------------------------------------
Sub TestBasic()
    WScript.Echo "--- Test Case: Basic Interaction ---"
    Dim mock: Set mock = New AdvancedMock
    mock.Connect "A"
    
    Dim buffer: buffer = String(24 * 80, " ")
    buffer = "COMMAND:" & Mid(buffer, 9)
    mock.SetBuffer buffer
    
    Dim content
    mock.ReadScreen content, 8, 1, 1
    Assert content = "COMMAND:", "Should read back the exact string set in buffer"
    
    mock.SendKey "123"
    Assert mock.GetSentKeys() = "123|", "Should track sent keys"
End Sub

' ------------------------------------------------------------------------------
' Test Case 2: Latency Simulation
' ------------------------------------------------------------------------------
Sub TestLatency()
    WScript.Echo "--- Test Case: Latency Simulation ---"
    Dim mock: Set mock = New AdvancedMock
    mock.Connect "A"
    mock.SetLatency 500 ' 500ms delay
    
    Dim startTime: startTime = Timer
    Dim content
    mock.ReadScreen content, 10, 1, 1
    Dim duration: duration = (Timer - startTime) * 1000
    
    Assert duration >= 450, "ReadScreen should take at least 500ms (measured: " & Round(duration) & "ms)"
End Sub

' ------------------------------------------------------------------------------
' Test Case 3: Partial Screen Load
' ------------------------------------------------------------------------------
Sub TestPartialLoad()
    WScript.Echo "--- Test Case: Partial Screen Load ---"
    Dim mock: Set mock = New AdvancedMock
    mock.Connect "A"
    
    ' Fill screen with 'X'
    mock.SetBuffer String(24 * 80, "X")
    
    ' Simulate only first 2 rows loaded
    mock.SetPartialLoad 2
    
    Dim content
    mock.ReadScreen content, 5, 1, 1
    Assert content = "XXXXX", "Row 1 should be visible"
    
    mock.ReadScreen content, 5, 3, 1
    Assert content = "     ", "Row 3 should be empty (partial load limit exceeded)"
End Sub

' ------------------------------------------------------------------------------
' Test Case 4: Interference Injection (Modals/Errors)
' ------------------------------------------------------------------------------
Sub TestInterference()
    WScript.Echo "--- Test Case: Interference Injection ---"
    Dim mock: Set mock = New AdvancedMock
    mock.Connect "A"
    mock.SetBuffer String(24 * 80, ".")
    
    ' Inject a modal error at row 23
    mock.InjectInterference "RO LOCKED BY USER", 23, 1
    
    Dim content
    mock.ReadScreen content, 17, 23, 1
    Assert content = "RO LOCKED BY USER", "Should read the injected interference text"
    
    mock.ReadScreen content, 5, 1, 1
    Assert content = ".....", "Rest of screen should remain unchanged"
End Sub

' --- Run Tests ---
TestBasic()
WScript.Echo ""
TestLatency()
WScript.Echo ""
TestPartialLoad()
WScript.Echo ""
TestInterference()

WScript.Echo ""
WScript.Echo "Summary: " & g_TestsPassed & " Passed, " & g_TestsFailed & " Failed"
If g_TestsFailed > 0 Then WScript.Quit 1
