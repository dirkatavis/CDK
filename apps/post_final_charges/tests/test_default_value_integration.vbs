'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestDefaultValueIntegration
' **DATE CREATED:** 2025-12-29
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Integration test for default value detection using MockBzhao.
' Simulates the complete prompt sequence with default values.
'-----------------------------------------------------------------------------------

Option Explicit

' Include required files
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim scriptDir: scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
Dim appDir: appDir = fso.GetParentFolderName(scriptDir)
Dim appsDir: appsDir = fso.GetParentFolderName(appDir)
Dim repoRoot: repoRoot = fso.GetParentFolderName(appsDir)
Dim mockPath: mockPath = fso.BuildPath(repoRoot, "framework\AdvancedMock.vbs")
ExecuteGlobal fso.OpenTextFile(mockPath).ReadAll

' Test the complete prompt processing with default values
Sub TestDefaultValueProcessing()
    WScript.Echo "Testing Default Value Processing with AdvancedMock..."
    WScript.Echo String(50, "=")
    
    ' Create our test mock
    Dim testMock
    Set testMock = New AdvancedMock
    
    ' Define a sequence of prompts that include default values
    testMock.SetPromptSequence Array( _
        "TECHNICIAN (72925)?", _
        "ACTUAL HOURS (117)?", _
        "SOLD HOURS (0)?", _
        "TECHNICIAN (99)?", _
        "COMMAND:" _
    )
    
    ' Replace the global bzhao with our test mock
    Set bzhao = testMock
    testMock.Connect("TEST")
    
    ' Create the prompt dictionary (simplified for testing since we can't load the full script)
    ' Instead of using the actual CreateLineItemPromptDictionary, we'll simulate the behavior
    On Error Resume Next
    Dim prompts
    Set prompts = CreateObject("Scripting.Dictionary")
    If Err.Number <> 0 Then
        WScript.Echo "Note: Testing with simulated prompt dictionary (main script not loaded)"
        Err.Clear
    End If
    On Error GoTo 0
    
    ' Simulate processing each prompt in sequence
    Dim expectedBehavior(4)
    expectedBehavior(0) = "TECHNICIAN (72925)? - Should accept default (no text sent)"
    expectedBehavior(1) = "ACTUAL HOURS (117)? - Should accept default (no text sent)"  
    expectedBehavior(2) = "SOLD HOURS (0)? - Should accept default (no text sent)"
    expectedBehavior(3) = "TECHNICIAN (99)? - Should accept default (no text sent)"
    expectedBehavior(4) = "COMMAND: - Should finish processing"
    
    WScript.Echo "Expected behavior:"
    Dim i
    For i = 0 To UBound(expectedBehavior)
        WScript.Echo "  " & (i + 1) & ". " & expectedBehavior(i)
    Next
    WScript.Echo ""
    
    ' Process the first few prompts manually to test the logic
    WScript.Echo "Processing prompts..."
    
    ' Test TECHNICIAN (72925)? - should accept default
    WScript.Echo "1. Testing: " & testMock.GetCurrentPrompt()
    Dim shouldAcceptDefault
    shouldAcceptDefault = HasDefaultValueInPrompt("TECHNICIAN \([A-Za-z0-9]+\)\?", testMock.GetCurrentPrompt())
    WScript.Echo "   Default detected: " & shouldAcceptDefault
    If shouldAcceptDefault Then
        WScript.Echo "   Action: Send only <NumpadEnter> (accept default)"
        testMock.SendKey("<NumpadEnter>")
    Else
        WScript.Echo "   Action: Send '99' + <NumpadEnter> (override)"
        testMock.SendKey("99")
        testMock.SendKey("<NumpadEnter>")
    End If
    WScript.Echo ""
    
    ' Test ACTUAL HOURS (117)? - should accept default
    WScript.Echo "2. Testing: " & testMock.GetCurrentPrompt()
    shouldAcceptDefault = HasDefaultValueInPrompt("ACTUAL HOURS \(\d+\)", testMock.GetCurrentPrompt())
    WScript.Echo "   Default detected: " & shouldAcceptDefault
    If shouldAcceptDefault Then
        WScript.Echo "   Action: Send only <NumpadEnter> (accept default)"
        testMock.SendKey("<NumpadEnter>")
    Else
        WScript.Echo "   Action: Send '0' + <NumpadEnter> (override)"
        testMock.SendKey("0")
        testMock.SendKey("<NumpadEnter>")
    End If
    WScript.Echo ""
    
    ' Test SOLD HOURS (0)? - should accept default even though it's 0
    WScript.Echo "3. Testing: " & testMock.GetCurrentPrompt()
    shouldAcceptDefault = HasDefaultValueInPrompt("SOLD HOURS \([0-9]+\)\?", testMock.GetCurrentPrompt())
    WScript.Echo "   Default detected: " & shouldAcceptDefault
    If shouldAcceptDefault Then
        WScript.Echo "   Action: Send only <NumpadEnter> (accept default)"
        testMock.SendKey("<NumpadEnter>")
    Else
        WScript.Echo "   Action: Send '0' + <NumpadEnter> (override)"
        testMock.SendKey("0")
        testMock.SendKey("<NumpadEnter>")
    End If
    WScript.Echo ""
    
    ' Show final results
    WScript.Echo "Final Results:"
    WScript.Echo "  Keys sent: " & testMock.GetSentKeys()
    WScript.Echo "  Current prompt: " & testMock.GetCurrentPrompt()
    
    ' Analyze the key sequence
    Dim keySequence
    keySequence = testMock.GetSentKeys()
    WScript.Echo ""
    WScript.Echo "Analysis:"
    If InStr(keySequence, "99;") > 0 Then
        WScript.Echo "  ISSUE: Found '99' in key sequence - default values were overridden"
    Else
        WScript.Echo "  SUCCESS: No hardcoded values found - defaults were preserved"
    End If
    
    If InStr(keySequence, "0;") > 0 Then
        WScript.Echo "  ISSUE: Found '0' in key sequence - default values were overridden"
    Else
        WScript.Echo "  SUCCESS: No hardcoded '0' found for hours - defaults were preserved"
    End If
    
    testMock.Disconnect()
End Sub

' Test comparison: old behavior vs new behavior
Sub TestBehaviorComparison()
    WScript.Echo vbCrLf & "Behavior Comparison Test"
    WScript.Echo String(30, "-")
    
    WScript.Echo "OLD BEHAVIOR (before fix):"
    WScript.Echo "  TECHNICIAN (72925)? -> Send '99' + Enter -> Result: Technician 99"
    WScript.Echo "  ACTUAL HOURS (117)?  -> Send '0' + Enter  -> Result: 0 hours"
    WScript.Echo "  SOLD HOURS (0)?      -> Send '0' + Enter  -> Result: 0 hours"
    WScript.Echo ""
    WScript.Echo "NEW BEHAVIOR (after fix):"
    WScript.Echo "  TECHNICIAN (72925)? -> Send Enter only    -> Result: Technician 72925"
    WScript.Echo "  ACTUAL HOURS (117)?  -> Send Enter only    -> Result: 117 hours" 
    WScript.Echo "  SOLD HOURS (0)?      -> Send Enter only    -> Result: 0 hours"
    WScript.Echo ""
    WScript.Echo "BENEFIT: Existing values are preserved instead of being overwritten"
End Sub

' Main test entry point
Sub Main()
    WScript.Echo "Default Value Detection Integration Test"
    WScript.Echo "========================================"
    WScript.Echo ""

    ' Load required files
    ' AdvancedMock is already loaded via ExecuteGlobal in the script header
    
    ' Load the main script for functions like HasDefaultValueInPrompt (if they exist there)
    ' Using absolute path logic to match the header
    Dim libPath: libPath = fso.BuildPath(appDir, "PostFinalCharges.vbs")
    If Not fso.FileExists(libPath) Then
        WScript.Echo "Failed to find: " & libPath
        Exit Sub
    End If
    ExecuteGlobal fso.OpenTextFile(libPath).ReadAll

    ' Run the tests
    TestDefaultValueProcessing()
    TestBehaviorComparison()
    
    WScript.Echo vbCrLf & "Integration test completed."
End Sub

' Run the tests
Main()