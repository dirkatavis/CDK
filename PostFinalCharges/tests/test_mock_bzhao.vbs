'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestMockBzhao
' **DATE CREATED:** 2025-11-19
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Simple test script to demonstrate MockBzhao functionality.
' Tests basic screen interactions without requiring BlueZone.
'-----------------------------------------------------------------------------------

Option Explicit

' Include MockBzhao for testing
Function IncludeFile(filePath)
    On Error Resume Next
    Dim fsoInclude, fileContent, includeStream

    Set fsoInclude = CreateObject("Scripting.FileSystemObject")

    If Not fsoInclude.FileExists(filePath) Then
        WScript.Echo "IncludeFile - File not found: " & filePath
        IncludeFile = False
        Exit Function
    End If

    Set includeStream = fsoInclude.OpenTextFile(filePath, 1)
    fileContent = includeStream.ReadAll
    includeStream.Close
    Set includeStream = Nothing

    ExecuteGlobal fileContent
    IncludeFile = True
End Function

Sub TestMockBzhao()
    WScript.Echo "Testing MockBzhao..."
    
    ' Load the mock
    If Not IncludeFile("../mocks/MockBzhao.vbs") Then
        WScript.Echo "Failed to load MockBzhao.vbs"
        Exit Sub
    End If
    
    ' Create mock instance
    Dim mock
    Set mock = New MockBzhao
    
    ' Test connection
    mock.Connect("")
    WScript.Echo "Connected: " & mock.IsConnected()
    
    ' Test setting up a scenario
    mock.SetupTestScenario("basic_command_prompt")
    
    ' Test reading screen
    Dim buffer
    mock.ReadScreen buffer, 10, 1, 1
    WScript.Echo "Screen content (first 10 chars): '" & buffer & "'"
    
    ' Test sending keys
    mock.SendKey("TEST")
    mock.SendKey("<NumpadEnter>")
    WScript.Echo "Sent keys: '" & mock.GetSentKeys() & "'"
    
    ' Test another scenario
    mock.SetupTestScenario("ro_ready_to_post")
    mock.ReadScreen buffer, 25, 1, 1
    WScript.Echo "Screen content after scenario: '" & buffer & "'"
    
    WScript.Echo "MockBzhao test completed successfully!"
End Sub

' Run the test
TestMockBzhao