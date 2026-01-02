'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestOpenStatus
' **DATE CREATED:** 2025-12-30
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Test script for OPEN status closeout functionality.
' Validates that OPEN status is properly recognized and handled.
'-----------------------------------------------------------------------------------

Option Explicit

' Test framework setup
Dim g_TestsPassed, g_TestsFailed, g_TotalTests

' Include CommonLib for testing
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

Sub InitializeTests()
    g_TestsPassed = 0
    g_TestsFailed = 0
    g_TotalTests = 0
End Sub

Sub AssertEqual(actual, expected, testName)
    g_TotalTests = g_TotalTests + 1
    If actual = expected Then
        g_TestsPassed = g_TestsPassed + 1
        WScript.Echo "PASS: " & testName
    Else
        g_TestsFailed = g_TestsFailed + 1
        WScript.Echo "FAIL: " & testName & " (Expected: " & expected & ", Actual: " & actual & ")"
    End If
End Sub

Sub AssertTrue(condition, testName)
    g_TotalTests = g_TotalTests + 1
    If condition Then
        g_TestsPassed = g_TestsPassed + 1
        WScript.Echo "PASS: " & testName
    Else
        g_TestsFailed = g_TestsFailed + 1
        WScript.Echo "FAIL: " & testName
    End If
End Sub

Sub PrintTestSummary()
    WScript.Echo ""
    WScript.Echo "Test Summary:"
    WScript.Echo "Total Tests: " & g_TotalTests
    WScript.Echo "Passed: " & g_TestsPassed
    WScript.Echo "Failed: " & g_TestsFailed
    If g_TestsFailed = 0 Then
        WScript.Echo "ALL TESTS PASSED!"
    Else
        WScript.Echo "SOME TESTS FAILED!"
    End If
End Sub

' Test the ValidCloseoutStatuses configuration reading
Sub TestValidCloseoutStatuses()
    WScript.Echo "Testing ValidCloseoutStatuses configuration..."
    
    ' Load the main script (we'll mock some dependencies)
    If Not IncludeFile("../PostFinalCharges.vbs") Then
        WScript.Echo "Failed to load PostFinalCharges.vbs"
        Exit Sub
    End If
    
    ' Mock the GetIniSetting function to return our test configuration
    ExecuteGlobal "Function GetIniSetting(section, key, defaultValue)" & vbCrLf & _
                  "    If section = ""Processing"" And key = ""ValidCloseoutStatuses"" Then" & vbCrLf & _
                  "        GetIniSetting = ""READY TO POST,PREASSIGNED,OPENED""" & vbCrLf & _
                  "    Else" & vbCrLf & _
                  "        GetIniSetting = defaultValue" & vbCrLf & _
                  "    End If" & vbCrLf & _
                  "End Function"
    
    ' Test that GetValidCloseoutStatuses returns the correct array
    Dim validStatuses
    validStatuses = GetValidCloseoutStatuses()
    
    AssertEqual UBound(validStatuses) + 1, 3, "ValidCloseoutStatuses should return 3 statuses"
    AssertEqual validStatuses(0), "READY TO POST", "First status should be READY TO POST"
    AssertEqual validStatuses(1), "PREASSIGNED", "Second status should be PREASSIGNED"
    AssertEqual validStatuses(2), "OPENED", "Third status should be OPENED"
End Sub

' Test the IsValidCloseoutStatus function
Sub TestIsValidCloseoutStatus()
    WScript.Echo "Testing IsValidCloseoutStatus function..."
    
    ' Mock the GetIniSetting function
    ExecuteGlobal "Function GetIniSetting(section, key, defaultValue)" & vbCrLf & _
                  "    If section = ""Processing"" And key = ""ValidCloseoutStatuses"" Then" & vbCrLf & _
                  "        GetIniSetting = ""READY TO POST,PREASSIGNED,OPENED""" & vbCrLf & _
                  "    Else" & vbCrLf & _
                  "        GetIniSetting = defaultValue" & vbCrLf & _
                  "    End If" & vbCrLf & _
                  "End Function"
    
    ' Test valid statuses
    AssertTrue IsValidCloseoutStatus("READY TO POST"), "READY TO POST should be valid"
    AssertTrue IsValidCloseoutStatus("PREASSIGNED"), "PREASSIGNED should be valid"
    AssertTrue IsValidCloseoutStatus("OPENED"), "OPENED should be valid"
    
    ' Test case insensitive matching
    AssertTrue IsValidCloseoutStatus("ready to post"), "ready to post (lowercase) should be valid"
    AssertTrue IsValidCloseoutStatus("Opened"), "Opened (mixed case) should be valid"
    
    ' Test invalid statuses
    AssertTrue Not IsValidCloseoutStatus("CLOSED"), "CLOSED should not be valid"
    AssertTrue Not IsValidCloseoutStatus("INVALID"), "INVALID should not be valid"
    AssertTrue Not IsValidCloseoutStatus(""), "Empty string should not be valid"
End Sub

' Main test execution
Sub RunAllTests()
    InitializeTests()
    
    WScript.Echo "Starting OPEN Status Tests..."
    WScript.Echo "=============================="
    
    TestValidCloseoutStatuses()
    TestIsValidCloseoutStatus()
    
    WScript.Echo "=============================="
    PrintTestSummary()
End Sub

' Execute tests
RunAllTests()