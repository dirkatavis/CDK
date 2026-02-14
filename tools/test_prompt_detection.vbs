'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestPromptDetection
' **DATE CREATED:** 2025-11-18
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Standalone test script for debugging prompt detection timing issues.
' Tests various prompt patterns and captures screen content for analysis.
'-----------------------------------------------------------------------------------

Option Explicit

' Include CommonLib for testing
Function IncludeFile(filePath)
    On Error Resume Next
    Dim fsoInclude, fileContent, includeStream

    Set fsoInclude = CreateObject("Scripting.FileSystemObject")

    If Not fsoInclude.FileExists(filePath) Then
        MsgBox "IncludeFile - File not found: " & filePath
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

' Test prompts that commonly cause issues
Sub TestCommonPrompts()
    Dim testPrompts, i, result, screenContent

    testPrompts = Array( _
        "TECHNICIAN (", _
        "TECHNICIAN", _
        "TECHNICIAN ", _
        "ACTUAL HOURS (", _
        "ACTUAL HOURS", _
        "ACTUAL HOURS ", _
        "SOLD HOURS", _
        "SOLD HOURS?", _
        "SOLD HOURS " _
    )

    MsgBox "Testing prompt detection. Make sure BlueZone is running and logged in."

    For i = LBound(testPrompts) To UBound(testPrompts)
        screenContent = GetScreenSnapshot(3)
        result = IsTextPresent(testPrompts(i))

        MsgBox "Testing: '" & testPrompts(i) & "'" & vbCrLf & _
               "Found: " & result & vbCrLf & vbCrLf & _
               "Screen content:" & vbCrLf & screenContent
    Next
End Sub

' Test optional prompt handling
Sub TestOptionalPrompts()
    Dim optionalPrompts, i, result, screenContent

    optionalPrompts = Array( _
        "ACTUAL HOURS (", _
        "SOLD HOURS", _
        "TECHNICIAN (" _
    )

    MsgBox "Testing optional prompt handling (these may or may not appear)."

    For i = LBound(optionalPrompts) To UBound(optionalPrompts)
        screenContent = GetScreenSnapshot(3)
        result = WaitForPrompt(optionalPrompts(i), "", False, 2000, "") ' Short timeout for testing

        MsgBox "Testing optional prompt: '" & optionalPrompts(i) & "'" & vbCrLf & _
               "Appeared: " & result & vbCrLf & vbCrLf & _
               "Screen content:" & vbCrLf & screenContent
    Next
End Sub

' Main test routine
Sub Main()
    If Not IncludeFile("../lib/CommonLib.vbs") Then
        MsgBox "Failed to load CommonLib.vbs"
        Exit Sub
    End If

    ' Connect to BlueZone
    If Not ConnectBlueZone() Then
        MsgBox "Failed to connect to BlueZone"
        Exit Sub
    End If

    TestCommonPrompts()
    TestOptionalPrompts()

    MsgBox "Testing complete. Check results above."
End Sub

' Connect to BlueZone (simplified version)
Function ConnectBlueZone()
    On Error Resume Next
    Set bzhao = CreateObject("BZWhll.WhllObj")
    If Err.Number <> 0 Then
        MsgBox "Failed to create BZWhll.WhllObj: " & Err.Description
        Err.Clear
        ConnectBlueZone = False
        Exit Function
    End If

    bzhao.Connect "A"
    If Err.Number <> 0 Then
        MsgBox "Failed to connect to session A: " & Err.Description
        Err.Clear
        ConnectBlueZone = False
        Exit Function
    End If

    On Error GoTo 0
    ConnectBlueZone = True
End Function

Main()</content>
<parameter name="filePath">c:\Temp\Code\Scripts\VBScript\CDK\PostFinalCharges\test_prompt_detection.vbs