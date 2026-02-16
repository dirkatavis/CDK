'=====================================================================================
' Test PFC Scrapper Extraction
' Purpose: Verify data extraction logic using mock screen buffers.
'=====================================================================================

Option Explicit

' Mock BZWhll.WhllObj
Class MockBzhao
    Public Buffer
    Public Sub ReadScreen(ByRef content, length, row, col)
        Dim startPos: startPos = ((row-1) * 80) + (col-1) + 1
        content = Mid(Buffer, startPos, length)
    End Sub
    Public Sub Pause(ms) : End Sub
    Public Sub SendKey(key) : End Sub
End Class

Dim bzhao: Set bzhao = New MockBzhao
Dim screenGrid(23)

Sub SetRow(rowIdx, content)
    screenGrid(rowIdx) = Left(content & String(80, " "), 80)
End Sub

' Functions copied from PFC_Scrapper.vbs for testing
Function GetROFromScreen()
    Dim buf, re, matches
    bzhao.ReadScreen buf, 240, 1, 1 ' Read top 3 lines
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "RO:\s*(\d{4,})"
    re.IgnoreCase = True
    If re.Test(buf) Then
        Set matches = re.Execute(buf)
        GetROFromScreen = Trim(matches(0).SubMatches(0))
    Else
        GetROFromScreen = "UNKNOWN"
    End If
End Function

Function GetOpenDateFromScreen()
    Dim buf, re, matches
    bzhao.ReadScreen buf, 240, 1, 1 ' Read top 3 lines
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(?:DATE|OPN):\s*([\d/]{6,10})"
    re.IgnoreCase = True
    If re.Test(buf) Then
        Set matches = re.Execute(buf)
        GetOpenDateFromScreen = Trim(matches(0).SubMatches(0))
    Else
        GetOpenDateFromScreen = "UNKNOWN"
    End If
End Function

Function GetRepairOrderStatus()
    Dim buf, prefix, pos
    bzhao.ReadScreen buf, 40, 5, 1 ' Read row 5
    prefix = "RO STATUS: "
    pos = InStr(1, buf, prefix, vbTextCompare)
    If pos > 0 Then
        GetRepairOrderStatus = Trim(Mid(buf, pos + Len(prefix), 15))
    Else
        GetRepairOrderStatus = "UNKNOWN"
    End If
End Function

Function GetLineDescription(letter)
    Dim row, buf
    GetLineDescription = ""
    For row = 7 To 22
        bzhao.ReadScreen buf, 1, row, 1
        If UCase(Trim(buf)) = UCase(letter) Then
            bzhao.ReadScreen buf, 25, row, 7
            GetLineDescription = Trim(buf)
            Exit Function
        End If
    Next
End Function

Sub TestExtraction()
    WScript.Echo "Starting Extraction Tests..."
    
    Dim i
    For i = 0 To 23
        screenGrid(i) = String(80, " ")
    Next
    
    SetRow 0, "   RO: 123456                        DATE: 01/20/26"
    SetRow 4, "RO STATUS: READY TO POST"
    SetRow 9, "A      OIL CHANGE"
    SetRow 11, "C      TIRE ROTATION"
    
    bzhao.Buffer = Join(screenGrid, "")
    
    ' Test RO Number
    Dim roNum: roNum = GetROFromScreen()
    WScript.Echo "Test RO Number: " & roNum
    If roNum = "123456" Then WScript.Echo "[PASS]" Else WScript.Echo "[FAIL]"
    
    ' Test Date
    Dim openDate: openDate = GetOpenDateFromScreen()
    WScript.Echo "Test Open Date: " & openDate
    If openDate = "01/20/26" Then WScript.Echo "[PASS]" Else WScript.Echo "[FAIL]"
    
    ' Test Status
    Dim roStatus: roStatus = GetRepairOrderStatus()
    WScript.Echo "Test Status: '" & roStatus & "'"
    if roStatus = "READY TO POST" Then WScript.Echo "[PASS]" Else WScript.Echo "[FAIL]"
    
    ' Test Line A
    Dim lineA: lineA = GetLineDescription("A")
    WScript.Echo "Test Line A: '" & lineA & "'"
    If lineA = "OIL CHANGE" Then WScript.Echo "[PASS]" Else WScript.Echo "[FAIL]"
    
    ' Test Line B (Missing)
    Dim lineB: lineB = GetLineDescription("B")
    WScript.Echo "Test Line B: '" & lineB & "'"
    If lineB = "" Then WScript.Echo "[PASS]" Else WScript.Echo "[FAIL]"
    
    ' Test Line C
    Dim lineC: lineC = GetLineDescription("C")
    WScript.Echo "Test Line C: '" & lineC & "'"
    If lineC = "TIRE ROTATION" Then WScript.Echo "[PASS]" Else WScript.Echo "[FAIL]"
    
    WScript.Echo "Tests Completed."
End Sub

TestExtraction
