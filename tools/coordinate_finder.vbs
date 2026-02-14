'=====================================================================================
' Coordinate Finder Utility
' Purpose: Captures a specific row and provides a column-indexed ruler to verify
'          exact positions of data fields in the terminal.
'=====================================================================================

Dim bzhao: Set bzhao = CreateObject("BZWhll.WhllObj")
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim rowToRead: rowToRead = 6 ' The row where Sequence 1 is expected

On Error Resume Next
bzhao.Connect ""
If Err.Number <> 0 Then
    MsgBox "Failed to connect to BlueZone terminal session.", vbCritical
Else
    On Error GoTo 0

    Dim line, header1, header2, header3, result, i
    bzhao.ReadScreen line, 80, rowToRead, 1

    ' Create a multi-line ruler
    header1 = "          1         2         3         4         5         6         7         8"
    header2 = "01234567890123456789012345678901234567890123456789012345678901234567890123456789"
    header3 = "--------------------------------------------------------------------------------"

    result = "TERMINAL COORDINATE CHECK (ROW " & rowToRead & ")" & vbCrLf & _
             "Generated: " & Now & vbCrLf & vbCrLf & _
             header1 & vbCrLf & _
             header2 & vbCrLf & _
             header3 & vbCrLf & _
             line & vbCrLf & _
             header3

    ' Generate output file in current directory
    Dim currentPath, filePath
    currentPath = fso.GetParentFolderName(WScript.ScriptFullName)
    filePath = fso.BuildPath(currentPath, "coordinate_check.txt")
    
    Set ts = fso.CreateTextFile(filePath, True)
    ts.Write result
    ts.Close

    MsgBox "Terminal coordinates captured for Row " & rowToRead & "." & vbCrLf & _
           "Please check: " & filePath, vbInformation, "Capture Complete"
End If
