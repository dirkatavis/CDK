'=====================================================================================
' Coordinate Finder Utility
' Purpose: Captures a specific row and provides a column-indexed ruler to verify
'          exact positions of data fields in the terminal.
'=====================================================================================

Dim bzhao: Set bzhao = CreateObject("BZWhll.WhllObj")
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim rowToRead: rowToRead = 4 ' The row where OPEN DATE is expected

On Error Resume Next
bzhao.Connect ""
If Err.Number <> 0 Then
    MsgBox "Failed to connect to BlueZone terminal session.", vbCritical
Else
' --- Load PathHelper for centralized path management ---
' This bootstrap handles the "No WScript" host environment in BlueZone
Dim sh_bs, base_bs, helper_path_bs
Set sh_bs = CreateObject("WScript.Shell")
base_bs = sh_bs.ExpandEnvironmentStrings("%CDK_BASE%")
If base_bs = "%CDK_BASE%" Then base_bs = sh_bs.Environment("USER")("CDK_BASE")

If base_bs = "" Then
    MsgBox "CDK_BASE environment variable not set. Please run setup_cdk_base.vbs.", 16, "Error"
Else
    helper_path_bs = fso.BuildPath(base_bs, "common\PathHelper.vbs")
    ExecuteGlobal fso.OpenTextFile(helper_path_bs).ReadAll

    On Error GoTo 0

    Dim line, header1, header2, header3, result, i
    bzhao.ReadScreen line, 80, rowToRead, 1

    ' Create a multi-line ruler
    header1 = "          1         2         3         4         5         6         7         8"
    header2 = "01234567890123456789012345678901234567890123456789012345678901234567890123456789"
    header3 = "--------------------------------------------------------------------------------"

    result = "TERMINAL COORDINATE CHECK (HEADER ROWS 1-5)" & vbCrLf & _
             "Generated: " & Now & vbCrLf & vbCrLf & _
             "     " & header1 & vbCrLf & _
             "     " & header2 & vbCrLf & _
             "     " & header3 & vbCrLf

    Dim row
    For row = 1 To 5
        bzhao.ReadScreen line, 80, row, 1
        result = result & "R" & row & " | " & line & vbCrLf
    Next
    
    result = result & "     " & header3

    ' Generate output file using PathHelper and config.ini
    Dim filePath
    filePath = GetConfigPath("Coordinate_Finder", "Output")
    
    Dim ts
    Set ts = fso.CreateTextFile(filePath, True)
    ts.Write result
    ts.Close

    MsgBox "Terminal coordinates captured for Row " & rowToRead & "." & vbCrLf & _
           "Path: " & filePath, vbInformation, "Capture Complete"
End If
End If
