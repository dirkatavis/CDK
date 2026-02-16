'=====================================================================================
' RO Screen Mapper
' Purpose: Captures rows 1-23 of the current screen with a column ruler.
'          Use this while looking at an RO to find exact coordinates for fields.
'=====================================================================================

Dim bzhao: Set bzhao = CreateObject("BZWhll.WhllObj")
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")

Sub MapScreen()
    On Error Resume Next
    bzhao.Connect ""
    If Err.Number <> 0 Then
        MsgBox "Failed to connect to BlueZone terminal session.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    Dim result, row, line
    Dim header1, header2, header3

    ' Create ruler headers
    header1 = "          1         2         3         4         5         6         7         8"
    header2 = "01234567890123456789012345678901234567890123456789012345678901234567890123456789"
    header3 = "--------------------------------------------------------------------------------"

    result = "RO SCREEN COORDINATE MAP" & vbCrLf & _
             "Generated: " & Now & vbCrLf & vbCrLf & _
             "      " & header1 & vbCrLf & _
             "      " & header2 & vbCrLf & _
             "      " & header3 & vbCrLf

    For row = 1 To 23
        bzhao.ReadScreen line, 80, row, 1
        ' Pad row number for alignment
        Dim rowLabel
        rowLabel = row
        If row < 10 Then rowLabel = "0" & row
        
        result = result & rowLabel & " | " & line & vbCrLf
    Next

    result = result & "      " & header3

    ' --- System Standard Path Discovery ---
    ' Resolve path using CDK_BASE to ensure output lands in the repo
    Dim shell, base_path, full_path
    Set shell = CreateObject("WScript.Shell")
    base_path = shell.ExpandEnvironmentStrings("%CDK_BASE%")
    If base_path = "%CDK_BASE%" Then base_path = shell.Environment("USER")("CDK_BASE")

    If base_path <> "" And fso.FolderExists(base_path) Then
        ' Check for .cdkroot to be sure
        If fso.FileExists(fso.BuildPath(base_path, ".cdkroot")) Then
            ' Try to use the Coordinate_Finder path from config.ini but with our filename
            ' For simplicity in this standalone tool, we anchor to the tools folder
            full_path = fso.BuildPath(base_path, "tools\ro_screen_map.txt")
        Else
            full_path = "ro_screen_map.txt" ' Fallback to CWD
        End If
    Else
        full_path = "ro_screen_map.txt" ' Fallback to CWD
    End If

    Set ts = fso.CreateTextFile(full_path, True)
    ts.Write result
    ts.Close

    MsgBox "RO Screen Map captured to: " & vbCrLf & full_path, vbInformation
End Sub

' Run it
MapScreen
