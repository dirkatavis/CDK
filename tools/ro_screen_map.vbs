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

    ' --- Load PathHelper for centralized path management ---
    Dim shell, base_path, helper_path
    Set shell = CreateObject("WScript.Shell")
    base_path = shell.ExpandEnvironmentStrings("%CDK_BASE%")
    If base_path = "%CDK_BASE%" Then base_path = shell.Environment("USER")("CDK_BASE")

    If base_path = "" Or Not fso.FolderExists(base_path) Then
        MsgBox "ERROR: CDK_BASE environment variable is missing or invalid." & vbCrLf & _
               "Please run tools\setup_cdk_base.vbs first.", 16, "Path Configuration Error"
        Exit Sub
    End If

    ' Load common library
    helper_path = fso.BuildPath(base_path, "framework\PathHelper.vbs")
    If Not fso.FileExists(helper_path) Then
        MsgBox "ERROR: Cannot find PathHelper.vbs at: " & helper_path, 16, "Validation Error"
        Exit Sub
    End If
    ExecuteGlobal fso.OpenTextFile(helper_path).ReadAll

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

    ' Save to file using configured path
    Dim filePath
    On Error Resume Next
    filePath = GetConfigPath("Coordinate_Finder", "Output")
    ' Swap the default filename for the mapper filename
    filePath = Replace(filePath, "coordinate_check.txt", "ro_screen_map.txt")
    On Error GoTo 0

    If filePath = "" Then filePath = fso.BuildPath(base_path, "tools\ro_screen_map.txt")

    Set ts = fso.CreateTextFile(filePath, True)
    ts.Write result
    ts.Close

    MsgBox "RO Screen Map captured to: " & vbCrLf & filePath, vbInformation
End Sub

' Run it
MapScreen
