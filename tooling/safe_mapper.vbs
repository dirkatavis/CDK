'=====================================================================================
' Safe Screen Mapper
' Purpose: Captures rows 1-23 of the current screen with a column ruler.
' Note: Uses WScript.Shell for environment path discovery.
'=====================================================================================

Dim bzhao_safe, fso_safe, ts_safe
Dim result_safe, row_safe, line_safe
Dim h1, h2, h3

Set bzhao_safe = CreateObject("BZWhll.WhllObj")
Set fso_safe = CreateObject("Scripting.FileSystemObject")

On Error Resume Next
bzhao_safe.Connect ""
If Err.Number <> 0 Then
    MsgBox "Failed to connect to BlueZone terminal session.", 16, "Error"
Else
    On Error GoTo 0

    ' Create ruler headers
    h1 = "          1         2         3         4         5         6         7         8"
    h2 = "01234567890123456789012345678901234567890123456789012345678901234567890123456789"
    h3 = "--------------------------------------------------------------------------------"

    result_safe = "RO SCREEN COORDINATE MAP" & vbCrLf & _
                  "Generated: " & Now & vbCrLf & vbCrLf & _
                  "      " & h1 & vbCrLf & _
                  "      " & h2 & vbCrLf & _
                  "      " & h3 & vbCrLf

    For row_safe = 1 To 23
        bzhao_safe.ReadScreen line_safe, 80, row_safe, 1
        ' Pad row number for alignment
        If row_safe < 10 Then
            result_safe = result_safe & "0" & row_safe & " | " & line_safe & vbCrLf
        Else
            result_safe = result_safe & row_safe & " | " & line_safe & vbCrLf
        End If
    Next

    result_safe = result_safe & "      " & h3

    ' --- System Standard Path Discovery ---
    ' This replicates the logic in PathHelper.vbs in a host-safe way
    Dim shell_safe, fso_safe, base_path, rel_path, full_path
    Set shell_safe = CreateObject("WScript.Shell")
    Set fso_safe = CreateObject("Scripting.FileSystemObject")
    
    ' 1. Get Repo Root (CDK_BASE)
    base_path = shell_safe.ExpandEnvironmentStrings("%CDK_BASE%")
    If base_path = "%CDK_BASE%" Then 
        base_path = shell_safe.Environment("USER")("CDK_BASE")
    End If
    
    ' Fail fast if CDK_BASE is missing or invalid (no fallbacks)
    If base_path = "" Or Not fso_safe.FolderExists(base_path) Then
        MsgBox "ERROR: CDK_BASE environment variable is missing or invalid." & vbCrLf & _
               "Value: " & base_path & vbCrLf & vbCrLf & _
               "Please run tools\setup_cdk_base.vbs first.", 16, "Path Configuration Error"
        Exit Sub
    End If

    ' Validate .cdkroot marker existence
    If Not fso_safe.FileExists(fso_safe.BuildPath(base_path, ".cdkroot")) Then
        MsgBox "ERROR: Missing .cdkroot marker in base path:" & vbCrLf & base_path & vbCrLf & vbCrLf & _
               "Ensure CDK_BASE points to the repository root.", 16, "Validation Error"
        Exit Sub
    End If

    ' 2. Read Output Path from config.ini (Coordinate_Finder section)
    ' This ensures we follow the project's folder structure
    Dim ini_path
    ini_path = fso_safe.BuildPath(base_path, "config\config.ini")
    rel_path = "tools\ro_screen_map.txt" ' Default
    
    ' Simple INI read for Coordinate_Finder:Output
    If fso_safe.FileExists(ini_path) Then
        Dim ts_ini, line_ini
        Set ts_ini = fso_safe.OpenTextFile(ini_path, 1)
        Do Until ts_ini.AtEndOfStream
            line_ini = Trim(ts_ini.ReadLine)
            If InStr(1, line_ini, "Output=", 1) > 0 Then
                rel_path = Trim(Split(line_ini, "=")(1))
                ' We'll change filename slightly so as not to overwrite core check
                rel_path = Replace(rel_path, "coordinate_check.txt", "ro_screen_map.txt")
                Exit Do
            End If
        Loop
        ts_ini.Close
    End If

    full_path = fso_safe.BuildPath(base_path, rel_path)

    ' 3. Save to file
    Dim ts_out
    Set ts_out = fso_safe.CreateTextFile(full_path, True)
    ts_out.Write result_safe
    ts_out.Close

    MsgBox "RO Screen Map captured!" & vbCrLf & _
           "Path: " & full_path, 64, "Success"
End If
