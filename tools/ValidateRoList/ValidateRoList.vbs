Option Explicit

' WRAPPER_TARGET: tools\validate_ro_list\ValidateRoList.vbs

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim sh: Set sh = CreateObject("WScript.Shell")
Dim basePath: basePath = sh.Environment("USER")("CDK_BASE")

If basePath = "" Or Not fso.FolderExists(basePath) Then
    Err.Raise 53, "Wrapper", "Invalid or missing CDK_BASE. Value: " & basePath
End If

If Not fso.FileExists(fso.BuildPath(basePath, ".cdkroot")) Then
    Err.Raise 53, "Wrapper", "Cannot find .cdkroot in base path:" & vbCrLf & basePath
End If

Dim targetPath: targetPath = fso.BuildPath(basePath, "tools\validate_ro_list\ValidateRoList.vbs")
If Not fso.FileExists(targetPath) Then
    Err.Raise 53, "Wrapper", "Wrapper target not found:" & vbCrLf & targetPath
End If

ExecuteGlobal fso.OpenTextFile(targetPath).ReadAll
