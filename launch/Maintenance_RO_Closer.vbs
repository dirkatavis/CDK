Option Explicit

' WRAPPER_TARGET: apps\maintenance_ro_closer\Maintenance_RO_Closer.vbs
' Legacy launch compatibility wrapper - REMOVABLE at sunset

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim sh: Set sh = CreateObject("WScript.Shell")
Dim basePath: basePath = sh.Environment("USER")("CDK_BASE")

If basePath = "" Or Not fso.FolderExists(basePath) Then
    Err.Raise 53, "Wrapper", "Invalid or missing CDK_BASE. Run: cscript tooling\setup_cdk_base.vbs"
End If

If Not fso.FileExists(fso.BuildPath(basePath, ".cdkroot")) Then
    Err.Raise 53, "Wrapper", "Cannot find .cdkroot in base path:" & vbCrLf & basePath
End If

Dim targetPath: targetPath = fso.BuildPath(basePath, "apps\maintenance_ro_closer\Maintenance_RO_Closer.vbs")
If Not fso.FileExists(targetPath) Then
    Err.Raise 53, "Wrapper", "Wrapper target not found:" & vbCrLf & targetPath
End If

ExecuteGlobal fso.OpenTextFile(targetPath).ReadAll()
