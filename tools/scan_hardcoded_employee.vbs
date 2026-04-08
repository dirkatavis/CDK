Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim sh: Set sh = CreateObject("WScript.Shell")
Dim baseDir: baseDir = sh.Environment("USER")("CDK_BASE")

ScanFolder baseDir

Sub ScanFolder(folderPath)
    Dim folder: Set folder = fso.GetFolder(folderPath)
    Dim file, ts, lineNum, line
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "vbs" Then
            Set ts = fso.OpenTextFile(file.Path, 1)
            lineNum = 0
            Do Until ts.AtEndOfStream
                lineNum = lineNum + 1
                line = ts.ReadLine
                If InStr(line, "18351") > 0 Then
                    WScript.Echo file.Path & " (" & lineNum & "): " & Trim(line)
                End If
            Loop
            ts.Close
        End If
    Next
    Dim sub_
    For Each sub_ In folder.SubFolders
        ScanFolder sub_.Path
    Next
End Sub
