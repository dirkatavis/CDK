Option Explicit

Const ROOT_OVERRIDE = "C:\Temp_alt\CDK"
Const ENV_BASE_NAME = "CDK_BASE"
Const OUTPUT_FILE_NAME = "hardcoded_paths_report.txt"

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim baseDir: baseDir = GetBaseDir()
Dim outPath: outPath = GetSafeOutputPath(OUTPUT_FILE_NAME)
Dim outFile: Set outFile = fso.OpenTextFile(outPath, 2, True)

Dim totalFiles: totalFiles = 0
Dim matchedLines: matchedLines = 0

Dim re: Set re = CreateObject("VBScript.RegExp")
re.IgnoreCase = True
re.Global = True
re.Pattern = "[A-Z]:\\[^\r\n""' ]+"

outFile.WriteLine "BaseDir | " & baseDir
outFile.WriteLine "File | Line | Path"

ScanFolder baseDir

outFile.WriteLine "Totals | Files: " & totalFiles & " | Matches: " & matchedLines
outFile.Close

Sub ScanFolder(folderPath)
    Dim folder: Set folder = fso.GetFolder(folderPath)
    Dim file
    For Each file In folder.Files
        If IsIncludedFile(file.Name) Then
            totalFiles = totalFiles + 1
            ScanFile file.Path
        End If
    Next

    Dim subFolder
    For Each subFolder In folder.SubFolders
        ScanFolder subFolder.Path
    Next
End Sub

Sub ScanFile(filePath)
    Dim ts
    On Error Resume Next
    Set ts = fso.OpenTextFile(filePath, 1, False)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Dim lineNumber: lineNumber = 0
    Do Until ts.AtEndOfStream
        lineNumber = lineNumber + 1
        Dim line: line = ts.ReadLine
        Dim matches: Set matches = re.Execute(line)
        If matches.Count > 0 Then
            Dim m
            For Each m In matches
                matchedLines = matchedLines + 1
                outFile.WriteLine filePath & " | " & lineNumber & " | " & m.Value
            Next
        End If
    Loop
    ts.Close
End Sub

Function IsIncludedFile(fileName)
    Dim ext: ext = LCase(fso.GetExtensionName(fileName))
    Select Case ext
        Case "vbs", "ps1", "md", "txt", "csv"
            IsIncludedFile = True
        Case Else
            IsIncludedFile = False
    End Select
End Function

Function GetBaseDir()
    Dim envBase: envBase = GetEnvBase()
    If envBase <> "" Then
        GetBaseDir = envBase
    ElseIf ROOT_OVERRIDE <> "" Then
        GetBaseDir = ROOT_OVERRIDE
    Else
        GetBaseDir = fso.GetAbsolutePathName(".")
    End If
End Function

Function GetEnvBase()
    Dim sh: Set sh = CreateObject("WScript.Shell")
    Dim val: val = sh.Environment("PROCESS")(ENV_BASE_NAME)
    If val = "" Then
        val = sh.Environment("USER")(ENV_BASE_NAME)
    End If
    If val = "" Then
        val = sh.Environment("SYSTEM")(ENV_BASE_NAME)
    End If
    GetEnvBase = val
End Function

Function GetSafeOutputPath(fileName)
    Dim tempDir: tempDir = ""
    On Error Resume Next
    tempDir = fso.GetSpecialFolder(2)
    On Error GoTo 0

    If tempDir <> "" Then
        GetSafeOutputPath = fso.BuildPath(tempDir, fileName)
    Else
        GetSafeOutputPath = fso.BuildPath(GetBaseDir(), fileName)
    End If
End Function
