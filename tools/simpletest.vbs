Dim g_fso, g_sh, g_root
Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_sh = CreateObject("WScript.Shell")
g_root = g_sh.Environment("USER")("CDK_BASE")

' Searching for duplicate LOCAL DEFINITIONS of shared functions.
' Strategy: search for the keyword, but strip "Function <keyword>" lines first.
' If the keyword still appears after stripping, that file has a local definition.
' Expected: each term should appear ONLY in its canonical home (BZHelper / PathHelper).
' BZWhll.WhllObj is searched as-is — any occurrence is a production instantiation.
Dim terms(4)
terms(0) = "WaitForPrompt"
terms(1) = "IsTextPresent"
terms(2) = "FindRepoRootForBootstrap"
terms(3) = "WaitForAnyOf"
terms(4) = "BZWhll.WhllObj"

Dim i
For i = 0 To 4
    Dim results: results = ""
    SearchFolder g_root, terms(i)
    MsgBox terms(i) & " found in:" & Chr(10) & results
Next

Sub SearchFolder(folderPath, searchTerm)
    Dim folder, file, subfolder, ts, content, stripped
    Set folder = g_fso.GetFolder(folderPath)
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".vbs" Then
            On Error Resume Next
            Set ts = g_fso.OpenTextFile(file.Path, 1)
            content = ts.ReadAll
            ts.Close
            If Err.Number = 0 Then
                ' For BZWhll.WhllObj search as-is (no definition/call distinction needed)
                ' For function names: strip "Function <term>" so only local definitions remain
                If searchTerm = "BZWhll.WhllObj" Then
                    If InStr(1, content, searchTerm, vbTextCompare) > 0 Then
                        results = results & file.Path & Chr(10)
                    End If
                Else
                    stripped = Replace(content, "Function " & searchTerm, "~~REMOVED~~", 1, -1, vbTextCompare)
                    If InStr(1, stripped, searchTerm, vbTextCompare) > 0 Then
                        results = results & file.Path & Chr(10)
                    End If
                End If
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next
    For Each subfolder In folder.SubFolders
        SearchFolder subfolder.Path, searchTerm
    Next
End Sub
