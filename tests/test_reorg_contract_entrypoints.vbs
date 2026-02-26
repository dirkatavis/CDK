Option Explicit

Dim fso
Dim shell
Dim repoRoot
Dim failures
Dim mapPath

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
failures = 0

repoRoot = shell.Environment("USER")("CDK_BASE")
If Len(repoRoot) = 0 Then
    WScript.Echo "FAIL: CDK_BASE is not set"
    WScript.Quit 1
End If

If Not fso.FolderExists(repoRoot) Then
    WScript.Echo "FAIL: CDK_BASE folder does not exist: " & repoRoot
    WScript.Quit 1
End If

If Not fso.FileExists(fso.BuildPath(repoRoot, ".cdkroot")) Then
    WScript.Echo "FAIL: .cdkroot marker not found under CDK_BASE"
    WScript.Quit 1
End If

mapPath = fso.BuildPath(repoRoot, "tools\reorg_path_map.ini")
If Not fso.FileExists(mapPath) Then
    WScript.Echo "FAIL: Migration path map not found: " & mapPath
    WScript.Quit 1
End If

Dim entrypoints: Set entrypoints = ReadIniSection(mapPath, "TargetEntrypoints")
If entrypoints.Count = 0 Then
    WScript.Echo "FAIL: No entries found in [TargetEntrypoints]"
    WScript.Quit 1
End If

' Always validate core map and config files as static contracts.
CheckFile "tools\reorg_path_map.ini"
CheckFile "framework\PathHelper.vbs"
CheckFile "config\config.ini"

Dim key
For Each key In entrypoints.Keys
    CheckFile entrypoints(key)
Next

If failures > 0 Then
    WScript.Echo ""
    WScript.Echo "Entrypoint contract test FAILED. Missing files: " & failures
    WScript.Quit 1
End If

WScript.Echo "Entrypoint contract test PASSED"
WScript.Quit 0

Sub CheckFile(relPath)
    Dim fullPath
    fullPath = fso.BuildPath(repoRoot, relPath)

    If fso.FileExists(fullPath) Then
        WScript.Echo "PASS: " & relPath
    Else
        WScript.Echo "FAIL: Missing " & relPath
        failures = failures + 1
    End If
End Sub

Function ReadIniSection(filePath, sectionName)
    Dim dict: Set dict = CreateObject("Scripting.Dictionary")
    Dim ts: Set ts = fso.OpenTextFile(filePath, 1, False)
    Dim currentSection: currentSection = ""
    Dim line, trimmedLine, eqPos, iniKey, iniValue

    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        trimmedLine = Trim(line)

        If Len(trimmedLine) = 0 Then
            ' Skip
        ElseIf Left(trimmedLine, 1) = "#" Or Left(trimmedLine, 1) = ";" Then
            ' Skip
        ElseIf Left(trimmedLine, 1) = "[" And Right(trimmedLine, 1) = "]" Then
            currentSection = Mid(trimmedLine, 2, Len(trimmedLine) - 2)
        ElseIf LCase(currentSection) = LCase(sectionName) Then
            eqPos = InStr(trimmedLine, "=")
            If eqPos > 0 Then
                iniKey = Trim(Left(trimmedLine, eqPos - 1))
                iniValue = Trim(Mid(trimmedLine, eqPos + 1))
                If iniKey <> "" And iniValue <> "" Then
                    dict(iniKey) = iniValue
                End If
            End If
        End If
    Loop

    ts.Close
    Set ReadIniSection = dict
End Function
