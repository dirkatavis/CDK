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

mapPath = fso.BuildPath(repoRoot, "tools\reorg_path_map.ini")
If Not fso.FileExists(mapPath) Then
    WScript.Echo "FAIL: Migration path map not found: " & mapPath
    WScript.Quit 1
End If

Dim entrypoints: Set entrypoints = ReadIniSection(mapPath, "LegacyEntrypoints")
If entrypoints.Count = 0 Then
    WScript.Echo "FAIL: No entries found in [LegacyEntrypoints]"
    WScript.Quit 1
End If

Dim key
For Each key In entrypoints.Keys
    ValidateLegacyEntrypoint entrypoints(key)
Next

If failures > 0 Then
    WScript.Echo ""
    WScript.Echo "Wrapper compatibility contract FAILED. Issues: " & failures
    WScript.Quit 1
End If

WScript.Echo "Wrapper compatibility contract PASSED"
WScript.Quit 0

Sub ValidateLegacyEntrypoint(relPath)
    Dim fullPath
    fullPath = fso.BuildPath(repoRoot, relPath)

    If Not fso.FileExists(fullPath) Then
        WScript.Echo "FAIL: Missing legacy entrypoint: " & relPath
        failures = failures + 1
        Exit Sub
    End If

    Dim content
    content = ReadAllText(fullPath)

    ' Contract rule:
    ' - Today: script can still be the real implementation (no wrapper marker needed).
    ' - During migration: if script is converted to wrapper, it MUST include
    '   a line like: ' WRAPPER_TARGET: scripts\...\NewScript.vbs
    '   and that target must exist.
    Dim marker
    marker = "WRAPPER_TARGET:"

    If InStr(1, content, marker, vbTextCompare) > 0 Then
        Dim targetRelPath
        targetRelPath = ExtractWrapperTarget(content)

        If Len(targetRelPath) = 0 Then
            WScript.Echo "FAIL: Wrapper marker present but no target in " & relPath
            failures = failures + 1
            Exit Sub
        End If

        Dim targetFullPath
        targetFullPath = fso.BuildPath(repoRoot, targetRelPath)

        If Not fso.FileExists(targetFullPath) Then
            WScript.Echo "FAIL: Wrapper target missing for " & relPath & " -> " & targetRelPath
            failures = failures + 1
            Exit Sub
        End If

        WScript.Echo "PASS: Wrapper target valid for " & relPath & " -> " & targetRelPath
    Else
        WScript.Echo "PASS: Legacy entrypoint present (non-wrapper): " & relPath
    End If
End Sub

Function ReadAllText(filePath)
    Dim ts
    Set ts = fso.OpenTextFile(filePath, 1, False)
    ReadAllText = ts.ReadAll
    ts.Close
End Function

Function ExtractWrapperTarget(content)
    Dim lines, i, line, pos
    ExtractWrapperTarget = ""

    lines = Split(content, vbCrLf)
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        pos = InStr(1, line, "WRAPPER_TARGET:", vbTextCompare)
        If pos > 0 Then
            ExtractWrapperTarget = Trim(Mid(line, pos + Len("WRAPPER_TARGET:")))
            Exit Function
        End If
    Next
End Function

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
