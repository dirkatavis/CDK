Option Explicit

Dim fso
Dim shell
Dim repoRoot
Dim helperPath
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

helperPath = fso.BuildPath(repoRoot, "framework\PathHelper.vbs")
If Not fso.FileExists(helperPath) Then
    WScript.Echo "FAIL: PathHelper.vbs not found: " & helperPath
    WScript.Quit 1
End If

ExecuteGlobal fso.OpenTextFile(helperPath).ReadAll

mapPath = fso.BuildPath(repoRoot, "tooling\reorg_path_map.ini")
If Not fso.FileExists(mapPath) Then
    WScript.Echo "FAIL: Migration path map not found: " & mapPath
    WScript.Quit 1
End If

Dim contracts: Set contracts = ReadIniSection(mapPath, "ConfigContracts")
If contracts.Count = 0 Then
    WScript.Echo "FAIL: No entries found in [ConfigContracts]"
    WScript.Quit 1
End If

Dim key, contractParts
For Each key In contracts.Keys
    contractParts = Split(contracts(key), "|")
    If UBound(contractParts) <> 2 Then
        WScript.Echo "FAIL: Invalid contract format for key " & key & " -> " & contracts(key)
        failures = failures + 1
    Else
        ValidatePath Trim(contractParts(0)), Trim(contractParts(1)), Trim(contractParts(2))
    End If
Next

If failures > 0 Then
    WScript.Echo ""
    WScript.Echo "Config path contract test FAILED. Broken contracts: " & failures
    WScript.Quit 1
End If

WScript.Echo "Config path contract test PASSED"
WScript.Quit 0

Sub ValidatePath(sectionName, keyName, expectedRelativePath)
    On Error Resume Next
    Dim resolved
    resolved = GetConfigPath(sectionName, keyName)

    If Err.Number <> 0 Then
        WScript.Echo "FAIL: GetConfigPath(" & sectionName & ", " & keyName & ") -> " & Err.Description
        failures = failures + 1
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    If Len(resolved) = 0 Then
        WScript.Echo "FAIL: Empty path for [" & sectionName & "] " & keyName
        failures = failures + 1
        Exit Sub
    End If

    Dim expectedAbsolute
    expectedAbsolute = fso.BuildPath(repoRoot, expectedRelativePath)

    If StrComp(LCase(resolved), LCase(expectedAbsolute), vbTextCompare) <> 0 Then
        WScript.Echo "FAIL: Path mismatch for [" & sectionName & "] " & keyName
        WScript.Echo "  Expected: " & expectedAbsolute
        WScript.Echo "  Actual:   " & resolved
        failures = failures + 1
        Exit Sub
    End If

    WScript.Echo "PASS: [" & sectionName & "] " & keyName & " -> " & resolved
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
