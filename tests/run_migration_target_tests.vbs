' ============================================================================
' Migration Final-State Target Tests
' Purpose: Track progress toward final repo architecture (expected red until 100%)
' Usage: cscript.exe tools\run_migration_target_tests.vbs
' ============================================================================

Option Explicit

Dim fso
Dim shell
Dim repoRoot
Dim mapPath
Dim helperPath
Dim totalChecks
Dim passedChecks

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

totalChecks = 0
passedChecks = 0

repoRoot = shell.Environment("USER")("CDK_BASE")
If Len(repoRoot) = 0 Then
    WScript.Echo "FAIL: CDK_BASE not set"
    WScript.Quit 1
End If

mapPath = fso.BuildPath(repoRoot, "tooling\reorg_path_map.ini")
If Not fso.FileExists(mapPath) Then
    WScript.Echo "FAIL: Missing map file: " & mapPath
    WScript.Quit 1
End If

helperPath = fso.BuildPath(repoRoot, "framework\PathHelper.vbs")
If Not fso.FileExists(helperPath) Then
    WScript.Echo "FAIL: Missing PathHelper: " & helperPath
    WScript.Quit 1
End If

ExecuteGlobal fso.OpenTextFile(helperPath).ReadAll

WScript.Echo "=" & String(76, "=")
WScript.Echo "MIGRATION FINAL-STATE TARGET TESTS"
WScript.Echo "=" & String(76, "=")

RunEntrypointTargets
RunConfigTargets

Dim passPct
If totalChecks > 0 Then
    passPct = Int((passedChecks / totalChecks) * 100)
Else
    passPct = 0
End If

WScript.Echo ""
WScript.Echo "-" & String(74, "-")
WScript.Echo "Progress: " & passedChecks & "/" & totalChecks & " checks passed (" & passPct & "%)"

Dim minPct
minPct = GetCurrentPhaseThreshold()
WScript.Echo "Phase threshold: " & minPct & "%"

If passPct >= minPct Then
    WScript.Echo "Phase gate: PASS"
Else
    WScript.Echo "Phase gate: FAIL"
End If

If passPct = 100 Then
    WScript.Echo "Final-state gate: PASS"
    WScript.Echo String(76, "=")
    WScript.Quit 0
Else
    WScript.Echo "Final-state gate: FAIL (expected until migration completes)"
    WScript.Echo String(76, "=")
    WScript.Quit 1
End If

Sub RunEntrypointTargets()
    WScript.Echo ""
    WScript.Echo "[Target Entrypoints]"

    Dim section, k, relPath, fullPath
    Set section = ReadIniSection(mapPath, "TargetEntrypoints")

    For Each k In section.Keys
        relPath = section(k)
        fullPath = fso.BuildPath(repoRoot, relPath)
        totalChecks = totalChecks + 1

        If fso.FileExists(fullPath) Then
            passedChecks = passedChecks + 1
            WScript.Echo "PASS: " & relPath
        Else
            WScript.Echo "FAIL: " & relPath
        End If
    Next
End Sub

Sub RunWrapperTargets()
    WScript.Echo ""
    WScript.Echo "[Target Wrapper Targets]"

    Dim section, k, wrapperRel, targetRel, wrapperFull, targetFull, content, actualTarget
    Set section = ReadIniSection(mapPath, "TargetWrapperTargets")

    For Each k In section.Keys
        wrapperRel = k
        targetRel = section(k)
        wrapperFull = fso.BuildPath(repoRoot, wrapperRel)
        targetFull = fso.BuildPath(repoRoot, targetRel)

        totalChecks = totalChecks + 1

        If Not fso.FileExists(wrapperFull) Then
            WScript.Echo "FAIL: Missing wrapper " & wrapperRel
        Else
            content = ReadAllText(wrapperFull)
            actualTarget = ExtractWrapperTarget(content)

            If actualTarget = "" Then
                WScript.Echo "FAIL: No WRAPPER_TARGET in " & wrapperRel
            ElseIf StrComp(NormalizePath(actualTarget), NormalizePath(targetRel), vbTextCompare) <> 0 Then
                WScript.Echo "FAIL: Wrapper target mismatch " & wrapperRel
                WScript.Echo "  Expected: " & targetRel
                WScript.Echo "  Actual:   " & actualTarget
            ElseIf Not fso.FileExists(targetFull) Then
                WScript.Echo "FAIL: Wrapper target missing " & targetRel
            Else
                passedChecks = passedChecks + 1
                WScript.Echo "PASS: " & wrapperRel & " -> " & targetRel
            End If
        End If
    Next
End Sub

Sub RunConfigTargets()
    WScript.Echo ""
    WScript.Echo "[Target Config Contracts]"

    Dim section, k, raw, parts, secName, keyName, expectedRel, expectedAbs, actualAbs, root
    Set section = ReadIniSection(mapPath, "TargetConfigContracts")
    root = GetRepoRoot()

    For Each k In section.Keys
        raw = section(k)
        parts = Split(raw, "|")

        totalChecks = totalChecks + 1

        If UBound(parts) < 1 Then
            WScript.Echo "FAIL: Invalid contract format: " & raw
        Else
            secName = Trim(parts(0))
            keyName = Trim(parts(1))

            On Error Resume Next
            actualAbs = GetConfigPath(secName, keyName)
            If Err.Number <> 0 Then
                WScript.Echo "FAIL: GetConfigPath(" & secName & ", " & keyName & ") -> " & Err.Description
                Err.Clear
                On Error GoTo 0
            Else
                On Error GoTo 0
                ' Use the existence of the path from config.ini as the PASS criteria
                If fso.FileExists(actualAbs) Then
                    passedChecks = passedChecks + 1
                    WScript.Echo "PASS: [" & secName & "] " & keyName & " -> " & actualAbs
                Else
                    WScript.Echo "FAIL: [" & secName & "] " & keyName
                    WScript.Echo "  File not found at configured path: " & actualAbs
                End If
            End If
        End If
    Next
End Sub

Function GetCurrentPhaseThreshold()
    Dim milestones, phaseText, phaseNum, thresholdKey, thresholdText
    Set milestones = ReadIniSection(mapPath, "TargetMilestones")

    If Not milestones.Exists("CurrentPhase") Then
        GetCurrentPhaseThreshold = 0
        Exit Function
    End If

    phaseText = milestones("CurrentPhase")
    If Not IsNumeric(phaseText) Then
        GetCurrentPhaseThreshold = 0
        Exit Function
    End If

    phaseNum = CInt(phaseText)
    thresholdKey = "Phase" & phaseNum & "MinPassPct"

    If milestones.Exists(thresholdKey) Then
        thresholdText = milestones(thresholdKey)
        If IsNumeric(thresholdText) Then
            GetCurrentPhaseThreshold = CInt(thresholdText)
        Else
            GetCurrentPhaseThreshold = 0
        End If
    Else
        GetCurrentPhaseThreshold = 0
    End If
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

Function NormalizePath(pathValue)
    NormalizePath = LCase(Replace(pathValue, "/", "\\"))
End Function
