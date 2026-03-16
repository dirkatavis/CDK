Option Explicit

Sub Main()
    Dim fso, shell, basePath, markerPath, configPath
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")

    basePath = shell.Environment("USER")("CDK_BASE")
    If Len(Trim(CStr(basePath))) = 0 Or Not fso.FolderExists(basePath) Then
        MsgBox "CDK_BASE is missing or invalid: '" & basePath & "'", vbCritical, "Blacklist Probe"
        Exit Sub
    End If

    markerPath = fso.BuildPath(basePath, ".cdkroot")
    If Not fso.FileExists(markerPath) Then
        MsgBox ".cdkroot not found under CDK_BASE:" & vbCrLf & basePath, vbCritical, "Blacklist Probe"
        Exit Sub
    End If

    configPath = fso.BuildPath(basePath, "config\config.ini")

    Dim blacklistCsv
    blacklistCsv = ReadIniValue(configPath, "PostFinalCharges", "blacklist_terms")
    If Len(Trim(CStr(blacklistCsv))) = 0 Then
        MsgBox "blacklist_terms is empty in:" & vbCrLf & configPath, vbCritical, "Blacklist Probe"
        Exit Sub
    End If

    Dim bzhao
    On Error Resume Next
    Set bzhao = CreateObject("BZWhll.WhllObj")
    If Err.Number <> 0 Then
        MsgBox "Could not create BZWhll.WhllObj: " & Err.Description, vbCritical, "Blacklist Probe"
        Exit Sub
    End If

    bzhao.Connect "A"
    If Err.Number <> 0 Then
        MsgBox "Could not connect to BlueZone session A: " & Err.Description, vbCritical, "Blacklist Probe"
        Exit Sub
    End If
    On Error GoTo 0

    Dim lines
    lines = ReadAllScreenLines(bzhao)

    Dim matchLine, matchedTerm
    matchLine = ""
    matchedTerm = GetMatchedBlacklistTermWithLine(blacklistCsv, lines, matchLine)

    Dim report
    report = "Manual Blacklist Live Probe (No WScript)" & vbCrLf & _
             "=====================================" & vbCrLf & _
             "Loaded blacklist_terms: " & blacklistCsv & vbCrLf & vbCrLf

    report = report & "--- Current Screen (24 lines) ---" & vbCrLf

    Dim i
    For i = 1 To 24
        report = report & Right("  " & CStr(i), 2) & ": " & CStr(lines(i)) & vbCrLf
    Next

    If Len(Trim(CStr(matchedTerm))) > 0 Then
        report = report & vbCrLf & "[PASS] Production-style match found" & vbCrLf & _
                 "Term: " & matchedTerm & vbCrLf & _
                 "Line: " & matchLine
        Call SaveProbeLog(basePath, report)
        MsgBox report, vbInformation, "Blacklist Probe"
    Else
        report = report & vbCrLf & "[FAIL] No production-style match found" & vbCrLf & vbCrLf & _
                 "Potential normalized-space matches (diagnostic only):" & vbCrLf & _
                 BuildNormalizedHints(blacklistCsv, lines)
        Call SaveProbeLog(basePath, report)
        MsgBox report, vbExclamation, "Blacklist Probe"
    End If
End Sub

Function ReadIniValue(iniPath, sectionName, keyName)
    ReadIniValue = ""

    Dim fso, ts, currentSection, line, trimmedLine, eqPos, iniKey
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(iniPath) Then Exit Function

    Set ts = fso.OpenTextFile(iniPath, 1)
    currentSection = ""

    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        trimmedLine = Trim(line)

        If Len(trimmedLine) = 0 Then
            ' skip
        ElseIf Left(trimmedLine, 1) = "#" Or Left(trimmedLine, 1) = ";" Then
            ' skip
        ElseIf Left(trimmedLine, 1) = "[" And Right(trimmedLine, 1) = "]" Then
            currentSection = Mid(trimmedLine, 2, Len(trimmedLine) - 2)
        ElseIf StrComp(currentSection, sectionName, vbTextCompare) = 0 Then
            eqPos = InStr(1, trimmedLine, "=", vbTextCompare)
            If eqPos > 0 Then
                iniKey = Trim(Left(trimmedLine, eqPos - 1))
                If StrComp(iniKey, keyName, vbTextCompare) = 0 Then
                    ReadIniValue = Trim(Mid(trimmedLine, eqPos + 1))
                    Exit Do
                End If
            End If
        End If
    Loop

    ts.Close
End Function

Function ReadAllScreenLines(bzhao)
    Dim arr(24)
    Dim row, buf

    For row = 1 To 24
        buf = ""
        On Error Resume Next
        bzhao.ReadScreen buf, 80, row, 1
        On Error GoTo 0
        arr(row) = RTrim(CStr(buf))
    Next

    ReadAllScreenLines = arr
End Function

Function GetMatchedBlacklistTermWithLine(blacklistTermsCsv, lines, ByRef matchedLine)
    Dim terms, i, row, term, lineText
    GetMatchedBlacklistTermWithLine = ""
    matchedLine = ""

    blacklistTermsCsv = Trim(CStr(blacklistTermsCsv))
    If Len(blacklistTermsCsv) = 0 Then Exit Function

    terms = Split(blacklistTermsCsv, ",")

    For i = LBound(terms) To UBound(terms)
        term = Trim(CStr(terms(i)))
        If Len(term) > 0 Then
            For row = 1 To 24
                lineText = CStr(lines(row))
                If InStr(1, lineText, term, vbTextCompare) > 0 Then
                    GetMatchedBlacklistTermWithLine = term
                    matchedLine = CStr(row) & ": " & lineText
                    Exit Function
                End If
            Next
        End If
    Next
End Function

Function BuildNormalizedHints(blacklistTermsCsv, lines)
    Dim output, terms, i, row, term, normTerm, normLine, hitCount
    output = ""
    hitCount = 0

    terms = Split(CStr(blacklistTermsCsv), ",")

    For i = LBound(terms) To UBound(terms)
        term = Trim(CStr(terms(i)))
        If Len(term) > 0 Then
            normTerm = NormalizeSpaces(UCase(term))
            For row = 1 To 24
                normLine = NormalizeSpaces(UCase(CStr(lines(row))))
                If InStr(1, normLine, normTerm, vbTextCompare) > 0 Then
                    hitCount = hitCount + 1
                    output = output & "  [HIT] term='" & term & "' on line " & row & " (normalized)" & vbCrLf
                End If
            Next
        End If
    Next

    If hitCount = 0 Then output = "  (none)"
    BuildNormalizedHints = output
End Function

Function NormalizeSpaces(text)
    Dim result
    result = Trim(CStr(text))

    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    NormalizeSpaces = result
End Function

Sub SaveProbeLog(basePath, reportText)
    On Error Resume Next

    Dim fso, logDir, logPath, ts
    Set fso = CreateObject("Scripting.FileSystemObject")

    logDir = fso.BuildPath(basePath, "runtime\logs\post_final_charges")
    If Not fso.FolderExists(logDir) Then fso.CreateFolder(logDir)

    logPath = fso.BuildPath(logDir, "blacklist_live_probe.log")
    Set ts = fso.OpenTextFile(logPath, 8, True)
    ts.WriteLine String(70, "=")
    ts.WriteLine Now()
    ts.WriteLine reportText
    ts.Close

    On Error GoTo 0
End Sub

Main()
