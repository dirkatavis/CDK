'-----------------------------------------------------------------------------------
' ScrapeSkipReasons.vbs
'
' Reads PostFinalCharges.log and writes one row per skipped RO to SkipReasons.csv.
' Columns: Sequence, RO, SkipReason
'
' Usage:
'   cscript ScrapeSkipReasons.vbs
'   cscript ScrapeSkipReasons.vbs "C:\path\to\PostFinalCharges.log"
'-----------------------------------------------------------------------------------
Option Explicit

' Force console (cscript) execution
If InStr(1, WScript.FullName, "wscript", vbTextCompare) > 0 Then
    Dim cmdLine, argIdx
    cmdLine = """" & Replace(WScript.FullName, "wscript.exe", "cscript.exe", 1, -1, vbTextCompare) & """ //nologo """ & WScript.ScriptFullName & """"
    For argIdx = 0 To WScript.Arguments.Count - 1
        cmdLine = cmdLine & " """ & WScript.Arguments(argIdx) & """"
    Next
    CreateObject("WScript.Shell").Run cmdLine, 1, True
    WScript.Quit
End If

Dim fso, logPath, csvPath, ts, line
Dim roBySeq     ' Dictionary: seqNum -> RO number string
Dim rows()      ' Array of "seq|ro|reason" strings
Dim rowCount

Set fso = CreateObject("Scripting.FileSystemObject")
Set roBySeq = CreateObject("Scripting.Dictionary")
roBySeq.CompareMode = 0  ' binary — sequence numbers are numeric strings

ReDim rows(0)
rowCount = 0

' Resolve log path
If WScript.Arguments.Count > 0 Then
    logPath = WScript.Arguments(0)
Else
    logPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "PostFinalCharges.log")
End If

If Not fso.FileExists(logPath) Then
    WScript.Echo "Log not found: " & logPath
    WScript.Quit 1
End If

' Patterns:
'   Processing line: "Sequence N (RO XXXXXX) - Processing"
'   Result line:     "Sequence N - Result: Skipped - ..."
Dim processingMarker : processingMarker = ") - Processing"
Dim resultMarker     : resultMarker     = "Result: Skipped -"
Dim seqPrefix        : seqPrefix        = "Sequence "

Set ts = fso.OpenTextFile(logPath, 1)
Do While Not ts.AtEndOfStream
    line = Trim(ts.ReadLine)

    ' Strip log prefix (timestamp + level tags) — content starts after last ']'
    Dim lastBracket : lastBracket = 0
    Dim bi : bi = 1
    Do
        bi = InStr(bi, line, "]", 1)
        If bi = 0 Then Exit Do
        lastBracket = bi
        bi = bi + 1
    Loop
    Dim content : content = Trim(Mid(line, lastBracket + 1))

    ' Match "Sequence N (RO XXXXXX) - Processing"
    If Left(content, Len(seqPrefix)) = seqPrefix And InStr(content, processingMarker) > 0 Then
        Dim seqStr, roStr
        ' Extract sequence number — between "Sequence " and " ("
        Dim parenPos : parenPos = InStr(content, " (")
        If parenPos > 0 Then
            seqStr = Trim(Mid(content, Len(seqPrefix) + 1, parenPos - Len(seqPrefix) - 1))
            ' Extract RO — between "(RO " and ")"
            Dim roStart : roStart = InStr(content, "(RO ")
            Dim roEnd   : roEnd   = InStr(roStart, content, ")")
            If roStart > 0 And roEnd > roStart Then
                roStr = Trim(Mid(content, roStart + 4, roEnd - roStart - 4))
                roBySeq(seqStr) = roStr
            End If
        End If

    ' Match "Sequence N - Result: Skipped - ..."
    ElseIf Left(content, Len(seqPrefix)) = seqPrefix And InStr(content, resultMarker) > 0 Then
        Dim dashPos : dashPos = InStr(content, " - ")
        If dashPos > 0 Then
            Dim seqNum : seqNum = Trim(Mid(content, Len(seqPrefix) + 1, dashPos - Len(seqPrefix) - 1))
            Dim reasonPos : reasonPos = InStr(content, "Result: ")
            Dim reason : reason = Trim(Mid(content, reasonPos + Len("Result: ")))
            Dim roNum : roNum = ""
            If roBySeq.Exists(seqNum) Then roNum = roBySeq(seqNum)
            If rowCount > UBound(rows) Then ReDim Preserve rows(rowCount + 99)
            rows(rowCount) = seqNum & "|" & roNum & "|" & reason
            rowCount = rowCount + 1
        End If
    End If
Loop
ts.Close

If rowCount = 0 Then
    WScript.Echo "No skipped ROs found in log."
    WScript.Quit 0
End If

' Write CSV
csvPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "SkipReasons.csv")
Dim csvTs : Set csvTs = fso.CreateTextFile(csvPath, True)
csvTs.WriteLine "Sequence,RO,SkipReason"

Dim i, parts
For i = 0 To rowCount - 1
    parts = Split(rows(i), "|", 3)
    csvTs.WriteLine _
        CsvField(parts(0)) & "," & _
        CsvField(parts(1)) & "," & _
        CsvField(parts(2))
Next
csvTs.Close

WScript.Echo rowCount & " skipped RO(s) written to:"
WScript.Echo csvPath
WScript.Echo String(60, "-")
WScript.Echo Left("Seq", 6) & Left("RO", 12) & "SkipReason"
WScript.Echo String(60, "-")
For i = 0 To rowCount - 1
    parts = Split(rows(i), "|", 3)
    WScript.Echo Left(parts(0) & "      ", 6) & Left(parts(1) & "            ", 12) & parts(2)
Next

Function CsvField(val)
    CsvField = """" & Replace(val, """", """""") & """"
End Function
