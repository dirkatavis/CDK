'-----------------------------------------------------------------------------------
' ScrapeSkipReasons.vbs
'
' Reads PostFinalCharges.log and prints each unique skip reason, sorted.
'
' Usage:
'   cscript ScrapeSkipReasons.vbs
'   cscript ScrapeSkipReasons.vbs "C:\path\to\PostFinalCharges.log"
'-----------------------------------------------------------------------------------
Option Explicit

' Force console (cscript) execution — prevents per-line popup dialogs under wscript
If InStr(1, WScript.FullName, "wscript", vbTextCompare) > 0 Then
    Dim cmdLine, argIdx
    cmdLine = """" & Replace(WScript.FullName, "wscript.exe", "cscript.exe", 1, -1, vbTextCompare) & """ //nologo """ & WScript.ScriptFullName & """"
    For argIdx = 0 To WScript.Arguments.Count - 1
        cmdLine = cmdLine & " """ & WScript.Arguments(argIdx) & """"
    Next
    CreateObject("WScript.Shell").Run cmdLine, 1, True
    WScript.Quit
End If

Dim fso, logPath, csvPath, ts, line, marker, reason
Dim seen, i, key
Dim reasons()
Dim count

Set fso = CreateObject("Scripting.FileSystemObject")
Set seen = CreateObject("Scripting.Dictionary")
seen.CompareMode = 1  ' vbTextCompare — case-insensitive dedup

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

marker = "Result: Skipped -"

Set ts = fso.OpenTextFile(logPath, 1)
Do While Not ts.AtEndOfStream
    line = ts.ReadLine
    Dim pos : pos = InStr(1, line, marker, 1)
    If pos > 0 Then
        reason = Mid(line, pos + Len("Result: "))
        reason = Trim(reason)
        If Not seen.Exists(reason) Then
            seen.Add reason, 1
        End If
    End If
Loop
ts.Close

If seen.Count = 0 Then
    WScript.Echo "No skipped ROs found in log."
    WScript.Quit 0
End If

' Sort keys alphabetically
Dim keys : keys = seen.Keys
count = seen.Count
ReDim reasons(count - 1)
For i = 0 To count - 1
    reasons(i) = keys(i)
Next

' Bubble sort
Dim j, tmp
For i = 0 To count - 2
    For j = 0 To count - 2 - i
        If LCase(reasons(j)) > LCase(reasons(j + 1)) Then
            tmp = reasons(j)
            reasons(j) = reasons(j + 1)
            reasons(j + 1) = tmp
        End If
    Next
Next

' Write CSV
csvPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "SkipReasons.csv")
Dim csvTs : Set csvTs = fso.CreateTextFile(csvPath, True)
csvTs.WriteLine "SkipReason"
For i = 0 To count - 1
    csvTs.WriteLine """" & Replace(reasons(i), """", """""") & """"
Next
csvTs.Close

WScript.Echo "Unique skip reasons (" & count & ") written to:"
WScript.Echo csvPath
WScript.Echo String(60, "-")
For i = 0 To count - 1
    WScript.Echo reasons(i)
Next
