Option Explicit

' Test to scan configuration files for common corruption patterns like git conflict markers
' or non-ANSI encoding artifacts.

Dim fso, repoRoot, configPath, content, lines, i, line, failures
Set fso = CreateObject("Scripting.FileSystemObject")
repoRoot = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
configPath = fso.BuildPath(repoRoot, "config\config.ini")

failures = 0

WScript.Echo "Running Configuration Integrity Scan..."
WScript.Echo "Target: " & configPath

If Not fso.FileExists(configPath) Then
    WScript.Echo "[FAIL] config.ini missing"
    WScript.Quit 1
End If

Set content = fso.OpenTextFile(configPath, 1)
lines = Split(content.ReadAll, vbCrLf)
content.Close

For i = 0 To UBound(lines)
    line = lines(i)
    
    ' Check for git conflict markers
    If InStr(line, "<<<<<<<") = 1 Or InStr(line, "=======") = 1 Or InStr(line, ">>>>>>>") = 1 Then
        WScript.Echo "[FAIL] Conflict marker found on line " & (i + 1) & ": " & line
        failures = failures + 1
    End If
Next

If failures = 0 Then
    WScript.Echo "[PASS] No corruption markers found in config.ini"
    WScript.Quit 0
Else
    WScript.Echo "[FAIL] Found " & failures & " integrity issues."
    WScript.Quit 1
End If
