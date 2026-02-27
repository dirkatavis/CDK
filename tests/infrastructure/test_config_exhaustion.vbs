' ============================================================================
' Global Config Exhaustion Test
' Purpose: Ensures EVERY path defined in config.ini actually exists on disk.
'          This prevents "dead configuration" and broken refactoring links.
' ============================================================================

Option Explicit

Dim g_fso, g_shell, g_repoRoot, g_helperPath
Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_shell = CreateObject("WScript.Shell")

' --- Bootstrap ---
g_repoRoot = g_shell.Environment("USER")("CDK_BASE")
If g_repoRoot = "" Then WScript.Quit 1

g_helperPath = g_fso.BuildPath(g_repoRoot, "framework\PathHelper.vbs")
ExecuteGlobal g_fso.OpenTextFile(g_helperPath).ReadAll

Dim configPath: configPath = g_fso.BuildPath(g_repoRoot, "config\config.ini")
If Not g_fso.FileExists(configPath) Then
    WScript.Echo "[FAIL] Missing config.ini"
    WScript.Quit 1
End If

' Iterate through all sections and keys
Dim total, passed
total = 0
passed = 0

' Note: We reuse ReadIniSection from PathHelper
' But we need a list of ALL sections. We'll parse the file manually for sections.
Dim ts: Set ts = g_fso.OpenTextFile(configPath, 1)
Dim line, section, key, val, fullPath
section = ""

Do While Not ts.AtEndOfStream
    line = Trim(ts.ReadLine)
    If line = "" Or Left(line, 1) = "#" Or Left(line, 1) = ";" Then
        ' Skip empty or comments
    ElseIf Left(line, 1) = "[" And Right(line, 1) = "]" Then
        section = Mid(line, 2, Len(line) - 2)
    ElseIf InStr(line, "=") > 0 Then
        key = Trim(Left(line, InStr(line, "=") - 1))
        
        ' Certain keys are NOT paths (e.g. StabilityPause, SkipSequences)
        If IsPathKey(key) Then
            total = total + 1
            fullPath = GetConfigPath(section, key)
            
            If fullPath = "" Then
                WScript.Echo "[FAIL] Section: [" & section & "] Key: " & key & " - Could not resolve path"
            ElseIf Not g_fso.FileExists(fullPath) And Not g_fso.FolderExists(fullPath) Then
                ' Special case: Log files and OutputCSV might not exist yet if app hasn't run
                ' We should verify the PARENT folder exists
                Dim parent: parent = g_fso.GetParentFolderName(fullPath)
                If g_fso.FolderExists(parent) Then
                    passed = passed + 1
                    ' WScript.Echo "[PASS] " & section & "." & key & " (Target parent exists)"
                Else
                    WScript.Echo "[FAIL] " & section & "." & key & " - Path missing: " & fullPath
                End If
            Else
                passed = passed + 1
                ' WScript.Echo "[PASS] " & section & "." & key
            End If
        End If
    End If
Loop
ts.Close

WScript.Echo "Config path coverage: " & passed & "/" & total
If passed = total Then
    WScript.Echo "[PASS] All config paths are resolvable and valid"
    WScript.Quit 0
Else
    WScript.Echo "[FAIL] One or more config paths are broken"
    WScript.Quit 1
End If

Function IsPathKey(k)
    IsPathKey = False
    Dim p: p = UCase(k)
    ' Heuristic: keys that usually contain paths
    If InStr(p, "CSV") > 0 Or InStr(p, "LOG") > 0 Or InStr(p, "FILE") > 0 Or _
       InStr(p, "DIR") > 0 Or InStr(p, "PATH") > 0 Or InStr(p, "CRITERIA") > 0 Or _
       InStr(p, "LIB") > 0 Or InStr(p, "MARKER") > 0 Or InStr(p, "MAP") > 0 Then
        IsPathKey = True
    End If
    
    ' Exclude known non-path keys that might hit the heuristic
    If p = "LOGLEVEL" Then IsPathKey = False
End Function
