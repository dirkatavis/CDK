' ============================================================================
' Scan Unconfigured Keys
' Purpose: Scans all VBScript files for GetConfigPath calls and verifies
'          that the referenced section/key pairs exist in config.ini.
' ============================================================================

Option Explicit

Dim g_fso, g_shell, g_repoRoot, g_configMap
Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_shell = CreateObject("WScript.Shell")
Set g_configMap = CreateObject("Scripting.Dictionary")

' --- Initialize ---
g_repoRoot = g_shell.Environment("USER")("CDK_BASE")
If g_repoRoot = "" Then
    WScript.Echo "ERROR: CDK_BASE environment variable not set."
    WScript.Quit 1
End If

LoadConfigMap()

WScript.Echo "============================================================================"
WScript.Echo "SCANNING FOR UNCONFIGURED PATH KEYS"
WScript.Echo "============================================================================"

Dim totalFiles, issuesFound
totalFiles = 0
issuesFound = 0

ScanFolder g_repoRoot

WScript.Echo ""
WScript.Echo "----------------------------------------------------------------------------"
WScript.Echo "Scan complete. Files scanned: " & totalFiles & " | Issues found: " & issuesFound
WScript.Echo "============================================================================"

If issuesFound > 0 Then WScript.Quit 1

' ============================================================================
' CORE LOGIC
' ============================================================================

Sub LoadConfigMap()
    Dim configPath: configPath = g_fso.BuildPath(g_repoRoot, "config\config.ini")
    If Not g_fso.FileExists(configPath) Then Exit Sub

    Dim ts: Set ts = g_fso.OpenTextFile(configPath, 1)
    Dim section, line
    section = ""
    
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If Left(line, 1) = "[" And Right(line, 1) = "]" Then
            section = LCase(Mid(line, 2, Len(line) - 2))
            If Not g_configMap.Exists(section) Then
                Set g_configMap(section) = CreateObject("Scripting.Dictionary")
            End If
        ElseIf InStr(line, "=") > 0 And section <> "" Then
            Dim key: key = LCase(Trim(Left(line, InStr(line, "=") - 1)))
            g_configMap(section)(key) = True
        End If
    Loop
    ts.Close
End Sub

Sub ScanFolder(path)
    Dim folder: Set folder = g_fso.GetFolder(path)
    Dim file, subFolder
    
    ' Skip known non-source or huge directories
    If folder.Name = ".git" Or folder.Name = ".venv" Or folder.Name = "runtime" Or folder.Name = "Temp" Then Exit Sub

    For Each file In folder.Files
        If LCase(g_fso.GetExtensionName(file.Name)) = "vbs" Then
            totalFiles = totalFiles + 1
            ScanFile file.Path
        End If
    Next

    For Each subFolder In folder.SubFolders
        ScanFolder subFolder.Path
    Next
End Sub

Sub ScanFile(filePath)
    Dim ts: Set ts = g_fso.OpenTextFile(filePath, 1)
    Dim content: content = ts.ReadAll
    ts.Close

    ' Regex to find GetConfigPath("Section", "Key")
    ' Handles single or double quotes, and optional spaces
    ' Excludes calls using variables (only matches literal strings)
    Dim re: Set re = New RegExp
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "GetConfigPath\s*\(\s*[""']([^""' &]+)[""']\s*,\s*[""']([^""' &]+)[""']\s*\)"

    Dim matches, match, section, key, relPath
    Set matches = re.Execute(content)
    
    relPath = Replace(filePath, g_repoRoot & "\", "")

    ' Skip this diagnostic script itself to avoid false positives from the regex pattern example
    If InStr(relPath, "scan_unconfigured_keys.vbs") > 0 Then Exit Sub

    For Each match In matches
        section = LCase(match.SubMatches(0))
        key = LCase(match.SubMatches(1))
        
        If Not g_configMap.Exists(section) Then
            WScript.Echo "[MISSING] " & relPath & ": Section [" & match.SubMatches(0) & "] not found in config.ini"
            issuesFound = issuesFound + 1
        ElseIf Not g_configMap(section).Exists(key) Then
            WScript.Echo "[MISSING] " & relPath & ": Key '" & match.SubMatches(1) & "' missing from section [" & match.SubMatches(0) & "]"
            issuesFound = issuesFound + 1
        End If
    Next
End Sub
