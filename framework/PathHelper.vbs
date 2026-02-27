Option Explicit

' ==============================================================================
' PathHelper.vbs - Centralized Path Management for CDK Scripts
' ==============================================================================
' This module provides path discovery and configuration reading for all CDK scripts.
' 
' Usage in your script:
'   1. Include this file: ExecuteGlobal fso.OpenTextFile("path\to\PathHelper.vbs").ReadAll
'   2. Get the repo root: root = GetRepoRoot()
'   3. Get a configured path: csvPath = GetConfigPath("Initialize_RO", "CSV")
' ==============================================================================

Dim g_fso
If IsObject(g_fso) = False Then
    Set g_fso = CreateObject("Scripting.FileSystemObject")
End If

Const BASE_ENV_VAR = "CDK_BASE"

' ------------------------------------------------------------------------------
' GetRepoRoot - Discovers the repository root directory
' ------------------------------------------------------------------------------
' Search strategy:
'   1. Read base path from CDK_BASE environment variable
'   2. Validate .cdkroot exists under that base path
'   3. FAIL with clear error - no fallbacks, no silent failures
' ------------------------------------------------------------------------------
Function GetRepoRoot()
    Dim basePath: basePath = ReadBasePath()
    Dim markerPath: markerPath = g_fso.BuildPath(basePath, ".cdkroot")

    If Not g_fso.FileExists(markerPath) Then
        Err.Raise 53, "GetRepoRoot", _
            "Cannot find .cdkroot marker file at:" & vbCrLf & markerPath & vbCrLf & vbCrLf & _
            "Solution: Ensure the CDK folder you copied includes the .cdkroot file."
    End If

    GetRepoRoot = basePath
End Function

' ------------------------------------------------------------------------------
' FindRepoRootForBootstrap - Bootstrap-friendly repo root helper
' ------------------------------------------------------------------------------
' Alias kept for scripts that still call the legacy bootstrap function name.
' ------------------------------------------------------------------------------
Function FindRepoRootForBootstrap()
    FindRepoRootForBootstrap = GetRepoRoot()
End Function

' ------------------------------------------------------------------------------
' ReadBasePath - Reads repo root from environment variable
' ------------------------------------------------------------------------------
Function ReadBasePath()
    Dim sh: Set sh = CreateObject("WScript.Shell")
    Dim basePath: basePath = sh.Environment("USER")(BASE_ENV_VAR)

    If Len(basePath) = 0 Then
        Err.Raise 53, "ReadBasePath", _
            "Missing environment variable: " & BASE_ENV_VAR & vbCrLf & vbCrLf & _
            "Solution: Set CDK_BASE to the full path of the CDK folder."
    End If

    If Not g_fso.FolderExists(basePath) Then
        Err.Raise 53, "ReadBasePath", _
            "Base path does not exist:" & vbCrLf & basePath & vbCrLf & vbCrLf & _
            "Solution: Fix the CDK_BASE environment variable."
    End If

    ReadBasePath = basePath
End Function

' ------------------------------------------------------------------------------
' GetConfigPath - Builds absolute path from config.ini settings
' ------------------------------------------------------------------------------
' Parameters:
'   section - INI section name (e.g., "Initialize_RO")
'   key     - Setting name (e.g., "CSV")
' Returns: Absolute path built from repo root + relative path from config
' ------------------------------------------------------------------------------
Function GetConfigPath(section, key)
    Dim root: root = GetRepoRoot()
    Dim configFile: configFile = g_fso.BuildPath(root, "config\config.ini")
    
    If Not g_fso.FileExists(configFile) Then
        Err.Raise 53, "GetConfigPath", "config.ini not found at: " & configFile
    End If
    
    Dim relativePath: relativePath = ReadIniValue(configFile, section, key)
    If relativePath = "" Then
        Err.Raise 5, "GetConfigPath", "Config key not found: [" & section & "] " & key
    End If
    
    GetConfigPath = g_fso.BuildPath(root, relativePath)
End Function

' ------------------------------------------------------------------------------
' ReadIniValue - Reads a value from an INI file
' ------------------------------------------------------------------------------
' Simple INI parser for [Section] Key=Value format
' ------------------------------------------------------------------------------
Function ReadIniValue(filePath, section, key)
    ReadIniValue = ""
    
    Dim ts: Set ts = g_fso.OpenTextFile(filePath, 1, False)
    Dim currentSection: currentSection = ""
    Dim line, trimmedLine
    
    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        trimmedLine = Trim(line)
        
        ' Skip comments and empty lines
        If Len(trimmedLine) = 0 Or Left(trimmedLine, 1) = "#" Or Left(trimmedLine, 1) = ";" Then
            ' Continue
        ElseIf Left(trimmedLine, 1) = "[" And Right(trimmedLine, 1) = "]" Then
            ' Section header
            currentSection = Mid(trimmedLine, 2, Len(trimmedLine) - 2)
        ElseIf currentSection = section Then
            ' Key=Value in the target section
            Dim eqPos: eqPos = InStr(trimmedLine, "=")
            If eqPos > 0 Then
                Dim iniKey: iniKey = Trim(Left(trimmedLine, eqPos - 1))
                If LCase(iniKey) = LCase(key) Then
                    ReadIniValue = Trim(Mid(trimmedLine, eqPos + 1))
                    ts.Close
                    Exit Function
                End If
            End If
        End If
    Loop
    
    ts.Close
End Function

' ------------------------------------------------------------------------------
' BuildLogPath - Helper to build a log file path with timestamp option
' ------------------------------------------------------------------------------
Function BuildLogPath(section, key)
    BuildLogPath = GetConfigPath(section, key)
End Function

' ------------------------------------------------------------------------------
' BuildCSVPath - Helper specifically for CSV files
' ------------------------------------------------------------------------------
Function BuildCSVPath(section, key)
    BuildCSVPath = GetConfigPath(section, key)
End Function
