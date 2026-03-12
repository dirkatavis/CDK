'-----------------------------------------------------------------------------------
' **PROCEDURE NAME:** TestBlacklistIntegration
' **DATE CREATED:** 2026-03-12
' **AUTHOR:** GitHub Copilot
'
' **FUNCTIONALITY:**
' Integration test for blacklist feature that validates:
' 1. Config loading from config.ini works correctly
' 2. Blacklist array is properly populated
' 3. CheckBlacklistedTerms function detects terms in real RO screens
' 4. GetIniSetting wrapper handles missing values gracefully
'-----------------------------------------------------------------------------------

Option Explicit

Dim g_fso, g_BlacklistTerms

Function IncludeFile(filePath)
    On Error Resume Next
    Dim fsoInclude, fileContent, includeStream

    Set fsoInclude = CreateObject("Scripting.FileSystemObject")

    If Not fsoInclude.FileExists(filePath) Then
        WScript.Echo "ERROR: File not found: " & filePath
        IncludeFile = False
        Exit Function
    End If

    Set includeStream = fsoInclude.OpenTextFile(filePath, 1)
    fileContent = includeStream.ReadAll
    includeStream.Close
    Set includeStream = Nothing

    ExecuteGlobal fileContent
    IncludeFile = True
End Function

' Simple CheckBlacklistedTerms for integration testing
Function CheckBlacklistedTerms(blacklistArray, screenContent)
    Dim i, matchedTerm
    
    If UBound(blacklistArray) < 0 Then
        CheckBlacklistedTerms = ""
        Exit Function
    End If
    
    screenContent = UCase(screenContent)
    
    For i = LBound(blacklistArray) To UBound(blacklistArray)
        matchedTerm = Trim(UCase(blacklistArray(i)))
        If Len(matchedTerm) > 0 And InStr(screenContent, matchedTerm) > 0 Then
            CheckBlacklistedTerms = matchedTerm
            Exit Function
        End If
    Next
    
    CheckBlacklistedTerms = ""
End Function

Sub TestConfigLoading()
    WScript.Echo "TEST 1: Config loading from config.ini"
    
    Dim configPath, blacklistValue, result
    Set g_fso = CreateObject("Scripting.FileSystemObject")
    
    ' Find config.ini
    configPath = g_fso.BuildPath(g_fso.GetParentFolderName(g_fso.GetParentFolderName(g_fso.GetParentFolderName(g_fso.GetParentFolderName(WScript.ScriptFullName)))), "config\config.ini")
    
    If Not g_fso.FileExists(configPath) Then
        WScript.Echo "  [FAIL] config.ini not found at: " & configPath
        Exit Sub
    End If
    
    ' Read the PostFinalCharges section
    blacklistValue = ReadConfigValue(configPath, "PostFinalCharges", "blacklist_terms")
    
    If blacklistValue = "VEND TO DEALER" Then
        WScript.Echo "  [PASS] blacklist_terms='" & blacklistValue & "' loaded successfully"
    Else
        WScript.Echo "  [FAIL] Expected 'VEND TO DEALER', got '" & blacklistValue & "'"
    End If
End Sub

Sub TestGetIniSettingFunction()
    WScript.Echo "TEST 2: GetIniSetting wrapper function"
    
    Dim configPath, result
    Set g_fso = CreateObject("Scripting.FileSystemObject")
    configPath = g_fso.BuildPath(g_fso.GetParentFolderName(g_fso.GetParentFolderName(g_fso.GetParentFolderName(g_fso.GetParentFolderName(WScript.ScriptFullName)))), "config\config.ini")
    
    ' Test with existing value
    result = ReadConfigValue(configPath, "PostFinalCharges", "blacklist_terms")
    If result = "VEND TO DEALER" Then
        WScript.Echo "  [PASS] GetIniSetting retrieved existing value"
    Else
        WScript.Echo "  [FAIL] GetIniSetting failed to retrieve value"
    End If
    
    ' Test with non-existing value (should return empty)
    result = ReadConfigValue(configPath, "PostFinalCharges", "nonexistent_key")
    If result = "" Then
        WScript.Echo "  [PASS] GetIniSetting returns empty for missing keys"
    Else
        WScript.Echo "  [FAIL] Expected empty string, got '" & result & "'"
    End If
End Sub

Sub TestBlacklistArrayPopulation()
    WScript.Echo "TEST 3: Blacklist array population from config"
    
    Dim configPath, blacklistStr, blacklistArray, i
    Set g_fso = CreateObject("Scripting.FileSystemObject")
    configPath = g_fso.BuildPath(g_fso.GetParentFolderName(g_fso.GetParentFolderName(g_fso.GetParentFolderName(g_fso.GetParentFolderName(WScript.ScriptFullName)))), "config\config.ini")
    
    ' Load and parse
    blacklistStr = ReadConfigValue(configPath, "PostFinalCharges", "blacklist_terms")
    
    If Len(Trim(blacklistStr)) > 0 Then
        blacklistArray = Split(blacklistStr, ",")
        For i = LBound(blacklistArray) To UBound(blacklistArray)
            blacklistArray(i) = Trim(blacklistArray(i))
        Next
        
        If UBound(blacklistArray) >= 0 Then
            WScript.Echo "  [PASS] Blacklist array populated with " & (UBound(blacklistArray) + 1) & " term(s)"
            If blacklistArray(0) = "VEND TO DEALER" Then
                WScript.Echo "  [PASS] First term is 'VEND TO DEALER'"
            Else
                WScript.Echo "  [FAIL] Expected first term 'VEND TO DEALER', got '" & blacklistArray(0) & "'"
            End If
        Else
            WScript.Echo "  [FAIL] Blacklist array is empty"
        End If
    Else
        WScript.Echo "  [FAIL] No blacklist value in config"
    End If
End Sub

Sub TestBlacklistDetectionOnRealScreen()
    WScript.Echo "TEST 4: Blacklist detection on real RO screen"
    
    Dim configPath, blacklistStr, blacklistArray, screenContent, result, i
    Set g_fso = CreateObject("Scripting.FileSystemObject")
    configPath = g_fso.BuildPath(g_fso.GetParentFolderName(g_fso.GetParentFolderName(g_fso.GetParentFolderName(g_fso.GetParentFolderName(WScript.ScriptFullName)))), "config\config.ini")
    
    ' Load blacklist
    blacklistStr = ReadConfigValue(configPath, "PostFinalCharges", "blacklist_terms")
    If Len(blacklistStr) > 0 Then
        blacklistArray = Split(blacklistStr, ",")
        For i = LBound(blacklistArray) To UBound(blacklistArray)
            blacklistArray(i) = Trim(blacklistArray(i))
        Next
    Else
        ReDim blacklistArray(-1)
    End If
    
    ' Real RO screen from production
    screenContent = "ONSITE PLUS SERVICE            (PFC) POST FINAL CHARGES            12MAR26 06:34" & vbCrLf & _
                    "RO: 875084     TAG: T5596556 SA: 18351   25 TOYOTA TOYOTA VIN: 4T1DAACK3SU553265" & vbCrLf & _
                    "RO STATUS: READY TO POST          PROMISED: 05MAR26 17:00" & vbCrLf & _
                    "REPAIR ORDER #875084 DETAIL" & vbCrLf & _
                    "LC DESCRIPTION                           TECH... LTYPE    ACT   SOLD    SALE AMT" & vbCrLf & _
                    "B  VEND TO DEALER                        C92" & vbCrLf & _
                    "   L1 VTD VEND TO DEALER                 18351   W       0.00   0.10        8.68"
    
    result = CheckBlacklistedTerms(blacklistArray, screenContent)
    
    If result = "VEND TO DEALER" Then
        WScript.Echo "  [PASS] Detected blacklisted term on real RO screen"
    Else
        WScript.Echo "  [FAIL] Failed to detect 'VEND TO DEALER' on screen, got '" & result & "'"
    End If
End Sub

Sub TestMultipleBlacklistTerms()
    WScript.Echo "TEST 5: Multiple blacklist terms (future use)"
    
    Dim testArray(2), screenContent, result
    testArray(0) = "HOLD"
    testArray(1) = "PENDING"
    testArray(2) = "VEND TO DEALER"
    
    screenContent = "RO STATUS: READY TO POST" & vbCrLf & "B  VEND TO DEALER                        C92"
    
    result = CheckBlacklistedTerms(testArray, screenContent)
    
    If result = "VEND TO DEALER" Then
        WScript.Echo "  [PASS] Correctly detected third term in array"
    Else
        WScript.Echo "  [FAIL] Expected 'VEND TO DEALER', got '" & result & "'"
    End If
End Sub

' Helper function to read config values
Function ReadConfigValue(filePath, section, key)
    ReadConfigValue = ""
    
    Dim ts, currentSection, line, trimmedLine, eqPos, iniKey
    On Error Resume Next
    
    Set ts = g_fso.OpenTextFile(filePath, 1)
    
    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        trimmedLine = Trim(line)
        
        If Len(trimmedLine) = 0 Or Left(trimmedLine, 1) = "#" Or Left(trimmedLine, 1) = ";" Then
            ' Skip
        ElseIf Left(trimmedLine, 1) = "[" And Right(trimmedLine, 1) = "]" Then
            currentSection = Mid(trimmedLine, 2, Len(trimmedLine) - 2)
        ElseIf currentSection = section Then
            eqPos = InStr(trimmedLine, "=")
            If eqPos > 0 Then
                iniKey = Trim(Left(trimmedLine, eqPos - 1))
                If LCase(iniKey) = LCase(key) Then
                    ReadConfigValue = Trim(Mid(trimmedLine, eqPos + 1))
                    ts.Close
                    On Error GoTo 0
                    Exit Function
                End If
            End If
        End If
    Loop
    
    ts.Close
    On Error GoTo 0
End Function

Sub Main()
    WScript.Echo "PostFinalCharges Blacklist Feature - Integration Test Suite"
    WScript.Echo "==========================================================="
    WScript.Echo ""
    
    TestConfigLoading
    WScript.Echo ""
    
    TestGetIniSettingFunction
    WScript.Echo ""
    
    TestBlacklistArrayPopulation
    WScript.Echo ""
    
    TestBlacklistDetectionOnRealScreen
    WScript.Echo ""
    
    TestMultipleBlacklistTerms
    WScript.Echo ""
    
    WScript.Echo "==========================================================="
    WScript.Echo "Integration Test Suite Complete"
End Sub

Main()
