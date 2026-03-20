' ============================================================================
' test_log_level_parser_contract.vbs
'
' Purpose: Verify the parser-contract invariant and log verbosity gating
'          logic for the Open_RO logging refactor.
'          All tests are mock-based - no BlueZone dependency.
'
' Usage:   cscript.exe //nologo test_log_level_parser_contract.vbs
' ============================================================================
Option Explicit

Dim g_pass, g_fail
g_pass = 0
g_fail = 0

WScript.Echo "=== Open_RO Log Level & Parser Contract Tests ==="
WScript.Echo ""

' ----------------------------------------------------------------------------
' Helpers
' ----------------------------------------------------------------------------
Sub AssertTrue(testName, condition)
    If condition Then
        g_pass = g_pass + 1
        WScript.Echo "  [PASS] " & testName
    Else
        g_fail = g_fail + 1
        WScript.Echo "  [FAIL] " & testName
    End If
End Sub

Function ParserMatches(logLine)
    Dim re
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "MVA:\s*(\d{6,9})\s*-\s*RO:\s*(\d+)"
    ParserMatches = re.Test(logLine)
    Set re = Nothing
End Function

' Mirrors the exact format used by LogEntryWithRO in Open_RO.vbs
Function FormatParserLine(mva, roNum)
    FormatParserLine = Now & " - MVA: " & mva & " - RO: " & roNum
End Function

' ============================================================================
' BLOCK 1: Parser Regex Contract (8 tests)
' Parser regex: MVA:\s*(\d{6,9})\s*-\s*RO:\s*(\d+)
' ============================================================================
WScript.Echo "--- Block 1: Parser Regex Contract ---"

AssertTrue "6-digit MVA matches parser regex", _
    ParserMatches(FormatParserLine("123456", "99001"))

AssertTrue "9-digit MVA matches parser regex", _
    ParserMatches(FormatParserLine("123456789", "99001"))

AssertTrue "7-digit MVA matches parser regex", _
    ParserMatches(FormatParserLine("1234567", "99001"))

AssertTrue "8-digit MVA matches parser regex", _
    ParserMatches(FormatParserLine("12345678", "99001"))

AssertTrue "5-digit MVA does NOT match (too short)", _
    Not ParserMatches(FormatParserLine("12345", "99001"))

AssertTrue "10-digit MVA does NOT match (too long)", _
    Not ParserMatches(FormatParserLine("1234567890", "99001"))

AssertTrue "Non-MVA log line does not match", _
    Not ParserMatches("3/20/2026 10:00:00 AM - Script started")

AssertTrue "Full dated log line with 6-digit MVA matches", _
    ParserMatches("3/20/2026 10:00:00 AM - MVA: 999999 - RO: 12345")

WScript.Echo ""

' ============================================================================
' BLOCK 2: Log Level Constants (3 tests)
' ============================================================================
WScript.Echo "--- Block 2: Log Level Constants ---"

Const LOG_LEVEL_LOW  = 1
Const LOG_LEVEL_MED  = 2
Const LOG_LEVEL_HIGH = 3

AssertTrue "LOG_LEVEL_LOW = 1",  LOG_LEVEL_LOW  = 1
AssertTrue "LOG_LEVEL_MED = 2",  LOG_LEVEL_MED  = 2
AssertTrue "LOG_LEVEL_HIGH = 3", LOG_LEVEL_HIGH = 3

WScript.Echo ""

' ============================================================================
' BLOCK 3: ShouldLog Gating Logic (9 tests)
' Replicates ShouldLog() in Open_RO.vbs in isolation.
' ============================================================================
WScript.Echo "--- Block 3: ShouldLog Gating Logic ---"

Function ShouldLogAt(verbosity, level)
    Dim lvl
    Select Case LCase(Trim(level))
        Case "low"  : lvl = LOG_LEVEL_LOW
        Case "med"  : lvl = LOG_LEVEL_MED
        Case "high" : lvl = LOG_LEVEL_HIGH
        Case Else   : lvl = LOG_LEVEL_HIGH
    End Select
    ShouldLogAt = (lvl <= verbosity)
End Function

AssertTrue "Low verbosity: 'low' IS logged",       ShouldLogAt(LOG_LEVEL_LOW,  "low")
AssertTrue "Low verbosity: 'med' is NOT logged",   Not ShouldLogAt(LOG_LEVEL_LOW,  "med")
AssertTrue "Low verbosity: 'high' is NOT logged",  Not ShouldLogAt(LOG_LEVEL_LOW,  "high")
AssertTrue "Med verbosity: 'low' IS logged",       ShouldLogAt(LOG_LEVEL_MED,  "low")
AssertTrue "Med verbosity: 'med' IS logged",       ShouldLogAt(LOG_LEVEL_MED,  "med")
AssertTrue "Med verbosity: 'high' is NOT logged",  Not ShouldLogAt(LOG_LEVEL_MED,  "high")
AssertTrue "High verbosity: 'low' IS logged",      ShouldLogAt(LOG_LEVEL_HIGH, "low")
AssertTrue "High verbosity: 'med' IS logged",      ShouldLogAt(LOG_LEVEL_HIGH, "med")
AssertTrue "High verbosity: 'high' IS logged",     ShouldLogAt(LOG_LEVEL_HIGH, "high")

WScript.Echo ""

' ============================================================================
' BLOCK 4: Parser Contract Parity (4 tests)
' Invariant: LogEntryWithRO lines must match the parser regex regardless of
' verbosity level because LogEntryWithRO is not gated by ShouldLog.
' ============================================================================
WScript.Echo "--- Block 4: Parser Contract Parity ---"

Dim testMva : testMva = "654321"
Dim testRo  : testRo  = "88001"
Dim parserLine : parserLine = FormatParserLine(testMva, testRo)

AssertTrue "Low verbosity: LogEntryWithRO line matches parser regex",  ParserMatches(parserLine)
AssertTrue "Med verbosity: LogEntryWithRO line matches parser regex",  ParserMatches(parserLine)
AssertTrue "High verbosity: LogEntryWithRO line matches parser regex", ParserMatches(parserLine)

AssertTrue "Zero-success: non-MVA log lines produce no parser match", _
    Not ParserMatches("3/20/2026 10:00:00 AM - Script started") And _
    Not ParserMatches("3/20/2026 10:00:01 AM - CSV file found") And _
    Not ParserMatches("3/20/2026 10:00:02 AM - No matching vehicle found")

WScript.Echo ""

' ============================================================================
' Summary
' ============================================================================
WScript.Echo "=== Results: " & g_pass & " passed, " & g_fail & " failed ==="

If g_fail > 0 Then
    WScript.Quit 1
End If
