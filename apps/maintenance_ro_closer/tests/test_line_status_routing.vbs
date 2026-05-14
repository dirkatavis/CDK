Option Explicit

' Test suite for line status detection and routing in Maintenance_RO_Closer
' Tests the status->action mapping, RO-level gates, and state transitions

Dim g_Pass, g_Fail
g_Pass = 0
g_Fail = 0

' Global variables and functions extracted from Maintenance_RO_Closer for testing
Dim g_CurrentPageLineRecords

Set g_CurrentPageLineRecords = CreateObject("Scripting.Dictionary")

' ============================================================================
' EXTRACTED FUNCTIONS FROM MAINTENANCE_RO_CLOSER
' ============================================================================

Function CreateLineRecord(lineLetter, statusCode, rowNumber, description)
    Dim record
    Set record = CreateObject("Scripting.Dictionary")
    record("lineLetter") = lineLetter
    record("statusCode") = statusCode
    record("row") = rowNumber
    record("description") = description
    Set CreateLineRecord = record
End Function

Function GetLineStatus(lineLetter)
    If g_CurrentPageLineRecords.Exists(lineLetter) Then
        Dim record
        Set record = g_CurrentPageLineRecords(lineLetter)
        GetLineStatus = record("statusCode")
    Else
        GetLineStatus = ""
    End If
End Function

Function CreateStatusActionMap()
    Dim m
    Set m = CreateObject("Scripting.Dictionary")
    m("C92") = "REVIEW"
    m("C93") = "SKIP_REVIEWED"
    Set CreateStatusActionMap = m
End Function

Function GetLineActionFromStatus(statusCode)
    Dim statusMap, action
    Set statusMap = CreateStatusActionMap()
    
    ' First try exact match
    If statusMap.Exists(statusCode) Then
        GetLineActionFromStatus = statusMap(statusCode)
        Exit Function
    End If
    
    ' Then try pattern matching
    If Left(statusCode, 1) = "I" Then
        GetLineActionFromStatus = "FINISH_AND_REROUTE"
    ElseIf Left(statusCode, 1) = "H" Then
        GetLineActionFromStatus = "SKIP_RO_ON_HOLD"
    Else
        GetLineActionFromStatus = "SKIP_UNKNOWN"
    End If
End Function

Function CheckRoLineStatuses()
    Dim recordKey, record, allC93, hasHold
    Dim statusCode
    
    ' If no lines found, no early gate
    If g_CurrentPageLineRecords.Count = 0 Then
        CheckRoLineStatuses = ""
        Exit Function
    End If
    
    allC93 = True
    hasHold = False
    
    ' Scan all lines
    For Each recordKey In g_CurrentPageLineRecords.Keys
        Set record = g_CurrentPageLineRecords(recordKey)
        statusCode = record("statusCode")
        
        ' Check for hold (Hxx)
        If Left(statusCode, 1) = "H" Then
            hasHold = True
            Exit For
        End If
        
        ' Check if all are C93
        If statusCode <> "C93" Then
            allC93 = False
        End If
    Next
    
    ' Return early gate result
    If hasHold Then
        CheckRoLineStatuses = "HOLD_DETECTED"
    ElseIf allC93 Then
        CheckRoLineStatuses = "ALL_REVIEWED"
    Else
        CheckRoLineStatuses = ""
    End If
End Function

Sub AssertEqual(ByVal label, ByVal expected, ByVal actual)
    If CStr(expected) = CStr(actual) Then
        g_Pass = g_Pass + 1
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "FAIL: " & label & " | expected=[" & expected & "] actual=[" & actual & "]"
    End If
End Sub

Sub AssertTrue(ByVal label, ByVal actual)
    If CBool(actual) Then
        g_Pass = g_Pass + 1
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "FAIL: " & label & " | expected=true actual=false"
    End If
End Sub

' ==============================================================================
' TEST 1: Status -> Action Mapping
' ==============================================================================

Sub Test_StatusAction_C92()
    Dim action
    action = GetLineActionFromStatus("C92")
    AssertEqual "C92 maps to REVIEW", "REVIEW", action
End Sub

Sub Test_StatusAction_C93()
    Dim action
    action = GetLineActionFromStatus("C93")
    AssertEqual "C93 maps to SKIP_REVIEWED", "SKIP_REVIEWED", action
End Sub

Sub Test_StatusAction_Ixx()
    Dim action1, action2
    action1 = GetLineActionFromStatus("I123")
    action2 = GetLineActionFromStatus("I91")
    AssertEqual "I123 maps to FINISH_AND_REROUTE", "FINISH_AND_REROUTE", action1
    AssertEqual "I91 maps to FINISH_AND_REROUTE", "FINISH_AND_REROUTE", action2
End Sub

Sub Test_StatusAction_Hxx()
    Dim action1, action2
    action1 = GetLineActionFromStatus("H20")
    action2 = GetLineActionFromStatus("H99")
    AssertEqual "H20 maps to SKIP_RO_ON_HOLD", "SKIP_RO_ON_HOLD", action1
    AssertEqual "H99 maps to SKIP_RO_ON_HOLD", "SKIP_RO_ON_HOLD", action2
End Sub

Sub Test_StatusAction_Unknown()
    Dim action1, action2
    action1 = GetLineActionFromStatus("X99")
    action2 = GetLineActionFromStatus("Z01")
    AssertEqual "X99 maps to SKIP_UNKNOWN", "SKIP_UNKNOWN", action1
    AssertEqual "Z01 maps to SKIP_UNKNOWN", "SKIP_UNKNOWN", action2
End Sub

Sub Test_StatusAction_Empty()
    Dim action
    action = GetLineActionFromStatus("")
    AssertEqual "Empty status maps to SKIP_UNKNOWN", "SKIP_UNKNOWN", action
End Sub

' ==============================================================================
' TEST 2: Line Record Creation and Line Status Query
' ==============================================================================

Sub Test_CreateLineRecord()
    Dim record
    Set record = CreateLineRecord("A", "C92", 10, "TEST DESCRIPTION")
    AssertEqual "Line letter", "A", record("lineLetter")
    AssertEqual "Status code", "C92", record("statusCode")
    AssertEqual "Row", 10, record("row")
    AssertEqual "Description", "TEST DESCRIPTION", record("description")
End Sub

Sub Test_GetLineStatus_Found()
    ' Clear and populate records
    Set g_CurrentPageLineRecords = CreateObject("Scripting.Dictionary")
    Set g_CurrentPageLineRecords("A") = CreateLineRecord("A", "C92", 10, "TEST")
    
    Dim status
    status = GetLineStatus("A")
    AssertEqual "Status for line A", "C92", status
End Sub

Sub Test_GetLineStatus_NotFound()
    ' Clear records
    Set g_CurrentPageLineRecords = CreateObject("Scripting.Dictionary")
    
    Dim status
    status = GetLineStatus("Z")
    AssertEqual "Status for non-existent line", "", status
End Sub

Sub Test_LineRecordsByLetter()
    ' Populate multiple records
    Set g_CurrentPageLineRecords = CreateObject("Scripting.Dictionary")
    Set g_CurrentPageLineRecords("A") = CreateLineRecord("A", "C92", 10, "TIRE PRESSURE")
    Set g_CurrentPageLineRecords("B") = CreateLineRecord("B", "I91", 11, "TIRE TREAD")
    Set g_CurrentPageLineRecords("C") = CreateLineRecord("C", "H20", 12, "BATTERY TEST")
    
    AssertEqual "Line A status", "C92", GetLineStatus("A")
    AssertEqual "Line B status", "I91", GetLineStatus("B")
    AssertEqual "Line C status", "H20", GetLineStatus("C")
End Sub

' ==============================================================================
' TEST 3: RO-Level Gate Logic
' ==============================================================================

Sub Test_CheckRoLineStatuses_AllC93()
    ' Simulate RO with all lines C93 (already reviewed)
    Set g_CurrentPageLineRecords = CreateObject("Scripting.Dictionary")
    Set g_CurrentPageLineRecords("A") = CreateLineRecord("A", "C93", 10, "LINE A")
    Set g_CurrentPageLineRecords("B") = CreateLineRecord("B", "C93", 11, "LINE B")
    
    Dim gateResult
    gateResult = CheckRoLineStatuses()
    AssertEqual "All C93 -> ALL_REVIEWED gate", "ALL_REVIEWED", gateResult
End Sub

Sub Test_CheckRoLineStatuses_HoldDetected()
    ' Simulate RO with hold
    Set g_CurrentPageLineRecords = CreateObject("Scripting.Dictionary")
    Set g_CurrentPageLineRecords("A") = CreateLineRecord("A", "C92", 10, "LINE A")
    Set g_CurrentPageLineRecords("B") = CreateLineRecord("B", "H20", 11, "LINE B")
    
    Dim gateResult
    gateResult = CheckRoLineStatuses()
    AssertEqual "Hold detected -> HOLD_DETECTED gate", "HOLD_DETECTED", gateResult
End Sub

Sub Test_CheckRoLineStatuses_Mixed()
    ' Simulate RO with mixed statuses (not all reviewed, no hold)
    Set g_CurrentPageLineRecords = CreateObject("Scripting.Dictionary")
    Set g_CurrentPageLineRecords("A") = CreateLineRecord("A", "C92", 10, "LINE A")
    Set g_CurrentPageLineRecords("B") = CreateLineRecord("B", "I91", 11, "LINE B")
    
    Dim gateResult
    gateResult = CheckRoLineStatuses()
    AssertEqual "Mixed statuses -> no early gate", "", gateResult
End Sub

Sub Test_CheckRoLineStatuses_Unknown()
    ' Simulate RO with unknown status (not C92/C93/I/H)
    Set g_CurrentPageLineRecords = CreateObject("Scripting.Dictionary")
    Set g_CurrentPageLineRecords("A") = CreateLineRecord("A", "X99", 10, "LINE A")
    
    Dim gateResult
    gateResult = CheckRoLineStatuses()
    ' Unknown status alone does not trigger RO-level gate
    AssertEqual "Unknown status -> no early gate", "", gateResult
End Sub

Sub Test_CheckRoLineStatuses_Empty()
    ' Simulate RO with no lines
    Set g_CurrentPageLineRecords = CreateObject("Scripting.Dictionary")
    
    Dim gateResult
    gateResult = CheckRoLineStatuses()
    AssertEqual "No lines -> no early gate", "", gateResult
End Sub

' ==============================================================================
' RUN TESTS
' ==============================================================================

WScript.Echo "Line Status Detection and Routing Tests"
WScript.Echo "========================================"
WScript.Echo ""

WScript.Echo "TEST GROUP: Status -> Action Mapping"
Test_StatusAction_C92
Test_StatusAction_C93
Test_StatusAction_Ixx
Test_StatusAction_Hxx
Test_StatusAction_Unknown
Test_StatusAction_Empty

WScript.Echo ""
WScript.Echo "TEST GROUP: Line Record Creation and Query"
Test_CreateLineRecord
Test_GetLineStatus_Found
Test_GetLineStatus_NotFound
Test_LineRecordsByLetter

WScript.Echo ""
WScript.Echo "TEST GROUP: RO-Level Gate Logic"
Test_CheckRoLineStatuses_AllC93
Test_CheckRoLineStatuses_HoldDetected
Test_CheckRoLineStatuses_Mixed
Test_CheckRoLineStatuses_Unknown
Test_CheckRoLineStatuses_Empty

WScript.Echo ""
WScript.Echo "========================================"
If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All " & g_Pass & " tests passed."
    WScript.Quit 0
Else
    WScript.Echo "FAIL: " & g_Fail & " test(s) failed, " & g_Pass & " passed."
    WScript.Quit 1
End If
