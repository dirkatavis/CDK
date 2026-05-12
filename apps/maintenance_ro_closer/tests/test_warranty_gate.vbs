' ============================================================================
' Maintenance RO Closer Warranty Processing Test
' Validates warranty lines are not pre-skipped and dialog detection is available.
' ============================================================================

Option Explicit

Dim g_Pass, g_Fail, g_fso, g_shell, g_repoRoot
g_Pass = 0
g_Fail = 0

Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_shell = CreateObject("WScript.Shell")
g_repoRoot = g_shell.Environment("USER")("CDK_BASE")

If g_repoRoot = "" Then
    WScript.Echo "FAIL: CDK_BASE is not set"
    WScript.Quit 1
End If

Class FakeBzhao
    Private m_page

    Public Sub SetPage(ByVal pageText)
        m_page = pageText
    End Sub

    Public Sub ReadScreen(ByRef outText, ByVal length, ByVal row, ByVal col)
        Dim pos
        pos = ((row - 1) * 80) + col
        If pos < 1 Then pos = 1
        outText = Mid(m_page, pos, length)
    End Sub

    Public Sub Pause(ByVal ms)
    End Sub

    Public Sub SendKey(ByVal keyText)
    End Sub

    Public Sub Disconnect()
    End Sub

    Public Sub StopScript()
    End Sub
End Class

Sub AssertTrue(ByVal label, ByVal value)
    If value Then
        g_Pass = g_Pass + 1
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "FAIL: " & label
    End If
End Sub

Sub AssertFalse(ByVal label, ByVal value)
    AssertTrue label, (Not value)
End Sub

Sub AssertEqual(ByVal label, ByVal expected, ByVal actual)
    If CStr(expected) = CStr(actual) Then
        g_Pass = g_Pass + 1
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "FAIL: " & label & " | expected=[" & expected & "] actual=[" & actual & "]"
    End If
End Sub

Function SetColText(ByVal rowText, ByVal colNum, ByVal textValue)
    Dim base
    base = Left(rowText & String(80, " "), 80)
    SetColText = Left(Left(base, colNum - 1) & textValue & Mid(base, colNum + Len(textValue)) & String(80, " "), 80)
End Function

Function SetRow(ByVal pageBuf, ByVal rowNum, ByVal rowText)
    Dim pos
    pos = (rowNum - 1) * 80 + 1
    SetRow = Left(pageBuf, pos - 1) & Left(rowText & String(80, " "), 80) & Mid(pageBuf, pos + 80)
End Function

Function BuildStatusPage(ByVal statusText, ByVal laborTypeCode)
    Dim pageBuf, row5, row11
    pageBuf = String(24 * 80, " ")

    row5 = String(80, " ")
    row5 = SetColText(row5, 1, statusText)
    pageBuf = SetRow(pageBuf, 5, row5)

    row11 = String(80, " ")
    row11 = SetColText(row11, 4, "L1")
    row11 = SetColText(row11, 50, laborTypeCode)
    pageBuf = SetRow(pageBuf, 11, row11)

    BuildStatusPage = pageBuf
End Function

Dim scriptPath, fileContent
scriptPath = g_fso.BuildPath(g_repoRoot, "apps\maintenance_ro_closer\Maintenance_RO_Closer.vbs")
fileContent = g_fso.OpenTextFile(scriptPath).ReadAll
fileContent = Replace(fileContent, "Set g_bzhao = CreateObject(""BZWhll.WhllObj"")", "Set g_bzhao = Nothing")
fileContent = Replace(fileContent, vbCrLf & "' Execute" & vbCrLf & "RunAutomation", vbCrLf & "' Execute disabled during tests")
ExecuteGlobal fileContent

Dim fake
Set fake = New FakeBzhao
Set g_bzhao = fake
BLACKLIST_TERMS = ""
OLD_RO_DAYS_THRESHOLD = 999
DEBUG_LEVEL = 0
WARRANTY_LTYPES_RAW = "WCH,WF,W"
InitializeSupportedWarrantyLTypes

AssertEqual "Detect FCA warranty dialog", "FCA", DetectMaintenanceWarrantyDialogFromText("LABOR OP: L1")
AssertEqual "Detect W warranty dialog", "W", DetectMaintenanceWarrantyDialogFromText("FAILURE CODE:")
AssertEqual "Detect FORD warranty dialog", "FORD", DetectMaintenanceWarrantyDialogFromText("MODIFY FORD REPAIR TYPE INFORMATION")
AssertEqual "No warranty dialog detected on plain screen", "", DetectMaintenanceWarrantyDialogFromText("READY TO POST")

fake.SetPage BuildStatusPage("READY TO POST", "WCH")
AssertTrue "WCH READY TO POST RO is not pre-skipped", ShouldProcessRoByBusinessRules("111111")

fake.SetPage BuildStatusPage("READY TO POST", "WF")
AssertTrue "WF READY TO POST RO is not pre-skipped", ShouldProcessRoByBusinessRules("222222")

fake.SetPage BuildStatusPage("READY TO POST", "CP")
AssertTrue "Non-warranty READY TO POST RO still proceeds", ShouldProcessRoByBusinessRules("333333")

fake.SetPage BuildStatusPage("READY TO POST", "WZZ")
AssertFalse "Unsupported W* labor type is skipped", ShouldProcessRoByBusinessRules("444444")

If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All " & g_Pass & " maintenance warranty processing tests passed."
Else
    WScript.Echo "FAIL: " & g_Fail & " maintenance warranty processing test(s) failed."
    WScript.Quit 1
End If

Function DetectMaintenanceWarrantyDialogFromText(ByVal markerText)
    Dim fakePage
    fakePage = String(24 * 80, " ")
    fakePage = SetRow(fakePage, 20, markerText)
    fake.SetPage fakePage
    DetectMaintenanceWarrantyDialogFromText = DetectMaintenanceWarrantyDialog()
End Function