Option Explicit

' Production-backed tests for Maintenance_RO_Closer line routing and review scan behavior.

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
    Private m_pages
    Private m_pageIndex
    Private m_pending
    Private m_commands

    Private Sub Class_Initialize()
        m_pageIndex = 0
        m_pending = ""
        m_commands = ""
    End Sub

    Public Sub SetPages(ByRef pages)
        m_pages = pages
        m_pageIndex = 0
        m_pending = ""
        m_commands = ""
    End Sub

    Public Sub ReadScreen(ByRef outText, ByVal length, ByVal row, ByVal col)
        Dim pageText, pos
        pageText = CStr(m_pages(m_pageIndex))
        pos = ((row - 1) * 80) + col
        If pos < 1 Then pos = 1
        outText = Mid(pageText, pos, length)
    End Sub

    Public Sub Pause(ByVal ms)
    End Sub

    Public Sub SendKey(ByVal keyText)
        If keyText = "<NumpadEnter>" Then
            ExecutePendingCommand
            Exit Sub
        End If

        m_pending = m_pending & CStr(keyText)
    End Sub

    Public Sub Disconnect()
    End Sub

    Public Sub StopScript()
    End Sub

    Public Function CommandsCsv()
        CommandsCsv = m_commands
    End Function

    Private Sub ExecutePendingCommand()
        Dim cmd
        cmd = UCase(Trim(m_pending))

        If cmd <> "" Then
            If m_commands <> "" Then
                m_commands = m_commands & "|"
            End If
            m_commands = m_commands & cmd

            If cmd = "N" Then
                If m_pageIndex < UBound(m_pages) Then
                    m_pageIndex = m_pageIndex + 1
                End If
            End If
        End If

        m_pending = ""
    End Sub
End Class

Sub AssertEqual(ByVal label, ByVal expected, ByVal actual)
    If CStr(expected) = CStr(actual) Then
        g_Pass = g_Pass + 1
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "FAIL: " & label & " | expected=[" & expected & "] actual=[" & actual & "]"
    End If
End Sub

Sub AssertTrue(ByVal label, ByVal value)
    If CBool(value) Then
        g_Pass = g_Pass + 1
    Else
        g_Fail = g_Fail + 1
        WScript.Echo "FAIL: " & label
    End If
End Sub

Sub AssertFalse(ByVal label, ByVal value)
    AssertTrue label, (Not value)
End Sub

Function ContainsToken(ByVal csvText, ByVal token)
    ContainsToken = (InStr(1, "|" & UCase(csvText) & "|", "|" & UCase(token) & "|", vbTextCompare) > 0)
End Function

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

Function BuildPageWithSingleLine(ByVal lineLetter, ByVal lineStatus, ByVal laborTypeCode, ByVal description, ByVal endOfDisplay)
    Dim pageBuf, row10, row24
    pageBuf = String(24 * 80, " ")

    row10 = String(80, " ")
    row10 = SetColText(row10, 1, lineLetter)
    row10 = SetColText(row10, 4, "L1 " & description)
    row10 = SetColText(row10, 42, lineStatus)
    row10 = SetColText(row10, 50, laborTypeCode)
    pageBuf = SetRow(pageBuf, 10, row10)

    row24 = String(80, " ")
    row24 = SetColText(row24, 1, "COMMAND:")
    If endOfDisplay Then
        row24 = SetColText(row24, 30, "(END OF DISPLAY)")
    End If
    pageBuf = SetRow(pageBuf, 24, row24)

    BuildPageWithSingleLine = pageBuf
End Function

Sub ConfigureTestEnvironment(ByRef fake)
    Set g_bzhao = fake
    DEBUG_LEVEL = 0
    STABILITY_PAUSE = 0
    LOOP_PAUSE = 0
    REVIEW_PAUSE = 0
    BLACKLIST_TERMS = ""
    WARRANTY_LTYPES_RAW = "WCH,WF"
    InitializeSupportedWarrantyLTypes
End Sub

Sub Test_StatusActionMap_Production()
    AssertEqual "C92 maps to REVIEW", "REVIEW", GetLineActionFromStatus("C92")
    AssertEqual "C93 maps to SKIP_REVIEWED", "SKIP_REVIEWED", GetLineActionFromStatus("C93")
    AssertEqual "Ixx maps to FINISH_AND_REROUTE", "FINISH_AND_REROUTE", GetLineActionFromStatus("I91")
    AssertEqual "Hxx maps to SKIP_RO_ON_HOLD", "SKIP_RO_ON_HOLD", GetLineActionFromStatus("H20")
End Sub

Sub Test_ProcessRoReview_MultiPage_QueuesNonC93Only()
    Dim fake, pages(1), ok, commands
    Set fake = New FakeBzhao

    pages(0) = BuildPageWithSingleLine("A", "C93", "WCH", "PAGE1", False)
    pages(1) = BuildPageWithSingleLine("C", "C92", "WCH", "PAGE2", True)
    fake.SetPages pages

    ConfigureTestEnvironment fake

    ok = ProcessRoReview()
    commands = fake.CommandsCsv()

    AssertTrue "ProcessRoReview returns success on actionable multipage", ok
    AssertEqual "Review phase result", "PROCEED", g_ReviewPhaseResult
    AssertTrue "Sends N to scan next page", ContainsToken(commands, "N")
    AssertTrue "Reviews line on page 2", ContainsToken(commands, "R C")
    AssertFalse "Does not review C93 line from page 1", ContainsToken(commands, "R A")
    AssertFalse "Does not skip review phase", ContainsToken(commands, "E")
End Sub

Sub Test_ProcessRoReview_MultiPage_AllC93Skips()
    Dim fake, pages(1), ok, commands
    Set fake = New FakeBzhao

    pages(0) = BuildPageWithSingleLine("A", "C93", "WCH", "PAGE1", False)
    pages(1) = BuildPageWithSingleLine("B", "C93", "WCH", "PAGE2", True)
    fake.SetPages pages

    ConfigureTestEnvironment fake

    ok = ProcessRoReview()
    commands = fake.CommandsCsv()

    AssertFalse "All-C93 multipage returns skipped", ok
    AssertEqual "Review phase result for all-C93", "SKIPPED", g_ReviewPhaseResult
    AssertTrue "Sends N before deciding all-C93", ContainsToken(commands, "N")
    AssertTrue "Sends E to exit skipped review", ContainsToken(commands, "E")
End Sub

Sub Test_ProcessRoReview_UnsupportedWarrantySkips()
    Dim fake, pages(0), ok, commands
    Set fake = New FakeBzhao

    pages(0) = BuildPageWithSingleLine("A", "C92", "WZZ", "WARRANTY", True)
    fake.SetPages pages

    ConfigureTestEnvironment fake

    ok = ProcessRoReview()
    commands = fake.CommandsCsv()

    AssertFalse "Unsupported warranty gate returns skipped", ok
    AssertEqual "Review phase result for unsupported warranty", "SKIPPED", g_ReviewPhaseResult
    AssertTrue "Sends E for unsupported warranty skip", ContainsToken(commands, "E")
    AssertFalse "Does not attempt review when unsupported warranty found", ContainsToken(commands, "R A")
End Sub

Dim scriptPath, fileContent, scriptStream
scriptPath = g_fso.BuildPath(g_repoRoot, "apps\maintenance_ro_closer\Maintenance_RO_Closer.vbs")
Set scriptStream = g_fso.OpenTextFile(scriptPath)
fileContent = scriptStream.ReadAll
scriptStream.Close
fileContent = Replace(fileContent, "Set g_bzhao = CreateObject(""BZWhll.WhllObj"")", "Set g_bzhao = Nothing")
fileContent = Replace(fileContent, vbCrLf & "' Execute" & vbCrLf & "RunAutomation", vbCrLf & "' Execute disabled during tests")
ExecuteGlobal fileContent

WScript.Echo "Maintenance RO line routing tests"
WScript.Echo "================================"

Test_StatusActionMap_Production
Test_ProcessRoReview_MultiPage_QueuesNonC93Only
Test_ProcessRoReview_MultiPage_AllC93Skips
Test_ProcessRoReview_UnsupportedWarrantySkips

WScript.Echo ""
If g_Fail = 0 Then
    WScript.Echo "SUCCESS: All " & g_Pass & " tests passed."
    WScript.Quit 0
Else
    WScript.Echo "FAIL: " & g_Fail & " test(s) failed, " & g_Pass & " passed."
    WScript.Quit 1
End If
