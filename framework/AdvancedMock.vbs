'=====================================================================================
' AdvancedMock.vbs - Robust BlueZone Terminal Mocking Framework
'=====================================================================================
' Purpose: Simulate terminal interactions with advanced "Real World" failure modes:
'   - Latency/Timeouts
'   - Partial Screen Loads
'   - Modal/Error Interference
'   - Buffer Fuzzing
'=====================================================================================

Class AdvancedMock
    Private m_Buffer
    Private m_Connected
    Private m_SentKeys
    Private m_LastReadTime
    Private m_LatencyMs
    Private m_PartialLoadRows
    Private m_InterferenceText
    Private m_InterferencePos
    Private m_PromptSequence
    Private m_PromptIndex
    
    Private Sub Class_Initialize()
        m_Buffer = String(24 * 80, " ")
        m_Connected = False
        m_SentKeys = ""
        m_LatencyMs = 0
        m_PartialLoadRows = 24 ' Default to full load
        m_InterferenceText = ""
        m_InterferencePos = 0
        m_PromptSequence = Array()
        m_PromptIndex = -1
    End Sub

    ' --- Configuration Methods ---
    
    Public Sub SetBuffer(content)
        m_Buffer = Left(content & String(24 * 80, " "), 24 * 80)
    End Sub

    Public Sub SetPromptSequence(promptArray)
        m_PromptSequence = promptArray
        m_PromptIndex = 0
        If UBound(m_PromptSequence) >= 0 Then
            UpdateBufferFromSequence()
        End If
    End Sub

    Private Sub UpdateBufferFromSequence()
        If m_PromptIndex >= 0 And m_PromptIndex <= UBound(m_PromptSequence) Then
            Dim prompt: prompt = m_PromptSequence(m_PromptIndex)
            ' Place prompt at Row 23 by default for CDK compatibility, or custom if it contains newline
            If InStr(1, prompt, vbCrLf) > 0 Then
                SetBuffer prompt
            Else
                ' Default to main prompt line (Row 23)
                Dim buf: buf = String(24 * 80, " ")
                Dim pos: pos = ((23 - 1) * 80) + 1
                buf = Left(buf, pos - 1) & prompt & Mid(buf, pos + Len(prompt))
                SetBuffer buf
            End If
        End If
    End Sub

    Public Sub SetLatency(ms)
        m_LatencyMs = ms
    End Sub

    Public Sub SetPartialLoad(rowCount)
        m_PartialLoadRows = rowCount
    End Sub

    Public Sub InjectInterference(text, row, col)
        m_InterferenceText = text
        m_InterferencePos = ((row - 1) * 80) + (col - 1) + 1
    End Sub

    ' --- BlueZone Compatibility Interface ---

    Public Sub Connect(session)
        m_Connected = True
        ' Log to console if available
        On Error Resume Next
        WScript.Echo "[AdvancedMock] Connected to session: " & session
        On Error GoTo 0
    End Sub

    Public Sub Disconnect()
        m_Connected = False
    End Sub

    Public Function IsConnected()
        IsConnected = m_Connected
    End Function

    Public Sub ReadScreen(ByRef content, length, row, col)
        If Not m_Connected Then
            content = ""
            Exit Sub
        End If

        ' Simulate Latency
        If m_LatencyMs > 0 Then
            WaitMs m_LatencyMs
        End If

        Dim startPos: startPos = ((row - 1) * 80) + (col - 1) + 1
        Dim effectiveBuffer: effectiveBuffer = m_Buffer

        ' Apply Partial Load (Clear rows beyond limit)
        If m_PartialLoadRows < 24 Then
            Dim limitPos: limitPos = m_PartialLoadRows * 80
            If startPos > limitPos Then
                content = String(length, " ")
                Exit Sub
            End If
            ' Truncate buffer for the read
            effectiveBuffer = Left(m_Buffer, limitPos) & String((24 - m_PartialLoadRows) * 80, " ")
        End If

        ' Apply Interference
        If m_InterferenceText <> "" Then
            Dim prefix: prefix = Left(effectiveBuffer, m_InterferencePos - 1)
            Dim suffix: suffix = Mid(effectiveBuffer, m_InterferencePos + Len(m_InterferenceText))
            effectiveBuffer = Left(prefix & m_InterferenceText & suffix & String(24 * 80, " "), 24 * 80)
        End If

        content = Mid(effectiveBuffer, startPos, length)
    End Sub

    Public Sub SendKey(key)
        If Not m_Connected Then Exit Sub
        m_SentKeys = m_SentKeys & key & "|"
        
        ' Auto-advance prompt sequence on Enter
        If (key = "<NumpadEnter>" Or key = "<Enter>") And m_PromptIndex >= 0 Then
            m_PromptIndex = m_PromptIndex + 1
            If m_PromptIndex <= UBound(m_PromptSequence) Then
                UpdateBufferFromSequence()
            End If
        End If
    End Sub

    Public Sub Pause(ms)
        WaitMs ms
    End Sub

    ' --- Mock Inspection Methods ---

    Public Function GetSentKeys()
        GetSentKeys = m_SentKeys
    End Function

    Public Sub ClearSentKeys()
        m_SentKeys = ""
    End Sub

    ' --- Utilities ---
    
    Private Sub WaitMs(ms)
        If ms <= 0 Then Exit Sub
        Dim endTime: endTime = Timer + (ms / 1000)
        Do While Timer < endTime
            ' Busy wait for mock simplicity
        Loop
    End Sub
End Class
