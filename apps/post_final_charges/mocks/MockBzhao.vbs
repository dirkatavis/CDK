' Redirect to central AdvancedMock
Class MockBzhao
    Private m_mock
    Private Sub Class_Initialize()
        Dim fsoLocal: Set fsoLocal = CreateObject("Scripting.FileSystemObject")
        Dim centralMock: centralMock = fsoLocal.BuildPath(fsoLocal.GetParentFolderName(fsoLocal.GetParentFolderName(fsoLocal.GetParentFolderName(fsoLocal.GetParentFolderName(WScript.ScriptFullName)))), "framework\AdvancedMock.vbs")
        If Not fsoLocal.FileExists(centralMock) Then centralMock = "C:\Temp_alt\CDK\framework\AdvancedMock.vbs"
        ExecuteGlobal fsoLocal.OpenTextFile(centralMock).ReadAll
        Set m_mock = New AdvancedMock
    End Sub
    Public Sub Connect(session) : m_mock.Connect session : End Sub
    Public Sub ReadScreen(ByRef buf, len, row, col) : m_mock.ReadScreen buf, len, row, col : End Sub
    Public Sub SendKey(key) : m_mock.SendKey key : End Sub
    Public Sub SetupTestScenario(name) : ' Compatibility stub : End Sub
    Public Sub Pause(ms) : m_mock.Pause ms : End Sub
    Public Function IsConnected() : IsConnected = m_mock.IsConnected() : End Function
    Public Function GetSentKeys() : GetSentKeys = m_mock.GetSentKeys() : End Function
End Class