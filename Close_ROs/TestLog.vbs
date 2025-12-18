Option Explicit

Dim g_LogBuffer
Dim LOG_FILE_PATH
Dim bzhao

' --- Configuration ---
LOG_FILE_PATH = "C:\Temp\Code\Scripts\VBScript\CDK\Close_ROs\TestLog.log"

' --- Main Execution ---
On Error Resume Next
Set bzhao = CreateObject("BZWhll.WhllObj")
If Err.Number <> 0 Then
    MsgBox "ERROR: Failed to create BZWhll.WhllObj. Ensure BlueZone is installed and running. " & Err.Description, vbCritical, "TestLog Error"
    bzhao.StopScript ' Use bzhao.StopScript for BlueZone compatibility
End If
On Error GoTo 0

g_LogBuffer = "" ' Initialize global log buffer

Call LogResult("INFO", "Test log entry 1: Script started.")
Call LogResult("DEBUG", "Test log entry 2: Debugging is fun!")
Call LogResult("ERROR", "Test log entry 3: Something went wrong.")

Call FlushLogBuffer()

bzhao.MsgBox "TestLog.vbs finished. Check " & LOG_FILE_PATH

Set bzhao = Nothing

' --- Subroutines ---

Sub LogResult(level, message)
    Dim logLine
    logLine = Now & " [" & level & "] " & message
    g_LogBuffer = g_LogBuffer & logLine & vbCrLf
End Sub

Sub FlushLogBuffer()
    If Len(g_LogBuffer) > 0 Then
        Dim fso, logFile
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        On Error Resume Next ' Enable error handling for file operations
        Set logFile = fso.OpenTextFile(LOG_FILE_PATH, 8, True) ' 8 = ForAppending, True = Create if not exists
        If Err.Number <> 0 Then
            bzhao.MsgBox "ERROR in FlushLogBuffer (OpenTextFile): " & Err.Description & " (Error #" & Err.Number & ")"
            Err.Clear ' Clear the error
            Set logFile = Nothing
            Set fso = Nothing
            Exit Sub ' Exit subroutine on error
        End If

        On Error GoTo 0 ' Disable Resume Next for the write operation to catch errors immediately
        logFile.Write g_LogBuffer
        
        On Error Resume Next ' Re-enable for close
        logFile.Close
        If Err.Number <> 0 Then
            bzhao.MsgBox "ERROR in FlushLogBuffer (Close): " & Err.Description & " (Error #" & Err.Number & ")"
            Err.Clear ' Clear the error
        End If
        On Error GoTo 0 ' Disable Resume Next

        g_LogBuffer = "" ' Clear buffer after writing
        Set logFile = Nothing
        Set fso = Nothing
    End If
End Sub