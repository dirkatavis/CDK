'====================================================================
' Script: HighestRoFinder.vbs - Binary Search Version
' Purpose: Find highest valid sequence number using binary search
'====================================================================

Option Explicit

' --- Load PathHelper for centralized path management ---
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Const BASE_ENV_VAR_LOCAL = "CDK_BASE"

Function FindRepoRootForBootstrap()
    Dim sh: Set sh = CreateObject("WScript.Shell")
    Dim basePath: basePath = sh.Environment("USER")(BASE_ENV_VAR_LOCAL)
    If basePath = "" Or Not g_fso.FolderExists(basePath) Then
        Err.Raise 53, "Bootstrap", "Invalid or missing CDK_BASE. Value: " & basePath
    End If
    If Not g_fso.FileExists(g_fso.BuildPath(basePath, ".cdkroot")) Then
        Err.Raise 53, "Bootstrap", "Cannot find .cdkroot in base path:" & vbCrLf & basePath
    End If
    FindRepoRootForBootstrap = basePath
End Function

Dim helperPath: helperPath = g_fso.BuildPath(FindRepoRootForBootstrap(), "framework\PathHelper.vbs")
ExecuteGlobal g_fso.OpenTextFile(helperPath).ReadAll

' --- Configuration ---
Dim LOG_FILE_PATH: LOG_FILE_PATH = GetConfigPath("HighestRoFinder", "Log")

' Constants
Const MIN_NUMBER = 0
Const MAX_NUMBER = 3000
Const SCREEN_SIZE = 1920

' Variables
Dim bzhao
Dim foundError
Dim low, high, mid, lastValid
Dim fsoLog, logFile
Dim SpeedMode: SpeedMode = 1
Set fsoLog = CreateObject("Scripting.FileSystemObject")
Set logFile = fsoLog.OpenTextFile(LOG_FILE_PATH, 8, True)


' Initialize
If Not Initialize() Then
    bzhao.StopScript
End If

' Binary search algorithm
low = MIN_NUMBER
high = MAX_NUMBER
lastValid = -1
LogResult "=============================================="
LogResult "          ***Starting New Run***"
LogResult "=============================================="
LogResult "Starting binary search between " & MIN_NUMBER & " and " & MAX_NUMBER
Do While low <= high
    mid = Int((low + high) / 2)
    
    ' Test the current number
    bzhao.SendKey CStr(mid)
    bzhao.SendKey "<NumpadEnter>"
    bzhao.Wait SpeedMode
    ' Use new screen search method for error detection
    foundError = FindStringOnScreen("DOES NOT EXIST")
    Dim roFoundMsg
    If foundError Then
        roFoundMsg = "ROFound: NO"
    Else
        roFoundMsg = "ROFound: YES"
    End If
    LogResult "RO Search: " & mid & " | " & roFoundMsg
    If foundError Then
        high = mid - 1
    Else
        bzhao.SendKey "E"
        bzhao.SendKey "<NumpadEnter>"
        bzhao.Wait SpeedMode
        lastValid = mid
        low = mid + 1
    End If
Loop

' Show result
If lastValid >= 0 Then
    LogResult "RESULT: The highest valid sequence number is: " & CStr(lastValid)
Else
    LogResult "RESULT: No valid numbers found in range " & MIN_NUMBER & " to " & MAX_NUMBER
End If

' Disconnect
bzhao.Disconnect
logFile.Close


'----------------------------------------------------
' LogResult subroutine for logging results/errors
'----------------------------------------------------
Sub LogResult(logMsg)
    logFile.WriteLine Now & " | " & logMsg
End Sub

'----------------------------------------------------
' Finds a string on the screen using Host.Search and returns True if found, False otherwise
'----------------------------------------------------
Function FindStringOnScreen(stringToFind)
    Dim row, col
    row = 1
    col = 1
    bzhao.Search stringToFind, row, col
    FindStringOnScreen = (row > 0 And col > 0)
End Function

'----------------------------------------------------
' Initializes BlueZone and logger, returns True if successful, False otherwise
'----------------------------------------------------
Function Initialize()
    On Error Resume Next
    Set bzhao = CreateObject("BZWhll.WhllObj")
    ' Logger already initialized at top
    If bzhao Is Nothing Then
        LogResult "ERROR: Failed to create BlueZone object."
        Initialize = False
        Exit Function
    End If
    If bzhao.Connect("") <> 0 Then
        LogResult "ERROR: Failed to connect to BlueZone session."
        Initialize = False
        Exit Function
    End If
    On Error GoTo 0
    Initialize = True
End Function