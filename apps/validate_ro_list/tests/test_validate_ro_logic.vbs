' ============================================================================
' ValidateRoList Logic Test
' Purpose: Exercises ValidateRoList.vbs in Mock Mode to verify its core logic.
' ============================================================================

Option Explicit

Dim g_fso, g_shell, g_repoRoot
Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_shell = CreateObject("WScript.Shell")

' --- Bootstrap ---
g_repoRoot = g_shell.Environment("USER")("CDK_BASE")
If g_repoRoot = "" Then WScript.Quit 1

' Prepare mock files
Dim scriptDir: scriptDir = g_fso.GetParentFolderName(WScript.ScriptFullName)
Dim testDir: testDir = g_fso.BuildPath(scriptDir, "test_artifacts")
If Not g_fso.FolderExists(testDir) Then GeneratePath(testDir)

Dim mockInput: mockInput = g_fso.BuildPath(testDir, "test_input.csv")
Dim mockOutput: mockOutput = g_fso.BuildPath(testDir, "test_output.txt")
Dim mockMap1: mockMap1 = g_fso.BuildPath(testDir, "mock_map_1.txt")
Dim mockMap2: mockMap2 = g_fso.BuildPath(testDir, "mock_map_2.txt")

' 1. Create Mock Screen Content (Simulating BlueZone screens)
Dim ts: Set ts = g_fso.CreateTextFile(mockMap1, True)
ts.WriteLine "Screen Content Here"
ts.WriteLine "Status Message: RO: 123456" ' Should match "RO:" -> "Open"
ts.Close

Set ts = g_fso.CreateTextFile(mockMap2, True)
ts.WriteLine "ERROR: NOT ON FILE" ' Should match "NOT ON FILE"
ts.Close

' 3. Run the app in mock mode
Dim env: Set env = g_shell.Environment("PROCESS")
env("MOCK_VALIDATE_RO") = "true"
env("MOCK_SCREEN_MAPS") = mockMap1 & ";" & mockMap2
env("MOCK_OUTPUT_FILE") = mockOutput

Dim appPath: appPath = g_fso.BuildPath(g_repoRoot, "apps\validate_ro_list\ValidateRoList.vbs")
Dim cmd: cmd = "cscript.exe //nologo " & Chr(34) & appPath & Chr(34)

WScript.Echo "Running ValidateRoList in Mock Mode..."
Dim exec: Set exec = g_shell.Exec(cmd)
Do While exec.Status = 0
    WScript.Sleep 10
Loop
WScript.Echo "APP: " & exec.StdOut.ReadAll()

' 4. Verify Output
' ... (cleanup) ...

' Read output file and verify contents
If Not g_fso.FileExists(mockOutput) Then
    WScript.Echo "[FAIL] Output file not created"
    WScript.Quit 1
End If

Dim outContent: outContent = g_fso.OpenTextFile(mockOutput, 1).ReadAll
If InStr(outContent, "mock_map_1.txt,Open") > 0 And _
   InStr(outContent, "mock_map_2.txt,NOT ON FILE") > 0 Then
    WScript.Echo "[PASS] ValidateRoList logic execution verified"
    WScript.Quit 0
Else
    WScript.Echo "[FAIL] Output content mismatch"
    WScript.Echo "Actual output:" & vbCrLf & outContent
    WScript.Quit 1
End If



Sub GeneratePath(p)
    If Not g_fso.FolderExists(g_fso.GetParentFolderName(p)) Then GeneratePath(g_fso.GetParentFolderName(p))
    If Not g_fso.FolderExists(p) Then g_fso.CreateFolder(p)
End Sub
