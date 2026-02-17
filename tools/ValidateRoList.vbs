Option Explicit

' ======================================================================
' ValidateRoList.vbs
' Reads a CSV of RO numbers and checks each RO in BlueZone.
' Writes results to utilities\ValidateRoList_Results.txt in format: RO,STATUS
' Expected statuses: "NOT ON FILE" or "(PFC) POST FINAL CHARGES"
' ======================================================================

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim sh: Set sh = CreateObject("WScript.Shell")

' --- Determine repo root and paths ---
Dim scriptPath: scriptPath = WScript.ScriptFullName
Dim scriptDir: scriptDir = fso.GetParentFolderName(scriptPath)
Dim repoRoot: repoRoot = fso.GetParentFolderName(scriptDir) ' tools/ -> repo root

Dim inputFile: inputFile = fso.BuildPath(repoRoot, "utilities\ValidateRoList_IN.csv")
Dim outputFile: outputFile = fso.BuildPath(repoRoot, "utilities\ValidateRoList_Results.txt")

If Not fso.FileExists(inputFile) Then
    MsgBox "ERROR: Input file not found: " & inputFile, vbCritical, "ValidateRoList"
    WScript.Quit 1
End If

' --- Prepare BlueZone object ---
Dim bzhao: Set bzhao = CreateObject("BZWhll.WhllObj")
On Error Resume Next
bzhao.Connect ""
If Err.Number <> 0 Then
    MsgBox "ERROR: Failed to connect to BlueZone terminal session. " & Err.Description, vbCritical, "ValidateRoList"
    WScript.Quit 1
End If
On Error GoTo 0

' --- Helper subs (copied pattern used across repo) ---
Sub PressKey(key)
    bzhao.SendKey key
    bzhao.Pause 100
End Sub

Sub EnterTextAndWait(text)
    bzhao.SendKey text
    bzhao.Pause 100
    Call PressKey("<NumpadEnter>")
    bzhao.Pause 500
End Sub

' Wait for any of the target texts to appear at bottom rows (23 or 24)
Function WaitForOneOf(targetsCSV, timeoutMs)
    Dim targets: targets = Split(targetsCSV, "|")
    Dim elapsed: elapsed = 0
    Dim col: col = 1
    Dim screenLength: screenLength = 80
    Dim buffer23, buffer24, screenBuffer, i

    Do
        bzhao.Pause 500
        elapsed = elapsed + 500
        bzhao.ReadScreen buffer23, screenLength, 23, col
        bzhao.ReadScreen buffer24, screenLength, 24, col
        screenBuffer = UCase(buffer23 & " " & buffer24)

        For i = 0 To UBound(targets)
            If InStr(screenBuffer, UCase(targets(i))) > 0 Then
                WaitForOneOf = Trim(targets(i))
                Exit Function
            End If
        Next

        If elapsed >= timeoutMs Then
            WaitForOneOf = "__TIMEOUT__"
            Exit Function
        End If
    Loop
End Function

' --- Process input file ---
Dim inTS: Set inTS = fso.OpenTextFile(inputFile, 1, False)
Dim outTS: Set outTS = fso.CreateTextFile(outputFile, True)

Dim line, ro, status, found
Dim timeoutMs: timeoutMs = 10000 ' 10 seconds per your choice
Dim targets: targets = "NOT ON FILE| (PFC) POST FINAL CHARGES"

Do Until inTS.AtEndOfStream
    line = inTS.ReadLine
    ro = Trim(line)
    If ro <> "" Then
        ' send RO to COMMAND prompt
        EnterTextAndWait ro

        ' wait for one of the two expected results
        found = WaitForOneOf(targets, timeoutMs)

        If found = "__TIMEOUT__" Then
            status = "TIMEOUT"
        Else
            ' Normalize match to canonical status strings
            If UCase(Trim(found)) = "NOT ON FILE" Then
                status = "NOT ON FILE"
            ElseIf InStr(UCase(found), "PFC") > 0 Then
                status = "(PFC) POST FINAL CHARGES"
            Else
                status = Trim(found)
            End If
        End If

        outTS.WriteLine ro & "," & status
    End If
Loop

inTS.Close
outTS.Close

MsgBox "ValidateRoList complete. Results: " & outputFile, vbInformation, "ValidateRoList"
WScript.Quit 0
