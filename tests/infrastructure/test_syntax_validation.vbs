' ============================================================================
' CDK Syntax & Environment Validator
' Purpose: Catches "DoEvents" errors and other VBScript/WSH incompatibilities
'          that aren't caught by logical unit tests.
' ============================================================================

Option Explicit

Dim g_fso, g_shell, g_repoRoot
Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_shell = CreateObject("WScript.Shell")

' --- Bootstrap ---
g_repoRoot = g_shell.Environment("USER")("CDK_BASE")
If g_repoRoot = "" Then WScript.Quit 1

WScript.Echo "============================================================================="
WScript.Echo "ENVIRONMENT & SYNTAX SCAN"
WScript.Echo "============================================================================="

Dim apps, app, failed
apps = Array( _
    "apps\post_final_charges\PostFinalCharges.vbs", _
    "apps\pfc_scrapper\PFC_Scrapper.vbs", _
    "apps\maintenance_ro_closer\Maintenance_RO_Closer.vbs" _
)

failed = False

For Each app In apps
    Dim fullPath: fullPath = g_fso.BuildPath(g_repoRoot, app)
    Dim padCount: padCount = 40 - Len(app)
    If padCount < 1 Then padCount = 1
    WScript.StdOut.Write "  Checking " & app & " " & String(padCount, ".") & " "
    
    If Not g_fso.FileExists(fullPath) Then
        WScript.Echo "[SKIP] File missing"
    Else
        If CheckForIncompatibilities(fullPath) Then
            WScript.Echo "[FAIL]"
            failed = True
        Else
            WScript.Echo "[PASS]"
        End If
    End If
Next

If failed Then
    WScript.Echo "============================================================================="
    WScript.Echo "[ERROR] Incompatibilities found. System unstable for CLI execution."
    WScript.Quit 1
Else
    WScript.Echo "============================================================================="
    WScript.Echo "[SUCCESS] No illegal environment keywords found."
    WScript.Quit 0
End If

Function CheckForIncompatibilities(path)
    Dim ts: Set ts = g_fso.OpenTextFile(path, 1)
    Dim content: content = ts.ReadAll: ts.Close
    
    Dim hasError: hasError = False
    
    ' Pattern 1: DoEvents (VBA keyword, undefined in WSH)
    If InStr(content, "DoEvents") > 0 Then
        ' Simple check for commented out line
        If InStr(content, "' DoEvents") = 0 And InStr(content, "'DoEvents") = 0 Then
            WScript.Echo ""
            WScript.Echo "    ! Error: Found active 'DoEvents' - Not supported in WSH scripts."
            hasError = True
        End If
    End If
    
    ' Pattern 2: MsgBox without /nologo detection (blocks CI)
    ' (Future proofing)
    
    ' Pattern 3: Option Explicit Placement (Must be before any logic)
    If Not CheckOptionExplicitPlacement(content) Then
        WScript.Echo ""
        WScript.Echo "    ! Error: 'Option Explicit' missing or not at top of file."
        hasError = True
    End If
    
    CheckForIncompatibilities = hasError
End Function

Function CheckOptionExplicitPlacement(content)
    Dim lines: lines = Split(content, vbCrLf)
    Dim i, line
    Dim foundOptionExplicit: foundOptionExplicit = False
    
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        
        ' Skip empty lines and comments
        If line <> "" And Left(line, 1) <> "'" Then
            ' If the first non-comment non-empty line isn't Option Explicit, it's a fail
            If InStr(UCase(line), "OPTION EXPLICIT") = 1 Then
                CheckOptionExplicitPlacement = True
                Exit Function
            Else
                ' Found executable code before Option Explicit
                CheckOptionExplicitPlacement = False
                Exit Function
            End If
        End If
    Next
    
    CheckOptionExplicitPlacement = False ' Never found it
End Function

Function IsInComment(content, word)
    ' Crude check: is there a ' earlier on the same line?
    ' For a syntax scanner, we'll keep it simple for now.
    IsInComment = False 
End Function
