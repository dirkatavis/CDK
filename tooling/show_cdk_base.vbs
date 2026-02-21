Option Explicit

' =====================================================================
' show_cdk_base.vbs - Show current CDK_BASE value
' =====================================================================

Dim sh: Set sh = CreateObject("WScript.Shell")
Dim val: val = sh.Environment("USER")("CDK_BASE")

If val = "" Then
    MsgBox "CDK_BASE is not set for this user.", vbExclamation, "CDK_BASE"
Else
    MsgBox "CDK_BASE = " & val, vbInformation, "CDK_BASE"
End If
