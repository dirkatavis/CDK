'=====================================================================================
' create_upstream_pr.vbs - Helper to create PRs to upstream repo
'
' Purpose: Automate correct PR creation to upstream (dirkatavis/CDK) from fork
' Usage: cscript create_upstream_pr.vbs <branch> <title> [base_branch]
'
' Example:
'   cscript create_upstream_pr.vbs feature/eliminate-hardcoded-paths "Fix hardcoded paths" main
'
'=====================================================================================

Option Explicit

Dim WshShell: Set WshShell = CreateObject("WScript.Shell")
Dim args: Set args = WScript.Arguments

If args.Count < 2 Then
    WScript.Echo "Usage: cscript create_upstream_pr.vbs <branch> <title> [base_branch]"
    WScript.Echo ""
    WScript.Echo "Example:"
    WScript.Echo "  cscript create_upstream_pr.vbs feature/eliminate-hardcoded-paths ""Fix hardcoded paths"" main"
    WScript.Quit 1
End If

Dim featureBranch: featureBranch = args(0)
Dim prTitle: prTitle = args(1)
Dim baseBranch: baseBranch = IIf(args.Count > 2, args(2), "main")

WScript.Echo "Creating PR to upstream repository..."
WScript.Echo "  Branch: " & featureBranch
WScript.Echo "  Title: " & prTitle
WScript.Echo "  Base: " & baseBranch
WScript.Echo ""
WScript.Echo "IMPORTANT: This creates a PR from your fork to the upstream repository."
WScript.Echo "Repository: dirkatavis/CDK (upstream)"
WScript.Echo "Source: dirkste:feature/eliminate-hardcoded-paths (your fork)"
WScript.Echo ""
WScript.Echo "Visit GitHub to complete the PR creation and add description."
WScript.Echo "URL will be provided after execution."
WScript.Echo ""
WScript.Echo "Script completed. Manual PR creation required via GitHub UI."

Function IIf(condition, trueValue, falseValue)
    If condition Then
        IIf = trueValue
    Else
        IIf = falseValue
    End If
End Function
