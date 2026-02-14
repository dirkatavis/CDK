# BlueZone Context Compatibility

## Overview

The CDK validation system is designed to work in both standalone and BlueZone embedded script contexts.

### Standalone Context
- Running scripts with: `cscript.exe script.vbs`
- Tools like: `tools\validate_dependencies.vbs`
- Environment: Full WScript object available

### BlueZone Context
- Running scripts from within BlueZone terminal emulator
- Scripts like: `PostFinalCharges.vbs` (executed in BlueZone)
- Environment: No WScript object available

## Validation System Design

### Dual-Context Validation Library: `common\ValidateSetup.vbs`

The validation library is designed to work in **both contexts** without requiring conditional imports:

```vbscript
' This function works in both contexts:
' - Standalone: Uses WScript.Echo and WScript.Quit
' - BlueZone: Uses LogInfo and sets g_ShouldAbort flag
Sub MustHaveValidDependencies()
    If Not ValidateScriptDependencies() Then
        Call SafeOutput("FATAL: Required dependencies not available.")
        
        ' Try WScript.Quit (works if standalone)
        WScript.Quit 1
        
        ' If we reach here, set BlueZone abort flag
        g_ShouldAbort = True
        g_AbortReason = "Dependency validation failed"
    End If
End Sub
```

### SafeOutput Function

Messages are output safely in both contexts:

```vbscript
' In standalone context: outputs to console
' In BlueZone context: logs to script log using LogInfo
Sub SafeOutput(msg)
    On Error Resume Next
    WScript.Echo msg           ' Standalone - succeeds
    On Error GoTo 0
    
    On Error Resume Next
    If g_CurrentCriticality >= 0 Then
        LogInfo msg, "Validation"  ' BlueZone - if LogInfo available
    End If
    On Error GoTo 0
End Sub
```

## Usage Contexts

### Context 1: Pre-Flight Validation (Recommended Before BlueZone Entry)

**Command Line (Standalone):**
```cmd
cscript.exe tools\validate_dependencies.vbs
```

- Runs with full WScript support
- Clear console output
- Exit codes indicate pass/fail
- Users fix any issues before entering BlueZone

### Context 2: Runtime Validation (Inside BlueZone)

**BlueZone Script:**
```vbscript
' At top of your script:
ExecuteGlobal g_fso.OpenTextFile(validateSetupPath).ReadAll

Sub StartScript()
    ' First thing: validate dependencies
    MustHaveValidDependencies
    
    ' If we reach here, all dependencies are OK
    ' (In BlueZone: if dependencies failed, g_ShouldAbort is True)
    If g_ShouldAbort Then
        LogEvent "Startup aborted: " & g_AbortReason, CRIT_CRITICAL, VERB_LOW
        Exit Sub
    End If
    
    ' Continue with script...
End Sub
```

## Handling Validation Failures

### In Standalone Context
```vbscript
MustHaveValidDependencies  ' Calls WScript.Quit 1 on failure
' Code here is never reached if validation fails
```

### In BlueZone Context
```vbscript
MustHaveValidDependencies  ' Sets g_ShouldAbort on failure
' Code continues to execute

' Check abort flag:
If g_ShouldAbort Then
    LogEvent "Validation failed: " & g_AbortReason, CRIT_CRITICAL, VERB_LOW
    ' Handle gracefully
End If
```

## Examples

### Correct BlueZone Usage

```vbscript
Option Explicit

Dim g_fso, g_shell, g_ShouldAbort, g_AbortReason
Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_shell = CreateObject("WScript.Shell")
g_ShouldAbort = False
g_AbortReason = ""

' Load validation library
ExecuteGlobal g_fso.OpenTextFile(...\ValidateSetup.vbs").ReadAll

Sub Main()
    MustHaveValidDependencies      ' Works in BlueZone context
    
    If g_ShouldAbort Then
        ' Handle validation failure
        Exit Sub
    End If
    
    ' Continue with script logic
End Sub

Main
```

### Correct Standalone Usage

```vbscript
Option Explicit

' This script uses WScript, no need for SafeOutput
' (ValidateSetup.vbs is included, but it detects standalone context)

Set g_fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal g_fso.OpenTextFile(validateSetupPath).ReadAll

MustHaveValidDependencies      ' Exits immediately if validation fails

' Code here only runs if validation passes
WScript.Echo "All dependencies validated successfully!"
```

## Error Handling Patterns

### Pattern 1: Silent Failure Detection

```vbscript
' Don't fail immediately, just check
If Not ValidateScriptDependencies() Then
    ' Handle missing dependencies
    LogEvent "Environment validation failed", CRIT_MAJOR, VERB_LOW
    g_ShouldAbort = True
End If
```

### Pattern 2: Strict Validation

```vbscript
' Fail immediately if dependencies missing
MustHaveValidDependencies

' If we reach here, all dependencies are guaranteed OK
```

### Pattern 3: Nested Script Context

```vbscript
' For scripts that might be called from multiple contexts:
If Not ValidateScriptDependencies() Then
    On Error Resume Next
    WScript.Quit 1                    ' Standalone: exit
    On Error GoTo 0
    
    g_ShouldAbort = True              ' BlueZone: set flag
    g_AbortReason = "Missing dependencies"
End If
```

## Migration Guide

### Updating Existing Scripts

To add validation support to an existing script:

1. **Include the validation library:**
```vbscript
Dim validateSetupPath
validateSetupPath = g_fso.BuildPath(GetRepoRoot(), "common\ValidateSetup.vbs")
ExecuteGlobal g_fso.OpenTextFile(validateSetupPath).ReadAll
```

2. **Call validation at startup:**
```vbscript
Sub Main()
    MustHaveValidDependencies  ' Works in both contexts
    ' ... rest of script
End Sub
```

3. **When running from BlueZone, check the abort flag:**
```vbscript
Sub Main()
    MustHaveValidDependencies
    
    If g_ShouldAbort Then
        LogEvent "Validation failed: " & g_AbortReason, CRIT_CRITICAL, VERB_LOW
        Exit Sub
    End If
    
    ' ... rest of script
End Sub
```

## Testing Validation in BlueZone

Unfortunately, you cannot fully test validation inside BlueZone without a live BlueZone terminal. However:

1. **Test standalone context:**
```cmd
cscript.exe tools\test_validation_positive.vbs
cscript.exe tools\test_validation_negative.vbs
```

2. **Validate file syntax:**
```cmd
cscript.exe //Nologo //D common\ValidateSetup.vbs
```

3. **In BlueZone, review logs** to see if SafeOutput messages appear when validation runs

## Troubleshooting

### "Variable is undefined: 'WScript'"

This means the script is running in BlueZone context and tried to use WScript directly. 

**Solution:** Use ValidateSetup.vbs functions instead (SafeOutput, MustHaveValidDependencies) which handle both contexts.

### Validation runs but output is not visible

In BlueZone context, validation messages go to the script log (via LogInfo), not console.

**Check:** Review the log file configured in config.ini for the script's `Log` entry.

### Abort flag not being checked

Make sure you include this after calling validation:
```vbscript
If g_ShouldAbort Then
    ' Handle abort
End If
```

## Design Notes

- **No Global State**: Validation doesn't modify global state except setting g_ShouldAbort
- **Error Resilience**: All operations wrapped in `On Error Resume Next` to handle missing LogInfo gracefully
- **Simple API**: Just one function to call: `MustHaveValidDependencies`
- **Context Detection**: Automatic - no need to specify which context you're in
