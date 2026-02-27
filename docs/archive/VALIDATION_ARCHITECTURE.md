# CDK Validation System Architecture

This document explains how CDK's dependency validation system works, how it's tested, and how it helps ensure scripts run successfully.

## Overview

The validation system is a three-layer approach:

1. **Validation Script** (`tools\validate_dependencies.vbs`) - Comprehensive pre-flight check
2. **Validation Library** (`common\ValidateSetup.vbs`) - Shareable validation routines
3. **Script Startup** (in each automation script) - Automatic validation before execution

## How It Works

### Layer 1: Standalone Validation Script

Users can run a complete environment check independently:

```cmd
cscript.exe tools\validate_dependencies.vbs
```

**Checks performed:**
- ✓ `CDK_BASE` environment variable is set
- ✓ `CDK_BASE` points to a valid folder
- ✓ `.cdkroot` marker file exists (repository structure validation)
- ✓ `PathHelper.vbs` exists (path resolution library)
- ✓ `config.ini` exists (central configuration)
- ✓ `config.ini` has valid INI format
- ✓ All paths referenced in `config.ini` are accessible

**Output:**
- Clear pass/fail status for each check
- Specific remediation instructions if a check fails
- Exit code 0 (success) or 1 (failure)

### Layer 2: Shared Validation Library

`common\ValidateSetup.vbs` provides reusable functions for any script:

```vbscript
' Strict validation (terminates on failure)
MustHaveValidDependencies

' Soft validation (returns True/False)
If Not ValidateScriptDependencies() Then
    WScript.Echo "Some dependencies may be missing"
End If

' Get repo root safely
repoRoot = GetRepoRootSafe()
```

**Available functions:**
- `MustHaveValidDependencies()` - Strict validation, exits on failure
- `ValidateScriptDependencies()` - Soft validation, returns boolean
- `GetRepoRootSafe()` - Returns repo root or empty string

### Layer 3: Automatic Startup Validation

Each critical script (e.g., `PostFinalCharges.vbs`) automatically validates dependencies:

```vbscript
Sub StartScript()
    ' FIRST THING: Validate all dependencies
    MustHaveValidDependencies
    
    ' Only reaches here if all dependencies pass
    Call LogInfo("PostFinalCharges script bootstrap starting", "Bootstrap")
    ' ... rest of script ...
End Sub
```

**Behavior:**
- If any dependency check fails → Script terminates immediately
- User sees clear error message with remediation steps
- No partial execution or silent failures

## Validation Levels

### Critical (Will Stop Script)
- `CDK_BASE` environment variable not set
- `CDK_BASE` points to non-existent folder
- `PathHelper.vbs` not found
- `config.ini` not found

### Warning (Script May Continue)
- `.cdkroot` marker not found
- Configured paths don't exist (but may be created on first run)

## Test Suite

The validation system includes comprehensive tests:

### Positive Tests
Verify validation passes with a complete setup:

```cmd
cscript.exe tools\test_validation_positive.vbs
```

Tests:
- ✓ CDK_BASE is valid
- ✓ .cdkroot marker exists
- ✓ PathHelper.vbs exists
- ✓ ValidateSetup.vbs exists
- ✓ config.ini exists
- ✓ config.ini has valid INI format
- ✓ Critical paths from config.ini exist
- ✓ Full validation passes

### Negative Tests
Verify validation detects missing/broken dependencies:

```cmd
cscript.exe tools\test_validation_negative.vbs
```

Tests:
- ✓ Detects missing CDK_BASE
- ✓ Detects invalid CDK_BASE path
- ✓ Detects missing .cdkroot marker
- ✓ Detects missing PathHelper.vbs
- ✓ Detects missing config.ini
- ✓ Handles corrupted config.ini

### Run All Tests
```cmd
cscript.exe tools\run_validation_tests.vbs
```

Executes both positive and negative test suites in sequence.

## Integration Points

### For New Scripts

To add validation to a new script:

**Option 1: Strict (Recommended)**
```vbscript
Option Explicit

' ... load PathHelper and ValidateSetup ...
Dim helperPath: helperPath = ...
ExecuteGlobal g_fso.OpenTextFile(helperPath).ReadAll

Dim validatePath: validatePath = ...
ExecuteGlobal g_fso.OpenTextFile(validatePath).ReadAll

Sub Main()
    MustHaveValidDependencies  ' Exit if any check fails
    ' ... rest of script ...
End Sub

Main
```

**Option 2: Soft (For Utilities)**
```vbscript
Sub Main()
    If Not ValidateScriptDependencies() Then
        WScript.Echo "Warning: Dependencies may be missing."
        ' Continue anyway, or:
        WScript.Quit 1
    End If
    ' ... rest of script ...
End Sub
```

### For Deployment/Packaging

When distributing CDK to other users:

1. Ensure these files are present:
   - `common\PathHelper.vbs`
   - `common\ValidateSetup.vbs`
   - `config.ini` (at repo root)
   - `.cdkroot` (marker file)

2. Have recipients run:
   ```cmd
   cscript.exe tools\setup_cdk_base.vbs
   ```

3. Have recipients verify:
   ```cmd
   cscript.exe tools\validate_dependencies.vbs
   ```

4. Run full test suite to ensure everything works:
   ```cmd
   cscript.exe tools\run_validation_tests.vbs
   ```

## Troubleshooting

### "All checks pass but script still won't run"

1. Check individual dependency:
   ```cmd
   cscript.exe tools\validate_dependencies.vbs
   ```

2. Check script startup behavior:
   ```cmd
   cscript.exe {script}.vbs 2>&1 | head -20
   ```

3. Review config.ini paths:
   ```cmd
   cscript.exe tools\test_path_helper.vbs
   ```

### "Validation checks failed"

1. Review specific errors from:
   ```cmd
   cscript.exe tools\validate_dependencies.vbs
   ```

2. Follow remediation steps shown for each failure

3. Run positive tests to re-verify:
   ```cmd
   cscript.exe tools\test_validation_positive.vbs
   ```

### "Tests are failing"

1. Run all tests with detailed output:
   ```cmd
   cscript.exe tools\run_validation_tests.vbs
   ```

2. Review which specific tests failed

3. Check environment setup:
   ```cmd
   cscript.exe tools\show_cdk_base.vbs
   ```

## Design Principles

The validation system follows these principles:

1. **Fail Fast**: Stop execution immediately if critical dependencies are missing
2. **Clear Messaging**: Tell users exactly what's wrong and how to fix it
3. **Tested**: Comprehensive test suites verify validation works correctly
4. **Shareable**: Validation library can be used by any script in the system
5. **Graceful Degradation**: Handle optional dependencies (like logs) without blocking
6. **Easy Distribution**: Minimal setup required - just set CDK_BASE environment variable

## Related Documentation

- [SETUP_VALIDATION.md](SETUP_VALIDATION.md) - User-facing validation guide
- [PATH_CONFIGURATION.md](PATH_CONFIGURATION.md) - config.ini structure and usage
- [tools/README.md](README.md) - Available tools and scripts
