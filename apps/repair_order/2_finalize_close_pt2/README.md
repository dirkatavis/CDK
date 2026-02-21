# Finalize Close Pt2 - Post-Manual RO Closure

## Purpose
Automates the post-manual processing steps to finalize and close Repair Orders (ROs), including printing and terminal cleanup.

## Entry Script
- `3_Finalize_Close_Pt2.vbs` - Main automation script

## Input Files
- `Finalize_Close.csv` - List of ROs to finalize with closing instructions

## Output/Logs
- Logs written to `runtime/logs/finalize_close_pt2/Finalize_Close_Pt2.log`
- Debug mode: Create `Finalize_Close_Pt2.debug` file in runtime log folder for slow-mode execution

## Usage
```cmd
cscript.exe 3_Finalize_Close_Pt2.vbs
```

## Dependencies
- BlueZone terminal emulator with active CDK session
- `framework/PathHelper.vbs` - Path resolution
- `framework/ValidateSetup.vbs` - Setup validation
- `config/config.ini` - Configuration file

## Testing
Run app-specific tests:
```cmd
cd apps\repair_order\finalize_close_pt2\tests
cscript.exe run_tests.vbs
```

## Notes
- Part of the "Sandwich Automation" Pt2 workflow - finalizes ROs after manual middle step
- Requires active BlueZone session at CDK main menu
- This script runs AFTER manual processing; Pt1 runs BEFORE
- Handles final RO closure and cleanup
