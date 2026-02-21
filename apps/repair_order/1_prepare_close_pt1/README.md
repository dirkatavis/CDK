# Prepare Close Pt1 - Pre-Manual RO Close Preparation

## Purpose
Automates the pre-manual processing steps required before closing Repair Orders (ROs), including seeding data and preparing the terminal state.

## Entry Script
- `2_Prepare_Close_Pt1.vbs` - Main automation script

## Input Files
- `Prepare_Close_Pt1.csv` - List of ROs to prepare for closing with required context

## Output/Logs
- Logs written to `runtime/logs/prepare_close_pt1/Prepare_Close_Pt1.log`
- Debug mode: Create `Prepare_Close_Pt1.debug` file in runtime log folder for slow-mode execution

## Usage
```cmd
cscript.exe 2_Prepare_Close_Pt1.vbs
```

## Dependencies
- BlueZone terminal emulator with active CDK session
- `framework/PathHelper.vbs` - Path resolution
- `framework/ValidateSetup.vbs` - Setup validation
- `config/config.ini` - Configuration file

## Testing
Run app-specific tests:
```cmd
cd apps\repair_order\prepare_close_pt1\tests
cscript.exe run_tests.vbs
```

## Notes
- Part of the "Sandwich Automation" Pt1 workflow - prepares ROs before manual middle step
- Requires active BlueZone session at CDK main menu
- This script runs BEFORE manual processing; Pt2 runs AFTER
