# Phase 1: Prepare Close Pt1 - Pre-Manual RO Processing

## Purpose
Automates the pre-manual steps to prepare Repair Orders for closing. This is the "before" part of the sandwich workflow - runs before manual work, preparing ROs for technician review.

## Entry Script
- `1_Prepare_Close_Pt1.vbs` - Main automation script

## Input Files
- `Prepare_Close_Pt1_in.csv` - List of RO numbers to prepare
  - **Shared with Phase 2** - Both pt1 and pt2 read the same input file

## Output Files
- `Prepare_Close_Pt1.log` - Transaction log

## Usage
1. Populate `Prepare_Close_Pt1_in.csv` with RO numbers
2. Open BlueZone with active CDK session
3. Run script:
```cmd
cscript.exe 1_Prepare_Close_Pt1.vbs
```
4. Perform manual work in CDK
5. Run Phase 2 finalize script (uses same input file)

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

## Workflow Position
**Phase 1** - Sandwich workflow pt1 (BEFORE manual work):
```
1_Prepare_Close_Pt1 → [Manual Work] → 2_Finalize_Close_Pt2
        ↓                                      ↓
  (reads Prepare_Close_Pt1_in.csv)  (reads same file)
```

## Notes
- Runs BEFORE manual processing; Phase 2 runs AFTER
- Both pt1 and pt2 share the same input file (sandwich pattern)
- All files stored locally in this folder
