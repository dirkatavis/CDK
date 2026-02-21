# Phase 2: Finalize Close Pt2 - Post-Manual RO Closure

## Purpose
Automates the post-manual steps to finalize and close Repair Orders. This is the "after" part of the sandwich workflow - runs after manual work to complete RO closure, printing, and cleanup.

## Entry Script
- `2_Finalize_Close_Pt2.vbs` - Main automation script

## Input Files
- **Shared Input**: `../1_prepare_close_pt1/Prepare_Close_Pt1_in.csv`
  - Reads the same RO list as Phase 1 (sandwich pattern)
  - Process same ROs that were prepared in pt1

## Output Files
- `Finalize_Close.log` - Transaction log (stored in this folder)

## Usage
1. Ensure Phase 1 has already run on `Prepare_Close_Pt1_in.csv`
2. Complete manual work in CDK
3. Open BlueZone with active CDK session
4. Run script:
```cmd
cscript.exe 2_Finalize_Close_Pt2.vbs
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

## Workflow Position
**Phase 2** - Sandwich workflow pt2 (AFTER manual work):
```
1_Prepare_Close_Pt1 → [Manual Work] → 2_Finalize_Close_Pt2
        ↓                                      ↓
  (reads Prepare_Close_Pt1_in.csv)  (reads same file)
```

## Notes
- Runs AFTER manual processing; Phase 1 runs BEFORE
- Reads input from Phase 1 folder (shared input file)
- Handles final RO closure, printing to printer 2, and cleanup
- Requires active BlueZone session at CDK main menu
