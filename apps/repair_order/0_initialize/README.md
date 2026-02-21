# Phase 0: Initialize RO - Repair Order Creation

## Purpose
Automates the creation of new Repair Orders (ROs) from a list of vehicles. Reads MVA (Motor Vehicle Account) and mileage data, creates ROs in CDK DMS, and outputs the generated RO numbers.

## Entry Script
- `0_Initialize_RO.vbs` - Main automation script

## Input Files
- `Initialize_RO_in.csv` - List of vehicles with MVA and mileage
  - Format: `MVA,Mileage`

## Output Files
- `Initialize_RO_out.csv` - Generated RO numbers (single column)
- `Initialize_RO.log` - Transaction log
- Debug mode: Create `Initialize_RO.debug` file in this folder for slow-mode execution

## Usage
1. Populate `Initialize_RO_in.csv` with vehicle data
2. Open BlueZone with active CDK session
3. Run script:
```cmd
cscript.exe 0_Initialize_RO.vbs
```
4. Review `Initialize_RO_out.csv` for generated RO numbers

## Dependencies
- BlueZone terminal emulator with active CDK session
- `framework/PathHelper.vbs` - Path resolution
- `framework/ValidateSetup.vbs` - Setup validation
- `config/config.ini` - Configuration file

## Testing
Run app-specific tests:
```cmd
cd apps\repair_order\initialize\tests
cscript.exe run_tests.vbs
```

## Workflow Position
**Phase 0** - Creates new ROs before the prepare/finalize workflow:
```
0_Initialize_RO → [Output RO list] → (Copy to Phase 1 input) → 1_Prepare_Close_Pt1
```

## Notes
- Outputs RO numbers only (single column CSV)
- All files stored locally in this folder (self-contained app)
- Requires active BlueZone session at CDK main menu
