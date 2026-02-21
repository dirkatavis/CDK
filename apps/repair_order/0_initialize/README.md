# Initialize RO - Repair Order Initialization

## Purpose
Automates the creation and initial setup of new Repair Orders (ROs) in the CDK DMS system via BlueZone terminal emulation.

## Entry Script
- `1_Initialize_RO.vbs` - Main automation script

## Input Files
- `Initialize_RO.csv` - List of ROs to initialize with required fields

## Output/Logs
- Logs written to `runtime/logs/initialize_ro/Initialize_RO.log`
- Debug mode: Create `Initialize_RO.debug` file in runtime log folder for slow-mode execution

## Usage
```cmd
cscript.exe 1_Initialize_RO.vbs
```

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

## Notes
- Part of the "Sandwich Automation" Pt1 workflow - sets up ROs before manual processing
- Requires active BlueZone session at CDK main menu
- See `PACKAGING_GUIDE.md` in repo root for distribution instructions
