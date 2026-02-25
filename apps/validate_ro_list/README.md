# Validate RO List - Data Validation Tool

## Purpose
Validates Repair Order lists against CDK system data to ensure accuracy before processing operations.

## Entry Script
- `ValidateRoList.vbs` - Main validation script

## Input Files
- `ValidateRoList_IN.csv` - List of ROs to validate

## Output/Logs
- Logs written to `runtime/logs/validate_ro_list/ValidateRoList.log`
- Screen mapping data: `ro_screen_map_pfc.txt`, `sample_map_pfc.txt`
- Validation results: `ValidateRoList_OUT.txt`
- Mock test artifacts: `ValidateRoList_mock_log.txt`, `ValidateRoList_mock_out.txt`

## Usage
```cmd
cscript.exe ValidateRoList.vbs
```

## Dependencies
- BlueZone terminal emulator with active CDK session
- `framework/PathHelper.vbs` - Path resolution
- `framework/ValidateSetup.vbs` - Setup validation
- `config/config.ini` - Configuration file

## Testing
Run app-specific tests:
```cmd
cd apps\validate_ro_list\tests
cscript.exe run_tests.vbs
```

## Notes
- Pre-processing validation to prevent errors in downstream automation
- Generates screen maps for diagnostic purposes
- Supports mock testing mode for development without live CDK connection
- Critical for data quality assurance
