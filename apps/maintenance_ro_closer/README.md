# Maintenance RO Closer

## Purpose
Automated closure of maintenance-type Repair Orders based on pattern matching criteria and exception lists.

## Entry Script
- `Maintenance_RO_Closer.vbs` - Main automation script

## Input Files
- `RO_List.csv` - List of ROs to evaluate for closure
- `PM_Match_Criteria.txt` - Pattern matching rules for maintenance RO identification

## Output/Logs
- Logs written to `runtime/logs/maintenance_ro_closer/Maintenance_RO_Closer.log`
- Output data written to `runtime/data/out/RO_Status_Report.csv`
- Exception list updated in `exception_list.csv`

## Usage
```cmd
cscript.exe Maintenance_RO_Closer.vbs
```

## Dependencies
- BlueZone terminal emulator with active CDK session
- `framework/PathHelper.vbs` - Path resolution
- `framework/ValidateSetup.vbs` - Setup validation
- `config/config.ini` - Configuration file

## Testing
Run app-specific tests:
```cmd
cd apps\maintenance_ro_closer\tests
cscript.exe run_tests.vbs
```

## Notes
- Uses pattern matching from `PM_Match_Criteria.txt` to identify maintenance ROs
- Maintains exception list for ROs that should not be auto-closed
- Generates status reports for auditing and verification
