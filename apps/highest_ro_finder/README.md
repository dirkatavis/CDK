# Highest RO Finder - Diagnostic Utility

## Purpose
Scans the CDK system to identify the highest currently active Repair Order number for diagnostics and planning.

## Entry Script
- `HighestRoFinder.vbs` - Main search script

## Input Files
- None (scans CDK system directly)

## Output/Logs
- Logs written to `runtime/logs/highest_ro_finder/HighestRoFinder.log`
- Results displayed in terminal output

## Usage
```cmd
cscript.exe HighestRoFinder.vbs
```

## Dependencies
- BlueZone terminal emulator with active CDK session
- `framework/PathHelper.vbs` - Path resolution
- `framework/ValidateSetup.vbs` - Setup validation
- `config/config.ini` - Configuration file

## Testing
Run app-specific tests:
```cmd
cd apps\highest_ro_finder\tests
cscript.exe run_tests.vbs
```

## Notes
- Diagnostic utility for finding RO number ranges
- Useful for planning bulk operations
- Quick execution - scans incrementally upward until no more ROs found
