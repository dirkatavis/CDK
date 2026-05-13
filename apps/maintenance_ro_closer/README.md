# Maintenance RO Closer

## Purpose
Automated closure of maintenance-type Repair Orders based on pattern matching criteria and exception lists.

## Entry Script
- `Maintenance_RO_Closer.vbs` - Main automation script

## Input Files
- `RO_List.csv` - List of ROs to evaluate for closure
- `PM_Match_Criteria.txt` - Pattern matching rules for maintenance RO identification

## Output/Logs
- Input/output/log locations are defined in `config/config.ini` (`[Maintenance_RO_Closer]` section).

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

## Recent Hardening Goals
- Prevent review-flow breaks from occasional/unexpected prompts by routing prompt handling through a single resolver path.
- Ensure fallback behavior is explicit for missing defaults (for example, `ACTUAL HOURS`/`SOLD HOURS` -> `0`, comeback prompt -> `Y`).
- Keep warranty handling conservative: process configured warranty labor types and skip unsupported `W*` types.
- Add deterministic tests for prompt actions so new prompt scenarios can be validated without long production runs.
- Capture unhandled prompt text in logs to accelerate adding new prompt rules and regression test cases.
