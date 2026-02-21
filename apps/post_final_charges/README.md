# Post Final Charges (PFC)

## Purpose
Advanced automation for posting final charges to Repair Orders using a state machine pattern to handle complex CDK screen flows and circular navigation.

## Entry Script
- `PostFinalCharges.vbs` - Main automation script with state machine logic

## Input Files
- `CashoutRoList.csv` - List of ROs to process with charge details

## Output/Logs
- Logs written to `runtime/logs/post_final_charges/PostFinalCharges.log`
- Debug mode: Create `PostFinalCharges.debug` file in runtime log folder for slow-mode execution
- Criticality levels: CRIT_COMMON (0) to CRIT_CRITICAL (3)
- Verbosity levels: VERB_LOW (0) to VERB_MAX (3)

## Supporting Files
- `CommonLib.vbs` - Shared helper functions and prompt detection logic

## Usage
```cmd
cscript.exe PostFinalCharges.vbs
```

## Dependencies
- BlueZone terminal emulator with active CDK session
- `framework/PathHelper.vbs` - Path resolution
- `framework/ValidateSetup.vbs` - Setup validation
- `config/config.ini` - Configuration file

## Testing
Run comprehensive app test suite:
```cmd
cd apps\post_final_charges\tests
cscript.exe run_all_tests.vbs
```

See `tests/README.md` for detailed test documentation covering:
- Default value detection (15+ test cases)
- Prompt pattern variations
- Integration tests with MockBzhao
- Bug prevention and regression tests

## Architecture
Uses a **state machine pattern** with `AddPromptToDictEx` to:
- Handle complex/circular screen flows
- Detect default values in prompts (e.g., "TECHNICIAN (72925)?")
- Support intervening text patterns (e.g., "OPERATION CODE FOR LINE A, L1 (I)?")
- Navigate CDK's non-linear prompt sequences

Key functions:
- `HasDefaultValueInPrompt(text)` - Detects default values in parentheses
- `AddPromptToDictEx(dict, trigger, response, key, isSuccess, acceptDefault)` - Advanced prompt handler
- `DiscoverLineLetters()` - Dynamic line letter discovery to prevent infinite loops

## Notes
- Most complex script in the repository - use with caution
- Extensively tested with 18+ test files covering edge cases
- Critical for production charge posting workflows
- See `tests/README.md` for test-driven prevention strategies
