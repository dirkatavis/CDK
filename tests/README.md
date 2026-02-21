# Tests - Global Repository Tests

## Purpose
Repository-level tests that validate cross-cutting concerns, infrastructure, and migration contracts. These tests ensure the **entire codebase** maintains integrity, not just individual apps.

## Contrast with App Tests
- **App tests** (`apps/*/tests/`) - Validate single app behavior, mocked inputs, unit logic
- **Repo tests** (`tests/`) - Validate infrastructure, path resolution, migration contracts, cross-app concerns

## Test Categories

### 🏗️ Infrastructure Validation
Tests that verify foundational components work correctly.

**`test_path_helper.vbs`** - Unit tests for PathHelper path resolution logic
**`test_hardcoded_paths_comprehensive.vbs`** - Scans entire codebase for hardcoded paths that should use PathHelper

### ✅ Environment Validation
Tests that ensure the environment is correctly configured.

**`test_validation_positive.vbs`** - Validates all dependencies are present (should PASS)
**`test_validation_negative.vbs`** - Simulates missing dependencies and verifies validation catches them (should FAIL appropriately)
**`test_reset_state.vbs`** - Preflight cleanup for deterministic test runs

### 📋 Migration Contract Tests
Tests that validate the repository reorganization maintains backward compatibility.

**`test_reorg_contract_entrypoints.vbs`** - Validates legacy entry paths still exist
**`test_reorg_contract_config_paths.vbs`** - Validates config.ini paths resolve correctly
**`test_reorg_contract_wrappers.vbs`** - Validates wrapper targets exist in new locations

### 🎯 Test Runners
Master test suites that orchestrate validation.

**`run_validation_tests.vbs`** - Current-state validation (MUST stay green)
  - Runs: Preflight Reset → Positive Tests → Negative Tests → Reorg Contracts
  - Exit 0 = all pass, Exit 1 = failure
  - Use for pre-commit validation

**`run_migration_target_tests.vbs`** - Final-state progress tracker (red→green)
  - Reports migration progress: % complete, phase gates
  - Validates target structure matches `tooling/reorg_path_map.ini`
  - Intentionally red/yellow until migration reaches 100%

## Running Tests

### Quick Pre-Commit Check
```cmd
cscript.exe tests\run_validation_tests.vbs
```
Should show: `6/6 tests passed` (or current count)

### Migration Progress Check
```cmd
cscript.exe tests\run_migration_target_tests.vbs
```
Shows: `XX/YY checks passed (ZZ%)` and phase gate status

### Individual Test
```cmd
cscript.exe tests\test_path_helper.vbs
```

## Exit Codes
- **0** = All tests passed
- **1** = One or more tests failed

## Test Data Sources
- **`tooling/reorg_path_map.ini`** - Migration contract definitions (LegacyEntrypoints, WrapperTargets, ConfigContracts, etc.)
- **`config/config.ini`** - Path configuration for contract validation

## Design Principles
- **Fast Feedback:** Tests run in seconds, not minutes
- **Fail Fast:** Clear error messages pointing to root cause
- **No Mocking (Infra Tests):** Infrastructure tests use real paths/files
- **Deterministic:** `test_reset_state.vbs` ensures clean slate

## When to Run
- **Before commits:** `run_validation_tests.vbs` ensures no regressions
- **After file moves:** Migration contract tests catch broken references
- **During reorganization:** `run_migration_target_tests.vbs` tracks progress
- **CI/CD pipelines:** Automated validation on every push

## Test Output Handling
Test output files should **never** be written to the repository root. Use proper locations:

- **Test logs:** `runtime/logs/tests/` - For detailed test execution logs
- **Shell redirection:** When capturing output via `>`, redirect to `runtime/logs/tests/test_name_output.txt`
- **Root directory:** Avoided - files here create clutter and are gitignored

Example proper usage:
```cmd
REM ❌ Wrong - writes to root
cscript test_validation.vbs > test_output.txt

REM ✅ Correct - writes to runtime/logs/tests/
cscript test_validation.vbs > runtime\logs\tests\validation_output.txt
```

## Adding New Global Tests
1. Test must validate a cross-cutting concern (affects multiple apps or core infrastructure)
2. Add test script to `tests/` folder
3. If part of validation suite, add to `run_validation_tests.vbs`
4. If part of migration tracking, update `tooling/reorg_path_map.ini` and `run_migration_target_tests.vbs`
5. Update this README with test description

## Notes
- Keep repo tests **minimal** - most testing should be app-local
- Focus on infrastructure, contracts, and cross-cutting concerns
- App-specific behavior belongs in `apps/*/tests/`
- See `PACKAGING_GUIDE.md` - repo tests are NOT distributed to end users (developers only)
