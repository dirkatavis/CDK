# CDK Tools - Validation & Setup

This directory contains validation and setup scripts for the CDK automation system.

## Usage

### 1. Initial Setup (First Time Only)

If this is your first time running CDK scripts on this machine:

```cmd
cscript.exe setup_cdk_base.vbs
```

This sets the `CDK_BASE` environment variable to point to your CDK repository root. **You may need to restart your terminal after running this.**

### 2. Validation

Before running any CDK scripts, validate your environment:

```cmd
cscript.exe validate_dependencies.vbs
```

This checks:
- ✓ CDK_BASE environment variable is set
- ✓ Repository structure is intact (.cdkroot marker)
- ✓ Common libraries are present (PathHelper.vbs, ValidateSetup.vbs)
- ✓ Central configuration (config.ini) exists
- ✓ All configured file paths are accessible

**Fix any issues reported before proceeding.**

### 3. Troubleshooting

View your current CDK_BASE setting:
```cmd
cscript.exe show_cdk_base.vbs
```

Scan for hardcoded paths in scripts (should use config.ini instead):
```cmd
cscript.exe scan_hardcoded_paths.vbs
```

or with PowerShell:
```cmd
powershell -ExecutionPolicy Bypass -File scan_hardcoded_paths.ps1
```

Test the path resolution system:
```cmd
cscript.exe test_path_helper.vbs
```

## Validation Tests

To verify the validation system is working correctly:

### Run all validation tests
```cmd
cscript.exe run_validation_tests.vbs
```

This runs both positive and negative tests to verify:
- ✓ Validation passes when all dependencies are present
- ✓ Validation detects missing CDK_BASE
- ✓ Validation detects invalid CDK_BASE paths
- ✓ Validation detects missing .cdkroot marker
- ✓ Validation detects missing PathHelper.vbs
- ✓ Validation detects missing config.ini
- ✓ Validation handles corrupted config.ini gracefully

### Run positive tests only
```cmd
cscript.exe test_validation_positive.vbs
```
Verifies all required dependencies exist and are accessible.

### Run negative tests only
```cmd
cscript.exe test_validation_negative.vbs
```
Simulates missing dependencies and verifies validation catches them.

## What These Scripts Do

### Operational Tools (User-Facing)
| Script | Purpose |
|--------|----------|
| `validate_dependencies.vbs` | Complete pre-flight check before running automation |
| `setup_cdk_base.vbs` | Set CDK_BASE environment variable (one-time setup) |
| `show_cdk_base.vbs` | Display current CDK_BASE setting |
| `scan_hardcoded_paths.vbs` | Find hardcoded paths that should use config.ini |
| `scan_hardcoded_paths.ps1` | PowerShell version of path scanner |
| `Coordinate_Finder.vbs` | Generate screen coordinate ruler for debugging |

### Infrastructure Tests (Developer-Facing)
| Script | Purpose |
|--------|----------|
| `test_path_helper.vbs` | Test path resolution system |
| `test_validation_positive.vbs` | Verify validation passes with all dependencies present |
| `test_validation_negative.vbs` | Verify validation detects missing/broken dependencies |
| `run_validation_tests.vbs` | Run all validation tests (positive + negative) |

### PostFinalCharges Tests (Developer-Facing)
| Script | Purpose |
|--------|----------|
| `test_default_value_detection.vbs` | Verify prompt default value parsing (15 test cases) |
| `test_default_value_integration.vbs` | Integration tests with MockBzhao |
| `test_bug_prevention.vbs` | Regression tests for known issues |
| `test_integration.vbs` | Full integration test flow |
| `test_mock_bzhao.vbs` | MockBzhao framework tests |
| `test_prompt_detection.vbs` | Prompt detection timing tests |
| `test_operation_code_*.vbs` | Operation code parsing tests |
| `test_open_status.vbs` | Open status validation |
| `run_all_tests.vbs` | Run all PostFinalCharges tests |
| `run_default_value_tests.vbs` | Run default value tests only |
| `run_tests.bat` | Batch file runner for tests |

## For New Users

1. Extract/download the CDK repository to your machine
2. Open a command prompt or PowerShell
3. Navigate to the CDK folder: `cd C:\Temp_alt\CDK` (or wherever you extracted it)
4. Run: `cscript.exe tools\setup_cdk_base.vbs`
5. Close and reopen your terminal
6. Run: `cscript.exe tools\validate_dependencies.vbs`
7. If all checks pass, you can now run CDK scripts

## For Administrators

When deploying CDK to multiple users:

1. Ensure the repository structure is complete (all folders and files intact)
2. Have each user run `tools\setup_cdk_base.vbs` during initial setup
3. Have each user run `tools\validate_dependencies.vbs` to verify installation
4. Consult [../docs/SETUP_VALIDATION.md](../docs/SETUP_VALIDATION.md) for detailed guidance

## Documentation

For comprehensive validation documentation, see:
- [../docs/SETUP_VALIDATION.md](../docs/SETUP_VALIDATION.md) - Detailed validation guide
- [../docs/PATH_CONFIGURATION.md](../docs/PATH_CONFIGURATION.md) - Path management & config.ini

## Questions?

If validation fails:
1. Review the error message carefully
2. Check [../docs/SETUP_VALIDATION.md](../docs/SETUP_VALIDATION.md) Troubleshooting section
3. Verify your CDK folder structure matches the expected layout
