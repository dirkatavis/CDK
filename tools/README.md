# CDK Utility Tools

This directory contains standalone utility scripts for setup, validation, and testing.

## Key Scripts

### Setup & Validation
*   **`setup_cdk_base.vbs`**: Sets the `CDK_BASE` environment variable to the current repository root. Run this once after cloning/moving the repo.
*   **`validate_dependencies.vbs`**: Checks that all dependencies (files, paths, config) are valid. Run this to troubleshoot path issues or before running automation.
*   **`show_cdk_base.vbs`**: Displays the currently configured `CDK_BASE` path.

### Testing
*   **`test_path_helper.vbs`**: Validates that `PathHelper.vbs` can correctly resolve paths from `config.ini`.
*   **`run_validation_tests.vbs`**: Runs the full suite of validation logic tests.

## Usage

Most tools are run via `cscript`:

```cmd
cscript tools\validate_dependencies.vbs
```
