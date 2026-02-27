# CDK Automation Setup & Verification Guide

This guide covers the initial installation and ongoing validation of the CDK automation environment.

## One-Time Installation

Follow these steps to set up the system on a new machine.

### 1. Extract the Files
Extract the repository folder to a permanent location (e.g., `C:\CDK`).

### 2. Configure Environment (`CDK_BASE`)
The system needs to know where the repository is located to find its configuration and helper files.
1. Navigate to the `tools/` folder.
2. Double-click `setup_cdk_base.vbs` (or run via terminal: `cscript setup_cdk_base.vbs`).
3. This sets the `CDK_BASE` user environment variable.
4. **Restart BlueZone** (or any open command prompts) to apply the change.

### 3. Verify the Setup
Run the pre-flight check to ensure all dependencies are met:
1. Double-click `tools/validate_dependencies.vbs` (or run `cscript tools/validate_dependencies.vbs`).
2. **Success**: You should see "All checks passed!"
3. **Failure**: Follow the on-screen remediation steps.

---

## Technical Validation (Under the Hood)

The system uses a three-layer validation approach to prevent script failures:

| Layer | Component | Purpose |
|-------|-----------|---------|
| **Manual** | `tools/validate_dependencies.vbs` | User-initiated pre-flight check. |
| **Library** | `framework/ValidateSetup.vbs` | Standardizes checks across all scripts. |
| **Startup** | Script Header | Every script calls `MustHaveValidDependencies` before execution. |

### Required Components for Portability
- **`.cdkroot`**: Found in the root folder. **DO NOT DELETE**.
- **`config/config.ini`**: Centralized path storage.
- **`framework/PathHelper.vbs`**: The engine resolving relative to absolute paths.

## Troubleshooting

- **Check Current Path**: Run `tools/show_cdk_base.vbs` to see where the system thinks it is installed.
- **Path Verification**: Run `tests/test_path_helper.vbs` to verify specific INI path resolution.
- **Missing Marker**: If you see "Repository root not found," ensure the file `.cdkroot` exists in your main folder.
