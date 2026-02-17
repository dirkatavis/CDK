# CDK Automation Deployment Guide

## Phase 1: Package Creation (Dev Machine)

**STEP 1**: Compress the `CDK` folder into a distribution ZIP.
- **EXPECTED RESULT**: A ZIP package is created containing the project files.

---

## Phase 2: Installation & Verification (Target Machine)

**STEP 1**: Extract the ZIP package to the target location (e.g., `C:\CDK`).
- **EXPECTED RESULT**: The `CDK` folder structure is visible in the file explorer.

**STEP 2**: Open a terminal and run `cscript.exe tools\setup_cdk_base.vbs`.
- **EXPECTED RESULT**: A popup appears confirming `CDK_BASE` is set. Click OK and **restart** your terminal or BlueZone.

**STEP 3**: Run `cscript.exe tools\validate_dependencies.vbs`.
- **EXPECTED RESULT**: The terminal displays **PASS** for all dependency checks.

**STEP 4**: Run `cscript.exe tools\test_path_helper.vbs`.
- **EXPECTED RESULT**: Terminal output confirms that `config.ini` paths are resolved correctly.
