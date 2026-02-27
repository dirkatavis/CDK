# Repository Test Suite

## Overview
This directory contains the global validation framework. It is organized into a **command-and-control** structure designed for immediate clarity, even for new developers.

## 🎯 Master Runners (Commanders)
These are your primary entry points. Use these to check the overall pulse of the repository.

- [run_all.vbs](run_all.vbs) - **The Grand Suite**. Orchestrates every test in the repo (Infra, Environment, Migration, and Apps).
- [run_validation_tests.vbs](run_validation_tests.vbs) - **CDK Grand Validation Suite**. The authoritative CI/CD runner for overall repository health.
- [run_stress_tests.vbs](run_stress_tests.vbs) - **Stress Suite**. Validates resilience against high terminal latency and partial screen loads.
- [run_migration.vbs](run_migration.vbs) - Tracks progress of the repository reorganization toward the target architecture.

## 📂 Categorized Tests (Workers)
Logic is grouped by purpose to simplify troubleshooting. If a master runner reports a failure in a specific section, you can find the relevant worker here.

### 🏗️ Infrastructure
Validates the core VBScript/PowerShell plumbing that powers the rest of the tools.
- [infrastructure/test_path_helper.vbs](infrastructure/test_path_helper.vbs) - Unit tests for PathHelper's relative-to-absolute resolution logic.
- [infrastructure/test_config_exhaustion.vbs](infrastructure/test_config_exhaustion.vbs) - Verifies that every single path defined in `config.ini` resolves to a real file.
- [infrastructure/test_hardcoded_paths.vbs](infrastructure/test_hardcoded_paths.vbs) - Scans the codebase for hardcoded absolute paths that should be using `PathHelper`.
- [infrastructure/test_syntax_validation.vbs](infrastructure/test_syntax_validation.vbs) - Scans for environment-breaking syntax (e.g., `DoEvents`, `MsgBox`, incorrect `Option Explicit` placement).

### 🌍 Environment
Checks the setup of the developer's machine and external dependencies.
- [environment/test_positive.vbs](environment/test_positive.vbs) - Verifies all required environment variables and folders exist.
- [environment/test_negative.vbs](environment/test_negative.vbs) - Simulates a "broken" environment to ensure the validation logic correctly catches it.
- [environment/test_reset.vbs](environment/test_reset.vbs) - Restores the workspace to a clean state (cleanup of log artifacts, etc.).

### 📋 Migration
Ensures that moving files doesn't break existing scripts or configuration.
- [migration/test_entrypoints.vbs](migration/test_entrypoints.vbs) - Validates that legacy entry points (stubs) still exist and are functional.
- [migration/test_config_paths.vbs](migration/test_config_paths.vbs) - Validates that `config.ini` paths align with the master reorg plan.
- [migration/test_wrappers.vbs](migration/test_wrappers.vbs) - Ensures shim/wrapper scripts point to their new targets correctly.

---

## 🚀 Usage Guide

### Full Repo Validation
Run this before submitting any PR. It must stay green at all times.
```cmd
cscript.exe //nologo tests\run_all.vbs
```

### Check Migration Progress
Use this to see how close the repository is to the final target structure.
```cmd
cscript.exe //nologo tests\run_migration.vbs
```

## Exit Codes
- **0** = SUCCESS: All tests passed.
- **1** = FAILURE: One or more tests failed or encountered an error.

## Design Principles
- **Fast Feedback:** The entire suite execution is measured in seconds.
- **Command & Control:** Root level contains "Commanders" (scripts you run); subfolders contain "Workers" (logic details).
- **Silent unless Failed:** Internal workers are typically verbose, but the Grand Suite provides a high-level table view.

