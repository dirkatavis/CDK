# Tooling - Setup, Diagnostics & Development Utilities

## Purpose
Developer and operator tooling for setup, validation, diagnostics, and migration tracking. These scripts are for **one-time setup**, **troubleshooting**, and **development workflows** - not production automation.

## Quick Start

### First-Time Setup
```cmd
# 1. Set CDK_BASE environment variable
cscript.exe tooling\setup_cdk_base.vbs

# 2. Validate environment
cscript.exe tooling\validate_dependencies.vbs

# 3. Test path resolution
cscript.exe tooling\test_path_helper.vbs
```

### Pre-Commit Validation
```cmd
# Run all validation tests to ensure no regressions
cscript.exe tooling\run_validation_tests.vbs
```

## Categories

### 🔧 Setup & Configuration
Scripts for initial environment setup and configuration.

**`setup_cdk_base.vbs`** - Sets `CDK_BASE` user environment variable (one-time setup)
**`show_cdk_base.vbs`** - Displays current `CDK_BASE` value
**`validate_dependencies.vbs`** - Comprehensive environment validation

### 🔍 Diagnostics & Scanning
Tools for analyzing the codebase and identifying issues.

**`scan_hardcoded_paths.vbs` / `scan_hardcoded_paths.ps1`** - Scans for hardcoded paths
**`coordinate_finder.vbs`** - Interactive screen coordinate discovery
**`ro_screen_map.vbs`** - Maps RO screen layouts
**`safe_mapper.vbs`** - Safe screen mapping with error recovery

### 🧪 Test Infrastructure
Global test runners and validation contracts.

**`run_validation_tests.vbs`** - Master test runner (current-state, must stay green)
**`run_migration_target_tests.vbs`** - Final-state progress tracker
**`test_reset_state.vbs`** - Preflight cleanup for deterministic tests
**`test_reorg_contract_*.vbs`** - Migration contract validation
**`test_validation_*.vbs`** - Positive/negative validation tests
**`test_path_helper.vbs`** - PathHelper unit tests

### 📋 Migration Tracking

**`reorg_path_map.ini`** - Single source of truth for migration contracts

### 🛠️ Development Tools
Ad-hoc utilities for development and troubleshooting.

**`close_single_ro.vbs`** - Manual single-RO closure
**`create_upstream_pr.vbs`** - Automates PR creation workflow

## Design Principles
- **One-Time or Rare Use:** Tooling is for setup/diagnostics, not daily operations
- **Developer-Focused:** Assumes technical expertise
- **Fast Feedback:** Quick validation loops
- **Clear Failures:** Fail fast with actionable errors

## Documentation
- [../docs/SETUP_VALIDATION.md](../docs/SETUP_VALIDATION.md) - Validation architecture
- [../docs/PATH_CONFIGURATION.md](../docs/PATH_CONFIGURATION.md) - Path management
- [../PACKAGING_GUIDE.md](../PACKAGING_GUIDE.md) - Note: tooling NOT distributed to end users
