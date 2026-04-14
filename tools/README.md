# Tools - Setup, Diagnostics & Development Utilities

## Purpose
Developer and operator tooling for setup, validation, diagnostics, and migration tracking. These scripts are for **one-time setup**, **troubleshooting**, and **development workflows** - not production automation.

## Quick Start

### First-Time Setup
```cmd
# 1. Set CDK_BASE environment variable
cscript.exe tools\setup_cdk_base.vbs

# 2. Validate environment
cscript.exe tools\validate_dependencies.vbs

# 3. Test path resolution
cscript.exe tools\test_path_helper.vbs
```

### Pre-Commit Validation
```cmd
# Run all validation tests to ensure no regressions
cscript.exe tools\run_validation_tests.vbs
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
**`scan_hardcoded_employee.vbs`** - Flags hardcoded employee numbers
**`scan_unconfigured_keys.vbs`** - Reports config keys referenced in scripts but missing from `config.ini`
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

See `tests/migration/reorg_path_map.ini` for migration contracts.

### 🛠️ Development Tools
Ad-hoc utilities for development and troubleshooting.

**`close_single_ro.vbs`** - Manual single-RO closure
**`create_upstream_pr.vbs`** - Automates PR creation workflow
**`mark_ini_assume_unchanged.ps1`** - Re-applies `assume-unchanged` to all tracked `.ini` files
**`upsert_pr_with_body.ps1`** - Creates or updates a PR and always applies body from file

Run it from repo root:
```powershell
powershell -ExecutionPolicy Bypass -File .\tools\mark_ini_assume_unchanged.ps1
```

Create or update a PR with guaranteed body text:
```powershell
# If PR exists for current branch, updates body (and title if provided)
powershell -ExecutionPolicy Bypass -File .\tools\upsert_pr_with_body.ps1 `
	-BodyFile .\Temp\pr_body.md `
	-Title "fix(post-final-charges): log RO at sequence start and add log overwrite toggle"

# Explicit target repo/branches
powershell -ExecutionPolicy Bypass -File .\tools\upsert_pr_with_body.ps1 `
	-Repo dirkatavis/CDK `
	-Head bugfix/pfc-log-ro-at-start `
	-Base main `
	-BodyFile .\Temp\pr_body.md `
	-Title "fix(post-final-charges): log RO at sequence start and add log overwrite toggle"
```

### 📊 Data Collection
BlueZone scrapers that feed the analysis pipeline.

**`labor_parts_scraper.vbs`** - Scrapes L-line and P-line data from PFC sequences into `runtime\data\Master_Labor_Log.csv`. Supports resume-on-abort via sequence tracking. Configure range and output path in `[LaborPartsScraper]` in `config\config.ini`.

**`mva_scrapper/`** - Looks up MVA numbers from VINs via the CDK screen. Configure via `[GetMvaFromVin]` in `config\config.ini`.

## Design Principles
- **One-Time or Rare Use:** Tooling is for setup/diagnostics, not daily operations
- **Developer-Focused:** Assumes technical expertise
- **Fast Feedback:** Quick validation loops
- **Clear Failures:** Fail fast with actionable errors

## Documentation
- [../docs/SETUP_VALIDATION.md](../docs/SETUP_VALIDATION.md) - Validation architecture
- [../docs/PATH_CONFIGURATION.md](../docs/PATH_CONFIGURATION.md) - Path management
- [../PACKAGING_GUIDE.md](../PACKAGING_GUIDE.md) - Note: tooling NOT distributed to end users
