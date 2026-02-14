# PostFinalCharges - RO Closeout Automation

## Overview
Automates the final charge posting and RO closeout process in CDK DMS through BlueZone terminal emulator. Processes multiple ROs from a CSV file, handles conditional prompts, and manages multi-line story processing.

## Production Status
✅ **Production Ready** - Successfully tested with live BlueZone terminal (2 ROs processed end-to-end)

## Key Features

### Core Capabilities
- **Batch Processing**: Reads RO numbers from CSV and processes sequentially
- **State Machine Logic**: Handles 30+ conditional prompts with intelligent default detection
- **Multi-Line Processing**: Automatically discovers and processes non-consecutive line letters (A, C, D, etc.)
- **Story Management**: FNL (finish line) then R (resume) workflow prevents "line not finished" errors
- **Smart Timing**: Adaptive wait logic with 2-second max timeout for prompt responses

### Recent Enhancements (Feb 2026)

#### Self-Contained Helper Functions
Added 7 terminal interaction functions (~400 lines) for BlueZone automation:

| Function | Purpose | Key Feature |
|----------|---------|-------------|
| `FastKey` | Send terminal keys | Medium verbosity logging shows actual key sent |
| `FastText` | Send text strings | Medium verbosity logging shows actual text sent |
| `WaitForPrompt` | Wait for prompts, send responses | 10-parameter support, comprehensive error handling |
| `IsTextPresent` | Screen text search | Scans all 24 lines, case-insensitive |
| `WaitMs` | Pause execution | Non-blocking sleep with DoEvents |
| `GetScreenSnapshot` | Capture screen for debugging | Default 24 lines, configurable range |
| `GetScreenLines` | Read multiple screen lines | Trim/format support |

#### Workflow Optimization
- **Line Processing Order**: FNL first, then R - prevents "line not finished" prompts
- **Smart Screen Waits**: Removed incorrect `WaitForScreenTransition` after R command (prompts appear immediately, not after screen transitions)
- **Adaptive Timing**: `WaitForScreenStable(2000, 300)` - smart 2-second max wait instead of fixed 500ms

#### Enhanced Logging
- Medium verbosity for key operations shows actual values being sent
- Screen snapshots available for debugging complex prompt sequences
- Comprehensive event logging with criticality levels (COMMON → CRITICAL)

## Configuration

### Input Files
- **CSV**: `PostFinalCharges\CashoutRoList.csv` (configured in `config.ini`)
- **Sequence Range**: Lines 30-50 (configured in `[Processing]` section)

### Dependencies
- BlueZone terminal emulator with active CDK DMS session
- `common\PathHelper.vbs` for path resolution
- `common\ValidateSetup.vbs` for dependency validation
- `config.ini` with `[PostFinalCharges_Main]` and `[Processing]` sections

## Running the Script

### From BlueZone (Recommended)
```vbscript
' Double-click PostFinalCharges.vbs or run:
cscript.exe PostFinalCharges.vbs
```

### Prerequisites
1. CDK_BASE environment variable set (run `tools\setup_cdk_base.vbs`)
2. Active BlueZone session connected to CDK DMS
3. CSV file with RO numbers to process
4. Sequence range configured in `config.ini`

## Testing

The script includes a comprehensive test suite covering:
- **Default value detection** (15 test cases including intervening text patterns)
- **Prompt handling** (conditional logic, error scenarios)
- **Integration tests** (MockBzhao simulations)
- **Bug prevention** (regression tests for known issues)

**Run all PostFinalCharges tests:**
```cmd
cd tools
cscript run_all_tests.vbs
```

**Or run specific test categories:**
```cmd
cscript run_default_value_tests.vbs     # Default value detection tests
cscript test_integration.vbs             # Integration tests
cscript test_bug_prevention.vbs          # Bug prevention tests
```

See [../docs/TESTING_STRATEGY.md](../docs/TESTING_STRATEGY.md) for testing documentation.

## Troubleshooting

**Script can't find config values:**
- Verify `[Processing]` section exists in `config.ini`
- Check `StartSequenceNumber` and `EndSequenceNumber` are defined
- See [CONFIG_FIX_SUMMARY.md](../CONFIG_FIX_SUMMARY.md)

**WaitForPrompt errors:**
- All required functions are now built-in (no external dependencies)
- See [WAITFORPROMPT_FIX_SUMMARY.md](../WAITFORPROMPT_FIX_SUMMARY.md)

**Line processing errors:**
- Script uses FNL→R workflow automatically
- Dynamic line letter discovery prevents infinite loops on non-consecutive lines

**Timing issues:**
- Default 2-second max wait with adaptive polling (300ms intervals)
- Create `.debug` file in same directory to enable slow-mode logging

## Legacy Notice
This script is part of a legacy system scheduled for sunset in 3-6 months. Prioritizes simplicity and immediate utility over long-term maintainability.
