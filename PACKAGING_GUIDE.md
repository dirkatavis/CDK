# CDK Automation System - Complete Packaging Guide

## Quick Start

For distributing CDK to others, the validation and configuration system provides complete self-service setup:

### What Recipients Need to Do

```cmd
# Step 1: Extract CDK package to desired location, then:

# Step 2: One-time setup (set CDK_BASE environment variable)
cscript.exe tools\setup_cdk_base.vbs

# Step 3: Close and reopen terminal

# Step 4: Validate everything is working
cscript.exe tools\validate_dependencies.vbs

# Step 5: Ready to use! Scripts run automatically
```

If validation passes, everything is ready. If it fails, validation provides clear remediation steps.

## The Three Pillars of the Packaging System

### 1. Centralized Configuration (`config.ini`)
- **What:** Single file defining all paths and parameters
- **Where:** At repository root
- **Purpose:** Scripts read from here instead of hardcoding paths
- **Benefit:** Change paths once in config.ini, all scripts use new paths

Example:
```ini
[PostFinalCharges_Main]
CSV=PostFinalCharges\CashoutRoList.csv
Log=PostFinalCharges\PostFinalCharges.log

[Close_ROs_Pt1]
CSV=Close_ROs\Close_ROs_Pt1.csv
Log=Close_ROs\Close_ROs_Pt1.log
```

### 2. Path Resolution Library (`common\PathHelper.vbs`)
- **What:** Shared library for reading config and resolving paths
- **Where:** In `common` folder
- **Purpose:** All scripts use this to get paths from config.ini
- **Benefit:** Consistent path handling across all scripts

Functions provided:
- `GetRepoRoot()` - Get base installation folder
- `GetConfigPath(section, key)` - Read value from config.ini
- `ResolvePath(relPath, baseFolder)` - Convert relative to absolute paths

### 3. Dependency Validation System (`common\ValidateSetup.vbs` + `tools\validate_dependencies.vbs`)
- **What:** Automatic checks that dependencies are available
- **Where:** ValidateSetup in `common` (included in scripts), tool in `tools`
- **Purpose:** Catch missing setup before scripts fail
- **Benefit:** Clear error messages guide users to fix issues

Checks performed:
- ✓ CDK_BASE environment variable is set
- ✓ Repository structure is intact (.cdkroot marker)
- ✓ Path resolution library exists (PathHelper.vbs)
- ✓ Central configuration exists (config.ini)
- ✓ All configured paths are accessible

## Key Files for Distribution

### Essential (Always Include)
| File | Purpose |
|------|---------|
| `.cdkroot` | Repository marker (validates structure) |
| `config.ini` | Central configuration (paths and parameters) |
| `common/PathHelper.vbs` | Path resolution library |
| `common/ValidateSetup.vbs` | Validation library (both contexts) |
| `tools/validate_dependencies.vbs` | Standalone validation tool |
| `tools/setup_cdk_base.vbs` | Environment setup script |

### Recommended (Enhance User Experience)
| File | Purpose |
|------|---------|
| `docs/SETUP_VALIDATION.md` | User setup and validation guide |
| `docs/PATH_CONFIGURATION.md` | config.ini structure documentation |
| `tools/show_cdk_base.vbs` | Check current CDK_BASE setting |

### Optional (For Testing)
| File | Purpose |
|------|---------|
| `tools/test_validation_positive.vbs` | Verify validation passes |
| `tools/test_validation_negative.vbs` | Verify validation catches failures |
| `tools/run_validation_tests.vbs` | Run all validation tests |

## How It Works: The Flow

### Before Installation
```
User receives: CDK package with all required files
```

### Step 1: Initial Setup
```
User runs: cscript.exe tools\setup_cdk_base.vbs
Result: CDK_BASE environment variable is set to CDK folder
```

### Step 2: Pre-Flight Check
```
User runs: cscript.exe tools\validate_dependencies.vbs

Validation checks:
✓ CDK_BASE environment variable is set
✓ CDK_BASE points to valid folder
✓ .cdkroot marker file exists
✓ PathHelper.vbs exists
✓ config.ini exists and is valid
✓ All configured paths are accessible

Output: PASS or FAIL with remediation steps
```

### Step 3: Script Execution
```
User runs: cscript.exe PostFinalCharges\PostFinalCharges.vbs

Script startup:
→ Load PathHelper.vbs (from common folder)
→ Load ValidateSetup.vbs (from common folder)
→ Call MustHaveValidDependencies() [FIRST THING]
  ✓ All dependencies pass → Continue
  ✗ Any fail → Show error, exit
→ Read paths from config.ini via GetConfigPath()
→ Execute automation logic
```

## BlueZone Compatibility

The validation system works in **two contexts:**

### Context 1: Standalone (Before BlueZone)
```cmd
cscript.exe tools\validate_dependencies.vbs
# Works: Full WScript support, clear console output
```

### Context 2: Inside BlueZone
```
[In BlueZone terminal]
cscript.exe C:\path\to\PostFinalCharges.vbs
# Works: SafeOutput adapts to BlueZone logging
```

**Key:** ValidateSetup.vbs automatically detects context using:
- `SafeOutput()` - Outputs to console (WScript) or log (BlueZone)
- `MustHaveValidDependencies()` - Exits (WScript) or sets flag (BlueZone)

See [docs/BLUEZONE_COMPATIBILITY.md](docs/BLUEZONE_COMPATIBILITY.md) for details.

## Configuration Management

### How Paths Work

**In config.ini:** All paths are **relative to repo root**
```ini
CSV=PostFinalCharges\CashoutRoList.csv
```

**When loaded:**
```vbscript
' Script code:
csvPath = GetConfigPath("PostFinalCharges_Main", "CSV")

' Result (at any installation):
' If installed at C:\Automation\CDK → C:\Automation\CDK\PostFinalCharges\CashoutRoList.csv
' If installed at D:\Tools\CDK → D:\Tools\CDK\PostFinalCharges\CashoutRoList.csv
```

**Benefit:** Move CDK folder anywhere, scripts automatically use correct paths.

### Adding New Paths

For new scripts or files:

1. **Add to config.ini:**
```ini
[MyNewScript]
DataFile=MyFolder\data.csv
OutputLog=MyFolder\Myscript.log
```

2. **In script, use:**
```vbscript
dataPath = GetConfigPath("MyNewScript", "DataFile")
logPath = GetConfigPath("MyNewScript", "OutputLog")
```

3. **No script modification needed** when paths change - just update config.ini.

## Validation Test Results

### Positive Tests (Should Pass)
✓ CDK_BASE is valid  
✓ .cdkroot marker exists  
✓ PathHelper.vbs exists  
✓ ValidateSetup.vbs exists  
✓ config.ini exists with valid format  
✓ Critical paths are accessible  
✓ Full validation passes  

### Negative Tests (Catch Problems)
✓ Detects missing CDK_BASE  
✓ Detects invalid CDK_BASE path  
✓ Detects missing .cdkroot  
✓ Detects missing PathHelper.vbs  
✓ Detects missing config.ini  
✓ Handles corrupted config.ini  

Run tests: `cscript.exe tools\run_validation_tests.vbs`

## Documentation Structure

```
docs/
├── DISTRIBUTION.md              # Guide for packaging and distributing
├── SETUP_VALIDATION.md          # User setup and troubleshooting
├── VALIDATION_ARCHITECTURE.md   # Technical design and testing
├── BLUEZONE_COMPATIBILITY.md    # How validation works in both contexts
├── PATH_CONFIGURATION.md        # config.ini structure and usage
└── ... other documentation
```

## Common Scenarios

### Scenario 1: New User Gets CDK

```
1. Extract CDK package
2. Run: tools\setup_cdk_base.vbs [folder path]
3. Restart terminal
4. Run: tools\validate_dependencies.vbs
5. If all ✓: Ready to use scripts
6. If ✗: Follow remediation steps shown
```

### Scenario 2: Move CDK to New Location

```
1. Move CDK folder to new location
2. Run: tools\setup_cdk_base.vbs [new path]
3. Run: tools\validate_dependencies.vbs
4. Scripts now use correct paths automatically
```

### Scenario 3: Update Paths in Configuration

```
1. Edit config.ini with new paths
2. Run: tools\validate_dependencies.vbs
3. If all ✓: Scripts automatically use new paths
4. If ✗: Fix path issues and try again
```

### Scenario 4: Debug Script Failure

```
Step 1: Run validation tool
  cscript.exe tools\validate_dependencies.vbs

Step 2: Check script logs
  Review log file configured in config.ini

Step 3: Verify paths exist
  Check folders/files referenced in config.ini

Step 4: Test path helper
  cscript.exe tools\test_path_helper.vbs

Step 5: Run full test suite
  cscript.exe tools\run_validation_tests.vbs
```

## Benefits for Distribution

✅ **Self-Service Setup** - Users can set up without admin help  
✅ **Automatic Diagnostics** - Tools show exactly what's wrong  
✅ **Flexible Paths** - Works from any installation location  
✅ **Centralized Config** - One file to update for all paths  
✅ **Robust Design** - Works in both standalone and BlueZone contexts  
✅ **Clear Documentation** - Guides users through setup and troubleshooting  
✅ **Tested System** - Validation tested with comprehensive test suites  

## Next Steps

1. **Prepare for distribution:**
   - Verify all files are present
   - Test validation: `tools\validate_dependencies.vbs`
   - Run test suite: `tools\run_validation_tests.vbs`

2. **Create distribution package:**
   - Include all essential files
   - Include relevant documentation
   - Follow checklist in `docs/DISTRIBUTION.md`

3. **Provide to users:**
   - Include setup instructions
   - Point to `docs/SETUP_VALIDATION.md`
   - Include contact info for support

4. **Monitor adoption:**
   - Collect setup feedback
   - Update documentation based on issues
   - Track automation success

## References

- [Path Configuration Guide](docs/PATH_CONFIGURATION.md)
- [Validation Architecture](docs/VALIDATION_ARCHITECTURE.md)
- [Setup & Validation Guide](docs/SETUP_VALIDATION.md)
- [BlueZone Compatibility](docs/BLUEZONE_COMPATIBILITY.md)
- [Distribution Guide](docs/DISTRIBUTION.md)
- [Tools README](tools/README.md)
