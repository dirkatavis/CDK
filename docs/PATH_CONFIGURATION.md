# CDK Centralized Path Configuration

## Overview
This feature centralizes all file path references in a single `config.ini` file, making it easy to:
- Move the repository to any location without editing scripts
- Distribute scripts to other users via sneakernet
- Reorganize folders by editing one config file

## How It Works

### 1. `CDK_BASE` Environment Variable (Required)
BlueZone runs scripts with a working directory of the BlueZone install folder. To locate the repo reliably, each machine must set the user environment variable:

```
CDK_BASE=C:\Temp_alt\CDK
```

### 2. `.cdkroot` Marker File (Required)
An empty file at the repository root that scripts use to validate the base path. **This file is required** - scripts will fail with a clear error if it's missing.

### 3. `config.ini` Configuration
All relative file paths organized by script/function:
```ini
[Close_ROs_Pt1]
CSV=Close_ROs\Close_ROs_Pt1.csv
Log=Close_ROs\Close_ROs_Pt1.log
```

### 4. `common\PathHelper.vbs` Helper Module
Reusable functions that scripts include to:
- Read the repo root from `CDK_BASE` and validate `.cdkroot`
- Read paths from `config.ini`
- Build absolute paths at runtime
- **Fail fast** with clear error messages if `.cdkroot` is missing (no silent fallbacks)

## For End Users (Non-Developers)

### Installation
1. Copy the entire CDK folder to your machine (anywhere you want)
2. Run `tools\setup_cdk_base.vbs` to set `CDK_BASE` automatically
3. Restart BlueZone so it picks up the new variable
4. Run `tools\test_path_helper.vbs` to verify setup
5. Scripts automatically find their files

### Moving the Repository
Just move the entire folder - scripts auto-adjust. No configuration needed.

## For Developers

### Testing the Path System
Run `tools\test_path_helper.vbs` from BlueZone:
- Shows discovered repo root
- Shows sample config paths
- Writes report to temp folder

### Quick Check
Run `tools\show_cdk_base.vbs` to display the current `CDK_BASE` value.

### Updating a Script to Use PathHelper

**Before:**
```vbscript
Const CSV_FILE = "C:\Temp_alt\CDK\Close_ROs\file.csv"
```

**After:**
```vbscript
' Load PathHelper (at script start)
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Const BASE_ENV_VAR = "CDK_BASE"
Dim sh: Set sh = CreateObject("WScript.Shell")
Dim basePath: basePath = sh.Environment("USER")(BASE_ENV_VAR)

Dim helperPath: helperPath = g_fso.BuildPath(basePath, "common\PathHelper.vbs")
ExecuteGlobal g_fso.OpenTextFile(helperPath).ReadAll

' Use config paths
Dim CSV_FILE: CSV_FILE = GetConfigPath("Close_ROs_Pt1", "CSV")
```

### Adding New Paths
1. Edit `config.ini` and add your section/keys
2. Use `GetConfigPath(section, key)` in your script

## Migration Status

### ✅ Updated Scripts
- Close_ROs\Close_ROs_Pt1.vbs

### ⏳ Pending Updates
- Close_ROs\Close_ROs_Pt2.vbs
- Close_ROs\PostFinalCharges.vbs
- Close_ROs\HighestRoFinder.vbs
- Close_ROs\TestLog.vbs
- CreateNew_ROs\Create_ROs.vbs
- CreateNew_ROs\Parse_Data.ps1
- Maintenance_RO_Closer\Maintenance_RO_Closer.vbs
- Maintenance_RO_Closer\Coordinate_Finder.vbs

## Files Created
- `.cdkroot` - Repo root marker
- `config.ini` - Centralized path configuration
- `common/PathHelper.vbs` - Path helper module
- `tools/test_path_helper.vbs` - Validation script
