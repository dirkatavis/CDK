# CDK Distribution Package Guide

This guide explains how to package and distribute the CDK automation system to other users.

## What to Include

### Required Files
```
CDK/
├── .cdkroot                           # Repo marker (create empty file)
├── config.ini                         # Central configuration
├── common/
│   ├── PathHelper.vbs                # Path resolution library
│   └── ValidateSetup.vbs             # Validation library (BlueZone-safe)
├── tools/
│   ├── validate_dependencies.vbs      # Pre-flight check tool
│   ├── setup_cdk_base.vbs            # Environment setup tool
│   ├── show_cdk_base.vbs             # Check current CDK_BASE
│   ├── test_validation_positive.vbs  # Test suite (optional)
│   ├── test_validation_negative.vbs  # Test suite (optional)
│   └── run_validation_tests.vbs      # Test runner (optional)
├── docs/
│   ├── SETUP_VALIDATION.md           # User validation guide
│   ├── VALIDATION_ARCHITECTURE.md    # Technical architecture
│   ├── BLUEZONE_COMPATIBILITY.md     # BlueZone context guide
│   └── PATH_CONFIGURATION.md         # config.ini documentation
└── [automation scripts and data folders]
```

### Core Automation Scripts (Examples)
- `utilities/PostFinalCharges.vbs`
- `workflows/repair_order/1_Initialize_RO.vbs`
- `workflows/repair_order/2_Prepare_Close_Pt1.vbs`
- `workflows/repair_order/3_Finalize_Close_Pt2.vbs`
- etc.

## Distribution Checklist

### Before Creating Package

- [ ] Ensure all paths in `config.ini` are correct for your environment
- [ ] Test validation: `cscript.exe tools\validate_dependencies.vbs`
- [ ] Test all tools: `cscript.exe tools\run_validation_tests.vbs`
- [ ] Verify all automation scripts have included ValidateSetup
- [ ] Confirm all required documentation is present

### Package Contents Checklist

**Critical (Must Include):**
- [ ] `.cdkroot` marker file
- [ ] `config.ini` with all required sections
- [ ] `common/PathHelper.vbs`
- [ ] `common/ValidateSetup.vbs`
- [ ] `tools/validate_dependencies.vbs`
- [ ] `tools/setup_cdk_base.vbs`
- [ ] `docs/SETUP_VALIDATION.md`

**Recommended (Should Include):**
- [ ] `tools/show_cdk_base.vbs`
- [ ] `docs/PATH_CONFIGURATION.md`
- [ ] `docs/VALIDATION_ARCHITECTURE.md`
- [ ] `docs/BLUEZONE_COMPATIBILITY.md`

**Optional (Nice to Have):**
- [ ] Test suites (`test_validation_*.vbs`)
- [ ] `tools/run_validation_tests.vbs`
- [ ] `docs/IMPLEMENTATION_SUMMARY.md`

## Installation Instructions for Recipients

### Step 1: Initial Setup (One-Time)

1. **Extract** CDK package to desired location (e.g., `C:\Automation\CDK`)

2. **Set up environment variable:**
   ```cmd
   cd C:\Automation\CDK
   cscript.exe tools\setup_cdk_base.vbs
   ```
   
   This sets `CDK_BASE` pointing to your CDK installation.

3. **Restart terminal** to apply environment changes

### Step 2: Validate Installation

1. **Run pre-flight check:**
   ```cmd
   cd C:\Automation\CDK
   cscript.exe tools\validate_dependencies.vbs
   ```

2. **Review output:**
   - If all ✓ PASS: Ready to use!
   - If warnings: May still work, but review them
   - If failures: Address issues before continuing

3. **Optional - Run full test suite:**
   ```cmd
   cscript.exe tools\run_validation_tests.vbs
   ```

### Step 3: Use Automation Scripts

1. **In BlueZone terminal:**
   ```
   cscript.exe Z:\path\to\CDK\PostFinalCharges\PostFinalCharges.vbs
   ```

2. **Scripts will automatically:**
   - Validate dependencies
   - Read configuration from `config.ini`
   - Execute automation

## Troubleshooting for Recipients

### Setup Phase Issues

**"CDK_BASE environment variable not set"**
```cmd
cscript.exe tools\setup_cdk_base.vbs
```
Then close and reopen terminal.

**"CDK_BASE points to non-existent folder"**
Edit user environment variables to correct the path.

**"PathHelper.vbs not found"**
Ensure `common\PathHelper.vbs` exists in your CDK installation.

**"config.ini not found"**
Ensure `config.ini` exists at CDK root (where `.cdkroot` is).

### Runtime Issues

**"All dependencies validated but script won't run"**
1. Check BlueZone is running
2. Check mapped drive/network path is accessible
3. Review script log files for errors

**"Scripts run but use wrong paths"**
1. Verify `config.ini` paths are correct
2. Check `CDK_BASE` points to correct location
3. Run: `cscript.exe tools\show_cdk_base.vbs`

## Configuration Notes

### Updating config.ini After Distribution

If you need to update `config.ini` for recipients:

1. **Provide updated `config.ini`** file
2. **Recipients:**
   - Backup existing: `copy config.ini config.ini.backup`
   - Replace with new: `copy new-config.ini config.ini`
   - Re-validate: `cscript.exe tools\validate_dependencies.vbs`

### Path Configuration

All paths in `config.ini` are **relative to repo root** (where `.cdkroot` is).

Example structure:
```ini
[PostFinalCharges_Main]
CSV=PostFinalCharges\CashoutRoList.csv        # Relative path
Log=PostFinalCharges\PostFinalCharges.log     # Relative path
```

Gets resolved to (for recipient at `C:\CDK`):
```
C:\CDK\PostFinalCharges\CashoutRoList.csv
C:\CDK\PostFinalCharges\PostFinalCharges.log
```

## Support

### For Recipients

1. Run validation tool: `tools\validate_dependencies.vbs`
2. Review documentation: `docs\SETUP_VALIDATION.md`
3. Check configuration: `config.ini`

### For Administrators

When distributing to multiple users:

1. **One-time admin setup:**
   - Ensure network paths are shared
   - Set up shared `config.ini` if needed
   - Verify all required files are present

2. **Per-user setup:**
   - Have each user run `setup_cdk_base.vbs`
   - Have each user run `validate_dependencies.vbs`
   - Document any custom configuration needed

3. **Ongoing maintenance:**
   - Update `config.ini` as paths change
   - Re-run validation after updates
   - Maintain backup of working config

## Validation System Benefits

When recipients follow this distribution model:

✅ **Automatic validation** - Scripts check dependencies before running  
✅ **Clear error messages** - Users know exactly what's wrong  
✅ **Self-service diagnostics** - `validate_dependencies.vbs` provides solutions  
✅ **Config centralization** - All paths in one `config.ini` file  
✅ **Path flexibility** - Works from any installation location  
✅ **BlueZone compatibility** - Works in both standalone and terminal contexts  

## Frequently Asked Questions

**Q: Do I need to run validation every time?**  
A: No. Validation runs automatically at script startup. The tool is for diagnostics if something fails.

**Q: Can I move the CDK folder after installation?**  
A: Yes, but you must update the `CDK_BASE` environment variable:
```cmd
cscript.exe tools\setup_cdk_base.vbs  [new-location]
```

**Q: What if I don't have admin rights to set environment variables?**  
A: Users can set `CDK_BASE` in their user environment variables (not system-wide). The tool handles this automatically.

**Q: Can multiple users share one CDK installation?**  
A: Yes, as long as they all have access to the same network path and each sets their own `CDK_BASE` to point to it.

**Q: What if config.ini paths need to be different per user?**  
A: Each user can have their own `config.ini`, or use relative paths that work from any location.

## Next Steps

After distribution:

1. **Monitor** initial user setup
2. **Collect feedback** on validation experience
3. **Update documentation** based on common issues
4. **Plan maintenance** for config.ini updates
5. **Track automation success** rates and issues
