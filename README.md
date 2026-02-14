# CDK DMS Automation Scripts

This repository contains VBScript automation for the CDK Dealership Management System (DMS) using BlueZone terminal emulator.

## Quick Start for End Users

1. **Copy the entire CDK folder** to your machine (anywhere you want)
2. **Run setup**: Double-click `tools\setup_cdk_base.vbs` to set `CDK_BASE`
3. **Restart BlueZone** so it picks up the new variable
4. **Test the setup**: Run `tools\test_path_helper.vbs` from BlueZone
5. **Run scripts**: Double-click any `.vbs` script or run from command line

The scripts automatically find their files - no configuration needed!

## Important: Moving the Repository

When you move or copy the CDK folder to a new location:
- ✅ All scripts automatically adjust
- ✅ No paths need to be edited
- ✅ No configuration changes required

Just copy the whole folder and run - it works immediately.

## Repository Structure

```
CDK/
├── .cdkroot                    # Don't delete - scripts need this!
├── config.ini                  # Edit this to change file locations
├── common/
│   └── PathHelper.vbs          # Shared path functions
├── Close_ROs/                  # RO closing automation
├── CreateNew_ROs/              # RO creation automation
├── Maintenance_RO_Closer/      # Maintenance RO automation
├── PostFinalCharges/           # Final charge posting
├── tools/                      # Utility scripts
└── docs/                       # Documentation
```

## For Developers

- **Path Configuration**: See [docs/PATH_CONFIGURATION.md](docs/PATH_CONFIGURATION.md)
- **Contributing**: See [.github/copilot-instructions.md](.github/copilot-instructions.md)
- **Testing**: Run `tools\test_path_helper.vbs` to validate path setup

## Distributing to Other Users

When sharing these scripts:
1. Copy the **entire CDK folder** (including `.cdkroot`)
2. Share via USB, network, email - any method works
3. Users copy to their machine and run - no setup needed

## Troubleshooting

**Scripts can't find files?**
- Run `tools\show_cdk_base.vbs` to confirm `CDK_BASE`
- Re-run `tools\setup_cdk_base.vbs` to reset `CDK_BASE`
- Restart BlueZone after setting the variable
- Verify `.cdkroot` file exists in the repo root
- Run `tools\test_path_helper.vbs` to diagnose
- Check `config.ini` has correct relative paths

**Want to reorganize files?**
- Edit `config.ini` (don't move files manually)
- All scripts will automatically use new locations

## Legacy Notice

This is a legacy system scheduled for sunset in 3-6 months. The codebase prioritizes simplicity and immediate utility over long-term maintainability.
