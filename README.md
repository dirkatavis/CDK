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
├── Close_ROs/                  # RO closing automation (Pt1/Pt2 split workflow)
├── CreateNew_ROs/              # RO creation from CSV (MVA/mileage entry)
├── Maintenance_RO_Closer/      # Automated PM RO processing
├── PostFinalCharges/           # Complete RO closeout (state machine)
├── tools/                      # Utility scripts
└── docs/                       # Documentation
```

### Script Purposes

**Close_ROs/**
- **Close_ROs_Pt1.vbs**: Pre-manual processing - Dynamic line discovery, initial closeout steps before manual intervention
- **Close_ROs_Pt2.vbs**: Post-manual processing - Finalize closeout, add stories for discovered lines
- Pattern: "Sandwich automation" with manual step in between

**PostFinalCharges/**
- **PostFinalCharges.vbs**: Complete automated RO closeout using state machine logic
- Handles 30+ conditional prompts, multi-line processing, FNL→R workflow
- Production ready - successfully tested with live BlueZone terminal
- See [PostFinalCharges/README.md](PostFinalCharges/README.md) for detailed features

**CreateNew_ROs/**
- **Create_ROs.vbs**: Automated RO creation from CSV input
- Reads vehicle data (MVA/mileage), creates new ROs via terminal automation
- Supports .debug flag for slow-mode execution

**Maintenance_RO_Closer/**
- **Maintenance_RO_Closer.vbs**: Automated PM (preventive maintenance) RO processing
- Matches criteria from PM_Match_Criteria.txt, processes matching ROs from list
- Generates status reports in RO_Status_Report.csv

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
