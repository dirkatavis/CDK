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
├── config/
│   └── config.ini              # Edit this to change file locations
├── common/
│   └── PathHelper.vbs          # Shared path functions
├── workflows/
│   └── repair_order/           # RO workflow automation scripts
│       ├── 1_Initialize_RO.vbs      # Create new ROs from CSV
│       ├── 2_Prepare_Close_Pt1.vbs  # Pre-closeout (before manual step)
│       └── 3_Finalize_Close_Pt2.vbs # Post-closeout (after manual step)
├── utilities/                  # Standalone utility scripts
│   ├── PostFinalCharges.vbs    # Complete RO closeout (state machine)
│   └── Maintenance_RO_Closer.vbs # Automated PM RO processing
├── tools/                      # Setup, testing, validation scripts
└── docs/                       # Documentation
```

### Script Purposes

**workflows/repair_order/** - Sequential RO processing workflow
- **1_Initialize_RO.vbs**: Create new repair orders from CSV input (MVA/mileage entry)
- **2_Prepare_Close_Pt1.vbs**: Pre-manual processing - Dynamic line discovery, initial closeout steps
- **3_Finalize_Close_Pt2.vbs**: Post-manual processing - Finalize closeout, add stories
- Pattern: "Sandwich automation" with manual intervention between Pt1 and Pt2

**utilities/** - Standalone automation tools
- **PostFinalCharges.vbs**: Complete automated RO closeout using state machine logic
  - Handles 30+ conditional prompts, multi-line processing, FNL→R workflow
  - Production ready - successfully tested with live BlueZone terminal
  - See [utilities/README.md](utilities/README.md) for detailed features
- **Maintenance_RO_Closer.vbs**: Automated PM (preventive maintenance) RO processing
  - Matches criteria from PM_Match_Criteria.txt, processes ROs from list
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
- Edit `config/config.ini` (don't move files manually)
- All scripts will automatically use new locations

## Legacy Notice

This is a legacy system scheduled for sunset in 3-6 months. The codebase prioritizes simplicity and immediate utility over long-term maintainability.
