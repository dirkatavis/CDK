# CDK DMS Automation Scripts

This repository contains VBScript automation for the CDK Dealership Management System (DMS) using BlueZone terminal emulator.

## Documentation

- **For End Users**: [docs/USER_SETUP.md](docs/USER_SETUP.md) (Installation & Verification)
- **For Developers**: [docs/DEVELOPER_HANDBOOK.md](docs/DEVELOPER_HANDBOOK.md) (Architecture & Patterns)
- **For Distribution**: [docs/DISTRIBUTION.md](docs/DISTRIBUTION.md) (Packaging Guide)

## Quick Start for End Users

1. **Extract the CDK folder** to your machine.
2. **Run setup**: Double-click `tools/setup_cdk_base.vbs` to set environment.
3. **Restart BlueZone** to refresh variables.
4. **Test the setup**: Run `tools/validate_dependencies.vbs`.

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
├── framework/
│   └── PathHelper.vbs          # Shared path functions
├── apps/
│   └── repair_order/           # RO workflow automation scripts
│       ├── initialize/         # Create new ROs from CSV
│       ├── prepare_close_pt1/  # Pre-closeout (before manual step)
│       └── finalize_close_pt2/ # Post-closeout (after manual step)
├── utilities/                  # Standalone utility scripts (legacy location)
├── tools/                      # Setup, testing, validation scripts
└── docs/                       # Documentation
```

### Script Purposes

**apps/repair_order/** - Sequential RO processing workflow
- **initialize/**: Create new repair orders from CSV input (MVA/mileage entry)
- **prepare_close_pt1/**: Pre-manual processing - Dynamic line discovery
- **finalize_close_pt2/**: Post-manual processing - Finalize closeout, add stories
- Pattern: "Sandwich automation" with manual intervention between Pt1 and Pt2

**apps/post_final_charges/** - State-machine based RO closeout
- Handles 30+ conditional prompts, multi-line processing, FNL→R workflow
- Production ready - successfully tested with live BlueZone terminal

**apps/maintenance_ro_closer/** - Automated PM (preventive maintenance) processing
- Matches criteria from PM_Match_Criteria.txt, processes ROs from list
- Generates status reports in RO_Status_Report.csv

## Technical Documentation

For developers, automated systems, or deep technical dives:
- **Architecture & Setup**: [docs/DEVELOPER_HANDBOOK.md](docs/DEVELOPER_HANDBOOK.md)
- **Path Configuration**: [docs/PATH_CONFIGURATION.md](docs/PATH_CONFIGURATION.md)
- **Validation Layers**: [docs/VALIDATION_ARCHITECTURE.md](docs/VALIDATION_ARCHITECTURE.md)

## Distributing to Other Users

When sharing these scripts:
1. Copy the **entire CDK folder** (including `.cdkroot`)
2. Share via USB, network, email - any method works
3. Users copy to their machine and run - no setup needed

## Troubleshooting

**Scripts can't find files?**
- Run `tools/show_cdk_base.vbs` to confirm `CDK_BASE`
- Re-run `tools/setup_cdk_base.vbs` to reset `CDK_BASE`
- Restart BlueZone after setting the variable
- Verify `.cdkroot` file exists in the repo root
- Run `tests/test_path_helper.vbs` to diagnose
- Check `config/config.ini` has correct relative paths

**Want to reorganize files?**
- Edit `config/config.ini` (don't move files manually)
- All scripts will automatically use new locations

## Legacy Notice

This is a legacy system scheduled for sunset in 3-6 months. The codebase prioritizes simplicity and immediate utility over long-term maintainability.
