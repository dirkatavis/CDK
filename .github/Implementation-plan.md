Implementation Plan - File Structure Reorganization
Goal
Simplify the project structure by enforcing a "Scripts / Lib / Config" pattern, leveraging the existing CDK_BASE strategy for robustness.

Proposed Structure
CDK/
├── .cdkroot
├── config/
│   └── config.ini          # [MOVED from root]
├── common/                 # [KEEP/MOVE PathHelper.vbs here]
│   └── PathHelper.vbs
├── scripts/
│   ├── close_ros/          # [MOVED from Close_ROs/]
│   │   ├── Close_ROs_Pt1.vbs
│   │   ├── Close_ROs_Pt2.vbs
│   │   └── ...
│   ├── create_ros/         # [MOVED from CreateNew_ROs/]
│   │   └── Create_ROs.vbs
│   ├── maintenance/        # [MOVED from Maintenance_RO_Closer/]
│   │   └── Maintenance_RO_Closer.vbs
│   └── final_charges/      # [MOVED from PostFinalCharges/]
│       ├── PostFinalCharges.vbs
│       └── tests/
├── tools/                  # [KEEP]
└── docs/                   # [KEEP]
Migration Steps
1. Structure Creation
Create config/, scripts/.
Create subdirectories in scripts/.
2. File Relocation
Move 
config.ini
 to config/.
Keep 
common/PathHelper.vbs
 in common/.
Move script directories to scripts/*.
Cleanup: Delete duplicate 
PostFinalCharges.vbs
 found in Close_ROs/ during move.
3. Code Updates
A. 
common/PathHelper.vbs
Update GetConfigPath to look for 
config.ini
 in config/config.ini (relative to root).
B. config/config.ini
Update all paths to reflect new locations (e.g., Close_ROs\ -> scripts\close_ros\).
C. All Scripts (
.vbs
)
Run all tests. if any test fails, fix it and run all tests again.   
Update the Bootstrap section to load 
common\PathHelper.vbs
 (no change needed if path is relative to root, but bootstrap logic might need tweak).
Validation Plan
Phase 0: Pre-Migration Validation
Baseline Check: Run 
tools\validate_dependencies.vbs
 to confirm current state is green.
Config Backup: Create a backup of 
config.ini
 (config.ini.bak).
Path Scan: Run 
tools\scan_hardcoded_paths.vbs
 to identify any existing hardcoded paths.
Full Backup: Create a zip or copy of the entire CDK folder (in case of catastrophic failure).
Phase 4: Post-Migration Validation
Dependency Check: Run 
tools\validate_dependencies.vbs
 (after updating it).
Expectation: All paths resolve, CDK_BASE is found.
Path Resolution Test: Run 
tools\test_path_helper.vbs
.
Expectation: PathHelper correctly finding 
config.ini
 in its new location.
Hardcoded Path Scan: Run 
tools\scan_hardcoded_paths.vbs
 again.
Expectation: Zero hardcoded paths found (except in comments/logs).
Script Launch Test:
Dry-run scripts\close_ros\Close_ROs_Pt1.vbs.
Rollback Plan
If any critical validation fails:

Restore 
config.ini
 from backup.
Move folders back to root.
Restore 
common\PathHelper.vbs
.