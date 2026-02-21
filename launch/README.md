# Launch - Legacy Compatibility Wrappers

## ⚠️ TEMPORARY - REMOVABLE AT SUNSET

This folder contains backward compatibility wrappers to preserve legacy launch paths while the repository is being restructured. **These wrappers are temporary and will be removed when the system is retired (3-6 months).**

## Purpose
Operators may have existing documentation, shortcuts, or muscle memory pointing to old script paths like:
- `workflows\repair_order\1_Initialize_RO.vbs`
- `utilities\PostFinalCharges.vbs`
- `tools\ValidateRoList\ValidateRoList.vbs`

These wrappers forward to the new locations in `apps/` without requiring updates to existing documentation or operator workflows.

## How Wrappers Work
Each wrapper:
1. Validates `CDK_BASE` environment variable
2. Verifies `.cdkroot` marker file exists
3. Checks target script exists in `apps/` folder
4. Loads and executes the target script via `ExecuteGlobal`

Example flow:
```
launch\PostFinalCharges.vbs
  ↓ (wrapper forwards to)
apps\post_final_charges\PostFinalCharges.vbs
  ↓ (actual implementation)
```

## Wrapper List

| Legacy Entry Point | Target Location |
|--------------------|-----------------|
| `1_Initialize_RO.vbs` | `apps/repair_order/initialize/1_Initialize_RO.vbs` |
| `2_Prepare_Close_Pt1.vbs` | `apps/repair_order/prepare_close_pt1/2_Prepare_Close_Pt1.vbs` |
| `3_Finalize_Close_Pt2.vbs` | `apps/repair_order/finalize_close_pt2/3_Finalize_Close_Pt2.vbs` |
| `PostFinalCharges.vbs` | `apps/post_final_charges/PostFinalCharges.vbs` |
| `Maintenance_RO_Closer.vbs` | `apps/maintenance_ro_closer/Maintenance_RO_Closer.vbs` |
| `PFC_Scrapper.vbs` | `apps/pfc_scrapper/PFC_Scrapper.vbs` |
| `HighestRoFinder.vbs` | `apps/highest_ro_finder/HighestRoFinder.vbs` |
| `ValidateRoList.vbs` | `apps/validate_ro_list/ValidateRoList.vbs` |

## Usage (Temporary)
During transition, operators can launch from either location:

**Option 1: Legacy path (via wrapper)**
```cmd
cd C:\Temp_alt\CDK
cscript.exe launch\PostFinalCharges.vbs
```

**Option 2: Direct path (recommended for new workflows)**
```cmd
cd C:\Temp_alt\CDK\apps\post_final_charges
cscript.exe PostFinalCharges.vbs
```

Both invoke the same script - wrappers just forward to `apps/`.

## Migration Plan
1. **Current state:** Wrappers active in `launch/` folder
2. **Transition period:** Operators can use either path
3. **Update documentation:** Point to `apps/` locations directly
4. **Sunset:** Remove `launch/` folder entirely when system is retired

## Notes
- Do **NOT** add business logic to wrappers - they only forward
- Do **NOT** extend wrappers for new features - they are frozen and temporary
- Do **NOT** test wrappers separately - they validate via repo-level contract tests
- **DO** update operator documentation to reference `apps/` directly
- **DO** remove this entire folder at sunset (3-6 months)

## Testing
Wrappers are validated by:
- `tooling/test_reorg_contract_wrappers.vbs` - Ensures all wrapper targets exist
- `tooling/run_validation_tests.vbs` - Includes wrapper compatibility checks

## For Developers
When adding new apps:
1. Create app in `apps/` folder
2. Do **NOT** create a wrapper in `launch/`
3. Reference the `apps/` location directly
4. This folder is for **existing legacy paths only**
