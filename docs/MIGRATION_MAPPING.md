# Repository Reorganization Migration Mapping

## Migration Date: February 21, 2026
## Goal: Domain-first folder structure (apps/framework/tooling)

---

## ✅ KEEP (New Structure - These Stay)

| Folder | Purpose | Contents |
|--------|---------|----------|
| `apps/` | Production workflows | Self-contained apps with tests/README |
| `framework/` | Shared reusable components | PathHelper.vbs, ValidateSetup.vbs, HostCompat.vbs |
| `tooling/` | Setup/diagnostics/testing | setup_cdk_base.vbs, validate_dependencies.vbs, test scripts |
| `tests/` | Repo-level global tests | Cross-cutting infrastructure tests |
| `runtime/` | Generated artifacts | logs/, data/out/ (created at runtime) |

---

## ❌ DELETE (Old Structure - Remove These)

| Folder | Reason | Replaced By |
|--------|--------|-------------|
| `common/` | Renamed for clarity | `framework/` |
| `tools/` | Renamed for domain separation | `tooling/` |
| `utilities/` | Reorganized to self-contained apps | `apps/` (post_final_charges, maintenance_ro_closer, etc.) |
| `workflows/` | Reorganized to self-contained apps | `apps/repair_order/` |
| `Close_ROs/` | Legacy runtime output (stale logs) | `runtime/logs/` |
| `Maintenance_RO_Closer/` | Legacy runtime output | `runtime/logs/maintenance_ro_closer/` |
| `PostFinalCharges/` | Legacy runtime output | `runtime/logs/post_final_charges/` |

---

## 🔄 RENAME MAPPINGS (Old → New)

### Framework Components
```
common/PathHelper.vbs           → framework/PathHelper.vbs
common/ValidateSetup.vbs        → framework/ValidateSetup.vbs
common/HostCompat.vbs           → framework/HostCompat.vbs
```

### Tooling Scripts
```
tools/setup_cdk_base.vbs        → tooling/setup_cdk_base.vbs
tools/validate_dependencies.vbs → tooling/validate_dependencies.vbs
tools/scan_hardcoded_paths.vbs  → tooling/scan_hardcoded_paths.vbs
tools/test_*.vbs                → tooling/test_*.vbs (all test scripts)
tools/run_*.vbs                 → tooling/run_*.vbs (test runners)
tools/reorg_path_map.ini        → tooling/reorg_path_map.ini
```

### Production Apps
```
utilities/PostFinalCharges.vbs              → apps/post_final_charges/PostFinalCharges.vbs
utilities/Maintenance_RO_Closer.vbs         → apps/maintenance_ro_closer/Maintenance_RO_Closer.vbs
utilities/PFC_Scrapper.vbs                  → apps/pfc_scrapper/PFC_Scrapper.vbs
utilities/HighestRoFinder.vbs               → apps/highest_ro_finder/HighestRoFinder.vbs
tools/ValidateRoList/ValidateRoList.vbs     → apps/validate_ro_list/ValidateRoList.vbs

workflows/repair_order/1_Initialize_RO.vbs  → apps/repair_order/initialize/1_Initialize_RO.vbs
workflows/repair_order/2_Prepare_Close_Pt1.vbs → apps/repair_order/prepare_close_pt1/2_Prepare_Close_Pt1.vbs
workflows/repair_order/3_Finalize_Close_Pt2.vbs → apps/repair_order/finalize_close_pt2/3_Finalize_Close_Pt2.vbs
```

### App-Specific Tests
```
utilities/tests/*               → apps/post_final_charges/tests/* (PFC test suite)
utilities/tests/test_pfc_scrapper.vbs → apps/pfc_scrapper/tests/test_pfc_scrapper.vbs
```

### Global Tests
```
tools/test_validation_*.vbs     → tests/test_validation_*.vbs
tools/test_reorg_contract_*.vbs → tests/test_reorg_contract_*.vbs
tools/test_path_helper.vbs      → tests/test_path_helper.vbs
tools/test_reset_state.vbs      → tests/test_reset_state.vbs
tools/run_validation_tests.vbs  → tests/run_validation_tests.vbs
tools/run_migration_target_tests.vbs → tests/run_migration_target_tests.vbs
```

---

## 🔒 STAYS AS-IS (No Changes)

| Folder/File | Purpose |
|-------------|---------|
| `.cdkroot` | Repo marker file |
| `.github/` | GitHub workflows and documentation |
| `config/` | Configuration files (config.ini) |
| `docs/` | Documentation |
| `Install.vbs` | Root-level installer script |
| `README.md` | Main documentation |
| `PACKAGING_GUIDE.md` | Distribution guide |
| `.venv*` | Python virtual environments (gitignored) |

---

## 📋 Reference Updates Required

### Code References to Update:
- **All `.vbs` files**: `common\` → `framework\`
- **All `.vbs` files**: `tools\` → `tooling\`
- **Documentation**: Update all path examples
- **config.ini**: Update all app paths to `apps/*`
- **reorg_path_map.ini**: Update all entrypoints to `apps/*` (direct paths, no fallbacks)

---

## ✅ Validation Checklist

After migration:
- [x] All `apps/` scripts load from `framework/`
- [x] All `tooling/` scripts load from `framework/`
- [x] `config.ini` paths resolve to `apps/`, `runtime/`
- [x] Tests pass: `cscript tests\run_validation_tests.vbs`
- [x] Migration complete: `cscript tests\run_migration_target_tests.vbs` (100%)
- [x] Old folders deleted: `common/`, `tools/`, `utilities/`, `workflows/`, `Close_ROs/`, `Maintenance_RO_Closer/`, `PostFinalCharges/`
- [x] Fallback patterns removed: `launch/` deleted (fail-fast instead of silent redirect)

---

## 🚀 Execution Order (Completed)

1. ✅ **Create** new folders (`apps/`, `framework/`, `tooling/`, `tests/`)
2. ✅ **Copy** files to new locations
3. ✅ **Update** all internal references
4. ✅ **Validate** tests pass
5. ✅ **Delete** old folders (validation passed, cleanup complete)
6. ✅ **Remove fallbacks** (deleted `launch/` - fail-fast pattern enforced)

---

**Status**: ✅ Migration complete - New structure active, legacy folders removed
