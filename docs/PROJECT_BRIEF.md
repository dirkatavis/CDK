# CDK Automation Project Brief

**Purpose:** Context document for all development chats on this project. Paste relevant sections at the start of each new conversation to establish context without re-discovery.

**Last Updated:** April 2026  
**Project Status:** Active development — sunset planned in 3–6 months  
**Repository Location:** `C:\Temp_alt\CDK` (may vary by machine — see CDK_BASE below)  
**Active Branch:** `feature/ws3-function-migration`

---

## 1. Execution Environment

### Host Application
- **Bluezone 6.2** terminal emulator hosting CDK Drive XL DMS
- **DMS:** CDK Drive XL — Version OF045, IR228 Portal (IE 11 / Chromium 96)
- **Working directory:** `C:\Program Files (x86)\BlueZone\6.2` (read-only, write-restricted)
- **Script engine:** Bluezone hosts its own VBScript runtime — NOT Windows Script Host (WSH)

### Confirmed Runtime Capabilities

| Capability | Status | Notes |
|---|---|---|
| `CreateObject` / `FileSystemObject` | WORKS | Core file I/O is available |
| `Environ()` | WORKS | Reads Windows environment variables |
| `ExecuteGlobal` | WORKS | Core library loading mechanism is viable |
| `Scripting.Dictionary` | WORKS | Available for load-guard pattern |
| `WScript.Shell` via `CreateObject` | WORKS | Use for `ExpandEnvironmentStrings` |
| `WScript` object (native) | NOT AVAILABLE | WSH is not the host — do not reference |
| `WScript.Quit` | NOT AVAILABLE | Use `Host_Quit` from HostCompat.vbs instead |
| Relative paths | UNRELIABLE | Working dir is BlueZone install folder — always use absolute paths |
| `MsgBox` | WORKS | Use for diagnostics instead of `WScript.Echo` |

### Writable Paths (Confirmed)

| Path | Status |
|---|---|
| `C:\Users\{user}\Documents` | Writable |
| `C:\Users\{user}\AppData\Local` | Writable |
| `C:\VBSLibrary` (or equivalent) | Writable |
| `C:\Program Files (x86)\BlueZone\6.2` | READ-ONLY — never write here |

### Environment Variable
- `CDK_BASE` — user-level environment variable pointing to the repository root
- Set via `setx CDK_BASE C:\Temp_alt\CDK`
- **Important:** `setx` does not update the current session — Bluezone must be restarted after setting
- Accessed in scripts via: `CreateObject("WScript.Shell").Environment("USER")("CDK_BASE")`

---

## 2. Repository Structure

```
CDK/
├── .cdkroot                          # Marker file — do not delete
├── config/
│   └── config.ini                    # Centralised path configuration
├── framework/
│   ├── PathHelper.vbs                # Shared path library (EXISTS)
│   ├── BZHelper.vbs                  # Bluezone terminal library (EXISTS — new)
│   └── HostCompat.vbs                # WScript compatibility shims
├── apps/
│   ├── repair_order/                 # RO workflow — initialize, prepare, finalize
│   ├── post_final_charges/           # State-machine RO closeout (30+ conditional prompts)
│   ├── maintenance_ro_closer/        # Automated PM processing
│   ├── pfc_scrapper/                 # PFC data scraping
│   ├── prescreened_ro_closer/        # Prescreened RO processing (new)
│   └── validate_ro_list/             # RO list validation
├── tools/                            # Setup, testing, validation scripts
├── tests/                            # Infrastructure and unit tests
└── docs/                             # Documentation
```

---

## 3. Existing Architecture

### PathHelper.vbs (framework\PathHelper.vbs)
Shared path library. Provides:
- `GetRepoRoot()` — resolves repository root via `CDK_BASE` + `.cdkroot` validation
- `FindRepoRootForBootstrap()` — legacy alias for `GetRepoRoot()` (kept for compatibility)
- `GetConfigPath(section, key)` — builds absolute path from `config.ini`
- `ReadIniValue(filePath, section, key)` — INI file parser

### BZHelper.vbs (framework\BZHelper.vbs) — NEW
Authoritative shared terminal automation library. Load AFTER PathHelper. Provides:
- `ConnectBZ()` / `DisconnectBZ()` — connection management; returns True/False
- `BZReadScreen(length, row, col)` — wrapper around `g_bzhao.ReadScreen`
- `IsTextPresent(searchText)` — pipe-delimited multi-target, case-insensitive, row-by-row
- `BZSendKey(keyValue)` — keystroke send; returns True/False
- `WaitMs(milliseconds)` — busy-wait; midnight rollover safe
- `WaitForPrompt(promptText, inputValue, sendEnter, timeoutMs, description)` — canonical prompt wait; **timeout in milliseconds**
- `BZH_Log` — internal shim; calls `LogResult` if defined, no-ops otherwise

**Usage pattern:**
```vbscript
' g_bzhao MUST be declared before loading BZHelper
Dim g_bzhao: Set g_bzhao = CreateObject("BZWhll.WhllObj")
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\BZHelper.vbs")).ReadAll
```

**Load guard:** safe to ExecuteGlobal multiple times.

### Standard Bootstrap Pattern (CURRENT — all migrated scripts)
```vbscript
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_sh: Set g_sh = CreateObject("WScript.Shell")
Dim g_root: g_root = g_sh.Environment("USER")("CDK_BASE")
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\PathHelper.vbs")).ReadAll
```

Scripts that also need BZHelper add two more lines after:
```vbscript
Dim g_bzhao: Set g_bzhao = CreateObject("BZWhll.WhllObj")
ExecuteGlobal g_fso.OpenTextFile(g_fso.BuildPath(g_root, "framework\BZHelper.vbs")).ReadAll
```

---

## 4. Duplication Audit — Current State

### Bootstrap (`FindRepoRootForBootstrap`)
| Script | Migrated? |
|---|---|
| `apps/maintenance_ro_closer/Maintenance_RO_Closer.vbs` | ✅ Done |
| `apps/pfc_scrapper/PFC_Scrapper.vbs` | ✅ Done |
| `apps/post_final_charges/Pfc_Summary.vbs` | ✅ Done |
| `apps/post_final_charges/PostFinalCharges.vbs` | ✅ Done |
| `apps/prescreened_ro_closer/Prescreened_RO_Closer.vbs` | ✅ Done |
| `apps/repair_order/0_Open_RO/Open_RO.vbs` | ✅ Done |
| `apps/repair_order/1_prepare_close_pt1/1_Prepare_Close_Pt1.vbs` | ✅ Done |
| `apps/repair_order/prepare_close_pt1/2_Prepare_Close_Pt1.vbs` | ✅ Done |
| `apps/repair_order/2_finalize_close_pt2/2_Finalize_Close_Pt2.vbs` | ✅ Done |
| `apps/validate_ro_list/ValidateRoList.vbs` | ✅ Done |
| `tools/mva_scrapper/get_mva_from_vin.vbs` | ✅ Done |
| `tools/simpletest.vbs` | ✅ Done (g_fso/g_sh/g_root only — no PathHelper load needed) |

### WaitForPrompt / IsTextPresent / WaitMs (local definitions)
| Script | Status |
|---|---|
| `apps/pfc_scrapper/PFC_Scrapper.vbs` | ✅ Removed — now uses BZHelper |
| `apps/post_final_charges/Pfc_Summary.vbs` | ✅ Removed — now uses BZHelper |
| `apps/repair_order/0_Open_RO/Open_RO.vbs` | ✅ Removed — migrated to BZHelper 5-arg signature |
| `apps/repair_order/2_finalize_close_pt2/2_Finalize_Close_Pt2.vbs` | ✅ Removed (was dead code) |
| `apps/post_final_charges/PostFinalCharges.vbs` | ⏳ **Pending — next session** |
| `apps/post_final_charges/tests/test_prompt_detection.vbs` | ⏳ Test file — defer |

### WaitForAnyOf
| Script | Status |
|---|---|
| `apps/pfc_scrapper/PFC_Scrapper.vbs` | ✅ Removed — now calls BZHelper; timeout updated to ms |

### Not yet assessed for full BZHelper migration
```
apps/post_final_charges/close_single_ro.vbs
apps/post_final_charges/tests/test_blacklist_runtime_detection_gap.vbs
apps/post_final_charges/tests/test_wch_skip_counter_runtime_regression.vbs
tools/close_single_ro.vbs
```

---

## 5. Known Issues

### High Priority
1. **`PostFinalCharges.vbs` WS3 migration pending** — largest and most complex script. Has local `WaitForPrompt`, `IsTextPresent`, `WaitMs`. Defer to next session with fresh context.

### Medium Priority
2. **`BuildLogPath` and `BuildCSVPath` in PathHelper.vbs** — pure aliases for `GetConfigPath`. Candidates for removal.

3. **`LOG` sub in `Open_RO.vbs` creates new FSO on every call** — logging inside CSV loop creates/destroys COM objects each iteration. Replace with `g_fso`.

4. **`Open_RO.vbs` WaitForPrompt 5th arg is `""`** — VBScript has no named parameters; empty string description is the idiomatic solution. Consider documenting the parameter contract in BZHelper comments.

### Low Priority
5. **`.cdkroot` marker** — retain given sunset timeline.

---

## 6. Workstreams

### Workstream 1 — BZHelper.vbs ✅ COMPLETE
`framework\BZHelper.vbs` created with canonical versions of:
`ConnectBZ`, `DisconnectBZ`, `BZReadScreen`, `IsTextPresent`, `BZSendKey`, `WaitMs`, `WaitForPrompt`, `BZH_Log`

**Drift analysis completed** — 4 different `WaitForPrompt` signatures found across the codebase. Canonical uses milliseconds (not seconds). Scripts previously using seconds-based calls updated at migration time.

### Workstream 2 — Bootstrap Standardisation ✅ COMPLETE
All 12 production scripts migrated to `g_fso`/`g_sh`/`g_root` pattern.
`FindRepoRootForBootstrap` now appears only in PathHelper (canonical), test_hardcoded_paths, and simpletest.

**Validation tool:** `tools\simpletest.vbs` — now cscript-friendly (`WScript.Echo`, no dialogs). Terms ordered by WS3 migration priority. Run via `cscript.exe //nologo tools\simpletest.vbs`.

**Full test suite:** `tests\run_all.vbs` — run via `cscript.exe //nologo tests\run_all.vbs`. Last result: **98 pass / 0 fail / 0 error**.

### Workstream 3 — Function Migration ⏳ IN PROGRESS
**Goal:** Replace local function bodies in production scripts with calls to BZHelper.

**Completed:**
- `WaitForAnyOf` lifted from `PFC_Scrapper.vbs` into BZHelper; timeout updated from seconds to ms
- `IsTextPresent` removed from `2_Finalize_Close_Pt2.vbs` (was dead code)
- `Open_RO.vbs` — local `WaitForPrompt` (Sub, 4-arg), `IsTextPresent`, `WaitMs` all removed; 16 call sites updated to BZHelper 5-arg signature; BZHelper now loaded in bootstrap

**Remaining:**
- `PostFinalCharges.vbs` — local `WaitForPrompt`, `IsTextPresent`, `WaitMs` — **next session, highest risk**
- Test files (`test_prompt_detection.vbs`, `test_blacklist_runtime_detection_gap.vbs`, `test_wch_skip_counter_runtime_regression.vbs`) — deferred

---

## 7. Design Decisions & Rationale

| Decision | Rationale |
|---|---|
| `CDK_BASE` environment variable for discovery | Only portable machine-level discovery mechanism without hardcoded paths |
| `.cdkroot` marker file as secondary validation | Prevents silent failures when `CDK_BASE` points to wrong directory |
| `ExecuteGlobal` for library loading | Only available mechanism for injecting definitions into global scope in VBScript |
| `g_fso.BuildPath` for all path construction | Handles trailing slash normalisation. Never use string concatenation for paths |
| Absolute paths only | Working directory is BlueZone install folder — relative paths resolve incorrectly |
| `WScript.Shell` via `CreateObject` | `WScript` native object unavailable in Bluezone |
| `g_bzhao` declared by calling script | Each script owns its connection — allows concurrent independent scripts |
| `WaitForPrompt` timeout in milliseconds | Canonical unit. Scripts previously using seconds updated at migration time |
| BZHelper load guard | `If Not IsObject(g_BZHelper_Loaded)` prevents double-execution via ExecuteGlobal |

---

## 8. How to Use This Document

**Starting a new chat:** Paste this entire document at the start of the conversation.

**Next session priorities:**
1. Migrate `PostFinalCharges.vbs` WS3 — remove local `WaitForPrompt`, `IsTextPresent`, `WaitMs`; load BZHelper in bootstrap
2. Run simpletest — expect `IsTextPresent` and `WaitForPrompt` in PostFinalCharges + test files only
3. Run full test suite — expect 98+ pass
4. Consider test file migrations (lower priority given sunset timeline)

**Section quick-reference:**
- New to the project? Read sections 1, 2, 3
- Working on BZHelper / WS3? Read sections 3, 4, 6
- Working on bootstrap / WS2? Read sections 3, 4, 6
- Debugging environment issues? Read section 1
