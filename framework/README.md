# Framework - Shared Reusable Components

## Purpose
Central location for shared, reusable building blocks used across all CDK automation scripts. These components provide core functionality that multiple apps depend on.

## Components

### PathHelper.vbs
**Purpose:** Centralized path resolution for all scripts
- Reads `CDK_BASE` environment variable (repo root)
- Validates `.cdkroot` marker file
- Resolves section/key pairs from `config/config.ini` to absolute paths
- Fails fast with clear errors instead of silent fallbacks

**Key Functions:**
- `GetRepoRoot()` - Returns absolute path to repo root from `CDK_BASE`
- `GetConfigPath(section, key)` - Resolves config.ini entries to absolute paths
- `ValidateRepoRoot(path)` - Ensures `.cdkroot` marker exists

**Usage Pattern:**
```vbs
' Load at top of every script
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(repoRoot & "\framework\PathHelper.vbs").ReadAll()

' Resolve paths from config.ini
csvPath = GetConfigPath("MyApp", "CSV")
logPath = GetConfigPath("MyApp", "Log")
```

### ValidateSetup.vbs
**Purpose:** Environment validation and dependency checking
- Validates BlueZone terminal availability
- Checks CDK_BASE environment variable
- Verifies config.ini structure
- Ensures runtime folders exist

**Key Functions:**
- `ValidateEnvironment()` - Comprehensive pre-flight checks
- `CheckBlueZoneConnection()` - Verifies terminal is ready
- `ValidateConfigFile()` - Ensures config.ini is valid

**Usage Pattern:**
```vbs
' Load after PathHelper
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(repoRoot & "\framework\ValidateSetup.vbs").ReadAll()

' Run validation before script execution
If Not ValidateEnvironment() Then
    WScript.Echo "Setup validation failed. Run tools\validate_dependencies.vbs"
    WScript.Quit 1
End If
```

### BZHelper.vbs
**Purpose:** Authoritative shared library for BlueZone terminal automation
- Centralises all terminal interaction: connection, screen reading, text detection, keystroke sending, prompt waiting
- Eliminates duplication of `WaitForPrompt`, `IsTextPresent`, `WaitMs`, `FastKey`/`FastText` across 13+ scripts
- `BZH_Log` shim bridges to the calling script's log system (`LogResult`) without coupling

**Key Functions:**
- `ConnectBZ()` — Connect `g_bzhao` to the active BlueZone session; returns True/False
- `DisconnectBZ()` — Cleanly disconnect and release `g_bzhao`
- `BZReadScreen(length, row, col)` — Wrapper around `g_bzhao.ReadScreen`
- `IsTextPresent(searchText)` — Pipe-delimited multi-target screen search; case-insensitive
- `BZSendKey(keyValue)` — Send keystroke or text to terminal; returns True/False
- `WaitMs(milliseconds)` — Busy-wait; midnight rollover safe
- `WaitForPrompt(promptText, inputValue, sendEnter, timeoutMs, description)` — Canonical prompt wait with optional input send; timeout in **milliseconds**

**Usage Pattern:**
```vbs
' 1. Declare and instantiate g_bzhao BEFORE loading BZHelper
Dim g_bzhao: Set g_bzhao = CreateObject("BZWhll.WhllObj")

' 2. Load BZHelper (after PathHelper)
Dim bzHelperPath: bzHelperPath = g_fso.BuildPath(g_root, "framework\BZHelper.vbs")
ExecuteGlobal g_fso.OpenTextFile(bzHelperPath).ReadAll

' 3. Connect and use
If ConnectBZ() Then
    If WaitForPrompt("R.O. NUMBER", roNumber, True, 30000, "Main prompt") Then
        ' prompt found, input sent, Enter pressed
    End If
    DisconnectBZ
End If
```

**Migration note:** Scripts previously calling `WaitForPrompt(text, timeoutSec)` (PFC_Scrapper, Pfc_Summary)
must update call sites — the canonical signature uses **milliseconds**, not seconds.

### HostCompat.vbs
**Purpose:** Dual-context execution compatibility (standalone vs BlueZone)
- Enables scripts to run in both BlueZone and standalone cscript.exe environments
- Provides host detection and compatibility layer

**Usage Pattern:**
```vbs
' Load for scripts that need host detection
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(repoRoot & "\framework\HostCompat.vbs").ReadAll()

' Check execution context
If IsBlueZoneHost() Then
    ' BlueZone-specific logic
Else
    ' Standalone execution logic
End If
```

### AdvancedMock.vbs
**Purpose:** High-fidelity terminal simulation for offline testing
- Simulates BlueZone COM objects (`BZWhll.WhllObj`)
- Provides latency, partial load, and auto-responder features
- Enables stress testing and race condition detection

**Key Features:**
- `SetLatency(ms)` - Simulates network/terminal delay
- `SetPartialLoad(bool)` - Simulates asynchronous screen rendering
- `SetPromptSequence(array)` - Defences a stateful conversation flow

**Usage Pattern:**
```vbs
' Load for offline test scripts
ExecuteGlobal fso.OpenTextFile(repoRoot & "\framework\AdvancedMock.vbs").ReadAll()

Dim bz: Set bz = New AdvancedMock
bz.SetLatency 1000
bz.SetPromptSequence Array(Array("COMMAND:", "S"), Array("R.O.", "1234"))
```

## Design Principles
- **Zero Hardcoded Paths:** All paths resolved via PathHelper from config.ini
- **Fail Fast:** Clear error messages instead of silent fallbacks
- **Single Responsibility:** Each component has one clear purpose
- **No App Logic:** Framework contains ONLY shared building blocks, never app-specific workflows

## Dependencies
- `config/config.ini` - Configuration file (PathHelper dependency)
- `CDK_BASE` environment variable - Set by `tools/setup_cdk_base.vbs`
- `.cdkroot` marker file at repo root - Validated by PathHelper

## Adding New Framework Components
1. Component must be truly shared (used by 2+ apps)
2. Must have single, well-defined responsibility
3. Must not contain app-specific business logic
4. Add comprehensive inline documentation
5. Update this README with usage patterns

## Testing
Framework components are tested via:
- App-level integration tests (all apps use these components)
- Repo-level global tests in `tests/` folder
- `tools/run_validation_tests.vbs` validates setup and paths

## Notes
- Keep framework minimal - prefer app-local helpers over framework bloat
- Breaking changes here affect ALL apps - test thoroughly
- See `docs/PATH_CONFIGURATION.md` for PathHelper architecture details
- See `docs/VALIDATION_ARCHITECTURE.md` for ValidateSetup design
