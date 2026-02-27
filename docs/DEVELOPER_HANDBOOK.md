# CDK Automation Developer Handbook

This handbook provides the technical architecture and standards for the CDK DMS Automation system. Given the systems legacy status (sunset in 3-6 months), this document focuses on core mechanics and immediate utility.

## 1. System Architecture: "Sandwich Automation"

The project automates the predictable "bookends" of a workflow that contains a manual middle step:
- **Pt 1 Scripts**: Pre-manual processing (seeding data, discovery).
- **Pt 2 Scripts**: Post-manual processing (finalizing, closing).

### Key Patterns
- **Procedural Logic**: Simple top-down logic for most scripts.
- **State Machine**: Used in `PostFinalCharges.vbs` to handle complex/circular screens.
- **Dynamic Discovery**: Scans the "LC" column for active line letters before processing to prevent infinite loops.

## 2. Path Configuration & Portability

The system is designed to be fully portable ("sneakernet" friendly).

### Centralized `config/config.ini`
All scripts use the `framework/PathHelper.vbs` library to resolve relative paths defined in `config/config.ini`.
- **NEVER** hardcode absolute paths.
- **NEVER** use silent fallbacks; fail fast with clear errors.

### Dependency Resolution
1. **`CDK_BASE`**: A user environment variable pointing to the repo root.
2. **`.cdkroot`**: A marker file in the root to verify the base path.
3. **`PathHelper.vbs`**: The engine that reads the INI and builds absolute paths.

## 3. Validation System

A three-layer approach ensuring the environment is "ready to run":
1. **`tools/validate_dependencies.vbs`**: Manual pre-flight check.
2. **`framework/ValidateSetup.vbs`**: Shared library for logic-based validation.
3. **Internal Startup Checks**: Every production script calls `MustHaveValidDependencies` on launch.

## 4. BlueZone Context Compatibility

Scripts are designed to work in two distinct contexts:

### Standalone Context (cscript.exe)
- Full `WScript` object availability.
- Used for setup, validation, and diagnostic tools.
- Uses `WScript.Echo` for console output.

### BlueZone Context (Embedded)
- No `WScript` object.
- Used for production automation (e.g., `PostFinalCharges.vbs`).
- Uses `LogInfo` and `g_ShouldAbort` flags for error handling instead of `WScript.Quit`.

### Safe Execution Library
`framework/ValidateSetup.vbs` provides the `MustHaveValidDependencies` routine, which detects the current context and handles failures safely (either quitting the console or setting an abort flag for the terminal).

## 5. Development Standards

- **Language**: VBScript (`.vbs`) with `Option Explicit`.
- **Logging**: Use `LogResult` or `LogEvent` for consistent audit trails in `runtime/logs/`.
- **Environment**: Use `tools/setup_cdk_base.vbs` to initialize a new machine.

---
*Last Updated: February 27, 2026*
