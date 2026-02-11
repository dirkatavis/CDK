# GitHub Copilot Instructions - CDK DMS Automation

## Project Overview
This codebase automates interactions with the CDK Dealership Management System (DMS) through the BlueZone terminal emulator using VBScript and PowerShell.

## Strategic Context
- **Legacy Sunset**: This system is legacy and scheduled for retirement in 3-6 months. **Prioritize simplicity and immediate utility** over long-term maintainability or complex abstractions. "Good enough" is the target.
- **Sandwich Automation**: The goal is to automate the manageable "bookends" of a workflow that contains a non-automatable manual middle step. 
    - `Pt1` scripts: Pre-manual processing (seeding data, initial setup).
    - `Pt2` scripts: Post-manual processing (finalizing, closing, printing).

## Big Picture Architecture
- **Terminal Automation**: Most scripts use the `BZWhll.WhllObj` COM object to interact with BlueZone.
- **Workflow Segregation**: Logic is separated by task (e.g., `CreateNew_ROs`, `Close_ROs`).
- **Script Patterns**:
    - **Procedural (Preferred)**: Simple top-down logic using `WaitForTextAtBottom` and `EnterTextAndWait`. Best for quick, high-value fixes.
    - **State Machine**: Found in `PostFinalCharges.vbs`. Use `AddPromptToDictEx` to handle complex/circular screens.
    - **Dynamic Discovery**: Use `DiscoverLineLetters()` pattern to scan the "LC" column (Column 1, Rows 7+) for active line letters (A-Z) before processing. This prevents infinite loops on non-consecutive lines.

## Technical Patterns & Conventions
- **Language**: Primary language is VBScript (`.vbs`). Use `Option Explicit`.
- **Terminal Interaction**:
    - Always wait for a specific prompt text (e.g., `COMMAND:`) before sending input.
    - Use `bzhao.ReadScreen` or `IsTextPresent(text)` to verify state.
    - Common screen row for prompts is 23 (`MainPromptLine = 23`).
- **Prompt Handling (State Machine)**:
    - `AddPromptToDict(dict, trigger, response, key, isSuccess)`: Always sends `response`.
    - `AddPromptToDictEx(dict, trigger, response, key, isSuccess, acceptDefault)`: If `acceptDefault=True`, it detects values in parentheses (e.g., `(72925)`) and sends ONLY the `key` (Enter) to accept it.
- **Prompt Detection (CRITICAL)**:
    - **Intervening Text**: Support patterns like `"OPERATION CODE FOR LINE A, L1 (I)?"` where descriptive text exists between the keyword and default value.
    - **Robust Regex**: Use `".*(\(.*\))?\?"` to capture optional default values safely. Example: `OPERATION CODE FOR LINE.*(\([A-Za-z0-9]*\))?\?`.
- **Common Terminal States**:
    - **Success/Entry**: `COMMAND:`, `R.O. NUMBER`, `SEQUENCE NUMBER`.
    - **Errors**: `NOT ON FILE`, `is closed`, `ALREADY CLOSED`, `VARIABLE HAS NOT BEEN ASSIGNED`.
- **Hardcoded Paths**: Scripts often use `C:\Temp\Code\Scripts\VBScript\CDK\...`. Ensure paths are consistent with the current environment (`c:\Temp_alt\CDK`). `PostFinalCharges.vbs` may use a `config.ini` for flexibility.
- **Logging**: 
    - Simple scripts use `LogResult(type, message)`.
    - `PostFinalCharges.vbs` uses `LogEvent` with `g_CurrentCriticality` (CRIT_COMMON=0 to CRIT_CRITICAL=3) and `g_CurrentVerbosity` (VERB_LOW=0 to VERB_MAX=3).
- **Classes in VBScript**: Used in advanced scripts for data structures (e.g., `Class Prompt`).

## Critical Domain Terms
- **RO**: Repair Order.
- **Story**: A segment of a repair (e.g., Story A, B, C).
- **CCC**: Command sequence to manage repair order lines.
- **FC**: Final Charge.
- **MVA**: Motor Vehicle Account/Asset (used in vehicle identification).

## Developer Workflows
- **Branching Policy**: **NEVER merge into `main` automatically.** All changes must be performed in a feature or bugfix branch.
- **Execution**: Run scripts using `cscript.exe` for console output.
  ```cmd
  cscript.exe Close_ROs_Pt2.vbs
  ```
- **Testing**: 
    - Use `PostFinalCharges/tests` for validating logic changes.
    - Run all tests: `cscript run_all_tests.vbs` from the `tests` directory.
    - Use `MockBzhao` to simulate terminal states without a live BlueZone connection.
    - **Regression Testing**: If fixing a prompt detection bug, add the specific screen pattern to `test_default_value_detection.vbs`.
- **Logging/Debugging**: Check the generated `.log` files in the respective folders. Some scripts support a `.debug` file flag to enable "slow mode".

## Integration Points
- **Input**: CSV files containing lists of RO numbers or vehicle data.
- **Output**: Logs and terminal state changes in BlueZone.
- **PowerShell**: Used for utility tasks like log parsing (`Parse_Data.ps1`).

## Code Review Instructions
The underlying application for this project will be deprecated soon. To minimize noise and focus only on essential changes, follow these rules:

- **What to IGNORE**:
    - **Refrain from raising NIT issues**: Do not comment on naming conventions, styling, or minor readability improvements.
    - **Ignore Security issues**: Standard security hardening is not required for this legacy app.
    - **Ignore Performance issues**: Optimization is not a priority.
- **What to FLAG**:
    - **Critical Bugs**: Only report issues that will cause immediate system failure or data corruption.
    - **Important Logic Errors**: Flag issues where the code does not perform the intended business logic.
    - **Breaking Changes**: Flag anything that breaks existing integrations or deployment pipelines.
- **Review Tone**: Be direct and concise. If an issue isn't Critical or Important, stay silent.
