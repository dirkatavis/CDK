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
    - **State Machine**: Found in `PostFinalCharges.vbs`. Use only if the screen logic is highly circular or unpredictable.

## Technical Patterns & Conventions
- **Language**: Primary language is VBScript (`.vbs`). Use `Option Explicit`.
- **Terminal Interaction**:
    - Always wait for a specific prompt text (e.g., `COMMAND:`) before sending input.
    - Use `bzhao.ReadScreen` to verify the state.
    - Common screen row for prompts is 23 (`MainPromptLine = 23`).
- **Hardcoded Paths**: Scripts often use `C:\Temp\Code\Scripts\VBScript\CDK\...`. Ensure paths are consistent with the current environment (`c:\Temp_alt\CDK`).
- **Logging**: 
    - Simple scripts use `LogResult(type, message)`.
    - `PostFinalCharges.vbs` uses a prioritized system: `LOG_LEVEL_ERROR` (1) to `LOG_LEVEL_TRACE` (5).
- **Classes in VBScript**: Used in advanced scripts for data structures (e.g., `Class Prompt`).

## Critical Domain Terms
- **RO**: Repair Order.
- **Story**: A segment of a repair (e.g., Story A, B, C).
- **CCC**: Command sequence to manage repair order lines.
- **FC**: Final Charge.
- **MVA**: Motor Vehicle Account/Asset (used in vehicle identification).

## Developer Workflows
- **Execution**: Run scripts using `cscript.exe` for console output.
  ```cmd
  cscript.exe Close_ROs_Pt2.vbs
  ```
- **Logging/Debugging**: Check the generated `.log` files in the respective folders. Some scripts support a `.debug` file flag to enable "slow mode".

## Integration Points
- **Input**: CSV files containing lists of RO numbers or vehicle data.
- **Output**: Logs and terminal state changes in BlueZone.
- **PowerShell**: Used for utility tasks like log parsing (`Parse_Data.ps1`).
