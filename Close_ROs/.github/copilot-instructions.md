# Copilot Instructions for Close_ROs VBScript Automation

## Project Overview
This workspace automates closing Repair Orders (ROs) in a legacy system using VBScript and BlueZone terminal emulation. Scripts interact with a mainframe via BlueZone Host Access, reading RO numbers from CSV files and executing command sequences.

## Key Files & Structure
- `Close_ROs_Pt1.vbs`: Main automation script. Reads RO numbers from `Close_ROs_Pt1.csv`, sends commands to BlueZone, handles errors, and logs results.
- `Close_ROs_Pt1.csv`: Input file with one RO number per line.
- `Close_ROs_Pt2.vbs`, `Create_ROs.vbs`, `HighestRoFinder.vbs`, `PostFinalCharges.vbs`, `TestLog.vbs`: Additional scripts for related automation tasks.
- `msgboxtext.txt`: Likely contains message box text templates.
- Log files are written to the same directory (e.g., `Close_ROs_Pt1.log`).

## Automation Flow
- **Input**: RO numbers are read line-by-line from CSV.
- **Processing**: For each valid 6-digit RO, the script:
  - Sends RO to BlueZone
  - Checks for "NOT ON FILE" error (via `CheckForROError`)
  - Executes CCC commands (A, B, C) with pauses and waits for story closure prompts
  - Logs results/errors using `LogResult`
- **Error Handling**: If "NOT ON FILE" is detected, logs and skips to next RO.
- **Screen Monitoring**: Uses `ReadScreen` to poll for specific text, e.g., story closure prompts.

## Patterns & Conventions
- **VBScript**: All scripts use classic VBScript with explicit variable declarations (`Option Explicit`).
- **BlueZone Automation**: Interactions are via `BZWhll.WhllObj` methods (`SendKey`, `Pause`, `ReadScreen`).
- **Logging**: Results/errors are appended to log files with timestamps.
- **Error Handling**: Critical errors trigger `MsgBox` and `bzhao.StopScript` to halt execution.
- **Subroutines**: Key logic is modularized into subs/functions (e.g., `ProcessRo`, `CheckForROError`, `WaitForStoryClosure`).
- **No Unit Tests**: No test framework or test scripts detected; debugging is via message boxes and log files.

## Developer Workflows
- **Run Scripts**: Execute `.vbs` files directly in Windows (double-click or via command line).
- **Debugging**: Use commented-out `bzhao.msgBox` lines for step-by-step inspection. Enable as needed.
- **Modify Input**: Update CSV files for different RO batches.
- **Logs**: Check `.log` files for results and troubleshooting.

## Integration Points
- **BlueZone**: Requires BlueZone Host Access installed and accessible via COM (`BZWhll.WhllObj`).
- **File System**: Reads/writes files in the script directory. Paths are hardcoded; update as needed for new environments.

## Example: Error Handling Pattern
```vbs
foundError = CheckForROError()
If foundError Then
    LogResult RoNumber, "RO NOT ON FILE - Skipping to next."
    Exit Sub
End If
```

## Recommendations for AI Agents
- Always use explicit variable declarations and modularize logic into subs/functions.
- Maintain hardcoded paths unless refactoring for portability.
- Follow the error handling/logging pattern for new scripts.
- Reference `Close_ROs_Pt1.vbs` for main automation flow and conventions.

---
For questions or unclear conventions, ask the user for clarification or examples from related scripts.
