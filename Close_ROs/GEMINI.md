# GEMINI.md

## Project Overview

This project consists of a set of VBScripts designed to automate the process of closing Repair Orders (ROs) in a CDK dealership management system (DMS) accessed via a BlueZone terminal emulator. The scripts read RO numbers from a CSV file and then navigate the terminal interface, entering commands and data to close out the ROs, post final charges, and handle various prompts and error conditions.

The workflow appears to be split into multiple parts, handled by different scripts:

*   `Close_ROs_Pt1.vbs`: Performs the initial steps of the closing process.
*   `Close_ROs_Pt2.vbs`: Continues the closing process, handling final closeout steps.
*   `PostFinalCharges.vbs`: A more robust script for posting final charges and closing ROs, which seems to be a more recent or alternative version of the process.

## Building and Running

These scripts are intended to be run directly on a Windows machine with the BlueZone terminal emulator software installed and configured.

### Running the Scripts

To run a script, you can use `cscript.exe` or `wscript.exe` from the command line. `cscript.exe` is generally preferred for unattended execution as it directs output to the console.

**Example:**

```shell
cscript.exe Close_ROs_Pt1.vbs
```

### Dependencies

*   **BlueZone Terminal Emulator:** The scripts rely on the `BZWhll.WhllObj` COM object, which is part of the BlueZone software.
*   **Input Files:** The scripts require a CSV file containing the list of RO numbers to process. The path to this file is hardcoded in the scripts.
*   **Configuration:** `PostFinalCharges.vbs` uses a `config.ini` file to configure file paths and other settings.

## Development Conventions

*   **Error Handling:** The scripts have varying levels of error handling. The more recent `PostFinalCharges.vbs` has more robust error checking and logging.
*   **Logging:** The scripts generate log files to record their progress and any errors encountered.
*   **Modularity:** The logic is broken down into subroutines and functions, separating tasks like connecting to BlueZone, reading the CSV, and processing individual ROs.
*   **Screen Scraping:** The scripts interact with the terminal by sending keystrokes and then "scraping" the screen content to check for specific text, which indicates the state of the application.
