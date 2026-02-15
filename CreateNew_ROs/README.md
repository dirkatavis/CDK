# CreateNew_ROs Automation

This module automates the creation of Repair Orders (ROs) within the CDK DMS using BlueZone terminal emulation.

## Purpose

The `Create_ROs.vbs` script reads vehicle data (MVA and Mileage) from a CSV file (`create_RO.csv`) and programmatically enters it into the terminal interface to create new ROs. This is particularly useful for batch processing new vehicles or transfers.

## Features

- **Prompt Detection:** Waits for specific terminal prompts to ensure synchronization.
- **Logging:** Maintains a log file (`VehicleData.log`) for auditing and error tracking.
- **Debugging Mode:** Supports a "slow mode" via a `.debug` file for visual debugging.

## Prerequisites

- **BlueZone Terminal Emulator:** Must be installed and configured.
- **CSV Data:** A file named `create_RO.csv` should exist in the script directory (or the configured path).

## Usage

1.  Ensure BlueZone is running and connected to CDK.
2.  Prepare your `create_RO.csv` with the required vehicle data.
3.  Run the script:

    ```cmd
    cscript Create_ROs.vbs
    ```
    
    Or simply double-click the `.vbs` file.

## Configuration

- **Constants:** File paths and wait times are defined as constants at the top of the script `Create_ROs.vbs`.
- **Debugging:** Create a file named `Create_RO.debug` in the script directory to enable slow-mode execution.
