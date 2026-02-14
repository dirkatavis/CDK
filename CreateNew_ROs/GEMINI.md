# GEMINI.md

## Project Overview

This project consists of a VBScript (`Create_ROs.vbs`) designed to automate the creation of Repair Orders (ROs) within the CDK system. The script utilizes the BlueZone terminal emulator to interact with the CDK application. It reads vehicle data (MVA and Mileage) from a CSV file (`Create_RO.csv`) and programmatically enters it into the terminal interface to create new ROs.

The script is designed to be robust, with features like:
- **Prompt Detection:** It waits for specific text prompts from the terminal before entering data, ensuring it stays in sync with the application.
- **Logging:** It maintains a log file (`VehicleData.log`) to record its operations and any potential errors.
- **Debugging Mode:** A "slow mode" can be enabled by creating a `Create_RO.debug` file, which introduces delays to make it easier to follow the script's execution.

## Running the Project

### Dependencies

- **BlueZone:** The script requires the BlueZone terminal emulator to be installed and accessible. It specifically uses the `BZWhll.WhllObj` COM object.
- **CSV Data:** A CSV file named `Create_RO.csv` must be present. The path is configured in `config.ini` under `[CreateNew_ROs] CSV=CreateNew_ROs\create_RO.csv`. See `docs/PATH_CONFIGURATION.md` for details on the centralized path system.

### Execution

To run the script, execute the `.vbs` file from the command line using `cscript`:

```shell
cscript Create_ROs.vbs
```

Alternatively, you can run it by double-clicking the file, which will use `wscript`.

## Development Conventions

- **`Option Explicit`:** The script uses `Option Explicit` to enforce variable declaration, which helps prevent typos and other common errors.
- **Configuration:** Key settings like file paths and wait times are defined as constants at the beginning of the script.
- **Structured Code:** The main logic is encapsulated in a `Main` subroutine, and helper functions are used for tasks like sending text, waiting for prompts, and logging.
- **Error Handling:** The script uses `On Error Resume Next` in a limited scope when creating the BlueZone object, but otherwise relies on prompt timeouts for error conditions.
- **Logging:** A custom `LOG` subroutine is used to write detailed information about the script's execution to the log file.
