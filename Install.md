# CDK Automation Setup Guide

This guide will help you set up and verify your CDK automation environment. Please follow the steps below in sequence.

## Phase 1: Preparation (Dev Machine Only)
If you are preparing this package for others:
1.  **Zip the Folder**: Compress the entire `CDK` folder into a ZIP file.
2.  **Distribute**: Send the ZIP file to the target machine.

---

## Phase 2: Installation (Target Machine)

### Step 1: Extract the Files
Extract the ZIP package to a permanent location on your computer (for example: `C:\CDK`).
- **EXPECTED RESULT**: You should see the `CDK` folder and its contents in your file explorer.

### Step 2: Run the Setup Utility
1.  Locate the `Install.vbs` file in the main folder.
2.  Double-click `Install.vbs` (or run `cscript.exe Install.vbs` from a terminal).
- **EXPECTED RESULT 1**: A popup will appear confirming your environment path is set. Click **OK**.
- **EXPECTED RESULT 2**: A black terminal window will show a series of checks (Initialization, Validation, and Path Test).
- **EXPECTED RESULT 3**: A final popup will appear stating: **"All deployment steps passed successfully!"**

### Step 3: Finalize
Once you see the success message, **close and reopen** your terminal or BlueZone session.
- **EXPECTED RESULT**: The automation environment is now active and ready for use.

---

## Troubleshooting
If any step displays a **CRITICAL ERROR** or a **FAILURE** message:
- Ensure the folder was fully extracted before running the script.
- Ensure you have permissions to set environment variables on your machine.
- Contact the system administrator with a screenshot of the error message.
