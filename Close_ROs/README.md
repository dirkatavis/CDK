# Close_ROs Automation

This module automates the closing of Repair Orders (ROs) in CDK using the "Sandwich Pattern".

## The "Sandwich Pattern" Workflow

This automation is designed to handle a process that cannot be fully automated due to a manual middle step.

1.  **Phase I (The Top Bun):** `Close_ROs_Pt1.vbs`
    *   Reads a list of ROs.
    *   Performs initial setup and discovery.
    *   **Stops** to allow for manual review or processing that requires human judgment.
2.  **Manual Step (The Meat):**
    *   User performs necessary manual actions in the terminal.
3.  **Phase II (The Bottom Bun):** `Close_ROs_Pt2.vbs`
    *   Resumes automation to finalize the Close RO process (FC, Mileage, Printing).

## Scripts

*   **`Close_ROs_Pt1.vbs`**:
    *   **Purpose**: Initial processing and line letter discovery.
    *   **Input**: CSV list of ROs.
*   **`Close_ROs_Pt2.vbs`**:
    *   **Purpose**: Finalizing the RO closeout (CCC, R, FC, Mileage, Print).
    *   **Input**: Same CSV list (typically).

## Configuration

Ensure `config.ini` has the following sections:

```ini
[Close_ROs_Pt1]
CSV=Close_ROs\Close_ROs_Pt1.csv
Log=Close_ROs\Close_ROs_Pt1.log

[Close_ROs_Pt2]
CSV=Close_ROs\Close_ROs_Pt1.csv  ; Usually shares the same input
Log=Close_ROs\Close_ROs_Pt2.log
```

## Usage

1.  Populate `Close_ROs_Pt1.csv` with RO numbers.
2.  Run **Phase I**:
    ```cmd
    cscript Close_ROs_Pt1.vbs
    ```
3.  Perform manual work in BlueZone.
4.  Run **Phase II**:
    ```cmd
    cscript Close_ROs_Pt2.vbs
    ```
