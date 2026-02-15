# Maintenance RO Auto-Closer

This module automates the closing of specific Maintenance Repair Orders (ROs) that match an exact "footprint" of service lines.

## Purpose

To bulk-close Maintenance ROs that fit a standard profile (e.g., specific PM services) without human intervention, while skipping any ROs that deviate from the expected criteria.

## How It Works

the script validates an RO against **Match Criteria** defined in `PM_Match_Criteria.txt`.

1.  **Reads RO List**: Getting RO numbers from `RO_List.csv`.
2.  **Validates RO**: Checks if "Ready to Post", not "Closed", etc.
3.  **Matches Footprint**:
    *   Finds "Line A" as an anchor.
    *   Verifies Line A, B, and C match the expected text variants (fuzzy matching).
    *   Ensures "Line D" does **not** exist (exclusion check).
4.  **Auto-Close**: If all criteria pass, it performs the full closeout sequence (CCC -> FC -> Mileage -> Print).

## Configuration

### `config.ini`
(If applicable, otherwise paths are currently defined in script constants - *Note: This script should be updated to use `PathHelper` if it doesn't already*)

### `PM_Match_Criteria.txt`
Defines the expected service line descriptions. Format:
```text
A = LUBE OIL FILTER | LOF | SYNTHETIC
B = TIRE ROTATION | ROTATE
C = MULTI POINT INSPECTION | MPI
```

## Usage

1.  Update `RO_List.csv` with target ROs.
2.  Run the script:
    ```cmd
    cscript Maintenance_RO_Closer.vbs
    ```
