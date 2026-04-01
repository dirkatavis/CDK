# Technical Procedure: RO Booking & Processing

**Process Overview:** This workflow outlines the steps for initiating an MVA request, assigning a Technician ID, and finalizing the Repair Order (RO).

---

### Phase 1: MVA Initiation
1. **Initialize MVA:** Type `mva` and press `Enter`.
2. **Selection Menu:** If the prompt *"Choose one of the following"* appears, input `1` and press `Enter`.
3. **Confirm Selection:** Press `Enter` to confirm.
4. **Mileage Entry:** Input the current **Miles** and press `Enter`.

### Phase 2: Transaction Coding
| Action | Input / Command | Key Stroke |
| :--- | :--- | :--- |
| Re-initialize MVA | `MVA` | `Enter` |
| Set Quick Code | `PMVEND` | `Enter` |
| Accept Confirmation | — | `F3` |

### Phase 3: Technician Assignment & Saving
* **Access Tech ID Field:** Press `F8`.
* **Enter Technician ID:** Input `99`.
* **Exit Field:** Press `F3` to escape the entry screen.
* **Commit Changes:** Press `F3` to **Save the RO**.

### Phase 4: Finalization & Data Capture
1. **Book RO:** Press `Enter`.
2. **Mileage Verification:** * Press `Enter` to accept **MILEAGE OUT**.
    * Press `Enter` to accept **MILEAGE IN**.
3. **Status Check:** When prompted to keep the RO open, press `N`.
4. **Data Extraction:** The system will automatically **Scrape the RO number** from Line 23.
5. **Reset:** Press `F3` to return to the main entry screen.

---
> **Note:** Ensure the **Tech ID (99)** is entered correctly before saving.