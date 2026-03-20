# Open_RO Expected Workflow

Purpose: provide a concise reference for the normal Open_RO flow so high-verbosity logs can be compared against the actual CDK behavior.

## Core Path

This is the normal expected sequence for Open_RO.

| Step | Expected user/action flow | Expected result | Status | Notes |
|---|---|---|---|---|
| 1 | Enter MVA + Enter | Vehicle is opened for processing | Confirmed | Current script prompt anchor: `Vehid.....` |
| 2 | Send Enter | Continue into RO setup flow | Confirmed | This is the normal path after MVA entry |
| 2a | Optional vehicle selection branch: choose `1` + Enter | Continue into RO setup flow | Confirmed | This happens less frequently; if vehicle selection appears, always choose `1` |
| 3 | At `Miles In`, enter mileage from CSV + Enter | Continue through RO creation flow | Confirmed | This is the next major expected data entry point |

## What Should Not Be Assumed Without Verification

These prompt expectations are currently suspect and should not be treated as trusted workflow anchors until reconfirmed against CDK.

| Prompt text | Current assessment | Reason |
|---|---|---|
| `Display them now` | Suspect | You do not recognize this prompt in the actual Open_RO flow |
| `greater than` | Suspect | May be conditional or may not be the exact prompt text |

## Logging Comparison Notes

When reviewing a High-verbosity log:

1. Look for the major flow checkpoints, not every low-level trace line.
2. The important sequence is:
   - MVA entered
   - optional vehicle selection handled if shown
   - Enter sent to continue
   - `Miles In` reached
3. If the log shows a prompt expectation between those major steps that you do not recognize in CDK, treat that expectation as suspect.
4. A timeout on a suspect expectation is more likely a script bug than an application issue.

## Current Review Direction

Based on current review:

1. `Display them now` is likely an incorrect expectation.
2. The vehicle-selection behavior should remain supported because it existed before the logging refactor.
3. The logging refactor should not have changed the core workflow described above.

## Validation Method

1. Run with `Verbosity=High`.
2. Compare the log to the three core workflow steps above.
3. Mark any extra prompt expectation that appears between those steps as either valid, conditional, or wrong.
4. Use this file as the source-of-truth checklist before adjusting prompt text in the script.