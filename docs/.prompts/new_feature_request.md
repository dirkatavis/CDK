Role: Expert VBScript developer for CDK terminal automation.

Task: Modify tools\close_single_ro.vbs to add a mandatory Review phase before sending FC.

Review sequence: Execute R A, R B, R C in order (or from DiscoverLineLetters output if applicable).
After each review command, verify: The prompt at the bottom of the screen shows, "COMMAND:" . If failed: stop/cancel the script,and log error message
Only after all review checks pass, send FC.
Reuse/adapt logic from PostFinalCharges.vbs and keep behavior/style consistent.
Constraints: Keep procedural style, Option Explicit, no hardcoded paths, no unrelated refactors.
Deliverables: code changes + brief summary + tests updated/added in tests\run_all.vbs.