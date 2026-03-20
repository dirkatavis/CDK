# GitHub Copilot Instructions - CDK DMS Automation

## Engineering Rules
- Legacy repo nearing retirement — optimize for simplicity and immediate value; do not over-engineer.
- Keep scripts procedural by default; use state-machine logic only where already established (notably `apps/post_final_charges/PostFinalCharges.vbs`).
- In `.vbs` files, use `Option Explicit`.
- Before sending terminal input, wait for known prompts and verify state with screen reads (`ReadScreen`/`IsTextPresent`).
- Use dynamic line-letter discovery (`DiscoverLineLetters`) on RO detail screens; do not assume contiguous line letters.

## Prompt Handling (Critical)
- Prompt regex must tolerate intervening text and optional defaults in parentheses.
- Use patterns equivalent to `.*(\(.*\))?\?` where appropriate (for example, operation-code prompts).
- For default-value prompts in state-machine flows, keep `AddPromptToDictEx(..., acceptDefault:=True)` behavior that accepts defaults via Enter only.

## Path and Config Rules (Non-Negotiable)
- Never hardcode absolute paths.
- Resolve paths through `framework/PathHelper.vbs` and `GetConfigPath(section, key)`.
- Add new file references to `config/config.ini` before use.
- Fail fast on missing config/path setup; do not add silent fallback behavior.

## Testing
- Run VBScript through `cscript.exe`.
- Use mock-based tests when possible (for BlueZone-dependent logic).
- When fixing prompt-detection bugs, add/adjust regression coverage in related test files.

## Code Review
- Ignore nits (style, naming, minor readability), performance tuning, and non-critical hardening.
- Flag only: critical bugs (failure/data corruption risk), important business-logic errors, breaking changes.

## Pull Request Workflow
- `gh` CLI is **not installed** on this machine. Do not use it.
- To open or update a PR, use `tools/upsert_pr_with_body.ps1` — it calls the GitHub REST API using the same credentials already stored for `git push` (no browser, no manual auth).
- Always write the PR body to a file under `Temp/` first, then invoke the script:
  ```powershell
  .\tools\upsert_pr_with_body.ps1 -Title "your title" -BodyFile "Temp\pr_body.md"
  ```
- The script auto-detects the current branch, detects existing open PRs (upserts), and uses `-Repo` when provided; otherwise it parses owner/repo from `git remote get-url origin`.
- `-Head`, `-Base`, `-Draft`, and `-Repo` are optional overrides.
- If no credentials are found, the script fails fast with an explicit error; it does not open a browser fallback.
