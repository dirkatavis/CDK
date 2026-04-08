# CDK Project — Claude Instructions

## GitHub CLI
`gh` is not installed on this machine. Do not attempt `gh pr create`, `gh pr edit`, or any other `gh` command.

**Always create and update PRs via:**
```
powershell -ExecutionPolicy Bypass upsert_pr_with_body.ps1 -Title "..." -BodyFile "tools\<body_file>.md"
```

Write the PR body to a `.md` file in `tools\` first, then pass it to the script.
The script handles both create and update (upsert) automatically.
`upsert_pr_with_body.ps1` is on the user PATH via `C:\code\Scripts\`. A local copy also exists in `tools\` as a fallback.
