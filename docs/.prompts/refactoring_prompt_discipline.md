# Prompt Discipline — Refactoring Tasks

Lessons learned from a logging-only refactor that unintentionally modified business logic.

## Rules for prompts where refactoring scope is limited

1. **Explicit scope exclusion.** State what must NOT be touched:
   > "Scope is limited to [X] only. Do not modify, remove, or reorder any existing business logic, control flow, or non-[X] calls. If a suspected issue is out of scope, flag it in a comment but do not fix it."

2. **Diff review gate before any code is applied.**
   > "Before writing any changes, produce a plain-English summary of every line that will be modified and the reason. Wait for approval before proceeding."

3. **Define functional parity operationally, not aspirationally.**
   > "Functional parity means: a `git diff` against `main` on the target file must show zero changes to any non-[X] line — no removals, no additions, no reorderings of existing logic."
