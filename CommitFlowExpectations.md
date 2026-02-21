Commit Segmentation Strategy
When dividing code changes into chunks, ignore line counts or time. Instead, use these three objective pillars:

1. The Isolation Pillar (Independence)
Rule: Every commit must be "bi-directionally stable."

Standard: If I check out only this commit, the build must pass and tests must run.

Subdivision Trigger: If a change can be reverted without breaking the rest of the new code, it must be its own commit.

2. The Intent Pillar (Separation of Concerns)
Rule: Never mix "What the code is" (Refactor) with "What the code does" (Feature/Fix).

Subdivision Tiers:

Tier 1: Mechanical (Low Risk): Renaming, reformatting, or moving files.

Tier 2: Structural (Medium Risk): Changing how a function is written without changing its output.

Tier 3: Functional (High Risk): Changing logic, fixing a bug, or adding a feature.

The Standard: A single commit may only contain changes from one Tier.

3. The Architectural Pillar (Layering)
Rule: Divide by the "Blast Radius" of the change.

Standard: If a bug fix or feature spans multiple layers, commit them in this specific order:

Contract/Data: Changes to API definitions, Types, or Database schemas.

Logic/Service: Changes to the "brain" of the application (Calculations, API calls).

UI/Presentation: Changes to the visual components and user interactions.

The "Changelog" Litmus Test
To verify a chunk before finalizing, ask: "Can I describe the value of this specific commit in one sentence without using the word 'and'?"

Good: "Refactor: Extract user validation to a standalone utility."

Bad: "Refactor user validation and fix the null pointer bug in the login form."

What this looks like in practice:
If you have a task to "Update the Login API and clean up the Auth hook," the AI should now automatically produce:

refactor(auth): Rename variables and clean up internal hook logic (Tier 2).

feat(api): Update endpoint from v1 to v2 in the config (Layer 1).

fix(auth): Implement the new API response handling in the UI (Layer 3).