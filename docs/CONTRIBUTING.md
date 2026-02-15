# Contributing to CDK Automation

## Philosophy
This is a legacy system with a short lifespan (3-6 months).
**Goal:** Simplicity and Stability.
**Non-Goal:** Perfect abstraction or extensive refactoring.

## How to Add a New Script

1.  **Create Directory**: Create a new folder for your script (e.g., `MyNewTask/`).
2.  **Create Script**: Copy `common/template.vbs` (if it exists) or use `Close_ROs_Pt1.vbs` as a base.
3.  **Update Config**: Add your paths to `config.ini` in the root.
4.  **Documentation**: Create `MyNewTask/README.md` explaining input/output.
5.  **Validation**: Run `tools\validate_dependencies.vbs` to ensure your new paths are reachable.

## Git Workflow
*   **Branching**: Create feature branches (`feature/my-task`). Never commit directly to `main`.
*   **Commit Messages**: Keep them clear and descriptive.

## Code Style
*   **VBScript**: Use `Option Explicit`.
*   **BlueZone**: Use `bzhao` object.
*   **Paths**: ALWAYS use `GetConfigPath()`. NEVER hardcode paths.

## Testing
*   Run the script in a safe terminal screen first.
*   Use `.debug` files to slow down execution for visual verification.
