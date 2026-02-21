# Pull Request Workflow

This document outlines the pull request submission workflow for the CDK DMS Automation project.

## Repository Structure

- **Upstream Repository**: `dirkatavis/CDK` (primary)
- **Your Fork**: `<your-username>/CDK` (where you push changes)

## Contribution Workflow

### 1. Set Up Your Local Environment

Clone from upstream and configure remotes:

```bash
git clone https://github.com/dirkatavis/CDK.git
cd CDK
git remote add upstream https://github.com/dirkatavis/CDK.git
git remote rename origin fork  # Optional: rename origin to fork for clarity
git remote add origin https://github.com/<your-username>/CDK.git
```

Verify your remotes:

```bash
git remote -v
# origin    https://github.com/<your-username>/CDK.git (fetch)
# origin    https://github.com/<your-username>/CDK.git (push)
# upstream  https://github.com/dirkatavis/CDK.git (fetch)
# upstream  https://github.com/dirkatavis/CDK.git (push)
```

### 2. Create a Feature Branch

Always branch from upstream main:

```bash
git fetch upstream
git checkout -b feature/your-feature-name upstream/main
```

### 3. Make Your Changes

Follow the development guidelines:
- Keep commits focused and descriptive
- Reference GitHub issues: `Fixes #35` or `Closes #42`
- Run tests before committing: `cscript utilities\tests\run_all_tests.vbs`

### 4. Push to Your Fork

```bash
git push origin feature/your-feature-name
```

### 5. Create a Pull Request

**IMPORTANT**: Create the PR from your fork **to upstream**, not to your own fork.

**GitHub Web Interface (Recommended)**:
1. Visit https://github.com/dirkatavis/CDK
2. Click "Compare & pull request" or go to Pull Requests → New
3. **Base repository**: dirkatavis/CDK (upstream)
4. **Base branch**: main
5. **Head repository**: <your-username>/CDK (your fork)
6. **Head branch**: feature/your-feature-name
7. Add description and submit

**Command Line** (if available):
```bash
# Verify you're pushing to the correct upstream
git push upstream feature/your-feature-name
# OR
hub pull-request -b dirkatavis:main -h <your-username>:feature/your-feature-name
```

### 6. Respond to Review Feedback

- Keep the feature branch intact (don't delete after merge)
- Make additional commits in response to feedback
- Push to origin again: `git push origin feature/your-feature-name`
- The PR automatically updates

## Key Rules

✅ **DO:**
- Push feature branches to **your fork** (origin)
- Create PRs **from your fork to upstream** (dirkatavis/CDK)
- Keep commits focused and descriptive
- Reference issues in commit messages
- Run tests before pushing

❌ **DON'T:**
- Push directly to upstream (you don't have access)
- Create PRs from your fork to your fork's main
- Force-push to shared branches
- Mix unrelated changes in one PR

## Testing

Before submitting a PR:

```bash
# Run all tests
cscript utilities\tests\run_all_tests.vbs

# Run specific test category
cscript utilities\tests\test_hardcoded_paths_comprehensive.vbs
cscript utilities\tests\run_default_value_tests.vbs
```

## Example PR Checklist

- [ ] Branch created from `upstream/main`
- [ ] Changes are focused and related
- [ ] Commits have clear, descriptive messages
- [ ] Tests pass locally (10+/11 expected)
- [ ] PR created **from your fork to upstream** (not fork to fork)
- [ ] PR description includes related GitHub issue(s)
- [ ] Code follows existing patterns and conventions

## Questions?

Refer to:
- `.github/copilot-instructions.md` - Project standards and conventions
- `docs/` folder - Technical documentation
- GitHub Issues - Ask questions or report problems
