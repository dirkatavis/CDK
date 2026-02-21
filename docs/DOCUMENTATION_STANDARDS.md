# Documentation Standards

## File Format Philosophy

This project follows a clear convention for choosing documentation formats based on the **target audience**:

### Format Selection Rules

| Audience | Format | Rationale | Examples |
|----------|--------|-----------|----------|
| **Human Users** | **HTML** | Styled, visually appealing, browser-friendly presentation. Better UX for non-technical users. | `Install.html` - Setup guide for end users |
| **AI Systems / Developers** | **Markdown** | Plain text, version control friendly, easier to parse programmatically, better for LLMs. | `README.md`, `MIGRATION_MAPPING.md`, all docs/ files |

### Decision Criteria

**Choose HTML when:**
- The audience is non-technical end users
- Visual presentation improves comprehension (styled sections, color-coded steps)
- The document is primarily consumed via browser
- The content is procedural/instructional with clear visual hierarchy needs
- Examples: Installation guides, user manuals, visual tutorials

**Choose Markdown when:**
- The audience is developers, AI agents, or technical users
- The content needs to be version controlled and diffed
- The document will be parsed programmatically
- The content is technical reference material
- Examples: Architecture docs, API references, contributing guides, changelogs

### Duplication Policy

**Avoid duplicating the same content in multiple formats** (e.g., `Install.md` + `Install.html`). Choose the format that best serves the primary audience and maintain only that version.

**Exception:** If you must support both audiences, maintain a single source of truth (typically Markdown) and generate the HTML version from it using a build process.

## Current Repository Examples

- ✅ **Correct**: `Install.html` (user-facing, styled setup guide)
- ❌ **Incorrect**: `Install.md` (duplicate of Install.html, marked for deletion)
- ✅ **Correct**: `README.md`, `MIGRATION_MAPPING.md`, `docs/*.md` (developer/AI audience)

## Migration Note

This standard was established during the February 2026 repository reorganization after discovering `Install.md` as an unnecessary duplicate of `Install.html`.
