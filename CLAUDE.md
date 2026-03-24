# CLAUDE.md — AI Assistant Guide for GetYoutubePlayCount

This file provides context and conventions for AI assistants (Claude Code and similar tools) working in this repository.

## Project Overview

**GetYoutubePlayCount** is a tool for retrieving play count (view count) information from YouTube videos. The project is currently in its initial scaffold stage — no implementation exists yet.

**Intended purpose:** Query YouTube's API or scrape public data to fetch video view/play counts, likely returning structured data that can be used programmatically or displayed to users.

## Repository State (as of 2026-03-25)

- Fully implemented as a Google Apps Script project
- Language: JavaScript (GAS environment)
- Deploy tooling: clasp (`npm run push` / `npm run deploy`)
- No tests configured

## Git Workflow

### Branch Naming

- Feature/fix branches follow: `claude/<description>-<session-id>`
- The main branch is `master`

### Pushing Changes

Always push using:

```bash
git push -u origin <branch-name>
```

Branch names used by Claude Code sessions start with `claude/` and end with the session ID.

### Commit Style

Write clear, descriptive commit messages in the imperative mood:

```
Add YouTube Data API v3 client wrapper
Fix rate-limit handling in view count fetcher
Update README with usage examples
```

Avoid vague messages like "fix stuff" or "update code".

## Development Setup

- **Language:** JavaScript (Google Apps Script)
- **Package manager:** npm
- **Key dependencies:** clasp, `@types/google-apps-script`

```bash
# Install dependencies (type definitions only — not needed for runtime)
npm install

# Push code to GAS
npm run push

# Push + git commit & push
npm run deploy
```

## Architecture

### Trigger Design

GAS has a 6-minute execution limit. To avoid timeouts, the script is split into two functions with separate triggers:

| Function | Frequency | Responsibility |
|----------|-----------|----------------|
| `main()` | Every 1 hour | Fetch & record view counts only |
| `updateAllCharts()` | Every 6 hours | Update charts, comparison sheet, sheet order |

Chart operations (`insertChart`, `updateChart`) are the primary cause of timeouts, so they are isolated from data collection.

### Key Files

- `main.gs` — All logic (single file for GAS compatibility)
- `appsscript.json` — GAS manifest (API scopes, runtime)
- `package.json` — clasp deploy scripts

## Code Conventions

- All logic lives in `main.gs` (GAS requires a single deployable file or flat structure)
- Private helper functions use a trailing underscore (`fetchVideoData_`, `runSampling_`, etc.)
- `CONFIG` at the top of `main.gs` is the only place users need to edit
- `SpreadsheetApp.flush()` is called after heavy sheet operations to prevent buffering-related timeouts
- `retryOnTimeout_()` wraps operations that are prone to transient GAS timeouts

## File Structure

```
GetYoutubePlayCount/
├── CLAUDE.md           # This file
├── README.md           # User-facing documentation
├── main.gs             # All GAS logic
├── appsscript.json     # GAS manifest
├── jsconfig.json       # Editor type support
├── package.json        # clasp scripts + type dep
└── package-lock.json
```

## Key Commands Reference

| Task | Command |
|------|---------|
| Install dependencies | `npm install` |
| Push to GAS | `npm run push` |
| Push + git deploy | `npm run deploy` |

## Notes for AI Assistants

- This repo is at the very start of development — no assumptions about language or framework should be made without checking for new files first
- Always read existing files before editing or creating new ones
- When adding implementation, update this CLAUDE.md to reflect actual conventions
- Do not commit `.env` files, API keys, or credentials
- Push changes to the `claude/...` branch, not to `master` directly
