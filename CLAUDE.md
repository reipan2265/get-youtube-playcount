# CLAUDE.md — AI Assistant Guide for GetYoutubePlayCount

This file provides context and conventions for AI assistants (Claude Code and similar tools) working in this repository.

## Project Overview

**GetYoutubePlayCount** is a tool for retrieving play count (view count) information from YouTube videos. The project is currently in its initial scaffold stage — no implementation exists yet.

**Intended purpose:** Query YouTube's API or scrape public data to fetch video view/play counts, likely returning structured data that can be used programmatically or displayed to users.

## Repository State (as of 2026-03-07)

- Single file: `README.md` (contains only a heading)
- No language, framework, or dependencies chosen yet
- No tests, CI/CD, or build tooling configured

When implementation begins, this file should be updated to reflect the actual stack and conventions.

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

## Development Setup (To Be Determined)

No language or framework has been selected. When the stack is chosen, document here:

- **Language:** TBD
- **Runtime/Version:** TBD
- **Package manager:** TBD
- **Key dependencies:** TBD

### Expected Setup Steps (Template)

```bash
# Clone and enter repo
git clone <repo-url>
cd GetYoutubePlayCount

# Install dependencies (update once stack is decided)
# e.g., pip install -r requirements.txt
# e.g., npm install

# Run the tool
# e.g., python main.py --video-id <VIDEO_ID>
```

## Likely Implementation Considerations

### YouTube Data API v3

The most reliable way to fetch view counts is via the YouTube Data API v3:

- Endpoint: `GET https://www.googleapis.com/youtube/v3/videos`
- Parameters: `part=statistics`, `id=<VIDEO_ID>`
- Requires a Google API key
- Returns `viewCount`, `likeCount`, `commentCount`, etc.

### API Key Handling

- Never hardcode API keys in source files
- Use environment variables (e.g., `YOUTUBE_API_KEY`)
- Add `.env` files to `.gitignore`

### Rate Limiting

YouTube Data API v3 has a daily quota (10,000 units by default). Each `videos.list` call costs 1 unit. Factor this into any batching or caching strategy.

## Testing

No tests exist yet. When added, document the test runner and how to run tests here. Preferred patterns:

- Unit tests for parsing/data transformation logic
- Integration tests (mocked or real API) for the HTTP client
- Test files should live in a `tests/` directory

```bash
# Example (update once stack is chosen)
pytest tests/
# or
npm test
```

## Code Conventions (To Be Established)

Once the language is chosen, add specific style/lint rules here. General principles to follow regardless of language:

- Keep functions small and focused on a single responsibility
- Validate inputs at system boundaries (user input, API responses)
- Prefer explicit error messages over silent failures
- Do not over-engineer: avoid abstractions that serve only one use case

## File Structure (Anticipated)

```
GetYoutubePlayCount/
├── CLAUDE.md           # This file
├── README.md           # User-facing documentation
├── .gitignore          # Excludes .env, build artifacts, etc.
├── src/ or main.*      # Core implementation
└── tests/              # Test files
```

Update this section once the actual structure is created.

## Key Commands Reference

| Task | Command (TBD) |
|------|--------------|
| Install dependencies | TBD |
| Run tool | TBD |
| Run tests | TBD |
| Lint/format | TBD |

## Notes for AI Assistants

- This repo is at the very start of development — no assumptions about language or framework should be made without checking for new files first
- Always read existing files before editing or creating new ones
- When adding implementation, update this CLAUDE.md to reflect actual conventions
- Do not commit `.env` files, API keys, or credentials
- Push changes to the `claude/...` branch, not to `master` directly
