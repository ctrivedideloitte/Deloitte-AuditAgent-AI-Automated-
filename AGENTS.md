# Project Custom Instructions

This file contains persistent rules and context for the AI agent working on this project.

## Environment & Infrastructure
- **Port Strategy**: Always use port 3000 for the dev server.
- **Model Selection**: Currently optimized for `gemini-3-flash-preview`. When updating AI logic, ensure the model ID is correct for the current AI Studio environment.
- **Verification Priority**: 
    - Always run `lint_applet` after any code change to catch syntax errors or type mismatches early.
    - Run `compile_applet` before concluding a task to ensure the build pipeline is healthy.

## Coding Standards
- **Type Safety**: Maintain strict TypeScript patterns. Use `as File[]` for file inputs and ensure `unknown` types from external libraries (like XLSX) are properly cast.
- **UI Logic**: Ensure modals and overlays (like `AnimatePresence`) have balanced tags and proper nesting to avoid hydration or parsing errors.

## Memory & Updates
- If a new architectural pattern or specific user preference is established during a turn, it MUST be recorded in this file (`AGENTS.md`) to serve as project memory for future agent sessions.
- Always check `.env.example` for required environment variables.
