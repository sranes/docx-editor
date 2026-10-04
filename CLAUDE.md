# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

`AGENTS.md` is the source of truth for this repository. It covers architecture, packages, verification gates, security rules, i18n, docs style, releasing, and conventions. Read it first and follow it. This file only adds command detail that `AGENTS.md` does not list.

@AGENTS.md

## Additional commands

The package manager and test runner is Bun. Root scripts live in `package.json`.

| Task | Command |
| --- | --- |
| Run one test file | `bun test path/to/file.test.ts` |
| Run tests by name | `bun test path/to/file.test.ts -t "name"` |
| Watch tests | `bun run test:watch` |
| Whole suite, one process | `bun run test:serial` |
| Fix lint findings | `bun run lint:fix` |
| Check formatting only | `bun run format:check` |
| Build all packages | `bun run build:packages` |
| Dev server, React and Vue demos | `bun run dev` |
| Dev server, one demo | `bun run dev:react`, `dev:vue`, `dev:igloo`, `dev:markdown`, `dev:collaboration`, … |
| Editor smoke E2E (Playwright) | `bun run test:e2e:editor` |
| Editor acceptance E2E | `bun run test:e2e:acceptance` |
| Collaboration E2E | `bun run test:e2e:collab` |
| Editor API compat report | `bun run --filter '@docx-editor.dev/editor-api' compat:report` |
| One package's script | `bun run --filter '@docx-editor.dev/<pkg>' <script>` |

E2E specs and their Playwright configs are in `e2e/`. Each `*.config.ts` there selects a spec group.

## Windows note

Many root scripts use POSIX syntax, such as `VAR=x cmd`, `cd … &&`, `rm -rf`, and `bash`. On Windows, run them from Git Bash, not PowerShell.
