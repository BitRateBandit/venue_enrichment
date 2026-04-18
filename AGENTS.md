# AGENTS.md

## Purpose
This file defines mandatory operating rules for AI agents working in this repository.

These rules are binding. When instructions conflict, prefer:
1. direct user request
2. this AGENTS.md
3. repository conventions already present in code/tests/docs

If a requested change would require violating these rules, explicitly say so.

---

## Operating Principles

- Make the smallest correct change that satisfies the request.
- Preserve existing behavior unless a behavior change is explicitly requested.
- Prefer local, deterministic, testable solutions.
- Prefer explicit, readable code over clever abstractions.
- Do not introduce new patterns when an existing repository pattern already solves the problem.

---

## Required Workflow

Before making changes, you must:

1. Read this file first.
2. Read the README.md and any files directly related to the requested change.
3. Inspect existing implementation patterns before introducing new ones.
4. Identify the minimum set of files that must change.
5. Avoid speculative edits outside the requested scope.

When finished, you must report:

- files changed
- what changed
- tests/validation run
- results
- any assumptions or limitations

---

## Scope Control

- Do not refactor unrelated code.
- Do not expand scope beyond the request.
- Do not rename files, functions, classes, flags, or fields unless required.
- Only modify files explicitly mentioned **plus** the minimal directly impacted supporting files required for correctness, such as:
  - tests
  - fixtures
  - schemas
  - documentation tied to changed behavior

If additional files are required, keep those edits minimal and explain why.

---

## Stability Requirements

Unless explicitly instructed otherwise, preserve:

- CLI interfaces and flag names
- default argument behavior
- file paths and output locations
- JSON / CSV / XLSX column names and schemas
- database schemas and table names
- checkpoint/resume behavior
- logging/report output shape
- public function signatures used elsewhere in the repo

Do not silently change data contracts.

---

## Code Philosophy

- Local-first
- Deterministic outputs
- No hidden side effects
- Explicit over implicit
- Readability over cleverness
- Small composable functions over broad rewrites

---

## Dependency Policy

- Prefer the Python standard library when practical.
- Do not add new dependencies unless they are clearly required.
- If adding a dependency, use the smallest suitable one and justify it.
- Do not introduce heavy frameworks for narrow tasks.
- Update dependency manifests only when necessary.

---

## Secrets Handling

- Follow SECRETS.md strictly.
- Never hardcode credentials.
- Never print or log secrets.
- Always use `get_secret()` where the repo expects it.
- Treat API keys, tokens, cookies, auth headers, signed URLs, and personal data as sensitive.

---

## Testing Requirements

- All changes must be testable.
- Prefer unit tests over manual validation.
- Reuse existing test patterns and fixtures when possible.
- No network calls in tests unless explicitly required.
- Add or update tests when behavior changes or bug fixes are introduced.
- Do not remove failing tests to make the suite pass.

If tests are not added, explain why they were unnecessary or impossible.

---

## State, Idempotency, and Safety

For pipelines, batch jobs, and enrichers:

- Preserve checkpoint/resume behavior.
- Do not duplicate writes on rerun.
- Ensure partial runs can recover safely when that pattern already exists.
- Skipped/completed work must remain skipped/completed unless explicitly changed.
- Do not double-count work, cost, or state during resume flows.
- Avoid destructive operations unless explicitly requested.

---

## Logging and Observability

- Keep logs useful, concise, and non-sensitive.
- Do not spam logs.
- Preserve existing status/reporting fields unless explicitly instructed otherwise.
- Prefer structured outputs when the repo already uses them.
- Add progress/cost/status instrumentation only in a way that does not break existing outputs.

---

## Error Handling

- Fail clearly, not silently.
- Reuse existing error-handling patterns where possible.
- Do not swallow exceptions without a reason.
- Include actionable error messages.
- Preserve partial progress when the repo already supports it.

---

## Output Expectations

Deliverables must be:

- clean
- readable
- minimal
- complete
- production-suitable

Do not leave:
- placeholders
- fake implementations
- stubbed logic
- TODO comments
- dead code

unless explicitly requested.

---

## Forbidden

- Hardcoding API keys or secrets
- Logging sensitive data
- Inventing requirements not supported by the request or repo patterns
- Changing data contracts without instruction
- Refactoring unrelated areas
- Adding broad abstractions for small changes
- Replacing stable existing patterns with personal preference

---

## When Uncertain

- Do not guess blindly.
- First inspect nearby code, tests, docs, and existing conventions.
- If ambiguity remains, choose the smallest implementation consistent with the current repository.
- State assumptions explicitly in the final report.

---

## Summary Rule

Make the smallest correct, testable, non-breaking change that satisfies the request and fits existing repository patterns.
