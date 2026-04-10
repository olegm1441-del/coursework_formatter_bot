# AGENTS.md

## Mission
You are working on a production Telegram bot with backend logic, conversation UX, and deployment through Railway.
Your job is to make safe, minimal, reviewable changes that preserve stability and avoid regressions.

The repository is production-sensitive.
Default behavior must be conservative.

---

## Absolute rules

### Branch and release safety
- Never push directly to `main`.
- Never merge into `main`.
- Never create or update release tags.
- Never deploy to production on your own.
- Never treat approval for coding as approval for merge, push, deploy, or release.
- Any action affecting `main`, production, or release flow requires an explicit direct command from me in the current thread.

### Sensitive business domains
- Never modify payment, billing, subscriptions, tariffs, invoices, checkout, credits, balance, referrals with payouts, or any money-related logic unless I explicitly authorize it in the current thread.
- If a requested change may indirectly affect payment-related flows, stop, explain the risk, and ask for confirmation before touching that area.
- If unsure whether a module is payment-related, assume it is sensitive and ask first.

### Scope control
- Do not perform broad rewrites.
- Do not rename files, move modules, or change public interfaces unless required by the task.
- Do not introduce new dependencies without explicit approval.
- Do not change environment variables, secrets handling, CI/CD, Railway config, webhook settings, or infrastructure unless the task explicitly requires it.
- Do not modify unrelated failing tests.
- Do not fix “nearby issues” unless I ask.

### Quality bar
- Prefer the smallest correct change.
- Prefer clarity over cleverness.
- Preserve backward compatibility unless the task explicitly requires a breaking change.
- Avoid hidden side effects.
- Keep diffs reviewable.

---

## Required workflow

Do not start implementation immediately on a new task.

### Phase 1 — Understanding
First, do all of the following:
1. Restate the task in your own words.
2. Identify affected modules, handlers, services, data flows, and external integrations.
3. Identify assumptions and uncertainties.
4. Ask clarifying questions when logic, UX, or constraints are not fully clear.
5. Explicitly call out whether the task may touch sensitive areas:
   - payment/billing
   - auth/security
   - data deletion
   - Telegram webhook/update routing
   - Railway deployment/runtime config

### Phase 2 — Test and acceptance design
Before writing code:
1. Define acceptance criteria.
2. Propose a test plan covering:
   - happy path
   - edge cases
   - regression risks
   - invalid input / malformed Telegram updates
   - retry / duplicate delivery behavior where relevant
   - permission / role boundaries where relevant
   - UX wording and conversation state transitions where relevant
3. Identify existing tests, commands, and files to inspect.
4. Propose the implementation plan.
5. Stop and wait for my approval before implementation, unless I explicitly say `implement now`.

### Phase 3 — Implementation
After approval:
- Implement in small, reviewable steps.
- Keep changes minimal and local.
- Reuse existing patterns in the repository.
- Avoid duplication.
- If logic repeats in multiple places and extraction is clearly justified, propose or perform a small focused extraction.
- Do not silently refactor large areas while implementing a feature.

### Phase 4 — Verification and review
After implementation:
1. Run relevant tests, linters, and type checks if available.
2. Summarize exactly what changed.
3. Summarize what was verified and how.
4. List residual risks and unverified assumptions.
5. Suggest a small follow-up refactor only if it is truly in scope.

Do not say “done” without verification results.

---

## Telegram bot specific rules

### Backend behavior
- Be careful with handler ordering, middleware side effects, callback data parsing, and conversation state transitions.
- Preserve idempotency where duplicate Telegram updates may occur.
- Validate user input defensively.
- Keep error handling explicit for network calls, Telegram API failures, and storage failures.
- Avoid breaking existing commands, buttons, callback payloads, and deep links unless explicitly requested.

### UX behavior
- Optimize for clear user flow, low confusion, and predictable next steps.
- Keep bot text concise, human, and unambiguous.
- If changing UX copy or conversation structure, explain:
  - what the old flow was
  - what the new flow is
  - why the new flow is better
  - what edge cases were considered
- Prefer improving existing UX over redesigning the entire flow.
- Do not make the bot more “clever” at the cost of consistency.

### Data and state
- Preserve conversation state integrity.
- Be explicit about timeout, cancellation, back-navigation, repeated button taps, and partial user input.
- If a migration is required, flag it before implementation.

---

## Architecture and coding standards

### General
- Follow the repository’s existing architecture and naming conventions.
- Prefer pure functions and isolated business logic where practical.
- Keep transport/framework code thin.
- Separate Telegram transport concerns from domain logic when possible.
- Do not duplicate domain rules across handlers.
- Centralize reusable validation and formatting logic if repetition is real.

### Refactoring policy
- Refactoring is allowed only in a narrow, task-adjacent way.
- After each completed task, do a small cleanup pass:
  - remove dead code introduced by the change
  - remove obvious duplication created by the change
  - improve names only where it materially improves readability
- If a larger refactor is beneficial, propose it separately instead of mixing it into the task.

### Logging and observability
- Add logs only where they help diagnose real operational issues.
- Do not log secrets, tokens, personal data, or payment information.
- Prefer structured, high-signal logging over noisy logging.

---

## Railway and deployment constraints
- Assume Railway deployment is production-sensitive.
- Do not modify Railway-related config, startup commands, env usage, health checks, background workers, or webhook URLs unless explicitly required.
- If a code change has deployment implications, state them clearly before implementation.
- Flag any change that may alter runtime behavior, cold start profile, process model, or background job execution.

---

## GitHub workflow
- Default target is a feature branch, not `main`.
- Prefer preparing a reviewable diff / PR-ready change set.
- Before proposing merge, provide:
  - changed files
  - purpose of each change
  - tests run
  - known risks
- Never assume permission to merge.
- Never assume permission to squash, rebase, or rewrite history.

---

## Response format for every substantial task

Use this structure:

### 1. Understanding
- what the task is
- what areas are affected
- what is unclear

### 2. Risks
- regressions
- sensitive modules
- deployment impact
- data/UX risks

### 3. Acceptance criteria
- clear “done when” checklist

### 4. Test plan
- happy path
- edge cases
- regressions

### 5. Implementation plan
- small numbered steps

### 6. Wait
- stop and wait for approval before coding

After coding, use this structure:

### 1. Changes made
### 2. Tests / checks run
### 3. Verification results
### 4. Residual risks
### 5. Optional small cleanup suggestions

---

## When to stop and ask me first
Always stop before proceeding if:
- payment or billing might be touched
- a migration may be needed
- public API or bot flow changes materially
- infrastructure or Railway config may change
- a dependency must be added
- the task conflicts with existing architecture
- the safest option is unclear

When in doubt, ask.
