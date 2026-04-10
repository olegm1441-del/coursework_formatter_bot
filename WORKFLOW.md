# WORKFLOW.md

## Purpose
This workflow defines the mandatory execution protocol for all non-trivial tasks in this repository.

The primary goal is not speed.
The primary goal is safe, minimal, verified changes with zero careless regressions.

The agent must behave conservatively at all times.

---

## Core principles

- Do not guess.
- Do not rush into coding.
- Do not skip phases.
- Do not broaden scope.
- Do not remove or rewrite existing behavior unless explicitly required.
- Do not destroy working logic in order to implement a new feature.
- Treat existing code, UX flows, handlers, and integrations as potentially important unless proven otherwise.
- Prefer a small correct change over a large elegant rewrite.
- Every non-trivial task must end with verification.
- Every completed task must end with a small cleanup/refactor pass, but only inside the task scope.

---

## Global execution rule

For every substantial task, the agent must follow all phases in order.

The agent must not move to the next phase without completing the current one.

The agent must not move from planning to implementation without explicit approval from the user, unless the user clearly says `implement now`.

If any requirement, behavior, constraint, or expected result is unclear, stop and ask.

---

## Phase 1 — Reformulation and system understanding

Before writing or editing any code, do all of the following:

1. Restate the task in precise technical terms.
2. Explain the user-visible goal and the system-level goal.
3. Identify the affected areas:
   - backend modules
   - Telegram handlers
   - callback flows
   - conversation state
   - storage/database
   - external APIs
   - deployment/runtime behavior
   - UX/copy/user flow
4. List assumptions.
5. List uncertainties.
6. Identify what must not be broken.
7. Identify sensitive zones that may be affected, including:
   - payment/billing
   - auth/security
   - webhook/update routing
   - Railway config/runtime
   - state persistence
   - migrations
   - admin-only flows
8. Ask clarifying questions if anything is ambiguous.

### Phase 1 output format
The agent must produce:
- task restatement
- affected areas
- assumptions
- uncertainties
- protected behaviors
- clarifying questions

The agent must then stop unless the task is fully clear.

---

## Phase 2 — Success criteria and failure prevention

Before implementation, define exactly what success means.

The agent must:

1. Write explicit acceptance criteria.
2. Write explicit non-goals.
3. Define what would count as failure, regression, or unacceptable side effect.
4. Identify existing functionality that could accidentally be damaged.
5. Identify what must remain unchanged.

This phase exists to prevent “technically implemented, but actually wrong” outcomes.

### Acceptance criteria must include where relevant:
- functional behavior
- UX behavior
- state transitions
- error handling
- invalid input handling
- backward compatibility
- no regression in existing flows

### Phase 2 output format
The agent must produce:
- acceptance criteria
- non-goals
- failure modes
- unchanged guarantees

The agent must then stop and wait if major ambiguity remains.

---

## Phase 3 — Test design before code

Before implementation, the agent must design verification first.

The agent must define:

1. Happy path tests
2. Edge case tests
3. Regression tests
4. Invalid input tests
5. Duplicate/retry/update replay tests where relevant
6. State transition tests where relevant
7. Role/permission tests where relevant
8. UX/copy validation checks where relevant
9. Integration checks where relevant
10. Manual verification steps if automated tests are insufficient

The agent must also identify:
- existing tests to reuse
- new tests to add
- commands to run
- mocks/stubs needed if applicable

The agent must not start implementation until there is a clear test strategy.

### Required mindset
Tests are not decoration.
Tests are proof that the change is correct and safe.

### Phase 3 output format
The agent must produce:
- test matrix
- commands to run
- coverage notes
- limitations of testing

---

## Phase 4 — Implementation plan and approval gate

Before changing code, the agent must propose a small, controlled implementation plan.

The plan must include:

1. Files likely to change
2. Purpose of each change
3. Order of changes
4. Risk level of each step
5. Whether any extraction or refactor is truly necessary
6. What will explicitly not be touched

### Rules for this phase
- The agent must prefer the smallest possible diff.
- The agent must prefer local changes over cross-cutting rewrites.
- The agent must not modify files outside the proposed list without approval.
- The agent must not add dependencies without approval.
- The agent must not introduce new abstractions unless clearly justified.
- The agent must not refactor preemptively.

### Approval gate
After presenting the implementation plan, the agent must stop and wait for explicit approval.

Allowed examples:
- `implement now`
- `go ahead`
- `approved`

Without approval, do not code.

### Phase 4 output format
The agent must produce:
- implementation plan
- file change list
- risk notes
- approval request

Then stop.

---

## Phase 5 — Implementation

Only after approval, the agent may implement.

### Implementation rules
- Make minimal, reviewable changes.
- Keep scope tight.
- Preserve existing architecture unless the task explicitly requires otherwise.
- Preserve public behavior unless the task explicitly requires otherwise.
- Reuse existing patterns from the repository.
- Avoid duplication.
- If repeated logic clearly emerges, a small focused extraction is allowed only if it directly supports the task and reduces risk.
- Never silently rewrite adjacent code.
- Never delete code unless its removal is necessary and justified.
- Never remove behavior without explicitly mentioning it.
- Be especially careful with:
  - handler order
  - callback parsing
  - message formatting
  - state mutation
  - retries
  - async side effects
  - persistence boundaries
  - environment-dependent behavior

### During implementation
The agent should keep a short running log of:
- what changed
- why it changed
- whether the change altered scope or not

If new uncertainty appears during implementation:
- stop
- explain
- ask before continuing

---

## Phase 6 — Verification

After implementation, verification is mandatory.

The agent must verify the result against:
1. Acceptance criteria
2. Regression risks
3. Existing behavior guarantees
4. Test plan from Phase 3

### Required verification actions
Run all relevant available checks, where applicable:
- unit tests
- integration tests
- end-to-end tests
- lint
- typecheck
- build
- smoke test
- manual flow verification

### Verification must include:
- expected result
- actual result
- pass/fail status
- anything not verified
- why it was not verified

### If any test fails
- stop
- investigate
- fix
- re-run verification
- do not claim completion until the failing path is understood

### If verification is incomplete
The agent must say so explicitly.
The agent must never pretend something was verified if it was not.

### Phase 6 output format
The agent must produce:
- checks run
- results
- acceptance criteria mapping
- unverified areas
- residual risks

---

## Phase 7 — Final review and scoped refactor

Only after the change is functionally correct and verified, perform a small cleanup pass.

This is not permission for a rewrite.

### Allowed refactoring
Only task-adjacent cleanup is allowed:
- remove dead code introduced by the change
- remove duplication introduced by the change
- improve naming where it directly improves clarity
- simplify a small local conditional or helper
- extract a small reusable helper if clearly justified by the new code

### Forbidden refactoring
- broad architectural changes
- folder restructuring
- unrelated cleanup
- style-only sweeping edits
- rewriting stable modules “for beauty”
- changing interfaces without need
- refactoring old code unrelated to the task

### Refactor check
Before applying a refactor, ask:
- does this reduce risk?
- does this improve clarity?
- is this still inside the exact task scope?
- can this be reviewed easily?

If not, do not do it.

### Phase 7 output format
The agent must produce:
- cleanup/refactor done
- why it was safe
- what was intentionally left untouched

---

## Mandatory consultation rules

The agent must stop and ask the user before proceeding if:

- requirements are ambiguous
- expected UX is ambiguous
- a state transition is unclear
- payment/billing may be affected
- auth/security may be affected
- a migration may be needed
- a dependency may be added
- Railway/runtime behavior may change
- public interfaces may change
- data may be deleted or transformed
- existing code seems strange but may be intentional
- the smallest safe implementation is unclear
- test expectations are unclear
- verification cannot be completed reliably

When in doubt, ask.

---

## Anti-destruction rules

The agent must assume that existing code may be important, even if it looks redundant, old, strange, or poorly designed.

Therefore:
- do not delete code just because it appears unused without checking references and purpose
- do not remove branches in logic without understanding why they exist
- do not simplify flows if that may remove business behavior
- do not collapse states without confirming semantics
- do not replace working code with a “cleaner” abstraction unless necessary
- do not remove fallback behavior unless explicitly approved
- do not change user text or UX ordering casually

If something looks unnecessary, verify first.

---

## Standard response structure for every substantial task

Before coding, respond in this order:

### 1. Reformulation
### 2. Affected areas
### 3. Assumptions and uncertainties
### 4. What must not break
### 5. Acceptance criteria
### 6. Test plan
### 7. Minimal implementation plan
### 8. Files likely to change
### 9. Risks
### 10. Wait for approval

After coding, respond in this order:

### 1. Changes made
### 2. Files changed
### 3. Tests and checks run
### 4. Verification against acceptance criteria
### 5. Residual risks
### 6. Small cleanup/refactor performed
### 7. What was intentionally not changed

---

## Hard rules

- Never skip phases.
- Never code before planning unless explicitly instructed.
- Never treat coding approval as merge approval.
- Never treat coding approval as deployment approval.
- Never claim success without verification.
- Never hide uncertainty.
- Never guess silently.
- Never broaden scope without approval.
- Never push or merge to `main`.
- Never touch payment-related logic without explicit permission.
- Never make destructive changes casually.
- Never prioritize elegance over safety.
