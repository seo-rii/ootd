# OOTD Roadmap

This document contains the active roadmap. Detailed historical steps remain in `PLAN.md`.
Milestones close only when their exit criteria are met; adding API names or synthetic happy-path
tests alone does not complete a compatibility milestone.

## Active Sequence

### M0 — Baseline And Scope Control

Status: complete.

- Align the declared MSRV and package metadata across all crates.
- Establish `STATUS.md` as the current compatibility source of truth.
- Establish this file as the active M0-M5 roadmap.
- Freeze new chart features, new array-function breadth, and typed pivot work until their
  prerequisite milestones close.

### M1 — Behavioral Excel Oracle

Status: in progress.

Implemented foundation:

- Versioned typed case, observation, suite, and exact-byte run-manifest contracts.
- Required-case completeness mapping into the existing differential report and blocking gate.
- An `ExcelRuntime` adapter for get/set/invoke/calculate cases with typed arrays and symbolic
  bindings; save cases remain gated on an external Excel normal-open verifier.
- A dependency-free .NET 10 runner with cross-platform contract tests, fake-backed lifecycle
  tests, late-bound COM automation, executable-part preflight, and a PID-scoped watchdog.

Still required before M1 closes:

- Execute the runner on the declared Windows/Excel profile and pin the actual fingerprint.
- Capture and replay at least 20 required cases twice, including save/reopen repair evidence and
  normalized package relationship artifacts.

- Keep OM TypeLib/PIA acquisition separate from behavioral Excel observation.
- Define versioned case, run-manifest, Oracle-observation, and runtime-observation contracts.
- Execute the same operation DSL through desktop Excel COM and `ExcelRuntime`.
- Compare typed return values, errors, arrays, symbolic object identity, workbook state,
  normalized package relationships, save/reopen results, and repair detection.
- Pin an initial required corpus and replay it without Excel in normal pull-request CI.

Exit criteria:

- One Excel build/channel/architecture/locale/timezone profile is pinned.
- At least 20 required cases have input SHA-256 and provenance metadata.
- Required cases contain no failed, missing, unsupported, or skipped result.
- Saved files reopen in Excel with repair explicitly observed as false.
- Pinned observations reproduce twice on the same host after normalization.

### M2 — Stable Internal Boundaries

Status: in progress after the synthetic M1 vertical slice; real Excel evidence remains pending.

Completed slices:

- Externalized all 2,826 `excel-xlsx` and 677 `excel-runtime` unit tests from their library roots
  while preserving both sorted test-name hashes and passing behavior.
- Routed the parser-backed formula coverage scan through an explicit implementation-source
  contract so calculation code can move without silently weakening the coverage gate.
- Moved the existing formula evaluator, parser, reference conversion, and calculation helpers
  behind `excel-runtime::calc` without changing public paths or test identities.
- Replaced the calculation module's wildcard parent import with an explicit internal dependency
  list, keeping its boundary reviewable before the M4 value-model redesign.
- Isolated workbook/sheet recalculation and dynamic-array spill writeback in a dedicated module so
  M4 can replace mutation semantics without reopening the evaluator implementation.
- Isolated shared-string parsing as the first `excel-xlsx` codec boundary.
- Isolated relationship parsing, part URI derivation, and target normalization before M3 changes
  their validation semantics.
- Isolated worksheet cell parsing, lossless XML rewriting, error conversion, and dimension helpers
  behind `excel-xlsx::worksheet` with the full test-name inventory unchanged.
- Grouped Application property and method dispatch in an explicit-dependency object-surface module,
  establishing the extraction pattern for the remaining runtime objects.
- Grouped Workbook and Workbooks collection get/invoke dispatch in the same object-surface pattern,
  leaving shared formula evaluation and operation routers in their current ownership.
- Grouped WorksheetFunction dispatch with its exclusive scalar/array/range argument rendering
  helpers, keeping formula evaluation itself in the calculation boundary.

- Apply the object-surface dispatch pattern to worksheet, range, names, and chart objects.

Exit criteria:

- Public root paths and test names remain stable.
- Move-only changes contain no semantic body changes beyond imports and visibility.
- Runtime state and package semantic snapshots are identical before and after each move.

### M3 — CI And Untrusted Input Safety

Status: pending; independent slices may run alongside M1 and M2.

- Enforce formatting, strict Clippy, MSRV, Linux stable, and Windows portability jobs.
- Add bounded ZIP loading and canonical OPC part/relationship identity.
- Apply shared XML depth, event, text, attribute, and collection budgets.
- Add dependency/license policy, property tests, scheduled fuzzing, and benchmark trends.

Exit criteria:

- Every default public workbook-open path is resource-bounded.
- Limit-plus-one, entry-flood, compression-bomb, malformed-target, and XML-budget tests fail
  with stable structured errors and do not expose partial models.
- General Windows CI does not require Excel; Excel automation remains a separate job.

### M4 — Formula2 Foundation

Status: pending M1 evidence and M2 calculation boundaries.

- Route all cell mutations through invariant-preserving model commands.
- Make spill replacement and obstruction atomic across all mutation paths.
- Introduce a common evaluation value model for scalar, array, reference, error, lambda, and
  omitted arguments.
- Centralize scalar, aggregate, array, and reference coercion.
- Complete `INDEX`, `INDIRECT`, `OFFSET`, `TRIMRANGE`, names, multi-area, 3D, `@`, and `#` semantics.
- Add dependency invalidation, cycle handling, and dynamic-array XLSX save/reopen metadata.
- Revalidate the existing 17 array functions before adding higher-order functions.

Exit criteria for every supported function:

`value + shape + reference + error + recalculation + mutation + save/reopen + Excel agreement`

### M5 — Pivot Preserve-Only

Status: pending M3 package safety and M4 calculation foundation.

- Inventory pivot table definitions, cache definitions and records, slicers, timelines,
  relationships, and content types.
- Preserve the graph across no-op save and unrelated cell edits.
- Define safe sheet rename/copy/move/delete and shared-cache ownership behavior.
- Reject unsupported destructive mutations before changing model or package state.

Exit criteria:

- The tracked pivot corpus retains its semantic part and relationship graph across supported
  operations and reopens in Excel without repair.
- Typed PivotTable mutation, refresh, and PivotChart binding remain out of scope until a later
  milestone.

## Development Policy During M0-M5

- Use test-driven development for behavior changes: failing regression, implementation, then
  focused and broader passing tests.
- Keep one independently reviewable and revertible work unit per commit.
- Never update Oracle golden results automatically in CI.
- Do not mix mechanical code movement with semantic changes.
- Continue to preserve unknown package data or return an explicit unsupported error when safe
  mutation cannot be proved.
