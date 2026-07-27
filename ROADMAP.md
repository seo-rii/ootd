# OOTD Roadmap

This document contains the active roadmap. Detailed historical steps remain in `PLAN.md`.
Milestones close only when their exit criteria are met; adding API names or synthetic happy-path
tests alone does not complete a compatibility milestone.

## 2026-07-27 Audit Priority Override

The repository-wide compatibility audit identified persistence correctness and silent-success
contracts as release blockers. Until Audit Wave 1 closes, new editing breadth, typed pivot work,
and new chart families remain frozen even when an older milestone below lists them as a next slice.

Active order:

1. `OOTD-001` commit every successful save as the next package/source baseline.
2. `OOTD-002`/`OOTD-046` separate prompt state from serialization dirty domains.
3. `OOTD-003`~`OOTD-006` make Save/SaveAs/SaveCopyAs/Close target-aware, read-only-safe, durable,
   atomic, and transactional through `OOTD-045`.
4. `OOTD-007`~`OOTD-009` persist formula caches and synchronize calculation metadata.
5. Continue with fail-closed public/security contracts, QName/reference/parser invariants, then
   cell/formula fidelity.
6. Close the compatibility loop with `OOTD-043`/`OOTD-085` pinned desktop Excel evidence before
   claiming practical chart/pivot/style parity.

Every numbered work unit starts with a failing regression and lands as its own reviewable commit.
The complete `OOTD-001`~`OOTD-086` ordering, regression inventory, and compatibility completion
definition are in `PLAN.md`; detailed active risks are in the local `RISK_REGISTER.md`.

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
- Grouped Worksheet and sheet-collection get/invoke dispatch by object surface while retaining
  shared sheet copy/move/delete lifecycle helpers in the runtime core.
- Grouped Names collection and Name object get/invoke dispatch with explicit model, formula, and
  metadata dependencies.
- Grouped Range and Areas get/set/helper dispatch by object surface; the inline Range method arm in
  the public invoke router remains a separately characterized follow-up.
- Grouped the chart-family helper surface, including ChartObjects, chart children, axes, legends,
  groups, series, labels, and points, in one explicit-dependency module.

- Characterize and extract the remaining inline object arms from the public get/set/invoke routers.

Exit criteria:

- Public root paths and test names remain stable.
- Move-only changes contain no semantic body changes beyond imports and visibility.
- Runtime state and package semantic snapshots are identical before and after each move.

### M3 — CI And Untrusted Input Safety

Status: core exit criteria complete; continuous hardening remains below.

Completed slices:

- Added finite default OPC ZIP budgets for archive bytes, central-directory entry count, part-name
  bytes, per-entry and total decompressed bytes, and compression ratio.
- Added an explicit `from_bytes_with_limits` override path and structured `ResourceLimit` errors;
  the default `from_bytes` path is bounded and preflights EOCD/ZIP64 entry counts before opening the
  central directory.
- Added canonical, ASCII-case-insensitive OPC part identities across load, lookup, mutation,
  content-type override resolution, and serialization; ambiguous case/percent-encoding duplicates
  and non-canonical URI spellings are rejected before package state is exposed.
- Made relationship parsing fail closed for missing required attributes, duplicate IDs, unknown
  target modes, malformed percent encodings, and internal targets that escape the package root or
  cannot identify a canonical part.
- Added shared ingress preflight for extension- and content-type-identified XML parts, with bounded
  depth, event count, text/CDATA bytes, cumulative attribute bytes, and attributes per element;
  malformed XML is rejected before specialized codecs expose partial state.
- Split CI portability coverage into explicit Ubuntu Rust 1.94, Ubuntu MSRV Rust 1.88, and Windows
  Rust 1.94 test lanes; the general Windows lane has no desktop Excel dependency.
- Added a bounded per-file rustfmt gate for 40 tracked Rust files, with four reviewed monolith
  exceptions guarded by path, minimum size, and individual growth ceilings.
- Enabled strict `-D warnings` Clippy for `office-idl`, `office-common`, `office-codegen`,
  `office-capture`, `office-opc`, and `excel-model` after clearing their existing warnings.

Continuous hardening:

- Ratchet strict Clippy across `excel-runtime` and `excel-xlsx`, and shrink the reviewed rustfmt
  exception set as M2 extraction continues.
- Add dependency/license policy, property tests, scheduled fuzzing, and benchmark trends.

Exit criteria:

- Every default public workbook-open path is resource-bounded.
- Limit-plus-one, entry-flood, compression-bomb, malformed-target, and XML-budget tests fail
  with stable structured errors and do not expose partial models.
- General Windows CI does not require Excel; Excel automation remains a separate job.

### M4 — Formula2 Foundation

Status: in progress on mutation invariants; real Excel evidence remains an external M1 dependency.

Completed slices:

- Added whole-batch spill-child preflight to model value, A1 formula/Formula2, and `ClearContents`
  commands so a later child target cannot leave earlier cells partially mutated.
- Characterized spill-anchor overwrite and clear: unstyled children are removed, styled children
  remain as blank shells, and owner/range/dynamic-formula metadata is cleared together.
- Routed single- and multi-area R1C1/Formula2R1C1 assignments through the same model formula
  command, removing two direct runtime mutation loops and extending the atomic child guard and
  anchor cleanup to their local aliases.
- Reconstructed array-formula anchor, extent, and materialized child ownership from worksheet XML,
  emitted `t="array"`/`ref` metadata for new Formula2 spills, and removed stale array attributes
  when an anchor becomes an ordinary formula. Spill ranges remain authoritative even when a blank
  child has no cached cell node.
- Added A1 spill-range references such as `J10#`, including explicit `#REF!` for non-spill anchors
  and extent lookup after a materialized spill changes shape.
- Split worksheet calculation into dynamic materialization followed by scalar evaluation, so a
  scalar `SUM(J10#)` dependent observes a changed spill shape in the same `Calculate` call without
  relying on cell order. Dynamic-to-dynamic dependencies and a general graph remain outstanding.

Next slices:

- Route all cell mutations through invariant-preserving model commands.
- Make spill replacement and obstruction atomic across all mutation paths.
- Introduce a common evaluation value model for scalar, array, reference, error, lambda, and
  omitted arguments.
- Centralize scalar, aggregate, array, and reference coercion.
- Complete `INDEX`, `INDIRECT`, `OFFSET`, `TRIMRANGE`, names, multi-area, 3D, `@`, and
  dynamic-to-dynamic `#` dependency semantics.
- Add dependency invalidation and cycle handling; validate dynamic-array extension metadata and
  save/reopen behavior against the pinned Excel Oracle profile.
- Revalidate the existing 17 array functions before adding higher-order functions.

Exit criteria for every supported function:

`value + shape + reference + error + recalculation + mutation + save/reopen + Excel agreement`

### M5 — Pivot Preserve-Only

Status: in progress on explicit package inventory and preservation gates; real Excel reopen evidence
remains an external M1 dependency.

Completed slices:

- Inventoried pivot table definitions, cache definitions and records, slicers and slicer caches,
  timelines and timeline caches by content type and relationship type.
- Followed each known seed's internal outgoing relationship closure as `OpaqueRelated`, retained
  external targets, and recorded workbook/worksheet incoming edges so shared-cache ownership is
  visible without traversing unrelated owner graphs.
- Snapshotted part bytes, current content type overrides, compression, owner `.rels` bytes, and
  normalized relationship identity; save validates the inventory before mutation and again before
  serialization.
- Proved synthetic clean save and unrelated cell edits retain the inventory, while changed parts,
  changed outgoing relationships, and dangling internal targets fail explicitly.
- Added runtime preflight for the `Worksheet` and sheet-collection OM surfaces: rename, delete, and
  copy are rejected when either involved workbook owns a preserved pivot graph; move to a new or
  different workbook is rejected before allocation or mutation, while same-workbook reorder is
  retained and passes save/reopen preservation.

Next slices:

- Extend the conservative preflight to chart-driven indirect sheet lifecycle paths before allowing
  them on pivot workbooks.
- Replace blanket rejection with owner-aware sheet rename/copy/delete behavior only after a real
  corpus proves shared-cache and relationship ownership semantics.
- Replace duplicate raw cache-record snapshots with a bounded digest or shared backing after the
  preservation contract is stable.

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
