# OOTD Status

This document is the current source of truth for implementation and verification status.
`PLAN.md` remains the historical implementation log, while `ROADMAP.md` defines active work.

Baseline date: 2026-07-27.

## Status Vocabulary

- **Implemented**: the stated bounded surface has model, runtime, codec, and regression coverage.
- **Partial**: useful behavior exists, but important lifecycle or compatibility semantics remain.
- **Preserve-only**: existing package data is retained where covered, without a general typed mutation API.
- **Unsupported**: no supported behavior is claimed; an explicit error is preferred over approximation.
- **Oracle-verified**: a supported behavior matches a pinned run from real desktop Excel.

No OOTD behavior is Oracle-verified yet. Existing compatibility evidence is synthetic and
contract-based until Milestone M1 pins the first behavioral Excel corpus.

## Compatibility Surface

| Area | Status | Current boundary | Next verification gate |
|---|---|---|---|
| Office OM source intake | Partial | Schema, template, capture planning, Windows launcher, receipt, and normalization paths exist; the real TypeLib/PIA bundle is not pinned | Real Windows capture bundle with build, channel, architecture, and locale |
| Behavioral Excel oracle | Partial | Typed cases, exact-byte run manifests, comparison/gate bridge, `ExcelRuntime` adapter, .NET contract tests, and an isolated COM runner/watchdog exist; no real Excel observation is pinned | Execute twice on the pinned Windows/Excel profile and commit the first required corpus |
| OPC package loading | Partial | ZIP parts and opaque bytes are retained; default loading enforces finite ZIP and XML depth/event/text/attribute budgets; canonical part identities and strict relationship parsing reject ambiguous duplicates, root escapes, invalid modes, malformed URIs, and malformed XML | M3 CI portability gates, property/fuzz coverage, and dependency policy |
| Workbook and worksheet model | Partial | Workbook, sheet, cell, name, chart, drawing, and basic dynamic-array state are modeled | Oracle-backed mutation and save/reopen cases |
| XLSX load/save | Partial | No-op and targeted dirty-save preservation have broad synthetic regression coverage; filesystem saves use verified preparation, same-directory durable temporary files, atomic replace/create-new, and post-write baseline commit; read-only Save cannot overwrite its source | Tracked real-world corpus and Excel reopen without repair |
| Runtime object model | Partial | Application, workbook, worksheet, range, names, selection, clipboard, and chart-related dispatch are available; `Workbook.Saved` uses prompt-only state, pathless `Workbook.Save` fails closed, and `Workbook.Close` has a deterministic headless save/discard/prompt state table | Complete dirty-domain taxonomy, generated member coverage, and behavioral Oracle cases |
| Scalar formula calculation | Partial | Broad deterministic function coverage exists behind an internal `calc` module, including Evaluate and Calculate paths; its value/coercion model is not yet unified | Shared coercion/reference model and Excel differential corpus |
| Formula2 and dynamic arrays | Partial | Seventeen array functions produce two-dimensional spill results; model value, A1/R1C1 formula families, and `ClearContents` commands reject spill-child batches atomically; worksheet array formula metadata restores and writes spill state across synthetic save/reopen; A1 `anchor#` resolves a materialized extent, and scalar dependents recalculate after dynamic materialization | Remaining runtime mutation paths, `@`, dynamic-to-dynamic dependency order/cycles, Excel-specific dynamic-array extension metadata, and Oracle agreement |
| Charts and drawings | Partial | Typed chart mutation and lossless-first relationship graph lifecycle cover a broad surface | Remain feature-frozen until PivotChart work; fix preservation regressions only |
| Styles and themes | Preserve-only | Raw bytes and typed summaries are retained; general typed style allocation and mutation are incomplete | Corpus preservation before broader typed editing |
| Macros and unsupported package parts | Preserve-only | OOXML macro-bearing variants and opaque parts are retained within the covered save paths | Real corpus and explicit capability matrix |
| Pivot tables, caches, slicers, and timelines | Preserve-only | Known pivot/cache/slicer/timeline parts, related opaque closure, external targets, shared-cache incoming edges, raw bytes, content types, compression, and owner relationships are inventoried; clean and unrelated-cell saves are guarded; direct Worksheet and sheet-collection rename/copy/delete and cross-workbook move fail before mutation while same-workbook reorder is retained | Chart-driven indirect lifecycle preflight, tracked real corpus, and Excel reopen without repair |

## Verification Baseline

- Rust MSRV: 1.88; development toolchain: 1.94.0.
- Linux workspace tests: enabled in CI.
- Current root test inventory: 704 `excel-runtime` tests and 2,838 `excel-xlsx` tests.
- M2 boundary progress: the `excel-xlsx` and `excel-runtime` unit tests now live outside their
  library roots with test identities unchanged; calculation and recalculation/writeback are
  isolated; shared strings, relationships, and worksheet cell codec logic are isolated; Application,
  Workbook/Workbooks, WorksheetFunction, and Worksheet/sheet-collection dispatch are grouped by
  object surface; Names/Name, Range/Areas, and the chart-family helper surface are also isolated;
  inline public-router arms remain as the final M2 dispatch debt.
- M3 input safety: default OPC ZIP loading is resource-bounded; canonical part identities are
  enforced across load/mutation/save; relationship attributes, IDs, target modes, and internal
  targets fail closed; XML-bearing parts receive a shared bounded well-formedness preflight.
- CI portability: Ubuntu Rust 1.94, Ubuntu MSRV Rust 1.88, and Windows Rust 1.94 run as independent
  test lanes. A bounded rustfmt gate covers 41 tracked files with four guarded monolith exceptions;
  strict Clippy is enforced for the six foundational/model crates, while runtime/XLSX warnings
  remain staged M3 debt.
- M4 spill lifecycle: model value, A1/R1C1 formula families, and `ClearContents` commands preflight
  all targets and reject spill children without partial mutation; single- and multi-area R1C1
  runtime dispatch now uses the same model command, and anchor edits atomically clear their current
  owned extent. Worksheet `t="array"`/`ref` metadata restores spill ranges on load, is emitted for
  newly calculated Formula2 extents, and is removed when an anchor becomes an ordinary formula;
  real Excel dynamic-array extension metadata remains Oracle-gated.
- Persistence dirty-state boundary: `Workbook.Saved` now changes only the prompt-facing dirty
  state. Serialization flags remain intact until a successful save commits its verified baseline,
  and runtime range mutations propagate prompt state only when model/package content actually
  changes.
- Save target boundary: `Workbook.Save` requires an existing source path and returns a stable
  `InvalidState` error before serialization when none exists, leaving both clean and dirty
  workbooks open and unchanged so callers must choose `SaveAs`.
- Read-only save boundary: `Workbook.Save` never overwrites a read-only source. `SaveAs` and
  `SaveCopyAs` use create-new targets for read-only workbooks; copy preserves the original
  read-only identity, while SaveAs detaches the open workbook to the new writable source.
- Close lifecycle boundary: all 48 combinations of prompt-dirty/source/SaveChanges/Filename/
  DisplayAlerts are regression-tested. Explicit save without a target and prompt-required
  headless close fail while leaving the workbook open; explicit discard and alerts-disabled
  omitted close are deterministic, and read-only close-save requires a create-new Filename.
- Durable save transaction: Save, SaveAs, SaveCopyAs, and Close(save) prepare and verify output,
  write a permission-preserving temporary file in the target directory, flush and sync it, then
  atomically replace or create the target and sync the parent before committing runtime state.
  Pre-replace fault injection preserves original bytes and dirty state; post-replace sync failure
  leaves a valid output and retryable dirty runtime. Host writers commit the baseline only after
  write and flush succeed.
- M5 pivot preservation: the codec inventories seven known pivot package kinds plus their internal
  opaque closure, incoming/shared and outgoing/external relationships, content types, compression,
  and raw bytes. Save-time gates protect clean and unrelated-cell edits and reject drift or dangling
  internal targets. Runtime preflight rejects unsafe direct Worksheet and sheet-collection
  lifecycle operations atomically and preserves same-workbook reorder across save/reopen; indirect
  chart-driven lifecycle, typed pivot mutation, and real Excel reopen evidence remain out of scope.
- Behavioral Oracle foundation: Rust and .NET contracts, runtime adapter, differential gate bridge,
  COM runner, and watchdog are implemented and synthetic/fake-backed tests pass.
- Real Excel behavioral cases: none pinned yet; the current Linux host cannot execute desktop Excel.
- Tracked corpus/golden XLSX fixtures: none yet; synthetic workbooks are generated inside tests.

## Stability

All crates are pre-1.0 and `publish = false`. Public APIs and serialized internal state may
change while M0-M5 establishes compatibility evidence and stable internal boundaries.
