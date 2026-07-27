# OOTD Status

This document is the current source of truth for implementation and verification status.
`PLAN.md` remains the historical implementation log, while `ROADMAP.md` defines active work.

Baseline date: 2026-07-26.

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
| OPC package loading | Partial | ZIP parts and opaque bytes are retained; default loading enforces finite archive/entry/name/decompression/ratio budgets with pre-central-directory entry-count checks; part identity validation remains narrow | M3 canonical part/relationship validation and shared XML budgets |
| Workbook and worksheet model | Partial | Workbook, sheet, cell, name, chart, drawing, and basic dynamic-array state are modeled | Oracle-backed mutation and save/reopen cases |
| XLSX load/save | Partial | No-op and targeted dirty-save preservation have broad synthetic regression coverage | Tracked real-world corpus, bounded parsing, and Excel reopen without repair |
| Runtime object model | Partial | Application, workbook, worksheet, range, names, selection, clipboard, and chart-related dispatch are available | Generated member coverage and behavioral Oracle cases |
| Scalar formula calculation | Partial | Broad deterministic function coverage exists behind an internal `calc` module, including Evaluate and Calculate paths; its value/coercion model is not yet unified | Shared coercion/reference model and Excel differential corpus |
| Formula2 and dynamic arrays | Partial | Seventeen array functions produce two-dimensional spill results with basic obstruction and recalculation handling | Full spill mutation lifecycle, `@`/`#`, dependency order, XLSX metadata, and Oracle agreement |
| Charts and drawings | Partial | Typed chart mutation and lossless-first relationship graph lifecycle cover a broad surface | Remain feature-frozen until PivotChart work; fix preservation regressions only |
| Styles and themes | Preserve-only | Raw bytes and typed summaries are retained; general typed style allocation and mutation are incomplete | Corpus preservation before broader typed editing |
| Macros and unsupported package parts | Preserve-only | OOXML macro-bearing variants and opaque parts are retained within the covered save paths | Real corpus and explicit capability matrix |
| Pivot tables, caches, slicers, and timelines | Unsupported | Generic opaque preservation may retain parts, but no pivot-specific inventory or lifecycle guarantee is claimed | M5 preserve-only inventory and ownership tests |

## Verification Baseline

- Rust MSRV: 1.88; development toolchain: 1.94.0.
- Linux workspace tests: enabled in CI.
- Current root test inventory: 677 `excel-runtime` tests and 2,826 `excel-xlsx` tests.
- M2 boundary progress: the `excel-xlsx` and `excel-runtime` unit tests now live outside their
  library roots with test identities unchanged; calculation and recalculation/writeback are
  isolated; shared strings, relationships, and worksheet cell codec logic are isolated; Application,
  Workbook/Workbooks, WorksheetFunction, and Worksheet/sheet-collection dispatch are grouped by
  object surface; Names/Name, Range/Areas, and the chart-family helper surface are also isolated;
  inline public-router arms remain as the final M2 dispatch debt.
- M3 input safety: default OPC ZIP loading is resource-bounded; canonical part names,
  relationship validation, and shared XML budgets remain in progress.
- Formatting, strict Clippy, MSRV, and Windows jobs: scheduled for M3.
- Behavioral Oracle foundation: Rust and .NET contracts, runtime adapter, differential gate bridge,
  COM runner, and watchdog are implemented and synthetic/fake-backed tests pass.
- Real Excel behavioral cases: none pinned yet; the current Linux host cannot execute desktop Excel.
- Tracked corpus/golden XLSX fixtures: none yet; synthetic workbooks are generated inside tests.

## Stability

All crates are pre-1.0 and `publish = false`. Public APIs and serialized internal state may
change while M0-M5 establishes compatibility evidence and stable internal boundaries.
