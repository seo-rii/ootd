# OOTD Status

This document is the current source of truth for implementation and verification status.
`PLAN.md` remains the historical implementation log, while `ROADMAP.md` defines active work.

Baseline date: 2026-07-28.

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
| OPC package loading | Partial | ZIP parts and opaque bytes are retained; default loading enforces finite ZIP and XML depth/event/text/attribute budgets; canonical part identities reject ambiguous duplicates and malformed URIs; `[Content_Types].xml` requires the exact package root expanded name, resolves prefixed direct declarations, rejects missing required attributes and case/canonical duplicates, and prevents opaque nested same-local nodes from changing typed resolution; `.rels` parsing likewise requires the exact package root namespace and direct typed children, rejects foreign/nested same-local edges, duplicate IDs, root escapes, invalid modes, and malformed XML, and accepts arbitrary bound prefixes | Strict-versus-repair parser separation, M3 property/fuzz coverage, and dependency policy |
| Workbook and worksheet model | Partial | Workbook, sheet, cell, name, chart, drawing, and basic dynamic-array state are modeled | Oracle-backed mutation and save/reopen cases |
| XLSX load/save | Partial | No-op and targeted dirty-save preservation have broad synthetic regression coverage; package-root `officeDocument` discovery supports nonstandard workbook/owner-relationships URIs for exact Transitional and Strict dialects; Strict main namespaces, core relationship types, both accepted main content types, relocated shared strings/calc chain, same-dialect cell edits, and runtime SaveAs are covered, while mixed/unknown dialect signals fail closed; workbook core, shared-string, and worksheet-cell parsing matches namespace URI plus local name, accepts arbitrary valid SpreadsheetML/relationship prefixes, preserves foreign same-local nodes, and inherits source prefixes for generated typed nodes; workbook sheet records require valid names, bounded unique IDs, and unique resolvable typed relationships instead of inventing fallback identities or empty part URIs; filesystem saves use verified preparation, same-directory durable temporary files, atomic replace/create-new, and post-write baseline commit; calculation-state rewrites update `calcPr` and remove stale calc chains; read-only Save cannot overwrite its source; codec options fail closed outside the implemented Excel365 lossless-preservation policy; encrypted OOXML CFB containers are detected before ZIP parsing; signed packages are readable but every rewrite is rejected before output; active-content paths, content types, and relationships are inventoried; typed preserve/refuse/strip snapshot policies include ownership-aware closure cleanup and deterministic audit output; external workbook/DDE/OLE/connection/query/data-model markers have a deterministic package inventory and cached parts survive unrelated edits; the Windows Oracle refuses known external-data markers and records source/copy preflight audits before COM activation | Remaining QName parser families, root/parent structure and package-metadata namespace validation, worksheet strict parsing, Markup Compatibility, Strict chart/drawing/structural mutation and cross-dialect conversion, tracked real-world corpus, Agile Encryption support, signed-package validation/strip policy, complex active-content Excel reopen evidence, external-data refresh/provider boundary, and repair-free Oracle output |
| Runtime object model | Partial | Application, workbook, worksheet, range, names, selection, clipboard, and chart-related dispatch are available; `Workbook.Saved` uses prompt-only state, typed workbook dirty domains have a command/save-failure transition contract, pathless `Workbook.Save` fails closed, `SaveAs` rejects unsupported options before write, `Workbooks.Open` implements read-only and rejects unsupported options before read, signed-package Save/SaveAs/SaveCopyAs/Close(save) paths fail before output, active-content macro-to-non-macro conversion fails before output by default, and an explicit snapshot API returns stripped bytes plus audit without mutating the open baseline; external data defaults to offline preserve, exposes a no-attempt report, and supports typed refuse-before-registration; backend-free refresh/spelling/fixed-format export/print methods return `Unsupported`; `Workbook.Close` has a deterministic headless state table; workbook calculation mode is synchronized with `Application.Calculation` | Remaining host callback/provider and Oracle isolation policies, generated member coverage, and behavioral Oracle cases |
| Scalar formula calculation | Partial | Broad deterministic function coverage exists behind an internal `calc` module; changed results are serialized as cached values, a public report classifies address-level outcomes without overwriting unresolved caches, and complete/partial/uncomputed states drive coherent `calcPr` metadata | Shared coercion/reference model and Excel differential corpus |
| Formula2 and dynamic arrays | Partial | Seventeen array functions produce two-dimensional spill results; model value, A1/R1C1 formula families, and `ClearContents` commands reject spill-child batches atomically; worksheet array formula metadata restores and writes spill state across synthetic save/reopen; A1 `anchor#` resolves a materialized extent, and scalar dependents recalculate after dynamic materialization | Remaining runtime mutation paths, `@`, dynamic-to-dynamic dependency order/cycles, Excel-specific dynamic-array extension metadata, and Oracle agreement |
| Charts and drawings | Partial | Typed chart mutation and lossless-first relationship graph lifecycle cover a broad surface | Remain feature-frozen until PivotChart work; fix preservation regressions only |
| Styles and themes | Preserve-only | Raw bytes and typed summaries are retained; general typed style allocation and mutation are incomplete | Corpus preservation before broader typed editing |
| Macros and unsupported package parts | Partial | OOXML macro-bearing variants and opaque parts are retained within covered unsigned same-format/macro-capable saves; VBA/XLM/dialog/ActiveX/OLE/custom UI markers are inventoried; macro-enabled to non-macro conversion fails closed instead of partially stripping VBA; explicit typed strip removes active roots, exclusive descendants, incoming anchors, relationship/content-type declarations, and returns an audit while retaining shared descendants; encrypted OOXML containers are identified without claiming their hidden inner package kind; OPC signature artifacts are inventoried but signed packages cannot be rewritten | Real Excel active-content corpus, complex VML/control cleanup evidence, Agile Encryption, and signature validation/strip |
| Pivot tables, caches, slicers, and timelines | Preserve-only | Known pivot/cache/slicer/timeline parts, related opaque closure, external targets, shared-cache incoming edges, raw bytes, content types, compression, and owner relationships are inventoried; clean and unrelated-cell saves are guarded; direct Worksheet and sheet-collection rename/copy/delete and cross-workbook move fail before mutation while same-workbook reorder is retained | Chart-driven indirect lifecycle preflight, tracked real corpus, and Excel reopen without repair |

## Verification Baseline

- Rust MSRV: 1.88; development toolchain: 1.94.0.
- Linux workspace tests: enabled in CI.
- Current root test inventory: 719 `excel-runtime` tests and 2,863 `excel-xlsx` tests.
- M2 boundary progress: the `excel-xlsx` and `excel-runtime` unit tests now live outside their
  library roots with test identities unchanged; calculation and recalculation/writeback are
  isolated; shared strings, relationships, and worksheet cell codec logic are isolated; Application,
  Workbook/Workbooks, WorksheetFunction, and Worksheet/sheet-collection dispatch are grouped by
  object surface; Names/Name, Range/Areas, and the chart-family helper surface are also isolated;
  inline public-router arms remain as the final M2 dispatch debt.
- M3 input safety: default OPC ZIP loading is resource-bounded; canonical part identities are
  enforced across load/mutation/save; relationship attributes, IDs, target modes, and internal
  targets fail closed; XML-bearing parts receive a shared bounded well-formedness preflight.
- OPC content-types boundary: `[Content_Types].xml` must have the package content-types `Types`
  expanded root name, while arbitrary bound prefixes are accepted. Only direct namespace-matching
  `Default` and `Override` children affect resolution. Required unqualified attributes, absolute
  canonical override names, case-insensitive default duplicates, canonical override duplicates,
  and empty typed content are enforced; unknown attributes/subtrees stay opaque and nested
  same-local poison declarations are ignored. The synthetic contract is documented in
  `docs/interfaces/opc_content_types.md`.
- OPC relationships boundary: parsed `.rels` parts must have the package relationships
  `Relationships` expanded root name and direct namespace-matching `Relationship` children;
  arbitrary bound prefixes are accepted. Wrong/missing root namespaces, foreign or nested
  same-local entries, and root text/CDATA fail closed. Required unqualified attributes, duplicate
  IDs across filtered external entries, target modes, and normalized internal targets retain their
  strict contract. Details are in `docs/interfaces/opc_relationships.md`.
- Workbook sheet-record boundary: each typed sheet requires a nonblank valid name, a unique
  nonzero unsigned 32-bit ID, and a unique relationship ID resolving to a supported internal sheet
  relationship. Names are unique under the runtime's ASCII-insensitive comparison. Missing,
  malformed, out-of-range, duplicate, or dangling declarations return `Parse`; load and save
  validation never invent IDs, `SheetN`, default kinds, or empty part URIs. An internal save rewrite
  can reconcile transient source-name lag by sheet identity, then the final workbook part must pass
  the full strict parser before serialization. Implicit repair remains unsupported. Details are in
  `docs/interfaces/workbook_sheet_records.md`.
- CI portability: Ubuntu Rust 1.94, Ubuntu MSRV Rust 1.88, and Windows Rust 1.94 run as independent
  test lanes. A bounded rustfmt gate covers 52 tracked files with four guarded monolith exceptions;
  strict Clippy is enforced for the six foundational/model crates, while runtime/XLSX warnings
  remain staged M3 debt.
- M4 spill lifecycle: model value, A1/R1C1 formula families, and `ClearContents` commands preflight
  all targets and reject spill children without partial mutation; single- and multi-area R1C1
  runtime dispatch now uses the same model command, and anchor edits atomically clear their current
  owned extent. Worksheet `t="array"`/`ref` metadata restores spill ranges on load, is emitted for
  newly calculated Formula2 extents, and is removed when an anchor becomes an ordinary formula;
  real Excel dynamic-array extension metadata remains Oracle-gated.
- Persistence dirty-state boundary: `Workbook.Saved` now changes only the prompt-facing dirty
  state. Public `WorkbookDirtyDomains` independently reports prompt, semantic, serialization,
  formula-cache, package-graph, and external-refresh state. Runtime mutations use a shared
  semantic marker, calculation writeback raises formula-cache/serialization state without
  changing prompt state, and calc-chain invalidation is visible as package-graph state. Only a
  successful verified baseline commit clears these domains and the consumed calculation snapshot.
  Filesystem and host-writer failures preserve the complete snapshot, while `SaveCopyAs` writes a
  copy without committing it. The command matrix is documented in
  `docs/interfaces/workbook_dirty_domains.md`; OOTD-065 inventory/reporting leaves external refresh
  inactive until an audited backend exists.
- Save target boundary: `Workbook.Save` requires an existing source path and returns a stable
  `InvalidState` error before serialization when none exists, leaving both clean and dirty
  workbooks open and unchanged so callers must choose `SaveAs`.
- SaveAs option boundary: `Password`, `WriteResPassword`, `ReadOnlyRecommended`, `CreateBackup`,
  `AccessMode`, `ConflictResolution`, `AddToMru`, `TextCodepage`, `TextVisualLayout`, and `Local`
  accept only omission-equivalent defaults. Unsupported values fail with stable diagnostics before
  package preparation or target creation and preserve source identity and all dirty domains.
- Open option boundary: `ReadOnly` is implemented and omitted/zero `UpdateLinks` is an explicit
  offline-preserve/no-update policy. External workbook, DDE/OLE, connection, query-table, and
  data-model markers are reported without contacting their targets; a typed `Refuse` policy fails
  before handle registration. Thirteen unsupported link/password/text-import/edit/notify/converter/
  MRU/locale/repair option classes fail with stable diagnostics before filesystem read.
- Codec option boundary: `LoadOptions` and `SaveOptions` accept only `Excel365`,
  unknown-part preservation, calc-chain inventory, and lossless save. Other profiles and
  destructive/skip/lossy modes return stable `Unsupported` before OPC parsing or serialization.
- OOXML dialect boundary: exact package-root relationship, workbook root namespace, and resolved
  main content type determine Transitional versus Strict. Known mixed-dialect relationships and
  unknown main types fail closed. Strict no-op, worksheet-cell edit, calc-chain invalidation, and
  same-format runtime save are synthetic preserve-only capabilities; Strict chart/drawing or
  worksheet-collection mutation and every Strict↔Transitional conversion return `Unsupported`.
- SpreadsheetML QName boundary: workbook core, shared-string, and worksheet-cell typed nodes are
  matched by the loaded dialect namespace URI plus local name. Arbitrary valid element prefixes,
  relationship-ID prefixes, and worksheet namespace redeclarations are accepted; generated dirty
  nodes inherit the nearest typed owner prefix while foreign same-local nodes and no-op source parts
  remain preserved. Broader parser migration, root/parent structure, Markup Compatibility, and
  real Excel evidence remain open. The current slice is documented in
  `docs/interfaces/spreadsheetml_qnames.md`.
- Execution backend boundary: correctly shaped refresh, spelling, fixed-format export, and print
  calls on Workbook, Worksheet, Chart, and sheet collections return stable `Unsupported` after
  object validation when no backend is configured. Malformed arguments retain their existing
  validation errors, and rejected calls create no output artifact or workbook state change;
  rejected refresh also records no external-access attempt.
- Encrypted OOXML boundary: a bounded CFB v3/v4 DIFAT/FAT/directory walk requires a root entry and
  both non-empty `EncryptionInfo` and `EncryptedPackage` streams before returning
  `EncryptedWorkbookUnsupported`. Legacy and partial compound containers are not misclassified;
  omitted/empty Password reaches this detection while non-empty Open/SaveAs Password remains
  rejected before read/prepare. Decryption, encryption, verifier, integrity, and inner-format
  identification remain unsupported.
- OPC digital-signature boundary: the codec inventories `_xmlsignatures` paths, signature
  content types, and origin/signature/certificate relationship markers while keeping the workbook
  readable. Any source or current artifact, including a content-type-only orphan, makes clean and
  dirty rewrites return `SignedPackageMutationUnsupported` before filesystem or host-writer output.
  The source inventory cannot be bypassed by manually replacing public package parts. Digest,
  certificate, and trust validation, explicit strip+audit, and re-signing remain unsupported.
- Active-content boundary: the codec inventories source/current VBA project/data/signature,
  XLM/dialog sheet, ActiveX/control property, OLE/embedded package, and custom UI markers from
  paths, content types, and relationships. Same-format and macro-capable saves preserve their part
  bytes and graph. XLSM/XLTM to XLSX/XLTX conversion returns
  `ActiveContentConversionUnsupported` before target creation instead of stripping only
  `vbaProject.bin`; the Windows Oracle rejects the same markers before Excel activation. Host code
  can explicitly select preserve/refuse/strip for snapshot serialization. Strip removes exclusive
  relationship closure, incoming XML anchors, active/orphan manifest entries, and macro/dialog sheet
  entries, retains shared descendants, and returns a deterministic audit without committing the
  open runtime baseline. Complex real-world ActiveX/VML/OLE/custom UI Oracle evidence remains open.
- External-data boundary: runtime opens preserve cached external workbook, DDE/OLE, connection,
  query-table, and Data Model artifacts without update, refresh, provider, filesystem-target, or
  network attempts; a typed refusal policy rejects markers before handle registration. The Windows
  Oracle applies a stricter refusal scan to the source and sandbox copy before COM construction and
  atomically records accepted/rejected preflight decisions. A host refresh callback and pinned
  isolated-Excel corpus remain unsupported.
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
- Formula cache persistence: scalar recalculation compares the evaluated value with the loaded
  cache and marks changed formula cells and their worksheet serialization-dirty without changing
  the prompt-facing `Workbook.Saved` state. A precedent-change → Calculate → Save → reopen
  regression verifies the new cached value and original formula text together.
- Calculation diagnostics: `calculate_workbook_with_report` returns deterministic one-based cell
  addresses for evaluated, unsupported, external-workbook, circular, volatile, and Excel-error
  outcomes. Unsupported and external formulas retain their previous cached values; volatile is an
  overlapping annotation, and unresolved categories make the report incomplete.
- Calculation metadata lifecycle: `calcPr` mode, source `calcId`, and cache-completion state are
  parsed into typed codec state. A complete workbook calculation records completed caches; partial
  or uncomputed inputs set `calcId=0` and force full recalculation on load. A SHA-256 digest of
  formula inputs prevents a completed snapshot from surviving later cell/name/sheet/date-system
  mutations. Rewritten calculation state removes calc-chain part/relationship/content-type
  artifacts, preserves unknown `calcPr` attributes, and inserts a missing element before later
  workbook extension children.
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
