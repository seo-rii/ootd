# OOTD Roadmap

This document contains the active roadmap. Detailed historical steps remain in `PLAN.md`.
Milestones close only when their exit criteria are met; adding API names or synthetic happy-path
tests alone does not complete a compatibility milestone.

## 2026-07-27 Audit Priority Override

The repository-wide compatibility audit identified persistence correctness and silent-success
contracts as release blockers. Until Audit Wave 1 closes, new editing breadth, typed pivot work,
and new chart families remain frozen even when an older milestone below lists them as a next slice.

Active order:

1. `OOTD-001` complete (2026-07-27): successful Save/SaveAs now commits the verified output as
   the next package/source baseline without replacing live model identities.
2. `OOTD-002` complete (2026-07-27): `Workbook.Saved` now controls an independent prompt state;
   it no longer clears serializable worksheet/name/chart/drawing deltas. This also lands the
   prompt-versus-serialization slice of `OOTD-046`; its remaining dirty domains stay open.
3. `OOTD-003` complete (2026-07-27): `Workbook.Save` now fails before serialization with a
   stable `InvalidState` error when no source path exists.
4. `OOTD-004` complete (2026-07-27): read-only `Save` cannot overwrite its source;
   `SaveAs`/`SaveCopyAs` accept only create-new targets, with `SaveAs` detaching to a writable
   source and `SaveCopyAs` preserving read-only identity.
5. `OOTD-005` complete (2026-07-27): `Workbook.Close` now follows a 48-case headless state table;
   explicit save requires a real target, prompt-required closes fail without invalidating the
   workbook, and read-only close-save accepts only a create-new Filename.
6. `OOTD-006`/`OOTD-045` complete (2026-07-27): all filesystem save APIs now use verified
   preparation, same-directory temporary write/flush/sync, atomic replace or create-new, and
   parent-directory sync before runtime commit; host-writer commit is success-gated too.
7. `OOTD-007` complete (2026-07-27): changed scalar formula results now mark their formula cells
   serialization-dirty, and Calculate → Save → reopen persists both formula text and cached value.
8. `OOTD-008` complete (2026-07-27): a public typed report classifies evaluated, unsupported,
   external, circular, volatile, and Excel-error formula cells by address while preserving
   unresolved cached values.
9. `OOTD-009` complete (2026-07-27): typed `calcPr` state now distinguishes complete,
   partial, and uncomputed caches; calculation-input digests prevent reuse after later edits,
   manual/automatic mode is synchronized, and every rewritten calculation state removes stale
   calc-chain artifacts.
10. `OOTD-046` complete (2026-07-27): typed prompt, semantic, serialization, formula-cache,
   package-graph, and external-refresh domains have a documented transition table. Semantic
   mutations and calculation writeback use separate markers, all save failure points preserve the
   complete snapshot, `SaveCopyAs` leaves it unchanged, and only a successful baseline commit
   clears it. External refresh remains inactive until OOTD-065 provides a supported backend.
11. `OOTD-010` complete (2026-07-27): `Workbook.SaveAs` accepts only implemented arguments or
   omission-equivalent defaults; ten unsupported option classes now fail before package
   preparation or file creation with stable diagnostics and unchanged runtime state.
12. `OOTD-011` complete (2026-07-27): `Workbooks.Open` now exposes a complete argument matrix;
   read-only is implemented, omitted/zero UpdateLinks is an explicit offline policy, and thirteen
   unsupported option classes fail before filesystem read.
13. `OOTD-012` complete (2026-07-27): codec load/save options now accept only the implemented
   Excel365 lossless-preservation policy; other profiles, unknown-part dropping, calc-chain
   skipping, and lossy save fail before parse/serialization.
14. `OOTD-013` complete (2026-07-27): Workbook, Worksheet, Chart, and sheet-collection refresh,
   spelling, fixed-format export, and print methods now retain argument/object validation but
   return stable `Unsupported` when no execution backend is configured; valid calls no longer
   return false `Empty` success or create output artifacts.
15. `OOTD-062` stage 1 complete (2026-07-27): bounded CFB v3/v4 FAT/directory discovery
   recognizes only containers with both required encrypted-OOXML streams and returns a dedicated
   `EncryptedWorkbookUnsupported` error before ZIP parsing. Legacy and partial compound files are
   not misclassified; Password arguments remain fail-closed. Agile Encryption remains open.
16. `OOTD-063` complete (2026-07-27): load exposes a source signature-artifact inventory and
   every codec/runtime/filesystem/host-writer rewrite refuses both linked and orphan OPC signature
   artifacts before output. Manual package replacement cannot bypass the source policy; signature
   verification, explicit strip+audit, and re-signing remain unsupported.
17. `OOTD-022` complete / `OOTD-064` stage 1 complete (2026-07-28): source/current package
   inventory covers VBA/XLM/dialog sheets, ActiveX/control properties, OLE/embedded packages, and
   custom UI by path, content type, and relationship. Same-format and macro-capable saves preserve
   artifacts; macro-enabled to non-macro conversion returns a dedicated fail-closed error before
   output. The Windows Oracle preflight blocks the same markers before Excel activation.
18. `OOTD-064` stage 2 complete (2026-07-28, synthetic): typed preserve/refuse/strip snapshot
   policies remove exclusive relationship closure, incoming XML anchors, active/orphan manifest
   entries, and macro/dialog sheets; retain shared descendants; and return a deterministic audit.
   Complex ActiveX/VML/OLE/custom UI real-Excel evidence remains required before Oracle-verified.
19. `OOTD-065` stage 1 complete (2026-07-28, synthetic): deterministic external-data inventory,
   typed offline-preserve/refuse open policy, access-attempt report, and unrelated-edit cache
   preservation are implemented.
20. `OOTD-041`/`OOTD-065` Oracle preflight stage complete (2026-07-28,
   synthetic/fake-backed): source and sandbox packages are audited before COM activation; known
   external-link, connection, query-table, and Data Model paths, content types, and relationships
   are refused while ordinary external hyperlinks remain allowed. **Active:** define the disabled
   host refresh/provider callback boundary and add pinned isolated-Excel evidence.
21. `OOTD-014` complete (2026-07-28, synthetic): workbook main-part and owner-relationship-part
   discovery now follows the single internal package-root `officeDocument` relationship. A
   nonstandard main URI survives no-op and targeted save/reopen, including relocated calc-chain
   invalidation and runtime format retagging; missing, external, dangling, and duplicate roots fail
   closed.
22. `OOTD-060`/`OOTD-020` complete (2026-07-28, synthetic): exact Strict/Transitional root,
   SpreadsheetML namespace, core relationship, and main-content-type tables now drive discovery
   and format detection. Strict same-dialect load/no-op/cell edit/save/reopen and runtime SaveAs are
   covered; mixed dialects and unknown content types fail closed. Cross-dialect conversion and
   Strict chart/drawing/structural mutation remain explicitly unsupported.
23. `OOTD-015`/`OOTD-049` stage 1 complete (2026-07-28, synthetic): workbook core, shared-string,
   and worksheet-cell parsing now uses namespace URI plus local name; arbitrary valid prefixes and
   relationship-attribute prefixes load correctly. Targeted rewrites inherit existing owner
   prefixes while foreign same-local elements and clean part bytes remain opaque and unchanged.
   Remaining parser families, structural validation, and Markup Compatibility stay open.
24. `OOTD-029` complete (2026-07-28, synthetic): `[Content_Types].xml` now requires the exact OPC
   root namespace/local name while accepting arbitrary bound prefixes. Only direct typed
   declarations affect resolution; duplicate defaults/overrides, missing required attributes,
   non-absolute or noncanonical override names, and nonempty declarations fail closed. Opaque
   extension subtrees cannot inject same-local declarations.
25. `OOTD-030` complete (2026-07-28, synthetic): `.rels` parsing now requires the exact package
   `Relationships` root expanded name and direct namespace-matching `Relationship` children while
   accepting arbitrary bound prefixes. Wrong/missing namespaces and foreign/nested same-local
   entries fail closed; required attributes, duplicate IDs, target modes, and normalized internal
   targets retain their strict contract.
26. `OOTD-016` complete (2026-07-28, synthetic): workbook sheet records now require a valid name,
   bounded nonzero sheet ID, and resolvable typed workbook relationship. Duplicate sheet IDs,
   ASCII-case-insensitive names, and relationship IDs fail closed; load/save validation no longer
   invents sequential IDs, `SheetN`, default kinds, or empty part URIs. Save rewrites may reconcile
   transient source-name lag by identity, but the final workbook part must pass the full strict
   parser before serialization. Implicit repair remains unsupported.
27. `OOTD-017` complete (2026-07-28, synthetic): malformed or zero row indices, invalid cell style
   lexicals, and row/cell address mismatches now fail closed with worksheet-part and cell context.
   Coordinate ownership is recorded before blank-cell elision, so duplicate empty or populated
   cells cannot overwrite each other. Numeric style-range diagnostics also identify their part and
   cell; implicit worksheet repair remains unsupported.
28. `OOTD-018` stage 1 complete (2026-07-29, synthetic): worksheet single-cell A1 parsing now uses
   an explicit ASCII grammar, checked base-26/base-10 accumulation, optional absolute markers, and
   the `XFD1048576` grid boundary. Malformed suffixes, qualification/range syntax, repeated markers,
   and long overflow inputs fail closed, while all 16,384 last-row column references round-trip.
   The remaining common range AST and `OOTD-048` consumer grammar migration stay open.
29. `OOTD-019` complete (2026-07-29, synthetic): `office_common::ExcelLimits` is the single source
   for `XFD1048576`; validated range/model commands and XLSX load/save preflight reject out-of-grid
   cell, dirty, and spill state. Checked `u64`/`usize` cardinality replaces range `u32` products,
   while `CountLarge` retains the full-grid value without materialization. Raw public DTO-field
   closure remains tracked by `OOTD-031`/`OOTD-054`; the contract is
   `docs/interfaces/excel_grid_limits.md`.
   `OOTD-031` stage 1 is complete (2026-07-29, synthetic): public package construction is fallible,
   canonicalizes root-relative names, and atomically rejects malformed or case/percent-equivalent
   duplicate identities. ZIP load, add, and serialization share the identity validator; the
   contract is `docs/interfaces/opc_package_construction.md`.
   `OOTD-031` stage 2 is complete (2026-07-29, synthetic): model load/save preflight rejects
   empty collections, duplicate or malformed model-internal worksheet identity, workbook ownership
   drift, partial/duplicate package bindings, missing worksheet data, malformed or dangling
   local-name scope, and inconsistent spill topology. The same gate precedes chart-graph
   materialization, while unbound chart-sheet records proceed to separate XLSX graph preflight.
   The contract is
   `docs/interfaces/workbook_state_save_validation.md`.
   `OOTD-031` stage 3 is complete (2026-07-29, synthetic): post-materialization save preflight
   compares the current discovered workbook package graph with model worksheet count/order/ID,
   relationship ID, normalized target and dialect-derived kind. Every target must exist and resolve
   to a uniquely owned canonical package part; name and visibility remain supported rewrites.
   `OOTD-031` stage 4 is complete (2026-07-29, synthetic): every recognized package-root or
   owner `.rels` part is parsed at the final Preserve/Refuse serialization boundary and again after
   active-content Strip. Normalized internal targets must resolve to canonical package parts;
   relationship-part recognition shares package lookup's case/percent canonical identity, external
   targets remain exempt, and Strip may clean orphan active markers before the final gate.
   `OOTD-031` stage 5 is complete (2026-07-29, synthetic): bounded manifest/cache validation now
   precedes active-content policy inventory, while final Preserve/Refuse and post-Strip gates require
   `[Content_Types].xml`, exact cache resolution, complete part coverage, and non-orphan canonical
   Overrides, while an Override cannot target the manifest itself. Default extension lookup uses
   only the canonical final path segment, Strip cleanup shares the package's case/percent canonical
   identity, relocated Transitional/Strict main parts remain protected during Strip, and clearing
   public preservation snapshots cannot bypass the gate.
   `OOTD-031` stage 6 is complete (2026-07-29, synthetic): one canonical relationship-part owner
   derivation now serves final closure and active-content traversal. Every non-root relationship
   part requires an existing non-relationship owner, while canonical root aliases stay exempt;
   empty/manifest/relationship owners, single-segment `.rels`, nested reserved `_rels` directories,
   malformed placement, and misplaced relationships MIME fail closed. Strip can remove root-level
   or percent-aliased active edges before the same final gate.
   `TEST-003` is complete (2026-07-29, synthetic): generic OPC serialization assigns every ZIP
   entry the canonical DOS epoch instead of the wall clock, making repeated package bytes and
   pinned artifact hashes stable across two-second timestamp boundaries.
   `OOTD-031` stage 7 is complete (2026-07-29, synthetic): public graph materialization and direct
   XLSX save validate fully materialized chart/drawing identity and workbook/host ownership against
   the actual serialization package. Canonical raw parts have one typed owner, host-sheet drawing
   relationships and drawing-to-chart bindings must resolve, drawingless package-bound chart sheets
   remain preserve-only, and loaded shared materialized charts remain valid while new shared graph
   allocation stays unsupported. The contract is
   `docs/interfaces/chart_drawing_graph_validation.md`.
   `OOTD-031` stage 8 is complete (2026-07-29, synthetic): codec load, public graph materialization,
   and direct save validate worksheet/drawing snapshot keys against live sheet owners, bind host and
   owner-relative relationship parts, and require canonical inventory, source-map, retained-summary,
   and internal-target coherence. Newly materialized graphs need no historical snapshot, while
   invalidated drawing source/summary and relationship-source subsets remain valid. The contract is
   `docs/interfaces/support_snapshot_validation.md`.
30. `OOTD-032` is complete (2026-07-29, synthetic): cell/source-XML mutation and worksheet-data
   access require a live worksheet owner and fail without state changes for unknown IDs. The
   auto-creating default-entry helper is gone, pre-seeded orphan entries are not mutable through the
   accessor, and save preflight rejects any extra data key introduced through public fields. The
   contract is `docs/interfaces/workbook_state_save_validation.md`.
31. `OOTD-033` is complete (2026-07-29, synthetic): workbook-ID reassignment prepares a rebound
   chart map, validates direct and full-reference ranges, and only then commits model, worksheet,
   chart, drawing, and chart-frame identities. Malformed deserialized ranges return contextual
   `InvalidState` without partial mutation; runtime registration consumes a handle only after
   success, and reload/save callers propagate failure. The contract is
   `docs/interfaces/workbook_state_save_validation.md`.
32. `OOTD-054` stages 1-19 are complete (2026-08-01, synthetic): the worksheet-data ownership map
   and live workbook-model metadata are private, decoded construction is validated, and runtime
   add/copy/delete use paired owner/data commands. Read-only access cannot insert or rekey orphan
   data; model metadata changes use explicit commands while workbook-ID changes remain atomic; and
   chart materialization fails rather than inventing a missing entry. Format metadata changes stay
   inside the broader package-retag transaction. Worksheet rename/visibility, chart-sheet package
   binding, and ordering now use validated commands; partial or retargeted chart bindings fail
   closed, and runtime reorder preflights the exact identity permutation before XML replacement. The
   live worksheet collection is private and exposes only an immutable slice; by-value
   reconstruction must pass through the validated parts constructor, with no mutable or unchecked
   access path. Runtime worksheet rename prepares all fallible defined-name and chart-source
   retargets on cloned substate and commits the live worksheet/name/chart state only after they
   succeed, so rejected renames preserve the complete model and dirty domains. Placement-target
   sheet-block Copy now prepares the entire batch in a cloned target runtime and restores package,
   support state, dirty domains, runtime object registries, handle allocators, and selection state
   after any later-sheet failure. Worksheet Add now uses the same runtime mutation snapshot across
   Count, template workbook, native chart/dialog/macro package graph creation, calc-chain
   invalidation, and handle/selection updates, so a late XML failure leaves no owner or orphan
   package residue. Individual worksheet/chart-sheet Delete uses that snapshot from first live
   owner removal through relationship/content-type/package rewrites, calc-chain invalidation, and
   stale-handle/selection cleanup; malformed content-types failure therefore restores model,
   package, support/pending graphs, dirty domains, and runtime session state. Collection Delete now
   adds an outer snapshot around the complete validated sheet block, so a later sheet failure also
   restores earlier successful deletions. Target-less sheet-collection Copy likewise starts an
   outer source-anchored snapshot before creating the destination workbook; a later copy failure
   removes that unpublished workbook and restores registries, allocators, active state, and
   selection. Single-area and multi-area `Range.Clear` now share a model command that preflights
   every spill child before atomically clearing cell/style, owned spill metadata, and dirty
   tracking. `Range.ClearFormats` uses a separate style-only command that permits spill children
   while preserving blank materialized child cells and all spill topology. Single-area and
   multi-area `Range.Replace` now prepare one immutable-snapshot replacement batch and commit it
   through a model command only after every changed coordinate passes spill-child validation.
   Replacing a dynamic anchor formula clears its old materialized extent while retaining the
   dynamic formula kind for the next calculation cycle. Row- and column-oriented `Range.Sort`
   likewise prepare one destination map from the immutable source and commit it through a
   rearrangement command that refuses every changed spill anchor or child before applying any
   plain-cell move. All four `Range.Fill` directions now simulate single- or multi-area operations
   in input order on a destination-only overlay, preserving overlap and formula-shift semantics
   before one model commit. The Fill command preflights every source and destination against spill
   topology, even when the resulting cell is unchanged, and creates no temporary per-area handles.
   Single-area `Range.Copy Destination` now snapshots source cells, refuses any source spill
   anchor/child, prepares formula-shifted destination replacements without mutation, and commits
   them through one destination command only after every target—including unchanged targets—passes
   spill-topology validation. Same- and cross-workbook failure preserves both workbook and session
   state; the existing styled-blank and A1-shift behavior remains intact.
   Single-area `Range.Cut Destination` now snapshots the complete source, builds destination and
   source-clear maps without live mutation, and applies every fallible spill-aware command to cloned
   workbook state before committing either owner. Same-sheet overlap keeps destination values,
   cross-sheet/workbook failures preserve every touched workbook and session state, and moved
   formulas remain exact rather than copy-shifted. Default all-like Cut-mode `PasteSpecial` delegates
   to this transaction and releases its internal temporary Range handles on both success and error.
   Cell-materializing custom `Range.PasteSpecial` now validates clipboard mode, writable owners,
   and the complete source spill topology before planning destination cells from immutable views.
   Late arithmetic/type errors cannot publish an earlier cell; Copy commits one destination command,
   while same- and cross-workbook Cut complete destination and style-preserving source-clear
   commands on cloned state before replacing either live owner. Every non-skipped destination is
   spill-preflighted even when unchanged, while `SkipBlanks` can leave a protected spill cell
   untouched. Same-sheet overlap, Copy-only A1 formula shifting, exact Cut formulas, and session
   cleanup semantics remain fixed by regressions.
   `Range.Insert` and `Range.Delete` now stage all four directional cell-payload shifts from one
   immutable sparse snapshot and publish one model command. Insert validates every shifted target
   before commit, and both operations reject any geometric intersection between the full shift
   corridor and dynamic spill topology, including an unmaterialized child. Formula/name/table/chart/
   drawing references and raw row/column metadata are not yet retargeted, so this remains a bounded
   atomicity contract rather than complete Excel structural-edit parity.
   The contract is
   `docs/interfaces/workbook_state_save_validation.md`.
   **Queued:** make the complete `WorksheetData` cell/spill/dirty payload private before building
   the common reference AST and completing `OOTD-048` consumer migration.
33. `BUG-004` is complete (2026-08-01, synthetic): metadata-only `Range.PasteSpecial` selectors
   `xlPasteComments`, `xlPasteValidation`, and `xlPasteColumnWidths` now return a named stable
   `Unsupported` before clipboard or owner-state handling. Copy/Cut and writable/read-only
   destination regressions preserve both workbook snapshots, dirty domains, and the complete
   Find/CutCopyMode/clipboard session.
34. `BUG-005` is complete (2026-08-01, synthetic): format/metadata-only `Chart.Paste` selectors
   `xlPasteFormats`, `xlPasteComments`, `xlPasteValidation`, and `xlPasteColumnWidths` return a
   named stable `Unsupported` before clipboard or owner-state handling. Copy/Cut and
   writable/read-only chart-destination regressions preserve source/destination workbook and chart
   state, dirty domains, object registries/allocators, and Find/CutCopyMode/clipboard session; the
   no-clipboard lane pins selector-first capability reporting.
35. `BUG-006` is complete (2026-08-01, synthetic): the `Chart.Paste xlPasteValues` branch opens a
   runtime workbook transaction before replacing series topology and rolls back chart state, dirty
   domains, stale/object handle registries and allocators, plus Find/CutCopyMode/clipboard when the
   later `Series.Values` dispatch rejects a non-finite source. Same- and cross-workbook Copy
   regressions retain the original Series handle and both persistence snapshots.
36. `BUG-007` is complete (2026-08-08, synthetic): values and formula `Chart.Paste` reuse the
   caller-visible chart handle and scope internal Range/SeriesCollection/Series handles to the
   call's allocator boundary. Success preserves only the intended stale transition for replaced
   caller-owned Series handles without registry or allocator growth; cross-workbook formula failure
   preserves both workbook snapshots and the complete session.
37. `BUG-008` is complete (2026-08-09, synthetic): `Chart.Paste` rejects a Cut range clipboard with
   stable `Unsupported` before checking destination mutability or changing chart/workbook/session
   state. Same-workbook all/formulas/values and cross-workbook read-only destination regressions
   preserve both workbook snapshots, dirty domains, object registry/allocator, and the complete
   Find/CutCopyMode/clipboard session.
38. `BUG-009` is complete (2026-08-09, synthetic): five all-like/number-format `Chart.Paste`
   selectors that previously discarded their format meaning now return named stable `Unsupported`
   before clipboard or owner-state handling. Together with the four metadata-only selectors, the
   Copy/Cut, writable/read-only, and no-clipboard matrix preserves both workbook snapshots, dirty
   domains, and the complete runtime session.
39. `OOTD-023` is complete (2026-08-09, synthetic): non-finite numbers are rejected by common
   cell coercion, direct/batch model mutations, save preflight, worksheet parse/rewrite, and
   `Range.PasteSpecial` arithmetic-result planning. NaN/positive/negative infinity and a late
   multiplication-overflow regression preserve workbook, dirty, and session state before any live
   commit.
40. `ARCH-024` formula-owner stage 1 is complete (2026-08-09, synthetic): a common bounded A1
   lexical detector skips quoted strings and non-reference tokens, while `Range.Insert`/`Delete`
   rejects any workbook-owned A1-reference or R1C1 formula before corridor planning and live
   mutation. Insert/Delete atomic regressions preserve workbook, dirty, and session snapshots;
   reference-free formula payload shifts still save/reopen.
41. `ARCH-024` defined-name owner stage 2 is complete (2026-08-09, synthetic): workbook- and
   worksheet-scoped A1-reference or R1C1 names now return a scope/name-bearing stable `Unsupported`
   before structural mutation, while A1-family reference-free constants survive Insert and
   save/reopen.
42. `ARCH-024` merged-cell owner stage 3a is complete (2026-08-09, synthetic): QName-aware load
   inventories bounded merged ranges, malformed references fail closed, and only a shift corridor
   intersecting a merge is rejected atomically with sheet/range diagnostics. Non-intersecting
   Insert preserves the merge through save/reopen.
43. `ARCH-024` standard data-validation owner stage 3b is complete (2026-08-09, synthetic):
   QName-aware direct-parent load expands bounded multi-area `sqref`, rejects malformed owner
   ranges, and atomically refuses only intersecting Insert/Delete corridors. Non-intersecting
   Insert preserves the validation XML and inventory through save/reopen.
44. `ARCH-024` standard data-validation formula-owner stage 3c is complete (2026-08-09,
   synthetic): QName-aware `formula1`/`formula2` text and CDATA are inventoried in source order;
   reference-bearing formulas refuse structural mutation workbook-wide before corridor planning,
   while reference-free formulas preserve the non-intersecting save/reopen path.
45. `ARCH-024` x14 data-validation owner stage 3d is complete (2026-08-09, synthetic): the exact
   worksheet `extLst`/`ext`/x14/xm owner path inventories bounded multi-area `xm:sqref` and nested
   `xm:f` formulas, rejects malformed or ambiguous owners, and reuses the standard range/formula
   structural preflight. A reference-free non-intersecting Insert preserves the x14 XML and
   inventory through save/reopen.
46. `ARCH-024` table relationship-owner stage 3e is complete (2026-08-09, synthetic): exact
   dialect-aware worksheet `tableParts`/`tablePart@r:id` owners are inventoried by QName and direct
   parent, malformed or ambiguous markers fail load, and any loaded table marker refuses structural
   mutation workbook-wide before table-part range and structured-formula semantics are modeled.
47. `ARCH-024` relationship-bound table owner stage 3f is complete (2026-08-09, synthetic): each
   worksheet marker resolves through an exact dialect table relationship to an internal, existing,
   correctly typed SpreadsheetML table part. Bounded `table@ref` and direct calculated/totals formulas
   become typed owners; malformed bindings fail load/save, A1-bearing formulas and intersecting table
   ranges refuse structural mutation atomically, and a non-intersecting reference-free table survives
   Insert plus save/reopen.
48. `ARCH-024` raw row/column metadata owner stage 3g is complete (2026-08-09, synthetic): exact
   direct row attributes/opaque children and bounded direct column ranges become full-axis owners;
   malformed column ranges and duplicate direct rows fail load, intersecting shift corridors refuse
   atomically, and a non-intersecting Insert preserves raw XML plus inventory through save/reopen.
49. `ARCH-024` chart-source owner stage 3h is complete (2026-08-09, synthetic): every series
   name/x-values/values/bubble-size source and full reference is checked as a typed workbook range;
   intersecting corridors, unresolved A1/R1C1 sources, invalid ownership, and 3D ranges fail before
   mutation, while a non-intersecting Insert preserves chart formulas through save/reopen.
50. `ARCH-024` drawing-anchor owner stage 3i is complete (2026-08-10, synthetic): typed one-cell
   and two-cell markers become bounded worksheet ranges, intersecting corridors refuse atomically,
   absolute/free-floating anchors remain eligible, and opaque/unresolved anchors fail closed on the
   host sheet. A non-intersecting Insert preserves the two-cell XML and typed anchor through
   save/reopen.
51. `OOTD-043`/`OOTD-085` replay foundation stage 1 is complete (2026-08-10, synthetic): a bounded
   filesystem loader resolves the fixed suite/run manifests, refuses symlinked, non-regular,
   oversized, or non-portable artifacts, and checks exact case/input/observation hashes before
   replay. No desktop Excel observation is claimed.
52. `OOTD-043`/`OOTD-085` repeated-capture foundation stage 2 is complete (2026-08-10,
   synthetic): two case-insensitively distinct run IDs, the suite's exact Excel fingerprint,
   complete required cases, stable statuses, and exact canonical typed observations are required
   before an evidence receipt is returned.
53. `OOTD-043`/`OOTD-085` suite-run assembly foundation stage 3 is complete (2026-08-10,
   synthetic): validated case-subset fragments sharing one exact run/profile/engine are assembled
   in suite order with canonical observation paths; missing, duplicate, cross-run, and tampered
   fragments fail closed.
54. `OOTD-043`/`OOTD-085` atomic materialization foundation stage 4 is complete (2026-08-10,
   synthetic): a fresh sibling temporary root receives create-new/synced observations and the
   manifest before one directory rename; existing destinations and tampered bundles leave no
   partial output.
55. **Active:** expose a bounded suite capture command around the Windows case watchdog, then
   collect two independent runs on the pinned desktop Excel host.

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
