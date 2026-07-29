# Workbook State Save Validation Contract

`WorkbookState::validate_for_save` is the model-level fail-closed boundary for topology that can
still be invalidated through public state fields. `XlsxCodec::load` validates the decoded state
before exposing it, and every lossless XLSX save validates the live state before package
serialization or graph materialization.

## Workbook Model Metadata

The live `WorkbookState` workbook model is private. Callers inspect it through `model()` and can
change individual metadata only through `set_display_name`, `set_date1904`, `set_is_addin`, and
`set_format`; each command returns whether the value changed. Workbook identity has no scalar
setter: `assign_workbook_id` remains the only supported identity command because it validates and
rebinds every owned model identity atomically.

The metadata commands do not own runtime dirty-state transitions. Runtime dispatch uses their
change result to mark semantic/serialization state and invalidate interaction state only when the
value actually changes. Reload, active-content strip, and successful save-baseline reconstruction
also use the commands instead of mutating metadata through a public field.

`set_format` is only the model-metadata leg of format retagging. Runtime keeps it as the final step
of the existing transaction that first validates conversion policy, rewrites package content types,
updates the detected format, and invalidates cached content-type summaries. Calling this command
alone does not convert an OOXML package; source/model/detected format unification remains tracked
by `OOTD-036`.

`WorkbookStateParts.model` remains a by-value construction field, but the resulting live state is
not exposed unless `WorkbookState::try_new` validates the complete parts DTO.

## Worksheet Collection

Each worksheet must satisfy all of the following:

- the workbook contains at least one worksheet record;
- its ID is nonzero, fits the supported unsigned 32-bit OOXML `sheetId` range, and is unique;
- its name is nonempty, at most 31 Unicode scalar values, free of Excel-forbidden/control
  characters, and unique under the runtime's ASCII-case-insensitive lookup policy;
- its `workbook_id` matches the owning `WorkbookModel`;
- nonempty relationship IDs and case-insensitive part URIs are unique;
- its relationship ID and part URI are both nonempty, except that an unbound chart-sheet record
  may omit both pending separate XLSX graph preflight; and
- a corresponding `WorksheetData` entry exists, while every data entry has exactly one worksheet
  owner.

The fully absent relationship/part pair is intentional. The model gate admits an unbound
chart-sheet record so the XLSX graph layer can separately validate and materialize the complete
chart/drawing binding. A partial pair is never serializable and returns
`OmErrorCode::InvalidState`.

The live worksheet vector is private. `worksheets()` exposes only an immutable slice, and there is
no mutable collection or element getter. `into_parts()` consumes the state when a caller needs to
transform its construction DTO; the result can become live state again only through
`WorkbookState::try_new`, which reruns the complete model invariant gate. This preserves package
mismatch tests without adding an unchecked mutation surface: tests reconstruct an internally valid
model and then verify that the separate XLSX package-to-model boundary rejects the mismatch.

### Worksheet metadata and ordering commands

Production mutation of worksheet metadata and order uses validated `WorkbookState` commands:

- `rename_worksheet` requires a live sheet, validates the candidate worksheet metadata, rejects an
  ASCII-case-insensitive duplicate name, and commits only the selected record;
- `set_worksheet_visibility` requires a live sheet and reports a no-op without exposing a mutable
  record. Read-only, last-visible-sheet, and active-selection policy remains a runtime concern;
- `bind_chart_sheet_package` accepts only a wholly unbound chart sheet or the exact same complete
  binding. It rejects partial bindings, retargeting, duplicate relationship IDs, and duplicate
  case-insensitive part URIs before atomically updating the worksheet relationship, worksheet part
  URI, and chart-sheet raw part URI; and
- `validate_worksheet_reorder` and `reorder_worksheets` require an exact permutation of the current
  unique sheet IDs. Runtime move paths preflight that permutation and prepare the workbook XML
  rewrite before committing the model order; collection move no longer drains the live vector
  before fallible XML work.

These commands own model invariants, while runtime owns the larger worksheet-rename transaction.
Rename constructs immutable old/new worksheet views, clones only the defined-name table and chart
map, and performs every fallible name resolution and chart-source rewrite against those prepared
copies. Only after preparation succeeds does it commit the live worksheet rename and replace the
name/chart substate; the replacements after the validated rename are infallible. A rejected rename
therefore preserves the complete `WorkbookState` and every dirty domain, including when a direct
chart source has already been rewritten in the prepared map before an invalid full-reference source
fails. The package is not touched until the later save transaction.

Placement-target sheet-block Copy now stages the complete target `RuntimeWorkbook` in a clone. Every
sheet in the requested block is copied into that prepared target; only an entirely successful block
remains live. A failure on any later sheet restores the original workbook/package/support state and
dirty domains, runtime object and stale registries, workbook/object handle allocators, active
workbook/chart and selection state, clipboard/find state, and removes temporary workbooks created
during preparation. The regression corrupts only the second of two source chart-sheet bindings and
requires the first copy to leave no target or runtime residue.

This is deliberately a correctness-first clone boundary. `OOTD-034`/`OOTD-056` retain the COW and
memory-reduction follow-up.

Worksheet Add uses the same `RuntimeWorkbookMutationSnapshot` boundary. Argument and placement
validation happen before the snapshot; Count iterations, template workbook registration and close,
native worksheet/chart-sheet/dialog-sheet/macro-sheet package graph creation, model and support
insertion, calc-chain invalidation, handle allocation, and selection changes happen against the
prepared target. A late failure restores the original target and session snapshot. The regression
uses malformed workbook XML so `Charts.Add` fails only after attempting to add chartsheet, drawing,
chart, relationship, and content-type artifacts, then requires exact workbook/package/dirty and
runtime-session equality.

Individual worksheet and chart-sheet Delete starts the same snapshot after read-only, last-sheet,
visibility, alert, and exclusively-owned package-graph preflight, immediately before the first live
owner removal. Worksheet/data, chart-sheet binding, drawing/chart and support/pending graphs,
workbook relationships, content types and package parts, calc-chain invalidation, stale handles,
selection, clipboard, and find state then form one transaction. The regression makes
`[Content_Types].xml` malformed so `Chart.Delete` fails only after removing model, support, and
package graph state, and requires exact restoration. The shared persistence snapshot compares
model/package, support and pending graphs, runtime chart support sources, calculation state, source
identity, and every dirty domain.

Collection Delete adds an outer snapshot after collection, pivot, last-sheet/visible-sheet, and
DisplayAlerts preflight and before the first item. Nested individual Delete transactions can commit
inside that prepared target, but any later item error restores the whole pre-call workbook and
session. The regression deletes a clean worksheet first and then rejects a worksheet whose opaque
drawing relationship target is missing; neither deletion, stale handle, package rewrite, nor
selection change remains.

Target-less `Sheets`/`Worksheets`/`Charts.Copy` starts an outer source-anchored runtime snapshot after
argument, collection, placement, and pivot preflight and before destination workbook creation. The
destination workbook, default-sheet cleanup, nested block copy, copied-sheet renames, object
handles, allocator counters, active workbook/chart, and selection remain prepared state until the
whole operation succeeds. If a later source sheet fails, the outer rollback removes the newly
registered workbook and restores the pre-call source and session state. The regression removes the
second of two source chart-sheet bindings, then requires the first copied chart, destination
workbook registration, handles, counters, and active state to leave no residue.

## Worksheet Data Mutation

Worksheet-data lookup and mutation require a worksheet record before consulting the data map.
`insert_cell` and `set_worksheet_source_xml` therefore return `OmErrorCode::NotFound` without
changing state when a `SheetId` is unknown; they never create a default data entry implicitly.
`set_worksheet_source_xml` is fallible for the same reason. A pre-seeded orphan map entry does not
make the ID mutable through `worksheet_data_for_sheet_mut`.

Worksheet creation, decode, and copy operations remain responsible for installing the worksheet
record and its data as one higher-level operation.

`WorkbookState::clear_range_with_change` owns the content-and-format semantics for both single-area
and multi-area `Range.Clear`. It resolves one live worksheet, validates every target coordinate
against spill-child ownership before mutation, and then clears each permitted cell plus any
anchor-owned spill extent and dynamic-array marker while deriving `dirty` and `dirty_cells`.
Validation failure therefore leaves earlier normal areas, spill metadata, cell/style payload, and
dirty state unchanged. Runtime uses this command for both Range object representations; the
regression covers direct `K10` and multi-area `A20,K10` attempts against a materialized `J10:K11`
spill.

`WorkbookState::clear_range_formats_with_change` owns the distinct style-only semantics for both
`Range.ClearFormats` representations. It validates the target worksheet/ranges, clears only
`style_id`, and derives dirty state without changing value, formula, dynamic-array marker, spill
range, or owner. A blank, non-formula cell is removed only when it does not participate in spill
topology. The regression seeds a styled blank child at `K10`, then verifies direct `K10` and
multi-area `A1,K10` calls preserve its materialized cell identity and leave the model valid for
save.

`WorkbookState::replace_cells_with_change` owns the final cell batch produced by
`Range.Replace`. Runtime evaluates every single-area or multi-area candidate against one immutable
worksheet snapshot and passes only changed final cells to the command. The command validates every
coordinate and every changed spill child before mutation, then derives cell, spill, dynamic-array,
and dirty state in one commit. A changed dynamic anchor clears its previously owned materialized
extent; formula-to-formula replacement restores the anchor's dynamic marker so the next
calculation cycle can materialize the new extent. Direct `K10` and multi-area `A20,K10`
spill-child failures preserve workbook persistence, dirty-domain, and session snapshots, while a
`J10` formula replacement is recalculated into its new spill shape.

### Worksheet-data ownership map

The `WorkbookState` worksheet-data map is private. External callers can inspect it through
`worksheet_data`, access a live owner through `worksheet_data_for_sheet` or
`worksheet_data_for_sheet_mut`, and change topology only through paired commands:

- `insert_worksheet_with_data` preflights the insertion index, worksheet metadata, workbook owner,
  ID, case-insensitive name, relationship ID, and part URI before adding the worksheet and data
  together;
- `replace_worksheet_data_for_sheet` requires both a live worksheet owner and an existing data
  entry;
- `remove_worksheet_with_data` checks the worksheet and data pair before removing either, and
  refuses to remove the workbook's only sheet; and
- `mark_worksheet_data_clean` changes dirty payload state without exposing map keys for mutation.

Codec construction uses `WorkbookState::try_new(WorkbookStateParts)`, so the public construction
DTO cannot become a live state until full model validation succeeds. Runtime add, copy, and delete
paths use the paired commands. Chart-sheet materialization no longer calls
`entry(...).or_default()`; a missing data owner now fails through `set_worksheet_source_xml`
instead of being invented.

This boundary closes external orphan-key insertion and rekeying. The live worksheet collection is
also private, so callers cannot insert, remove, reorder, or mutate worksheet identity through a
borrowed element. `WorksheetData` payload fields remain public through the existing live-owner
accessor; `Range.Clear`, `ClearFormats`, and `Replace` no longer use that bypass, but structural,
copy/paste, sort/fill, and calculation writeback paths remain explicit `OOTD-054` follow-ups.
Invalid state continues to be rejected by save preflight. Production worksheet metadata and
ordering paths use the command boundary above; workbook model metadata is separately private.

## Workbook Identity Reassignment

`WorkbookState::assign_workbook_id` is fallible and atomic. Its prepare phase clones only the chart
map, then reconstructs every range-bearing chart source for the target workbook ID:

- `name`, `x_values`, `values`, and `bubble_size`;
- each source's direct `resolved` range; and
- each source's `full_reference.resolved` range.

Malformed deserialized ranges return `OmErrorCode::InvalidState` with chart, one-based series,
source-slot, and direct/full-range context. No model, worksheet, chart, drawing, or chart-object ID
changes on failure. After every chart source validates, commit updates the workbook model, all
worksheet owners, the prepared chart map, every drawing owner, and each chart-frame object owner.
Defined names and raw formula strings carry no workbook ID and are not rewritten.

Runtime open/register increments its next handle only after reassignment succeeds. Active-content
strip/reload and prepared-save baseline reconstruction propagate the same error instead of
publishing a partially rebound state. The prepare phase deliberately avoids cloning worksheet cell
maps, source XML, or opaque package parts.

## Scoped Names And Spill Topology

A worksheet-scoped defined name must reference an ID in the worksheet collection. Save no longer
silently converts a dangling local name to workbook scope by omitting `localSheetId`. On load, a
present `localSheetId` must be unsigned decimal and index an existing worksheet; malformed or
out-of-range values fail parsing instead of being re-scoped.

For each worksheet data entry:

- every dynamic-array marker must identify a formula cell;
- every materialized spill range must belong to a dynamic-array marker and contain its anchor;
- materialized spill ranges cannot overlap; and
- every spill child must have a non-formula materialized cell and reference an existing owner range
  that contains the child without treating the anchor itself as a child.

A dynamic-array formula may have no materialized range. This represents an uncalculated or blocked
spill and remains valid.

## Package-To-Model Worksheet Binding

After state-only chart graphs have been materialized, save discovers the current workbook main
part and reparses its owner-relative relationships. The live model and package must then agree on:

- worksheet count and ordinal `sheetId` identity;
- relationship ID, normalized target part URI, and dialect-derived `SheetKind` at each ordinal;
- existence of every target part; and
- one-to-one ownership of the resolved package part, including targets that use different percent
  spellings for the same canonical OPC identity.

Worksheet name and visibility are deliberately excluded from this equality check because they are
supported workbook-XML rewrites and may lag the model until save. Discovery uses the actual root
`officeDocument` relationship and active Strict/Transitional dialect rather than assuming
`xl/workbook.xml` or Transitional relationship types.

## Deliberate Follow-up Boundaries

The worksheet-data ownership map and live workbook model metadata are the first two `OOTD-054`
private-field stages; the third routes production worksheet metadata, package binding, and order
changes through validated commands; and the fourth makes the live worksheet collection immutable
outside the model. The fifth makes runtime worksheet rename an atomic prepared substate commit.
Stages six through ten close sheet Copy/Add/Delete runtime transactions, stage eleven introduces the
first spill-aware payload mutation command for `Range.Clear`, and stage twelve separates
spill-preserving style-only `Range.ClearFormats`. Stage thirteen makes `Range.Replace` one
immutable-snapshot, spill-aware cell batch. Defined-name, chart, drawing, chart-sheet, opaque-part,
and the remaining `WorksheetData` payload fields remain public. Callers can still create malformed
state through those surfaces, but model save and identity-reassignment boundaries reject it
deterministically.

Manifest/content-type coherence and typed chart/drawing model-to-package ownership are enforced by
later `OOTD-031` stages. The chart/drawing boundary is documented in
`chart_drawing_graph_validation.md`; loaded worksheet/drawing/chart snapshot ownership and internal
inventory coherence are documented in `support_snapshot_validation.md`. Generic internal
relationship-target closure is enforced at the final package save boundary and documented in
`opc_relationships.md`. Excel grid coordinates and style indices retain their existing codec
preflight contracts. Spill-overlap validation currently has quadratic worst-case behavior, shared
with the existing worksheet parser; resource/performance hardening remains a later gate.

The current evidence is synthetic. No desktop Excel Oracle compatibility claim is attached to this
model validation boundary.
