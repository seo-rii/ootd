# Workbook State Save Validation Contract

`WorkbookState::validate_for_save` is the model-level fail-closed boundary for topology that can
still be invalidated through public state fields. `XlsxCodec::load` validates the decoded state
before exposing it, and every lossless XLSX save validates the live state before package
serialization or graph materialization.

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

## Worksheet Data Mutation

Worksheet-data lookup and mutation require a worksheet record before consulting the data map.
`insert_cell` and `set_worksheet_source_xml` therefore return `OmErrorCode::NotFound` without
changing state when a `SheetId` is unknown; they never create a default data entry implicitly.
`set_worksheet_source_xml` is fallible for the same reason. A pre-seeded orphan map entry does not
make the ID mutable through `worksheet_data_for_sheet_mut`.

Worksheet creation, decode, and copy operations remain responsible for installing the worksheet
record and its data as one higher-level operation.

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

This boundary closes external orphan-key insertion and rekeying. The worksheet collection and the
fields inside each `WorksheetData` remain public in this stage, so worksheet-ID drift and
cell/spill/dirty-state invariant bypasses remain explicit `OOTD-054` follow-ups and continue to be
rejected by save preflight.

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

The worksheet-data ownership map is the first `OOTD-054` private-field stage. The worksheet
collection, workbook model, defined-name, chart, drawing, chart-sheet, opaque-part, and
`WorksheetData` payload fields remain public. Callers can still create malformed state through
those surfaces, but model save and identity-reassignment boundaries reject it deterministically.

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
