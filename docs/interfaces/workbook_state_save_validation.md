# Workbook State Save Validation Contract

`WorkbookState::validate_for_save` is the model-level fail-closed boundary for topology that can be
invalidated through the currently public state fields. `XlsxCodec::load` validates the decoded
state before exposing it, and every lossless XLSX save validates the live state before package
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
- a corresponding `WorksheetData` entry exists.

The fully absent relationship/part pair is intentional. The model gate admits an unbound
chart-sheet record so the XLSX graph layer can separately validate and materialize the complete
chart/drawing binding. A partial pair is never serializable and returns
`OmErrorCode::InvalidState`.

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

## Deliberate Follow-up Boundaries

This stage does not reject extra `worksheet_data` keys that have no worksheet owner; `OOTD-032`
owns removing the auto-creating mutation path and orphan state. It also does not make public model
fields private (`OOTD-054`) or make workbook-ID reassignment atomic (`OOTD-033`).

Manifest/content-type coherence, generic package relationship closure, chart/drawing ownership,
support-snapshot graph validation, and exact package-versus-model sheet identity comparison remain
later `OOTD-031` stages. A caller can therefore still replace a complete model binding with another
internally consistent binding until that comparison lands. Excel grid coordinates and style
indices retain their existing codec preflight contracts. Spill-overlap validation currently has
quadratic worst-case behavior, shared with the existing worksheet parser; resource/performance
hardening remains a later gate.

The current evidence is synthetic. No desktop Excel Oracle compatibility claim is attached to this
model validation boundary.
