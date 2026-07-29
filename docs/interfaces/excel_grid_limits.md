# Excel Grid Limit And Range Cardinality Contract

OOTD uses `office_common::ExcelLimits` as the source of truth for the worksheet grid. The bounded
grid is rows `1..=1_048_576` and columns `1..=16_384` (`A..=XFD`), for a maximum of
`17_179_869_184` cells in one rectangular area.

## Validation Boundaries

`ExcelLimits::validate_cell` and `ExcelLimits::validate_rect` enforce positive 1-based coordinates,
ordered rectangle endpoints, and the Excel grid maximum. They are used at these operational
boundaries:

- `RangeArea` and `RangeSet` construction;
- all `WorkbookState` range reads, writes, formula writes, and clear operations;
- direct `WorkbookState::insert_cell` and `WorksheetData::clear_owned_spill` mutation;
- worksheet A1 and row parsing, runtime range parsing/expansion, chart references, and defined-name
  reference-like-name validation through constants derived from `ExcelLimits`;
- XLSX load and save state preflight, including cells, dirty cells, dynamic-array anchors, spill
  ranges, spill children, and spill owners.

Public model operations reject invalid coordinates with `OmErrorCode::InvalidArgument` before any
mutation. XLSX lexical input remains a `Parse` error with part/cell context. If callers bypass model
commands through currently public worksheet fields, save rejects that state with `InvalidState`
instead of emitting an out-of-grid cell reference.

## Checked Cardinality

`Rect` and `RangeArea` expose `checked_cell_count` for `u64` and
`checked_cell_count_usize` for allocation/indexing. Both validate the rectangle first. The `usize`
form returns `ResourceLimit` on platforms that cannot represent the count; range materialization
paths no longer multiply `u32` height and width before widening. Runtime `CountLarge` continues to
represent the full grid as `17_179_869_184` without allocating it.

These cardinality methods establish arithmetic safety, not a memory budget. A separate resource
policy may reject a valid but impractically large materialization before allocation.

## Remaining Invariant Work

`Rect`, `RangeRef`, `WorksheetData`, and other model structures still expose compatibility fields
or infallible DTO constructors. Every supported operation validates them before use, and the XLSX
save preflight prevents invalid output, but fully preventing construction or deserialization of raw
invalid state requires the private-field/validated-command work in OOTD-031 and OOTD-054.

The current evidence is synthetic. Desktop Excel differential observations remain required before
this behavior is marked Oracle-verified.
