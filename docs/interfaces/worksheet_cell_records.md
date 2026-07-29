# Worksheet Row And Cell Record Parse Policy

OOTD treats worksheet row and cell coordinates as declared identities. The default XLSX load path
must reject ambiguous or malformed declarations before it admits any cell from that worksheet into
the workbook model.

## Implemented Strict Boundary

For every typed SpreadsheetML `row` and `c` element recognized by namespace URI plus local name:

- `row@r` may be omitted, but when present it must be a positive decimal `u32`;
- every cell must provide a nonempty coordinate declaration in `c@r`; the complete A1 grammar is
  intentionally deferred to the checked reference work below;
- when the containing row declares `row@r`, the row component of `c@r` must match it;
- `c@s`, when present, must be a decimal `u64`; the existing workbook/style preflight separately
  rejects numeric IDs outside the loaded `styles.xml` `cellXfs` range;
- each normalized `(row, column)` coordinate may occur only once in the worksheet.

Duplicate detection happens before blank-cell elision. Consequently, an unstyled empty cell still
claims its coordinate and cannot hide a later populated duplicate. Both start/end cells and empty
cell elements use the same checks.

Violations return a deterministic `OmErrorCode::Parse` for lexical, coordinate, and duplicate
failures. Diagnostics include the worksheet part name and, whenever a reference is available, the
cell address. Numeric style references outside `cellXfs` retain the save/load invariant error but
now also identify the worksheet part and cell.

## Repair Boundary

There is no implicit worksheet repair mode. OOTD does not discard malformed row/style attributes,
move a cell to its container row, or choose one duplicate value. A malformed worksheet prevents the
workbook from loading and no partially repaired model is returned.

A future repair capability must be explicit and return a structured report for every discarded or
rewritten declaration. It must not weaken the default strict path.

## Remaining Boundaries

The worksheet single-cell grammar, checked arithmetic, absolute markers, and grid bounds now have
the bounded first-stage contract in `worksheet_a1_cell_references.md`. Common range endpoints,
whole-row/column references, sheet qualification, and consumer grammar migration remain
OOTD-018/OOTD-048. Central model limits and range cardinality are defined in
`excel_grid_limits.md`. Full worksheet parent/content-model ordering and Markup Compatibility
remain OOTD-049/OOTD-061.

The current evidence is synthetic. Desktop Excel open/save/reopen observations remain required
before this behavior is marked Oracle-verified.
