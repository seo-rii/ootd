# Worksheet A1 Cell Reference Parse Policy

OOTD uses a bounded parser for single-cell coordinates read from worksheet XML. This is the first
stage of OOTD-018 and fixes the XLSX ingestion boundary without claiming that every runtime, chart,
or defined-name consumer already shares one reference AST.

## Implemented Grammar

The worksheet cell parser accepts:

- one or more ASCII column letters followed by one or more decimal row digits;
- ASCII upper- or lowercase column letters;
- at most one absolute marker before the column and at most one before the row;
- a column-only form only when the enclosing row supplies the current row index.

Examples include `A1`, `xfd1048576`, `$XFD$1048576`, `$A1`, `A$1`, and an in-row `BC`.
Standalone row-only references, ranges, sheet-qualified references, whitespace, trailing text,
misplaced/repeated absolute markers, and mixed endpoint forms are rejected at this boundary.

## Arithmetic And Grid Limits

Column base-26 and row base-10 accumulation use checked arithmetic. Parsing stops with
`OmErrorCode::Parse` as soon as a coordinate exceeds Excel's worksheet grid:

- rows: `1..=1_048_576`;
- columns: `1..=16_384` (`A..=XFD`).

The original lexical input is retained in every grammar, overflow, and out-of-grid diagnostic.
Worksheet load wraps that error with the owning part name. Declared `row@r` values use the same row
limit, including rows that contain no materialized cells.

Synthetic evidence includes the first and last cell, absolute and lowercase forms, beyond-grid
rows and columns, numeric and alphabetic overflow strings, malformed suffix/qualification/marker
cases, invalid current rows, and an exhaustive formatter/parser round trip across all 16,384 Excel
columns at row 1,048,576.

## Remaining Reference Work

OOTD-018 remains open for the common A1 endpoint/range AST, including standalone whole-row and
whole-column ranges and sheet quoting. OOTD-019 now provides the shared `ExcelLimits` and checked
range-cardinality contract described in `excel_grid_limits.md`. OOTD-048 will migrate runtime,
defined-name, chart, formula, and codec grammar consumers to one shared reference AST.

The current evidence is synthetic. Desktop Excel differential observations remain required before
this behavior is marked Oracle-verified.
