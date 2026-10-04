# Calculation Order And Memoization

Status: synthetic contract (`OOTD-050`). Desktop Excel evidence has not been captured for any row
below.

A calculation cycle is one `CalcContext` (see `docs/interfaces/runtime_environment.md`). Within a
cycle the runtime evaluates every formula cell at most once and orders evaluation by the lexical
dependency graph, so deep or cyclic workbooks cannot exhaust the native stack.

## Memoization

- Every formula cell result, whether a value or an error, is memoized in the context under
  `(sheet, row, column)`.
- Every reference to a cell observes the memoized result, so a volatile cell such as
  `=RAND()` has one value per cycle for all of its dependents. Shared precedents are evaluated
  once, so doubling chains grow linearly rather than exponentially.
- The memo is discarded after the cycle commits new spill ranges, because formulas evaluated
  before that may have read the old spill children.

## Dependency Order

Before evaluating a formula cell, the evaluator scans the formula for lexical references
(A1 and sheet-qualified areas, 3D references, and defined names). It skips string literals and
tokens followed by `(`, so `LOG10(` is a function call rather than a cell. Every formula cell
inside a referenced area becomes a precedent.

Precedents are evaluated in post-order with an explicit heap stack, and each result is memoized.
When the cell itself is evaluated, its references hit the memo, so native recursion stays shallow
however long the chain is.

## Cycles

A lexical back edge does not prove a circular reference, because reference-only arguments
(`AREAS`, `ROWS`, `ROW`, `COLUMNS`, `ISREF`) never read the values they name. So the back-edge
target reads as circular only provisionally, while the rest of its cycle is evaluated. This matches
what recursive evaluation observes through its visiting guard. The target is then evaluated
normally:

- In a genuine value cycle, every member and every dependent evaluates to `#CALC!` and appears in
  `CalculationReport::circular`.
- In a reference-only cycle, every member computes normally.
- `IFERROR` around a cyclic reference catches the provisional circular result, as it does under
  recursion.

Cycle results are deterministic for a given workbook, independent of which cell starts the
evaluation. The regressions are in `crates/excel-runtime/src/tests/runtime_environment.rs`
(`calculation_*`).

## Iterative Calculation

`calcPr@iterate`, `iterateCount`, and `iterateDelta` load into `WorkbookIteration`
(`ExcelRuntime::workbook_iteration`/`set_workbook_iteration`). Invalid values fail closed: a count
outside 1..=32,767, or a negative or non-finite delta, is a load `Parse` error and a setter
`InvalidArgument`. Changed settings rewrite only these three attributes, omit Excel's defaults
(off, 100, 0.001), and keep every other `calcPr` attribute.

With iteration enabled, a workbook calculation runs up to `max_iterations` passes over every
worksheet inside one `CalcContext`:
- a circular reference reads the referenced cell's value from the previous pass instead of
  failing, and lexical back edges provisionally read that value too;
- the memo is discarded between passes;
- calculation stops early once no formula result moved by more than `max_change`;
- converged or capped cycles are not reported as circular.

Without iteration, cycles still evaluate to `#CALC!`. Worksheet and range `Calculate` run a single
pass.

The regressions are in `crates/excel-runtime/src/tests/iterative_calculation.rs`.

## Remaining Boundaries

- `Application.Iteration`, `MaxIterations`, and `MaxChange` are not dispatchable members yet,
  because the pinned object-model template has no captured evidence for them. Hosts use the Rust
  API.
- Passes evaluate sheets in collection order, so the iteration count Excel needs to converge may
  differ from this runtime's.
- References computed at run time (`INDIRECT`, `OFFSET`, `INDEX` returning a reference) are not
  in the lexical graph, so chains built only from them still evaluate recursively.
- The graph is rebuilt per cycle. There is no persistent dependency index or dirty-set
  propagation yet.
- A full-column or full-row reference scans every materialized cell in the covered rows for each
  referencing formula.
