# Formula File Grammar

Status: synthetic contract (part of `OOTD-076`/`OOTD-077`). Desktop Excel evidence has not been
captured for any row below.

SpreadsheetML stores cell formulas in a file grammar that differs from the grammar users type.
The model (`CellData.formula`, `Range.Formula`, and evaluation) always holds the typed grammar. The
XLSX codec converts at the package boundary in `crates/excel-xlsx/src/worksheet/formula_grammar.rs`.

| Typed grammar | File grammar |
| --- | --- |
| `XLOOKUP(…)` and every other post-2007 function | `_xlfn.XLOOKUP(…)` |
| `FILTER(…)`, `SORT(…)` | `_xlfn._xlws.FILTER(…)`, `_xlfn._xlws.SORT(…)` |
| `LET(x,1,x+1)`, `LAMBDA(v,v*2)` parameters | `_xlpm.x`, `_xlpm.v` |
| `@A1:A10`, `@(A1:A2+1)` | `_xlfn.SINGLE(A1:A10)`, `_xlfn.SINGLE(A1:A2+1)` |
| `D1#`, `'My Sheet'!D1#` | `_xlfn.ANCHORARRAY(D1)`, `_xlfn.ANCHORARRAY('My Sheet'!D1)` |

## Rules

- Only function names in the codec's future-function list are converted. An unknown
  `_xlfn.NAME(…)` stays byte-for-byte in both directions, so functions newer than the list
  survive a round trip.
- String literals, quoted sheet names, and bracketed structured or external references are copied
  verbatim. So `"a@b#c"` and `Table1[@Col]` are never rewritten.
- `LET` names enter scope at their declaration and `LAMBDA` parameters for the whole call. Only
  identifiers inside that scope gain `_xlpm.`, so an outer name that matches a parameter
  elsewhere in the formula is left alone.
- Load converts shared-formula masters, array formulas, data-table formulas, and ordinary
  formulas alike. Untouched cells keep their source bytes on save. A rewritten cell converts its
  typed text back, so loading and re-saving file-grammar formulas reproduces them exactly.

The regressions are the unit tests in `formula_grammar.rs` and
`crates/excel-xlsx/src/tests/formula_grammar.rs`.

## Remaining Boundaries

- Defined names, data-validation formulas, conditional formats, and chart formulas still keep
  their file text without conversion.
- The runtime does not evaluate `@` yet.
