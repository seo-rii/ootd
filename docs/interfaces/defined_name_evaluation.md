# Defined-Name Evaluation

Status: synthetic contract (`OOTD-078`, in progress). Desktop Excel evidence has not been captured
for any row below.

## Resolution

A name in a formula resolves through `DefinedNameTable::lookup`. A name scoped to the evaluating
worksheet shadows a workbook-scoped name with the same spelling, and case is ignored. A name whose
formula refers to itself, directly or through other names, evaluates to `#CALC!`.

## Relative References

A1 name formulas are stored relative to A1, as Excel writes them. When a name is used, its
relative row and column parts move by the calling cell's offset from A1, and absolute (`$`) parts
stay fixed. A reference that leaves the grid wraps around it, so a name stored as
`Sheet1!A1048576` means "the cell one row up" and `Sheet1!XFD1` "the cell one column left" from
every caller. `Sheet1!$P$1:$P1048576` means "column P from row 1 to the row above". Without a
calling cell (`Application.Evaluate`) A1 is the caller. R1C1 name formulas already resolve at the
caller.

## Name Formulas

| Name formula | Use | Result |
| --- | --- | --- |
| `0.5`, `"total"` | `=P1*Rate`, `=Label&"!"` | the constant |
| `Rate*2` | `=Doubled` | the formula, evaluated in the caller's context |
| `Sheet1!$P$1:$P$3` | `=SUM(Values)`, `=ROWS(Values)`, `=INDEX(Values,2)` | the range |
| `Sheet1!$P$1,Sheet1!$P$3` | `=SUM(Corners)` | every area |
| `LAMBDA(x,x*3)` | `=Triple(4)` | the lambda, callable by name |
| `{1,2,3}` | `=SUM(Constants)` | the array constant |

## Grammar Added With This Item

- **Concatenation.** `&` joins operands into text in cell formulas and in function arguments,
  for example `=P1&" / "&P2+1` or `=LEN("ab"&Label)`. Arithmetic binds tighter than `&`. Blanks
  add nothing, numbers and booleans use their text forms, and the first error operand is the
  result.
- **Array constants.** `{1,2;3,4}` uses `,` between columns and `;` between rows. Elements are
  signed numbers, strings, `TRUE`/`FALSE`, and error literals, and every row must be the same
  width. Numeric aggregates (`SUM`, `AVERAGE`, `MAX`, and relatives) accept an array constant
  written inline or through a name. Like a reference, it contributes only its numbers.

The regression is `crates/excel-runtime/src/tests/defined_name_semantics.rs`.

## Structural Retargeting

`Range.Insert` and `Range.Delete` of whole rows (`A2:XFD3`) or whole columns (`B1:C1048576`)
retarget every cell formula and defined name in the workbook that references the edited sheet:
- On insert, endpoints at or after the insertion point move. A range end pushed past the grid
  stays on the last row or column; any other endpoint pushed past it becomes `#REF!`.
- On delete, endpoints after the deleted span move back. A range that loses only part of its span
  shrinks, and a reference entirely inside the span becomes `#REF!`.
- References to other sheets, string literals, structured references, and external-workbook
  references are unchanged. Edited cells are marked dirty at their new positions, and names are
  rewritten through the name table.

The shift is refused atomically, as before, when it would need to rewrite any of:
- an R1C1 formula;
- a reference to an unknown sheet, or a 3D reference;
- a member of a shared-formula group;
- a name whose moving reference is relative or unqualified;
- data-validation, table, or chart formulas.

Partial-width or partial-height `Insert`/`Delete` still refuse any reference-bearing formula. The
rules live in `office_common::retarget_formula_references`; the regressions are its unit tests
and `whole_row_and_column_shifts_retarget_formulas_and_names`.

## Remaining Boundaries

- `Names.Add` and `Name.RefersTo` treat A1 text as relative to A1. Desktop Excel reads and
  reports it relative to the active cell, so the two differ when relative names are added or
  read with another cell active.
- Built-in `_xlnm.` names (`Print_Area`, `_FilterDatabase`, `Print_Titles`) are preserved but not
  given special evaluation meaning.
- External-workbook references in name formulas are not evaluated.
- A deleted reference is written as `#REF!` without its sheet qualifier (Excel shows
  `Sheet1!#REF!`).
- A text argument passed straight to a numeric aggregate (`SUM("x"&1)`) is unsupported rather
  than `#VALUE!`.
