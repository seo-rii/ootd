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

## Remaining Boundaries

- `Names.Add` and `Name.RefersTo` treat A1 text as relative to A1. Desktop Excel reads and
  reports it relative to the active cell, so the two differ when relative names are added or
  read with another cell active.
- Built-in `_xlnm.` names (`Print_Area`, `_FilterDatabase`, `Print_Titles`) are preserved but not
  given special evaluation meaning.
- External-workbook references in name formulas are not evaluated.
- A text argument passed straight to a numeric aggregate (`SUM("x"&1)`) is unsupported rather
  than `#VALUE!`.
