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
- An R1C1 formula is retargeted in A1 form at its cell's old position and written back relative
  to where the cell lands, so `R[-1]C` keeps reading the same cell as both move. An R1C1 name is
  retargeted relative to A1, where name text is anchored, and stays R1C1.

Worksheet structure moves with the cells. Each shift is recorded on the worksheet
(`WorksheetData::structural_shifts`), and the XLSX rewriter replays the recorded shifts:
- onto source-keyed rows and cells (row attributes such as `ht`, and cell attributes);
- onto `mergeCell@ref`, `dataValidation@sqref`, `conditionalFormatting@sqref`, `hyperlink@ref`,
  `autoFilter@ref`, `selection@sqref`/`activeCell`, `pane@topLeftCell`, and `col@min`/`max`.

Comment anchors (`comment@ref` in the comments part) and legacy VML shapes (zero-based
`x:Row`/`x:Column` and the `x:Anchor` box) move too. The save writes the moved comment and VML
parts next to the rewritten worksheet.

Tables follow too. Their parts replay the recorded shifts onto `table@ref`, `autoFilter@ref`,
`sortState@ref`, and `sortCondition@ref`:
- Rows may be inserted anywhere: a table moves, or grows when the insertion is inside its body.
- Body rows may be deleted while the header row and one data row survive.
- A column insert or delete may only move a table that lies wholly to one side of it, because
  adding or removing table columns would change the table's column definitions.
- Table formulas must not need retargeting.

Charts follow too:
- **Series sources:** the `name`, `x-values`, `values`, and `bubble-size` sources on the edited
  sheet rewrite their reference text and resolved ranges. The chart is patched in place, so its
  formatting is kept.
- **Chart frames:** frames hosted on the sheet move their two-cell or one-cell anchors. A
  move-only frame keeps its extent, and an edge inside a deleted span lands on the first row or
  column after it.
- **Pictures, shapes, and other drawing objects:** the save replays the recorded shifts onto
  their `from`/`to` markers in the drawing part, editing each anchor in place and leaving the rest
  of the part byte-identical. `oneCellAnchor` and `editAs="oneCell"` objects keep their size;
  `editAs="absolute"` and `absoluteAnchor` objects do not move.

Ranges grow, shrink, or move as references do. Ranges on deleted rows or columns are removed, and
`mergeCells@count` follows. A selection or pane origin on deleted cells returns to A1. The model's
structural inventory moves with the shift, save checks the worksheet's hyperlink snapshot against
the replayed refs, and a successful save rebases the source XML and clears the record.

A shared-formula group that a whole-axis shift moves, or whose member formulas it rewrites, is
unshared: every member becomes an ordinary formula with its own text (the runtime expands
children from their master at open), retargets like any other formula, and is saved as a plain
`<f>`. Groups the shift neither moves nor rewrites stay shared. A child without expanded formula
text, as in codec-only use, refuses the shift.

The shift is refused atomically, as before, when it would need to rewrite any of:
- a reference to an unknown sheet, or a 3D reference;
- a name whose moving reference is relative or unqualified;
- a data-validation formula that references the moved area;
- table or chart formulas.

It is also refused when it would:
- delete through part of a merged range;
- remove a whole data-validation range, a hyperlink, or a commented cell;
- delete a table's header row or all of its data rows, or insert or delete columns inside a
  table;
- leave a chart source whose whole range is deleted, or one that is unresolved or 3D;
- move a sheet that hosts a drawing object whose anchor kind is not two-cell, one-cell, or
  absolute, or a chart frame with a cell-bound absolute anchor.

### Partial Corridors

`Insert`/`Delete` of a partial-width or partial-height range (`B2:C3` shifted down) is a banded
shift: it moves only the cells in the band of columns (or rows) the range spans, from the edited
row (or column) onward, as Excel shifts cells.
- A reference or range moves only when its whole span on the other axis lies inside the band;
  `A1:C9` stays when only `B:C` shift. Whole-row and whole-column references never move.
- Merged cells, validation ranges, comments, notes, and chart series sources inside the band move.
  A band that cuts through a merged range or a validation range is refused.
- A conditional format or selection the band cuts through is split, as Excel splits it: the part
  inside the band moves and the parts beside it stay. A band through part of a hyperlink or
  auto-filter range is refused, because a single range cannot hold the split.
- Row heights and column widths stay, because no whole row or column moves.
- Tables and chart frames stay put; a band that reaches one is refused. A sheet that hosts
  pictures or shapes refuses partial corridors.

The rules live in `office_common::retarget_formula_references` and `StructuralShift`. The
regressions are their unit tests, `partial_corridor_shifts_retarget_formulas_inside_the_band`,
`partial_corridor_shifts_move_worksheet_structure_inside_the_band`,
`structural_shifts_retarget_r1c1_formulas`, `structural_shifts_retarget_r1c1_formulas_and_names`,
`whole_row_and_column_shifts_retarget_formulas_and_names`,
`whole_row_and_column_shifts_move_worksheet_structure`,
`whole_row_shifts_move_comments_notes_and_hyperlinks`,
`whole_row_shifts_move_chart_sources_and_frames`, `whole_row_and_column_shifts_move_tables`, and
`whole_row_shifts_move_shapes_and_keep_absolute_objects`, with shared formulas pinned by
`whole_row_shifts_unshare_moved_shared_formulas` and
`whole_row_shifts_keep_untouched_shared_formulas_shared`.

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
