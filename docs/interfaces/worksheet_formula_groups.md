# Worksheet Formula Groups

Status: synthetic contract (`OOTD-027`/`OOTD-028`/`OOTD-067`). Desktop Excel evidence has not been
captured for any row below.

A formula group is a set of worksheet cells whose formula meaning is defined by one `<f>` element.
The codec models every group kind SpreadsheetML defines, so load, save, and mutation handle each
group as a unit instead of as independent cells.

## Group Kinds

| Source `<f>` | Model | Owned cells |
| --- | --- | --- |
| no `t`, or `t="normal"` | `CellData.formula` only | the cell itself |
| `t="shared" ref=R si=N` + text | `FormulaGroup { kind: Shared, range: R, shared_index: N }` keyed by the master | the master plus every child that references `si=N` |
| `t="shared" si=N` without `ref` | child entry in the master's `members` | — |
| `t="array" ref=R` | `dynamic_array_formulas` + `spill_ranges` (both legacy CSE and dynamic arrays) | every cell in `R` |
| `t="dataTable" ref=R` | `FormulaGroup { kind: DataTable, range: R }` keyed by the top-left cell | every cell in `R` |

`WorksheetData.formula_groups` holds shared and data-table groups keyed by anchor.
`FormulaGroup::owns` and `WorksheetData::formula_group_owner_for_key` answer whether a coordinate
belongs to a group. A shared child keeps `CellData.formula = None` and its cached value; the
child's formula text is not synthesized. A data-table anchor usually has an empty `<f/>` element
and therefore also has no formula text.

Legacy CSE arrays and dynamic arrays share the spill model: both reject edits to non-anchor
members and both round-trip their `t="array"`/`ref` metadata. Telling the two apart requires
resolving the anchor's `cm` attribute against the workbook cell-metadata part (`XLDAPR`), which the
codec does not read yet. The `cm` attribute itself is preserved on every rewrite.

## Load Validation

Load fails closed with the worksheet part and cell reference when:

- `t` is not `normal`, `shared`, `array`, or `dataTable`;
- a shared formula has no `si`, or `si` is not a decimal `u32`;
- two shared masters declare the same `si`;
- a shared master has no formula text or lies outside its own `ref`;
- a shared child references an `si` with no master, or lies outside the master's `ref`;
- a data-table formula has no `ref`, or its cell is not the top-left of `ref`;
- a cell owned by a shared or data-table group also lies inside an array spill range or another
  data table.

## Mutation Preflight

Every public cell-payload command reaches the group check through the same spill preflight, so a
rejected command changes nothing:

- value, formula, dynamic-array, and R1C1 formula writes;
- `ClearContents` and `Clear`;
- Replace, rearrange, sort, fill (sources and destinations), copy/cut destinations and cut
  sources, and paste-special materialization;
- `Range.Insert`/`Delete`, whose shift corridor may not intersect a group-owned cell.

The error is `InvalidState` with the message
`cannot <operation> <shared formula master|shared formula child|data table cell> R<r>C<c>; formula group anchor is R<r>C<c>`.
Format-only commands (`ClearFormats`, style changes) are allowed and keep the group intact.

Save preflight rejects in-memory groups whose shared master has no formula cell, whose master or
members lie outside the group range, whose data-table anchor is not the top-left of its range,
whose range leaves the grid, or whose cells overlap an array spill range.

## Save Guarantees

1. **No-op save** keeps the worksheet bytes identical.
2. **Unrelated edit** keeps every group cell's source bytes verbatim, and reopen restores the same
   groups.
3. **Touched rewrite** (a group cell marked dirty with an unchanged payload, for example after a
   format change) keeps `t`, `ref`, `si`, and every other `<f>` attribute on masters, and re-emits
   the original formula-less `<f/>` element for shared children and data-table anchors. Reopen
   restores the same groups.

The regression matrix is `crates/excel-xlsx/src/tests/formula_groups.rs`; model preflight is
covered by `formula_group_members_reject_structural_and_transfer_commands` and
`save_preflight_rejects_incoherent_formula_groups` in `excel-model`.

## Remaining Boundaries

- Shared children do not carry synthesized formula text, so calculation treats them as their
  cached values.
- No command removes or replaces a whole shared group or data table yet, so group-owned cells stay
  read-only for payload edits.
- Legacy-array versus dynamic-array classification waits for cell-metadata (`cm` → `XLDAPR`)
  support.
