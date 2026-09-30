# Dynamic-Array Cell Metadata

Status: synthetic contract (part of `OOTD-075`/`OOTD-076`/`OOTD-077`). Desktop Excel evidence has
not been captured for any row below.

SpreadsheetML writes every array formula as `<f t="array" ref="…">`. Desktop Excel treats one as a
dynamic array only when its cell carries a `cm` index into the workbook cell-metadata part, and the
entry at that index marks it with `XLDAPR` dynamic-array properties. Without that metadata Excel
opens the formula as a fixed-size legacy (CSE) array shown as `{=…}`.

## The Metadata Part

The part is reached through the workbook relationship of type `…/relationships/sheetMetadata`
(Transitional or Strict), normally at `xl/metadata.xml`, with content type
`application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml`.

A `cm` value `N` selects the `N`-th `cellMetadata/bk` block. The block counts as dynamic when any
of its `rc` records has:
- `t` naming (one-based) a `metadataType` called `XLDAPR`;
- `v` selecting (zero-based) an `XLDAPR` `futureMetadata/bk` block;
- that block containing `xda:dynamicArrayProperties` with `fDynamic="1"`.

`WorkbookSupportParts::cell_metadata` records the part URI and the set of dynamic indices.

## Load

- A `t="array"` formula whose cell `cm` is a dynamic index becomes a dynamic array
  (`dynamic_array_formulas` plus `spill_ranges`).
- Every other array formula becomes a `LegacyArray` formula group. That includes a missing `cm`,
  a `cm` that resolves to other metadata, and a `cm` with no metadata part at all. See
  `docs/interfaces/worksheet_formula_groups.md`.
- These fail closed with the part URI and cell context:
  - more than one `sheetMetadata` relationship, an external or missing target, or the wrong root
    namespace;
  - an `rc` record without `t`/`v`, a non-numeric index, or a `t` that names no metadata type;
  - a non-numeric `c@cm`.

## Save

When a dirty worksheet holds a dynamic-array anchor:
- If the workbook already defines a dynamic index, that index is reused.
- Otherwise, when the package has no metadata part, the codec adds `metadata.xml` next to the
  workbook part (in the single-block form Excel writes, so `cm="1"`), a workbook relationship with
  the next free `rId`, and a content-type override.
- Otherwise, when the existing metadata part has no dynamic block, the save is refused with
  `Unsupported` rather than rewriting a part the codec does not model.

The rewriter then:
- sets `cm` to the dynamic index on every dynamic-array anchor, replacing any other `cm`;
- removes a dynamic `cm` from rewritten cells that are no longer dynamic anchors, such as an
  anchor overwritten with an ordinary formula or a cleared cell;
- keeps every other `cm` value unchanged.

Untouched cells and parts keep their bytes, and saving a reopened workbook again without edits
reproduces the same package.

The regressions are in `crates/excel-xlsx/src/tests/dynamic_array_metadata.rs`.

## Remaining Boundaries

- Appending an `XLDAPR` block to an existing cell-metadata part that lacks one (for example, a
  workbook with only rich-value metadata) is not implemented.
- `fCollapsed` and other dynamic-array properties are preserved but not modeled.
