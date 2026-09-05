# Worksheet Cell Value Channel Policy

OOTD loads every SpreadsheetML `c` element into one `CellData` whose value, formula, and style
channels are independent. This document fixes how the codec distinguishes a missing cell from a
blank cell, how each `t` cell type maps its `v`/`is` content onto a typed value, how a formula
cell is loaded with and without a cached result, and what a dirty rewrite emits for each case.

## Presence Model

| Source cell | Loaded model | Dirty rewrite of an unchanged value |
|---|---|---|
| `<c r="A1"/>` or `<c r="A1"></c>` without `s`, `f`, `v`, or `is` | not materialized | element removed |
| the same with any `t` attribute | not materialized; a `t` without a value channel carries no value | element removed |
| `<c r="A1" s="n"/>` or `<c r="A1" t="…" s="n"></c>` | `Blank` with style `n` | `<c r="A1" s="n"/>` |
| `<c r="A1"><f>…</f></c>` with any `t` and no `v` | `Blank` with formula: no cached result | `<c r="A1"><f>…</f></c>` |
| `<c r="A1" t="str"><v></v></c>` | `Text("")`: an empty string constant or cached result | constant: `t="inlineStr"` with `<is><t></t></is>`; formula cache: `t="str"` with `<v></v>` |

`Blank` with a formula therefore means "no cached result" and never "cached empty string". The
runtime must evaluate such a cell before its value is meaningful.

Both element forms (`<c/>` and `<c></c>`) follow the same rules, and a cell may declare at most one
`v` element.

## Typed Value Lexicals

| `t` | Accepted lexical | Typed value | Canonical rewrite lexical |
|---|---|---|---|
| absent or `n` | finite `xsd:double` without surrounding whitespace | `Number(f64)` | the shortest decimal that reparses to the identical binary64 value (`1E+20` → `100000000000000000000`, `.5` → `0.5`); `t` is omitted |
| `b` | `0`, `1`, `true`, `false` | `Bool` | `0` or `1` |
| `e` | non-empty token; known Excel errors are typed, any other token keeps its exact lexical | `Error` | exact token |
| `d` | validated ISO date or date-time | `IsoDateTime` | exact source lexical |
| `s` | decimal index into the shared string table | `Text`, or `RichText` sharing the item payload | `t="inlineStr"` materialization |
| `str` | any string, including empty | `Text` | formula cache: `t="str"` with `<v>`; constant: `t="inlineStr"` |
| `inlineStr` | one `is` item; a `v` element is rejected | `Text` or `RichText` | reused or materialized inline item |
| any other | — | fails closed | — |

Entity and character references inside `f` and `v` text are resolved on load (`&amp;`, `&lt;`,
`&gt;`, `&quot;`, `&apos;`, `&#n;`, `&#xh;`); unknown entities fail closed. Rewrites re-escape the
resolved text, so `=A1&"x"` survives load, dirty save, and reopen.

## Formula Cache Model

- `Blank` + formula: uncomputed cache; no `v` element is written.
- `Number`, `Bool`, `Error`, or `IsoDateTime` + formula: cached result typed by `t`.
- `Text` + formula: cached string result written as `t="str"`, including the empty string.
- A shared-string cache (`t="s"` with a formula) loads as the referenced text and rewrites as
  `t="str"`.
- `RichText` cannot be a formula cache; the rewriter refuses it.

## Fail-Closed Cases

Every failure returns `OmErrorCode::Parse` with the worksheet part URI and the cell address:

- an unknown `t` lexical in either element form;
- more than one `v` element in one cell;
- a `v` element inside an `inlineStr` cell;
- an empty `v` for `b`, `d`, `e`, `n`, or `s`;
- a boolean lexical other than the four accepted forms, including `TRUE`;
- a numeric lexical outside the finite `xsd:double` grammar (`1,5`, ` 1`, `0x10`) or a
  non-finite one (`INF`, `NaN`);
- a non-decimal or out-of-range shared string index;
- an invalid ISO date lexical;
- an unknown XML entity reference.

An in-memory `UnknownLexical("")` error value is rejected by `CellValue::validate`, by model
load/save preflight, and by the worksheet rewriter, so it can never be serialized as an empty `v`.

## Fidelity Matrix

The synthetic regression module `cell_value_fidelity` loads one row containing every channel above
and checks:

1. **Load**: every coordinate produces the typed model in the tables above.
2. **No-op save**: worksheet and shared-string part bytes are unchanged.
3. **Unrelated edit**: editing one cell in the same row keeps every other cell's source bytes
   verbatim, and reopen restores the same typed model.
4. **Touched rewrite**: marking every cell dirty with unchanged values emits the canonical forms
   above, and reopen produces the identical typed model.

## Remaining Boundaries

- A rewritten numeric or boolean cell does not retain its source spelling; the guarantee is
  typed-value identity, not byte identity. Untouched cells keep their bytes.
- Shared, legacy array, and data-table formula groups (`f@t`, `si`, `ref`) are outside this
  contract and remain `OOTD-027`/`OOTD-028`/`OOTD-067`.
- Desktop Excel evidence remains required before any row above is Oracle-verified. Excel's own
  tolerance of `true`/`false` and of empty `v` lexicals is unobserved, so those rows are
  synthetic policy decisions.
