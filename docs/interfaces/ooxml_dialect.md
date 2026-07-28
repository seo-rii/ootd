# Strict And Transitional OOXML Dialects

OOTD treats the OOXML dialect as package-derived state. It does not infer Strict from a file
extension or convert namespaces by changing only the workbook content type.

## Detection Contract

The codec resolves three independent signals before parsing workbook state:

1. the exact package-root `officeDocument` relationship type;
2. the discovered workbook root namespace;
3. the discovered workbook part's resolved content type.

The relationship and workbook namespace must identify the same dialect. Exactly one recognized
Strict or Transitional `officeDocument` relationship may exist. Mixed, duplicate, unknown, or
missing signals fail before a workbook model is returned.

| Dialect | Workbook namespace | Relationship base |
|---|---|---|
| Transitional | `http://schemas.openxmlformats.org/spreadsheetml/2006/main` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` |
| Strict | `http://purl.oclc.org/ooxml/spreadsheetml/main` | `http://purl.oclc.org/ooxml/officeDocument/relationships` |

Transitional main content types map to `Xlsx`, `Xlsm`, `Xltx`, or `Xltm`. Strict workbooks accept
both the ECMA Strict `application/vnd.openxmlformats-officedocument.spreadsheetml.main+xml` value
and Excel's `...spreadsheetml.sheet.main+xml` value, and map them to `StrictXlsx`. Strict templates
cannot be represented by the current `FileFormat` enum, and macro-enabled content types cannot be
combined with the Strict relationship dialect; both combinations return `Unsupported`. Unknown
main content types return a parse error instead of falling back to XLSX/XLSM.

## Relationship Contract

An explicit table maps the following core relationship kinds for both dialects:

- office document, worksheet, chart sheet;
- shared strings, styles, theme, and calculation chain;
- drawing, chart, hyperlink, and comments.

Dialog-sheet and VML relationship types are Transitional-only. Known relationship types from the
opposite dialect fail closed. Unknown vendor relationship types remain opaque unless a typed owner
requires them, so lossless preservation does not turn into suffix-based type guessing.

Shared strings are resolved from the workbook relationship graph and may live at a nonstandard
part URI. The historical fixed `xl/sharedStrings.xml` fallback remains Transitional-only for
synthetic and legacy packages that omit the relationship.

## Supported Strict Save Boundary

The current Strict capability is preserve-only:

- load and no-op save preserve the discovered main part, owner relationships, namespaces, and
  unchanged part bytes;
- a worksheet cell edit rewrites the existing Strict worksheet without changing its dialect;
- calculation-state rewrite and calculation-chain removal use the Strict relationship graph;
- runtime `Save`, same-format `SaveAs`, and `.xlsx` path inference retain `StrictXlsx`;
- explicit or inferred Strict-to-Transitional and Transitional-to-Strict conversion returns
  `Unsupported` before output.

Strict chart/drawing graph decoding and mutation, state-only graph materialization, and worksheet
collection restructuring are not yet dialect-aware and return `Unsupported`. `supports_format`
continues to report `false` for `StrictXlsx` because it represents general target-format capability,
not this bounded same-dialect preservation path.

## Verification

Synthetic regressions use a nonstandard `documents/book/main.xml`, Strict root/workbook/worksheet/
shared-string/calculation-chain namespaces, a relocated shared-string part, and a relocated
calculation chain. They cover clean save, targeted cell edit, chain invalidation, Strict URI
retention, reopen, both accepted Strict main content types, mixed-dialect rejection, duplicate
root relationships, unknown content type, and runtime SaveAs behavior.

No real desktop Excel observation is pinned yet. This surface remains `Preserve-only`, not
`Oracle-verified`. Workbook core, shared-string, and worksheet-cell arbitrary prefixes are covered
by the bounded OOTD-015/OOTD-049 slice in `spreadsheetml_qnames.md`; remaining parser families and
structural validation stay open. Broader Strict object graphs and conversion remain OOTD-021.
