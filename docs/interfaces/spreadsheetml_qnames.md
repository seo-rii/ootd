# SpreadsheetML QName And Prefix Preservation

OOTD matches the typed SpreadsheetML surface by expanded XML name. Prefix spelling is not part of
the semantic identity of a workbook, worksheet, shared-string item, or relationship identifier.

## Implemented Boundary

The first QName-aware codec slice covers:

- workbook `workbookPr`, `calcPr`, `sheets`/`sheet`, `definedNames`/`definedName`, and the
  relationship-ID attribute;
- shared-string `si` and `t` elements;
- worksheet cell data: `dimension`, `sheetData`, `row`, `c`, `f`, `v`, `is`, and `t`.

Elements are recognized only when both their namespace URI and local name match the loaded
Strict or Transitional dialect. Unqualified attributes are likewise distinguished from qualified
attributes. A workbook sheet relationship ID may use any prefix bound to the dialect's Office
Document relationships namespace; it is not limited to the conventional `r:id` spelling.

A foreign namespace element with the same local name is not interpreted as SpreadsheetML. Its raw
XML remains part of the preservation template, so a targeted cell edit does not turn, for example,
a vendor `v` element into a cell cached value.

## Write Contract

No-op save retains unchanged workbook, shared-string, and worksheet part bytes.

When a typed mutation requires XML generation, the writer derives the qualified name from the
nearest typed owner already present in the source:

- workbook metadata inherits the workbook element prefix;
- a generated worksheet dimension inherits the worksheet prefix;
- generated rows inherit the `sheetData` prefix;
- generated cells, formulas, cached values, inline-string containers, and text inherit their row
  or cell prefix.

Namespace declarations and opaque source fragments remain in place. The writer does not introduce
an unqualified typed element into a source subtree that uses a valid SpreadsheetML prefix.

## Verification

The synthetic regression uses independent `book`, `link`, `text`, `ws`, and locally redeclared
`data` prefixes, plus foreign same-local-name poison elements. It proves clean byte preservation,
existing-cell replacement, new inline-string and formula-cell insertion, calculation metadata
insertion, save/reopen, and exclusion of foreign text/value nodes.

No desktop Excel observation is pinned yet. This boundary remains synthetic and partial.

## Remaining Work

OOTD-015/OOTD-049 remain open for parser families outside this slice, stricter parent/root
structure validation, unknown-prefix and duplicate-expanded-attribute diagnostics, and a shared
owner-aware raw-fragment splice abstraction. OOTD-029 now provides QName-strict
`[Content_Types].xml` parsing; OOTD-030 still owns package relationship root/namespace parsing.
OOTD-061 owns Markup Compatibility, `AlternateContent`, and `extLst` preservation semantics.
