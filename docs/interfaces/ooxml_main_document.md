# OOXML Workbook Main-Document Discovery

OOTD discovers the workbook main document through the package-level `officeDocument`
relationship in `/_rels/.rels`. It does not assume that the part is named `xl/workbook.xml`.

## Load Contract

- `[Content_Types].xml` and `/_rels/.rels` must exist.
- Exactly one recognized Strict or Transitional `officeDocument` relationship must be present at
  the package root.
- Its target must be internal, normalize to a canonical package part, and exist.
- The workbook root namespace must match the relationship dialect, and its resolved main content
  type must be valid for that dialect. Unknown or contradictory combinations fail closed.
- The workbook relationship-part URI is calculated from the discovered owner. For example,
  `documents/book/main.xml` owns `documents/book/_rels/main.xml.rels`.
- Relative workbook relationships are resolved against the discovered owner's parent path.

Missing, external, dangling, or duplicate main-document relationships fail with deterministic parse
errors. Root relationship root/namespace validation belongs to OOTD-030 and is not claimed here.

## Save Contract

No-op and targeted saves rewrite the discovered workbook part, never an invented
`xl/workbook.xml`. The package-root relationship bytes and calculated workbook relationship part are
kept in the lossless support snapshot. Calculation-chain invalidation also uses the discovered owner
and the relationship-resolved calculation-chain part, so relocated metadata does not become stale.
Runtime format retagging and SaveAs update the discovered main part's content-type override and
in-memory OPC content type instead of inventing an `xl/workbook.xml` entry.

The Transitional and Strict synthetic regressions use `documents/book/main.xml`, a worksheet
relationship that climbs back to `xl/worksheets/sheet1.xml`, and a relocated calculation chain.
They cover sniff, load, no-op save, single-cell edit, calculation metadata rewrite, chain removal,
and reopen. Strict-specific format and relationship rules are documented in
`docs/interfaces/ooxml_dialect.md`.

## Remaining Dialect Scope

Strict and Transitional main-document discovery is implemented, including exact relationship,
namespace, and content-type agreement. General Strict graph mutation and cross-dialect conversion
remain OOTD-021 work. Arbitrary XML prefixes below the validated roots remain OOTD-015/OOTD-049.
