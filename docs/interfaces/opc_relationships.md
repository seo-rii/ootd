# OPC Package Relationships

OOTD resolves package relationship XML by expanded name. The package namespace URI and tree
position determine whether an element is a typed relationship; prefix spelling alone never does.

## Root Contract

Every parsed `.rels` document must have a `Relationships` root in
`http://schemas.openxmlformats.org/package/2006/relationships`. A default namespace or any bound
prefix is accepted. A missing namespace, wrong namespace, wrong local name, missing root, second
root, direct non-whitespace text, or direct CDATA returns a deterministic `Parse` error.

## Entry Contract

A typed `Relationship` must be a direct child of the root in the package relationships namespace.
A same-local-name element in another namespace or below an extension wrapper is ambiguous and is
rejected. Other unknown attributes and extension elements remain opaque in the retained source
part, and non-`Relationship` inner XML inside a typed entry does not become another edge.

Each typed entry requires unqualified, non-empty `Id`, `Type`, and `Target` attributes. IDs are
unique across internal, included external, and filtered external entries. `TargetMode` accepts only
case-insensitive `Internal` or `External`; external targets are returned only when the caller opts
in, while still participating in duplicate-ID validation.

Internal targets pass the shared package-target normalizer. It resolves owner-relative and absolute
paths and rejects root escape, empty segments, forbidden URI characters, malformed percent
encoding, encoded separators, control bytes, and dot-ending segments. External targets remain
lexically unchanged.

## Verification

Synthetic regressions cover default and prefixed roots/entries, wrong and missing namespaces,
wrong roots, foreign same-local entries, nested same-local entries, required attributes, duplicate
IDs, target modes, internal normalization, and external inclusion/filtering. The complete XLSX
suite exercises package-root, workbook, worksheet, drawing, chart, pivot, and opaque relationship
graphs.

This contract is synthetic, not Oracle-verified. Markup Compatibility policy and broader shared
QName/raw-fragment infrastructure remain OOTD-061/OOTD-049 work.
