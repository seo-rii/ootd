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

## Package Closure At Save

The final Preserve and Refuse serialization boundary enumerates the package root relationship part
and every recognized owner relationship part, including root-level owners. Each document is parsed
with the contract above. A canonical relationship-part owner derivation distinguishes exactly the
package-root identity from `prefix/_rels/name.rels` owner parts. Only the package root is ownerless;
every other relationship part must resolve to an existing canonical non-relationship owner. Empty
owner stems, `[Content_Types].xml` owners, relationship-on-relationship ownership, `.rels` or
`_rels` placement outside that grammar (including single-segment `.rels` and nested reserved
`_rels` directories), and relationships MIME at a non-relationship URI return `InvalidState`.

Every normalized internal target must resolve to an existing canonical OPC part; otherwise save
returns `InvalidState` identifying the relationship part, relationship ID, and resolved missing
target. Relationship-part and owner classification uses the same case/percent-normalized OPC
identity as package lookup, so alternate unreserved percent spellings cannot bypass either gate. An
external target never requires a ZIP part.

Active-content Strip intentionally applies this gate after cleanup. This lets Strip remove an
active relationship marker whose target was already absent, while ensuring that neither a missed
relationship nor removal of an active target leaves an ownerless part or dangling internal edge in
the returned package. Its traversal uses the same owner derivation, so root-level and percent-aliased
relationship parts participate in cleanup before the final gate.

## Verification

Synthetic regressions cover default and prefixed roots/entries, wrong and missing namespaces,
wrong roots, foreign same-local entries, nested same-local entries, required attributes, duplicate
IDs, target modes, internal normalization, external inclusion/filtering, opaque dangling targets,
parent-relative closure, external target exemption, canonical owner aliases, root alias exemption,
empty/manifest/relationship owners, single-segment `.rels`, nested reserved `_rels` directories,
misplaced relationships MIME, and post-Strip owner/target closure. The complete XLSX suite
exercises package-root, workbook, worksheet, drawing, chart, pivot, and opaque relationship graphs.

This contract is synthetic, not Oracle-verified. Markup Compatibility policy and broader shared
QName/raw-fragment infrastructure remain OOTD-061/OOTD-049 work. Chart/drawing/support ownership
graphs remain a separate OOTD-031 validation stage.
