# OPC Content Types Manifest

OOTD treats `[Content_Types].xml` as typed OPC package metadata. Declaration names are resolved by
namespace URI and local name; prefix spelling has no semantic meaning.

## Root Contract

The document root must be `Types` in
`http://schemas.openxmlformats.org/package/2006/content-types`. A default namespace or any bound
prefix is accepted. A missing namespace, wrong namespace, wrong local name, missing root, or second
root returns a deterministic `Parse` error before package content types are assigned.

## Declaration Contract

Only direct children of the root in the package content-types namespace are interpreted as typed
declarations:

- `Default` requires unqualified, non-empty `Extension` and present `ContentType` attributes;
- `Override` requires unqualified `PartName` and `ContentType` attributes;
- an override `PartName` must be absolute and pass the canonical OPC part-name validator;
- `Default` extensions are keyed case-insensitively and duplicates are rejected;
- `Override` part names are keyed by canonical package identity, so case- or percent-equivalent
  duplicates are rejected;
- typed declarations may use empty-element syntax or an empty start/end pair, but nested elements,
  non-whitespace text, and CDATA are rejected.

Unknown root attributes and unknown direct-child extension subtrees remain opaque in the original
part bytes. A `Default` or `Override` local name nested inside such a subtree is not interpreted and
cannot replace a direct typed declaration.

## Verification

Synthetic package regressions cover default and prefixed manifests, wrong root/namespace cases,
case-equivalent `Default` duplicates, canonical `Override` duplicates, required attributes,
absolute part names, nonempty typed declarations, and a nested same-local-name poison declaration.
The existing ZIP/XML resource preflight still runs before this parser.

This contract is synthetic, not Oracle-verified. Media-type grammar validation and a separate
repair mode remain future work; package relationship root/namespace validation is tracked by
OOTD-030.
