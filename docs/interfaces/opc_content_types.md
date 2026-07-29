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

- `Default` requires unqualified, non-empty `Extension` and non-empty `ContentType` attributes;
- `Override` requires unqualified `PartName` and non-empty `ContentType` attributes;
- an override `PartName` must be absolute and pass the canonical OPC part-name validator;
- `Default` extensions are keyed case-insensitively and duplicates are rejected;
- `Override` part names are keyed by canonical package identity, so case- or percent-equivalent
  duplicates are rejected;
- `Default` lookup derives the extension from the final path segment of that same canonical
  identity, so percent aliases such as `%78ml` and `%2Exml` cannot bypass the `xml` declaration and
  a dotted parent directory cannot give an extensionless part a type;
- typed declarations may use empty-element syntax or an empty start/end pair, but nested elements,
  non-whitespace text, and CDATA are rejected.

Unknown root attributes and unknown direct-child extension subtrees remain opaque in the original
part bytes. A `Default` or `Override` local name nested inside such a subtree is not interpreted and
cannot replace a direct typed declaration.

## Save Coherence Contract

XLSX save does not rely on the optional preservation snapshot to establish manifest validity.
Before active-content policy inventory, `OpcPackage::validate_content_type_cache` performs the
bounded manifest preflight and strict parse again, then compares every non-manifest part's cached
content type with the manifest result. Both sides may be absent at this early policy boundary, but
a present/missing or unequal value is `InvalidState`; a fabricated active cache therefore cannot
change Preserve/Refuse/Strip classification.

Immediately before Preserve/Refuse serialization, `validate_content_types_for_save` additionally
requires all of the following:

- the canonical `[Content_Types].xml` part exists and remains within default XML resource limits;
- every actual package part other than the manifest resolves to a non-empty `Default` or `Override`
  content type;
- every cached content type exactly equals that resolved value; and
- every canonical `Override` identity names an existing package part and does not name the manifest
  itself.

The manifest item itself is excluded from part coverage and cache comparison. Active-content
`Strip` deliberately applies the complete gate after removing active parts and their manifest
entries, so repairable orphan active markers can be cleaned while a stale non-active declaration
or uncovered remaining part can never reach output. Override cleanup uses canonical package
identity, including case and percent aliases. Workbook protection and final sheet-visibility checks
use the main part discovered from the package-root `officeDocument` relationship, including
relocated Transitional and Strict workbooks.

## Verification

Synthetic package regressions cover default and prefixed manifests, wrong root/namespace cases,
case-equivalent `Default` duplicates, canonical `Override` duplicates, required attributes,
absolute part names, nonempty typed declarations, canonical percent-aliased extensions, and a
nested same-local-name poison declaration. Dotted-directory/extensionless lookup and manifest-self
Override regressions close the final selector edge cases. Save regressions cover missing manifests,
uncovered parts, orphan overrides, cache drift before policy inventory, Preserve/Refuse final
validation, post-Strip validation, canonical-alias cleanup, and relocated Transitional/Strict
workbook cleanup. The existing ZIP/XML resource preflight runs on load and is repeated for public
in-memory manifest mutation at the save boundary.

This contract is synthetic, not Oracle-verified. Media-type grammar validation and a separate
repair mode remain future work. Package relationship root/namespace validation is now covered by
OOTD-030 and documented in `opc_relationships.md`.
