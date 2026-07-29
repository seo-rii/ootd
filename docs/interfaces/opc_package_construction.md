# OPC Package Construction Contract

`OpcPackage::try_new` is the public constructor for an in-memory OPC part collection. Construction
is atomic: every part name is checked and every canonical identity is collected before the package
is returned.

## Part Identity Rules

- Part names use OPC package-relative spelling. One leading `/` is accepted and removed.
- Empty names, empty path segments, dot segments, trailing dots, forbidden URI characters, and
  malformed or forbidden percent encodings return `OmErrorCode::InvalidArgument`.
- Identity comparison is ASCII case-insensitive and normalizes percent-encoded unreserved
  characters. Case-equivalent and percent-equivalent duplicates are rejected.
- `OpcPackage::canonical_part_identity` exposes that exact fallible identity transform for package
  graph validators, so part classification cannot drift from lookup and duplicate detection.
- Successful construction preserves part order, content type, compression method, and bytes while
  storing the normalized package-relative name.

`OpcPackage::from_bytes` applies the same constructor after ZIP/XML loading and maps an in-memory
identity failure to `Parse`. `add_part` and `to_bytes` use the same canonical-identity validator;
serialization therefore remains defensive against any internally corrupted package state.

There is no public infallible `OpcPackage::new`. Repository fixtures with pinned part names call
`try_new(...).expect(...)`; production paths either propagate the structured error or use a
compile-time fixture whose failure is a programming error.

## Remaining OOTD-031 Work

Construction alone still closes only part-name and identity invariants. `OpcPart` payload fields,
`OpcPackage::default`, and raw add/replace/remove mutations may temporarily produce an incomplete
graph, because generic OPC assembly and active-content cleanup need staged mutation. XLSX policy
inventory now calls the bounded `validate_content_type_cache`, and final Preserve/Refuse or
post-Strip serialization calls `validate_content_types_for_save` to require manifest/part/cache
coherence, reject manifest-self Overrides, and resolve Default extensions from only the canonical
final path segment. Strip protects the root-relationship-discovered workbook main part rather than
assuming `xl/workbook.xml`. Internal relationship-target closure is enforced at the same final XLSX
boundaries, where every non-root relationship part also requires a canonical existing
non-relationship owner, while single-segment `.rels` and nested reserved `_rels` directories fail
closed. Generic OPC construction remains staged; the stricter owner and target graph is an XLSX
final-save contract. Chart/drawing/support ownership graphs, workbook/model topology, and public
field encapsulation remain separately tracked by OOTD-031 and OOTD-054.

The current evidence is synthetic; no desktop Excel Oracle claim is attached to this constructor
contract.
