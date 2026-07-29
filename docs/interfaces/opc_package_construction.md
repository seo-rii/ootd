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
- Successful construction preserves part order, content type, compression method, and bytes while
  storing the normalized package-relative name.

`OpcPackage::from_bytes` applies the same constructor after ZIP/XML loading and maps an in-memory
identity failure to `Parse`. `add_part` and `to_bytes` use the same canonical-identity validator;
serialization therefore remains defensive against any internally corrupted package state.

There is no public infallible `OpcPackage::new`. Repository fixtures with pinned part names call
`try_new(...).expect(...)`; production paths either propagate the structured error or use a
compile-time fixture whose failure is a programming error.

## Remaining OOTD-031 Work

This stage closes only part-name and identity construction invariants. `OpcPart` payload fields,
`OpcPackage::default`, and raw add/replace/remove mutations do not yet prove that a package has a
valid `[Content_Types].xml` manifest, coherent cached content types, or a closed relationship graph.
Those checks belong to the next OOTD-031 save-validation stages. Workbook/model topology and public
field encapsulation remain separately tracked by OOTD-031 and OOTD-054.

The current evidence is synthetic; no desktop Excel Oracle claim is attached to this constructor
contract.
