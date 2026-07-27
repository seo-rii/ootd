# OPC Digital Signature Policy

OOTD can read an OOXML workbook that carries OPC digital-signature artifacts, but it does not
validate, remove, replace, or create signatures. Every package rewrite is fail-closed so a caller
cannot silently leave an invalid signature attached to modified workbook content.

## Detection Boundary

On load, the XLSX codec builds a `DigitalSignatureInventory` from:

- parts whose first path segment is `_xmlsignatures`, compared case-insensitively;
- the OPC digital-signature origin, XML-signature, and certificate content types; and
- relationship parts containing the OPC digital-signature `origin`, `signature`, or `certificate`
  relationship type.

The standard `schemas.openxmlformats.org` relationship URIs and the legacy
`schemas.microsoft.com` package URI spellings are recognized. The relationship types are the
package-wide identifiers documented by
[Microsoft's `Package.GetRelationshipsByType` reference](https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging.package.getrelationshipsbytype).

The public inventory exposes sorted part URIs and relationship-part URIs. It reports artifacts,
not cryptographic validity: OOTD does not currently parse `SignedInfo`, validate digests or
certificates, evaluate trust, or distinguish a valid signature from a malformed/orphan artifact.
Content-type-only orphan artifacts are therefore still inventoried and protected.

## Rewrite Policy

`XlsxCodec::save`, direct runtime serialization, host-writer save, `Workbook.Save`,
`Workbook.SaveAs`, `Workbook.SaveCopyAs`, and `Workbook.Close(SaveChanges=true)` return:

```text
code: SignedPackageMutationUnsupported
message: XlsxCodec::save refuses to rewrite packages containing OPC digital-signature artifacts
```

The error occurs during package preparation, before any writer or filesystem target is touched.
The source bytes, open workbook, and dirty domains remain unchanged. This applies to clean saves as
well as semantic edits because OOTD does not yet have a verification step that can prove a rewrite
preserved the signed package contract.

The source inventory is retained separately and the current package is scanned again before every
save. Replacing or removing public package parts manually therefore cannot bypass the source
signature policy. Closing with an explicit discard does not rewrite the package and remains
allowed.

## Remaining Scope

Explicit signature removal needs a typed option, complete relationship/content-type closure
removal, and an audit manifest. Signature validation needs digest/reference transforms,
certificate-chain and trust policy, timestamp handling, and pinned real-Excel signed fixtures.
Re-signing requires a separate key-provider boundary. None of those capabilities is currently
claimed.
