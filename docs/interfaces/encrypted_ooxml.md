# Encrypted OOXML Capability Contract

OOTD does not yet decrypt or encrypt OOXML workbooks. It does distinguish an encrypted OOXML
compound-file container from an ordinary invalid ZIP or legacy compound workbook before OPC/ZIP
parsing.

## Detection Boundary

The detector validates the CFB v3/v4 signature and sector geometry, collects the FAT through the
header DIFAT and any DIFAT sector chain, follows the directory stream through the FAT, and walks
the root storage's reachable child/sibling tree. Orphan directory entries do not count. A container
is classified as encrypted OOXML only when that root tree contains:

- a root storage entry;
- a non-empty `EncryptionInfo` stream entry; and
- a non-empty `EncryptedPackage` stream entry.

Traversal, visited-sector tracking, and allocations are bounded by the input's physical sector
count. A legacy compound workbook containing a `Workbook` stream, or a compound file containing
only one encryption stream, is not classified as encrypted OOXML.

## Public Behavior

`XlsxCodec::load`, direct runtime open, and `Workbooks.Open` with an omitted or empty `Password`
return:

```text
code: EncryptedWorkbookUnsupported
message: encrypted OOXML compound-file containers are not supported
```

The result is produced after codec-option validation but before OPC/ZIP parsing. The failed open
does not add a workbook to the runtime.

Non-empty `Workbooks.Open Password` and `Workbook.SaveAs Password` values remain fail-closed at
their existing public argument boundaries, before input read or output preparation respectively.
They are never silently ignored. Write-reservation passwords remain a separate unsupported
capability.

`XlsxCodec::sniff` returns `false` because encrypted bytes are not directly loadable by the XLSX
codec. The typed load error is the authoritative diagnostic.

## Remaining Stage

Agile Encryption decryption/encryption, password verification, integrity checking, and encrypted
SaveAs round trips remain unsupported. The inner package kind (`xlsx`, `xlsm`, template variants)
cannot be reported until successful decryption, so this first stage makes no inner-format claim.
Those capabilities require password-result and real Excel corpus cases before OOTD-062 can close.
