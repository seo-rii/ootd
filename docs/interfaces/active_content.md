# Active-Content Inventory And Save Policy

OOTD preserves active-content package artifacts when the target format can carry the source
contract. It does not silently remove a partial subset when converting a macro-enabled workbook
to a non-macro target. Such a conversion is fail-closed until the caller can choose an explicit,
audited strip policy.

## Inventory Boundary

On load, the XLSX codec creates an `ActiveContentInventory` for these categories:

- VBA project, VBA data, and classic/Agile/V3 VBA project signatures;
- Excel 4 macro sheets (including international macro sheets) and dialog sheets;
- ActiveX control XML/binary data and control properties;
- OLE objects and embedded packages; and
- Office 2006/2007 custom UI and user customization relationships.

The inventory combines three independent marker classes:

- canonical and case-insensitive package paths such as `xl/vbaProject.bin`, `xl/activeX/`,
  `xl/embeddings/`, and `customUI/`;
- active content types declared in `[Content_Types].xml`, including orphan declarations; and
- active relationship types found in any relationship part, including Transitional and currently
  relevant Strict relationship spellings.

The public inventory exposes sorted kinds, matching part URIs, relationship-part URIs, and whether
the content-type manifest declared an active marker. This is a structural security inventory, not
proof that a binary payload is executable, well formed, signed, or safe.

The identifiers follow Microsoft's published package contracts for
[ActiveX controls](https://learn.microsoft.com/en-us/openspecs/office_file_formats/MS-XLSB/71dd26ec-9725-49e7-83e7-52a8213b492e),
[VBA signatures](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/8fbb98c8-bf03-429e-b6e8-ae024539b5b9),
[Excel 4 macro sheets](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-offmacro/b8bee527-ef5a-4734-bb8c-6eae4166b6c9), and
[Ribbon extensibility](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/52faf7b6-fecc-48d9-96db-ee80a631a5ac).
Embedded packages are included because Microsoft documents that Office can activate them through
[Windows OLE technologies](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/32be961f-d71a-4812-913f-b675c79aa88a).

## Save Policy

Same-format saves, copies, and macro-capable `SaveAs` targets preserve active-content part bytes,
relationships, and content-type declarations through the existing lossless package path.

When an XLSM or XLTM source has any source or current active-content marker and the requested
target is XLSX or XLTX, save preparation returns:

```text
code: ActiveContentConversionUnsupported
message: macro-enabled workbook conversion to a non-macro format requires an explicit active-content policy
```

The failure happens before target creation or runtime baseline commit. Clean and dirty workbooks
retain their format, object identity, active-content inventory, and dirty domains. The source
inventory is retained independently and the current package is scanned again, so direct removal or
late insertion of public package parts cannot bypass the conversion boundary.

Macro-enabled sources with no active-content marker can still be retagged. OOTD does not infer that
a macro-capable file necessarily contains executable content.

Hosts that intentionally want a sanitized snapshot use
`ExcelRuntime::save_workbook_with_active_content_policy` (or the codec-level
`XlsxCodec::save_with_active_content_audit`) with one of three typed policies:

- `Preserve` is the default. It preserves artifacts for compatible targets and retains the
  fail-closed macro-to-non-macro conversion boundary above.
- `Refuse` rejects any source/current active-content marker with stable
  `ActiveContentPolicyRefused`, even when the target could preserve it.
- `Strip` serializes the current semantic edits, removes the active package graph, and then retags
  a requested non-macro target. The runtime method returns bytes without committing them as the
  open workbook baseline, matching the existing snapshot-only `save_workbook` contract.

`Strip` discovers internal descendants through every canonical relationship part. Active roots and
exclusive descendants are removed with their relationship parts. A descendant with an incoming
owner outside the removal set is retained, while every incoming edge to a removed root is deleted.
For retained XML owners, elements whose relationship-namespace `id` points at a deleted edge are
removed; this covers workbook macro/dialog sheet entries and worksheet control/OLE anchors in the
synthetic contract corpus. Active and removed-part overrides, active defaults, and orphan active
content-type/relationship declarations are removed as part of the same transformation. The
resulting bytes are reloaded before they are returned from the prepared-save path.

Every successful policy call returns a deterministic `ActiveContentAuditManifest`. It records the
sorted detected categories, removed part URI/content type/byte length, removed relationship owner,
ID/type/target/mode, removed Default/Override entries, rewritten XML owners, and shared descendants
that were intentionally retained. Repeating a strip against the same snapshot produces identical
bytes and an equal manifest in the synthetic regression corpus.

## Excel Oracle Boundary

The Windows Excel Oracle rejects all inventoried path, content-type, and relationship markers
before starting Excel. Its XML scan prohibits DTD processing and applies the existing per-entry
character bound. This intentionally favors false-positive refusal over activating an unreviewed
control, embedding, or callback inside the Oracle host.

External connections, query tables, linked-workbook refresh, data-model providers, and network
isolation belong to the separate OOTD-065 offline policy and are not claimed by this inventory.

## Remaining Scope

OOTD-022 is closed by selecting the safe default refusal instead of the former partial VBA-only
strip. OOTD-064 now has typed preserve/refuse/strip behavior, ownership-aware package closure
deletion, deterministic audit output, orphan-marker cases, shared-descendant cases, arbitrary
relationship-prefix owner cleanup, and macro/dialog-sheet removal in synthetic fixtures.

OOTD-064 remains `Partial` until a pinned desktop Excel corpus proves that complex real-world
ActiveX/VML drawing anchors, form controls, OLE previews, custom UI callback graphs, and mixed
macro/dialog sheet metadata open/save/reopen without a repair dialog. Signed packages continue to
hit the independent signature-mutation refusal before active-content stripping; signature removal
or re-signing requires the explicit OOTD-063 follow-up policy. The Excel Object Model `SaveAs`
surface does not invent a non-Excel optional argument: hosts must select destructive stripping
through the typed runtime/codec API.
