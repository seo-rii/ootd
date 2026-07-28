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

## Excel Oracle Boundary

The Windows Excel Oracle rejects all inventoried path, content-type, and relationship markers
before starting Excel. Its XML scan prohibits DTD processing and applies the existing per-entry
character bound. This intentionally favors false-positive refusal over activating an unreviewed
control, embedding, or callback inside the Oracle host.

External connections, query tables, linked-workbook refresh, data-model providers, and network
isolation belong to the separate OOTD-065 offline policy and are not claimed by this inventory.

## Remaining Scope

OOTD-022 is closed by selecting the safe `refuse` behavior instead of the former partial VBA-only
strip. OOTD-064 remains open for an explicit typed `strip` option, relationship/content-type closure
deletion, a deterministic audit manifest, policy-specific fixtures, and isolated real-Excel
evidence. No automatic strip capability is currently exposed.
