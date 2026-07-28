# Workbooks.Open Capability Contract

OOTD validates every public `Workbooks.Open` argument before reading the source path. The argument
order follows Microsoft's official
[`Workbooks.Open`](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open)
contract. Normal load value `xlNormalLoad = 0` follows
[`XlCorruptLoad`](https://learn.microsoft.com/en-us/office/vba/api/excel.xlcorruptload).

## Capability Matrix

| Argument | Accepted | Rejected before read |
|---|---|---|
| `Filename` | A string path | Missing filename or a non-string value |
| `UpdateLinks` | Omitted, `0`, or `false`; OOTD performs no external link update | `3`, `true`, or any other nonzero value |
| `ReadOnly` | Omitted/`false` or `true`; this controls runtime mutation and save policy | Non-boolean values |
| `Format` | Omitted | Any provided text-import format |
| `Password` | Omitted or an empty string | Any non-empty password until encrypted OOXML support exists |
| `WriteResPassword` | Omitted or an empty string | Any non-empty write-reservation password |
| `IgnoreReadOnlyRecommended` | Omitted or `false` | `true` |
| `Origin` | Omitted | Any provided text-file platform |
| `Delimiter` | Omitted | Any provided delimiter |
| `Editable` | Omitted or `false` | `true` |
| `Notify` | Omitted or `false` | `true` |
| `Converter` | Omitted | Any converter index |
| `AddToMru` | Omitted or `false` | `true` |
| `Local` | Omitted or `false` | `true` |
| `CorruptLoad` | Omitted or `xlNormalLoad` (`0`) | `xlRepairFile`, `xlExtractData`, or any other value |

`Missing`, `Empty`, and `Null` are treated as omitted optional values. Explicit default-equivalent
arguments are accepted without claiming the associated feature is implemented.

Excel normally prompts for link-update policy when `UpdateLinks` is omitted. OOTD is a headless,
offline runtime, so omission deterministically selects `ExternalDataPolicy::OfflinePreserve` and
means no external access. The caller can inspect the package-only inventory and attempt flags or
use the typed host API's `Refuse` policy; see `docs/interfaces/external_data.md`. Until an audited
host callback exists, any request to update links is `Unsupported`.

The rejection matrix uses a nonexistent source path to prove all unsupported options fail before
filesystem read. It also locks the exact error code/message and verifies that no workbook enters
the collection.
