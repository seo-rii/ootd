# Workbook.SaveAs Capability Contract

OOTD validates the public Excel `Workbook.SaveAs` signature before preparing or writing a target.
Only behavior that is implemented, or an explicit value equivalent to omission, is accepted.
Everything else fails closed with `OmErrorCode::Unsupported`.

The argument order and defaults follow the official Microsoft
[`Workbook.SaveAs`](https://learn.microsoft.com/en-us/office/vba/api/Excel.Workbook.SaveAs)
contract. `xlNoChange = 1` comes from
[`XlSaveAsAccessMode`](https://learn.microsoft.com/en-us/office/vba/api/excel.xlsaveasaccessmode);
`xlUserResolution = 1` comes from
[`XlSaveConflictResolution`](https://learn.microsoft.com/en-us/office/vba/api/excel.xlsaveconflictresolution).

## Capability Matrix

| Argument | Accepted | Rejected before write |
|---|---|---|
| `Filename` | A string path | Missing filename or a non-string value |
| `FileFormat` | Omitted, or a supported integral OOXML `XlFileFormat` value | Invalid types/values and unsupported formats |
| `Password` | Omitted or an empty string | Any non-empty password until OOXML encryption exists |
| `WriteResPassword` | Omitted or an empty string | Any non-empty write-reservation password |
| `ReadOnlyRecommended` | Omitted or `false` | `true` |
| `CreateBackup` | Omitted or `false` | `true` |
| `AccessMode` | Omitted or `xlNoChange` (`1`) | `xlShared`, `xlExclusive`, or any other value |
| `ConflictResolution` | Omitted or `xlUserResolution` (`1`) | Automatic local/other-session resolution or any other value |
| `AddToMru` | Omitted or `false` | `true` |
| `TextCodepage` | Omitted | Any provided value |
| `TextVisualLayout` | Omitted | Any provided value |
| `Local` | Omitted or `false` | `true` |

`Missing`, `Empty`, and `Null` are treated as omitted optional values. Explicit accepted defaults
do not claim that OOTD implements the corresponding feature; they are accepted because they do not
request behavior beyond the existing save path.

The rejection matrix verifies the stable error code/message, absence of a target file, unchanged
source identity, and unchanged workbook dirty domains for every unsupported argument.
