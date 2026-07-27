# XLSX Codec Option Contract

`XlsxCodec::load` and `XlsxCodec::save` accept only option combinations that have an implemented
policy. Unsupported profiles and lossy/skip modes fail before package parsing or serialization.

## LoadOptions

| Field | Supported value | Current behavior for other values |
|---|---|---|
| `profile` | `Excel365` | `Excel2016` and `Excel2021` return `Unsupported` |
| `preserve_unknown_parts` | `true` | `false` returns `Unsupported`; OOTD has no reviewed destructive-part filter |
| `read_calc_chain` | `true` | `false` returns `Unsupported`; calc-chain inventory remains required for coherent save invalidation |

## SaveOptions

| Field | Supported value | Current behavior for other values |
|---|---|---|
| `profile` | `Excel365` | `Excel2016` and `Excel2021` return `Unsupported` |
| `lossless` | `true` | `false` returns `Unsupported`; no canonical/lossy writer contract exists |

The profile restriction is a capability statement, not a claim that Excel 2016 or Excel 2021
files are categorically unreadable. OOTD has not yet defined profile-dependent namespaces,
features, or Oracle gates, so accepting those enum values would currently be a silent no-op.

Regression coverage verifies each unsupported field/profile produces a stable error before invalid
input reaches OPC parsing and before a loaded workbook reaches serialization. Default options
continue through the existing lossless package-preservation path.
