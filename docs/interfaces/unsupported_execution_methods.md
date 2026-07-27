# Unsupported Execution Method Contract

OOTD recognizes the method names and validates their public argument shapes, but it does not
currently configure refresh, spelling, fixed-format export, or print backends. These calls fail
closed with `OmErrorCode::Unsupported`; they never return `Empty` as if work had completed.

## Capability Matrix

| Backend | Methods that fail closed |
|---|---|
| Refresh | `Workbook.RefreshAll`, `Chart.Refresh` |
| Spelling | `Workbook.CheckSpelling`, `Worksheet.CheckSpelling`, `Chart.CheckSpelling` |
| Fixed-format export | `Workbook.ExportAsFixedFormat`, `Worksheet.ExportAsFixedFormat`, `Chart.ExportAsFixedFormat` |
| Print | `Workbook.PrintPreview`, `Workbook.PrintOut`, `Worksheet.PrintPreview`, `Worksheet.PrintOut`, and `PrintPreview`/`PrintOut` on `Worksheets`, `Sheets`, `Charts`, and `Chart` |

Chart sheets use the worksheet dispatch contract for spelling, fixed-format export, and print
methods. The failure message is stable:

```text
<Object>.<Method> is unsupported because no <capability> backend is configured
```

## Error Precedence

Arguments are validated before backend availability is reported. A malformed call therefore keeps
the existing `InvalidArgument` or `TypeMismatch` result, while a correctly shaped call reaches the
stable `Unsupported` boundary. Object and collection ownership are also resolved before returning
`Unsupported`, so stale handles and invalid chart-sheet bindings are not hidden by backend status.

Rejected calls do not create export/print artifacts, mutate selection, or change workbook dirty
domains. `Chart.Export` is a separate Boolean-returning contract: the current headless path returns
`false`, which explicitly reports that no image was exported rather than reporting success.

Support can be enabled only after a backend produces a testable file, state, or event result and
has both synthetic behavior tests and the applicable desktop Excel Oracle evidence.
