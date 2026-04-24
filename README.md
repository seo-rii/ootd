# OOTD

OOTD is a Rust workspace for building an Office-style application stack around an application object model rather than a file-format-only API.

The current repository focuses on the Excel slice first. The near-term goal is not a complete spreadsheet product, but a compatibility-oriented core that can:

- preserve XLSX packages with a lossless-first bias,
- normalize machine-readable Excel object model inputs,
- generate code and metadata from pinned contracts, and
- expose a minimal runtime surface for workbook/session and range operations.

## Status

This project is still in an early implementation phase.

Implemented so far:

- pinned documentation and capture contracts for Excel object model source acquisition,
- `office-idl` loading for canonical schema-backed metadata,
- `office-codegen` paths for summary generation and capture bundle normalization,
- `office-opc` package loading with raw entry preservation,
- `excel-model` workbook, worksheet, and cell state,
- `excel-xlsx` load/save support for a focused vertical slice,
- `excel-runtime` orchestration for workbook/session and range access,
- `excel-runtime` formula evaluation for arithmetic, comparisons, logical/control helpers, aggregates, criteria aggregates, lookup/reference and reference metadata helpers, date/time serial helpers, error/info helpers, focused text helpers, and `Application`/`Worksheet.Evaluate`,
- `office-capture` support for Windows capture bundle planning and materialization.

Not implemented yet:

- a complete Excel object model facade,
- full Excel calculation parity beyond the current formula subset,
- a fully pinned real-world Windows capture bundle checked into the repository,
- broader Office applications beyond the current Excel-first slice.

## Workspace Layout

- `crates/office-idl`: schema-backed types and JSON loaders for canonical Office IDL datasets.
- `crates/office-codegen`: code generation and capture-bundle normalization helpers.
- `crates/office-common`: shared primitives, ids, metadata, and error types.
- `crates/office-opc`: Open Packaging Conventions support with raw package entry preservation.
- `crates/office-capture`: Windows capture bundle planning, output layout, and execution artifacts.
- `crates/excel-model`: canonical workbook, worksheet, and cell state.
- `crates/excel-xlsx`: focused XLSX reader/writer support on top of the model and OPC layers.
- `crates/excel-runtime`: runtime orchestration and workbook/session APIs.
- `docs/`: repository documentation and spec intake notes.
- `specs/`: canonical schemas, pinned templates, and generated spec assets.
- `fixtures/`: synthetic and golden inputs for regression coverage.
- `excel_compatibility_bundle/`: archived background bundle that informed the current layout.

## Getting Started

### Requirements

- Rust `1.85` or newer

### Common Commands

```bash
cargo test --workspace --quiet
```

Run a focused crate test pass while iterating:

```bash
cargo test -p excel-xlsx
cargo test -p office-capture
```

## Continuous Integration

GitHub Actions runs the workspace test suite on pushes to `main`, pull requests, and manual dispatches.

```bash
cargo test --workspace --locked --quiet
```

## Documentation

- [Documentation index](docs/README.md)
- [Phase 1 plan](PLAN.md)
- [Excel runtime calculation surface](docs/interfaces/excel_runtime_calculation.md)
- [Spec roots](docs/spec_roots.md)
- [OM source acquisition](docs/specs/om_source_acquisition.md)
- [Windows capture runner contract](docs/specs/windows_capture_runner.md)

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development expectations and contribution workflow.

## License

MIT. See [LICENSE](LICENSE).
