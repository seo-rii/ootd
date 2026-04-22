# Contributing

OOTD is still in an early build-out phase. Contributions are most useful when they tighten compatibility contracts, improve the current Excel vertical slice, or reduce ambiguity in the object model intake path.

## Before You Start

Read these documents first:

- [README.md](README.md)
- [PLAN.md](PLAN.md)
- [docs/README.md](docs/README.md)
- [docs/spec_roots.md](docs/spec_roots.md)

If your work changes how object model data is acquired or normalized, also read:

- [docs/specs/om_source_acquisition.md](docs/specs/om_source_acquisition.md)
- [docs/specs/windows_capture_runner.md](docs/specs/windows_capture_runner.md)

## Development Expectations

- Keep the application object model boundary separate from file-format codecs.
- Preserve opaque XLSX/package content whenever the current architecture claims a lossless-first path.
- Treat `specs/pinned/` as repository contracts, not scratch files.
- Prefer small, focused changes over sweeping refactors across multiple crates.
- Update documentation when you change a boundary, capture contract, or repository workflow.

## Toolchain

- Rust `1.85` or newer

## Common Commands

Run the full workspace tests:

```bash
cargo test --workspace --quiet
```

GitHub Actions CI currently runs the locked workspace test command:

```bash
cargo test --workspace --locked --quiet
```

Run focused tests while iterating:

```bash
cargo test -p office-idl
cargo test -p office-codegen
cargo test -p office-capture
cargo test -p excel-model
cargo test -p excel-xlsx
cargo test -p excel-runtime
```

Format the workspace before sending a change:

```bash
cargo fmt --all
```

## Scope Guidance

Good contribution areas right now:

- Excel object model capture and normalization contracts,
- XLSX compatibility and round-trip preservation,
- runtime surface expansion that matches the pinned model,
- documentation that reduces ambiguity in repository direction or source-of-truth boundaries.

Areas that still need careful design before broad implementation:

- full Excel automation surface coverage,
- behavior-complete calculation and recalc semantics,
- non-Excel Office application support.

## Tests

Add targeted regression coverage when you change behavior or a repository contract. The most useful tests are the ones that lock down a concrete compatibility rule, capture shape, or persistence guarantee.

## Pull Requests

When opening a change, keep the description specific about:

- what behavior or contract changed,
- which crate or boundary owns the change,
- what verification you ran.
