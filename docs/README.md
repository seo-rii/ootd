# OOTD Documentation

This directory is the documentation entry point for `ootd`.

Start here if you want to understand the repository before jumping into crate code.
Some detailed implementation notes currently remain in Korean because they were written directly alongside development work.

## Core Entry Points

- [Repository overview](../README.md)
- [Current compatibility status](../STATUS.md)
- [Active roadmap](../ROADMAP.md)
- [Historical implementation plan](../PLAN.md)
- [Contributing guide](../CONTRIBUTING.md)

## Current Runtime Surfaces

- [Excel runtime calculation surface](interfaces/excel_runtime_calculation.md): current formula evaluator scope, supported functions, and known gaps.

## Spec Intake And Contracts

- [Spec roots](spec_roots.md): source-of-truth references for XLSX, OPC, and Excel object model intake.
- [OM source acquisition](specs/om_source_acquisition.md): how Excel COM type library and PIA inputs are pinned and normalized.
- [Windows capture runner](specs/windows_capture_runner.md): capture bundle layout, output contract, and current execution boundary.
- [Differential regression reports](specs/differential_regression_reports.md): Oracle/runtime comparison report and gate summary JSON contracts.

## Repository Data Sources

- [`../specs/`](../specs/): canonical schemas, generated assets, and pinned templates used by the codebase.
- [`../specs/pinned/`](../specs/pinned/): versioned contract templates and pinned capture placeholders.
- `../fixtures/`: reserved for the tracked corpus and golden inputs introduced by M1. Current
  synthetic workbooks are generated inside crate tests, so a clean clone does not yet contain
  tracked fixture files.

## Background Material

- [`../excel_compatibility_bundle/`](../excel_compatibility_bundle/): original bundle of architecture notes and references that informed the current workspace split.

## Reserved Documentation Areas

The following directories exist for focused documentation as the project grows:

- `architecture/`
- `interfaces/`
- `test-protocols/`

When a change adds a new long-lived design boundary, prefer placing the document under one of those directories instead of growing a single catch-all note.
