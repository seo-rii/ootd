# Behavioral Excel Oracle Protocol

Milestone M1 compares a versioned operation DSL through desktop Excel and `ExcelRuntime`. The
cross-platform Rust contract lives in `crates/excel-oracle`; the desktop runner lives in
`tools/excel-oracle-win`.

## Artifact Layers

The protocol deliberately separates four artifact kinds.

1. A case pins its profile, input path, exact input SHA-256, provenance, ordered operations, and
   complete probes.
2. A run manifest pins the actual engine fingerprint and exactly one record for every suite case.
3. An observation contains typed semantic values only: `Void`, `Missing`, `Empty`, `Null`, Boolean,
   finite number, text, cell error, symbolic object, and rectangular row-major array.
4. The existing `office-codegen` differential report is a summary/gate. Its artifact links point to
   the typed Oracle and runtime observations rather than flattening them into strings.

Case, input, and observation hashes cover the exact checked-in bytes. Missing, extra,
case-insensitively duplicated, skipped, failed, or unsupported `mustMatch` cases block the gate.
Native COM messages and HRESULTs remain diagnostic data; comparison uses the canonical error kind
and code so locale-specific messages cannot change a golden result. Number comparison is exact by
default and tolerances must be opted into per comparison policy.

The checked-in suite root uses `manifest/suite_manifest.json`; case and input paths are relative to
that root. Each captured run root uses `manifest/run_manifest.json`; observation paths are relative
to the run root. `PinnedSuiteArtifacts` loads this graph without following a root, parent, or file
symlink, rejects non-regular artifacts and non-portable Windows aliases, and validates the exact
case, input, and observation bytes before replay. JSON case/manifest artifacts are capped at 16 MiB,
observations at 64 MiB, and individual workbook inputs at 512 MiB.

`verify_repeated_excel_runs` promotes only completed observations from two independently named run
manifests whose exact engine fingerprint matches the suite's pinned desktop Excel profile. Every
`mustMatch` case must complete in both runs; failed cases, status drift, reused run IDs (including
ASCII case aliases), and canonical typed-value drift reject the evidence. The repeated comparison
is exact by default while still ignoring native diagnostic message text in the same way as the
Excel-to-OOTD comparator.

The Windows watchdog still emits one run root per case. `load_run_fragment` accepts a non-empty,
suite-declared subset while applying the same manifest, hash, engine, path, and typed-observation
checks as a complete run. `assemble_run_fragments` requires every fragment to share one exact run
ID, profile, and engine fingerprint; it rejects missing or duplicate case records and revalidates
observation bytes before ordering records by the suite. Completed observations receive canonical
`observations/<case-id>/oracle.json` paths in the assembled in-memory bundle.

`write_run_bundle` refuses an existing destination or symlinked parent, revalidates complete suite
coverage plus every canonical observation path, hash, size, engine, and typed case before creating
output, and materializes into a unique sibling temporary directory. Observation files are
create-new, flushed, and synced first; the manifest is written last. A final destination recheck
precedes one directory rename, and any pre-publication error removes the temporary root, so callers
never receive a partially populated run root.

The cross-platform assembly command exposes that same contract without weakening it:

```text
cargo run -p excel-oracle -- capture-plan --suite-root PATH

cargo run -p excel-oracle -- assemble-run \
  --suite-root PATH \
  --fragment-root PATH \
  --fragment-root PATH \
  --output-root FRESH_PATH
```

`capture-plan` validates the suite manifest contract and exact-hash-validates every case and input
before emitting suite-ordered relative case/input paths, artifact hashes, and the expected Excel
profile. It emits no plan when any artifact is missing, unsafe, oversized, or changed. A Windows
suite launcher consumes this preflight output before starting any case watchdog.

It accepts at most 4,096 fragment roots, emits a deterministic JSON receipt on success, returns exit
code 2 for command-line contract errors and exit code 1 for artifact validation or publication
errors. It does not launch Excel or the Windows watchdog; every fragment must already exist and the
set must exactly cover the pinned suite before the fresh output root is created.

## Excel Execution

One runner process owns and records every Excel Application it activates. The host must have no
pre-existing Excel process.
The runner configures automation security and manual calculation before opening the case-local
workbook copy, executes every operation, records OM errors without aborting later probes, and then
closes Excel even after failure.

Save cases are reopened in a new Excel session with `CorruptLoad=xlNormalLoad`. A successful normal
open is recorded as `repairDetected=false`; failure, crash, or timeout records an unknown repair
state and cannot pass a required case. The runner never retries with `xlRepairFile`.

ZIP bytes are not compared directly. Future corpus snapshots normalize semantic workbook state,
XML, and the OPC part/relationship graph while ignoring entry order, compression output, and ZIP
timestamps.

## Current Verification Boundary

Rust contract, manifest, bounded filesystem/fragment loader, exact suite capture-plan preflight,
deterministic suite-run assembler, atomic publisher and assembly CLI, repeated-capture gate,
comparator, report bridge, `ExcelRuntime` adapter, .NET contract normalization, and fake-backed
runner lifecycle tests execute in this repository. The preflight and assembly commands do not yet
iterate the Windows watchdog. The COM session and watchdog compile on Linux but require a real
Windows Excel host for execution. No real Excel observation or 20-case required corpus is pinned as
of 2026-08-11, so no behavior is yet Oracle-verified.
