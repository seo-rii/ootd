# Differential Regression Reports

This document fixes the JSON artifact boundary used after Windows capture and
runtime execution have produced comparable Oracle/runtime case results.

The capture runner is responsible for producing the raw Office OM capture
bundle. Differential regression artifacts start after that point: a runner
compares Excel Oracle behavior with `ootd` runtime behavior, writes a report,
derives a gate summary, and lets CI consume the gate summary.

## Source Registry Context

Every CI-grade differential report must be tied to the current source registry.
The registry is [sources.toml](/home/seorii/dev/hancomac/ootd/specs/sources.toml).

The report context is derived from `SourceRegistrySummary` and includes:

- `projectName`
- `defaultProfile`
- `defaultMode`
- `primaryOmArtifact`
- `primaryOoxmlSource`
- `enabledCorpusGroups`
- `enabledCorpusSourceCount`
- `validationModes`

`office-codegen` provides these entry points:

- `summarize_source_registry_toml`
- `differential_artifact_contract`
- `build_differential_report_with_source_context`
- `validate_differential_report_source_context`

CI runners should not compute a gate summary from a report that lacks this
context. Context mismatch is a contract error, not a test failure.
`validate_differential_report_source_context` validates the report contract
before comparing the source registry context.

`summarize_source_registry_toml` rejects malformed registry machine keys before
building the context. `project.name` is a non-blank ASCII token.
`project.default_profile`, `project.default_mode`, OM artifact ids, OOXML
source ids, and profile keys are non-blank ASCII identifiers without leading or
trailing whitespace; `project.default_profile` must match a declared profile
key. Auxiliary `binary_formats` and `behavior` map keys are ASCII identifiers
without ASCII case-insensitive collisions, and their values are non-blank ASCII
tokens. Profile keys also cannot collide under ASCII case-insensitive matching.
The registry must also enable at least one corpus group, provide enough enabled
corpus sources to cover those groups, and enable at least one validation mode.
Corpus groups with child source toggles must enable at least one child source
when the group is enabled, and cannot leave child sources enabled when the group
is disabled.

## Differential Report JSON

The report contains:

- `library`
- `version`
- `profile`
- `context`
- `caseCount`
- `statusCounts`
- `cases`

Each case has:

- `name`
- `surface`
- `member`
- `status`
- `expected`
- `actual`
- `message`
- `artifacts`

Supported statuses are:

- `passed`
- `failed`
- `missingOracle`
- `missingRuntime`
- `unsupported`
- `skipped`

`office-codegen` validates that:

- `library`, `version`, and `profile` are non-blank and contain no leading or
  trailing whitespace; `library` and `version` are ASCII tokens, and `profile`
  is an ASCII identifier
- `caseCount` equals `cases.length`
- `cases` is non-empty
- `case.name` is non-blank, contains no leading or trailing whitespace, and is
  unique within the report under exact and ASCII case-insensitive matching
- present `case.surface` and `case.member` values are non-blank ASCII
  identifiers without leading or trailing whitespace
- case artifact keys are non-blank ASCII identifiers without leading or
  trailing whitespace and without ASCII case-insensitive collisions within a
  case; artifact paths are non-blank forward-slash relative paths without `.`,
  `..`, empty path segments, Windows-reserved portable path characters,
  device-name segments, trailing-dot segments, or leading/trailing whitespace on
  any path segment that stay within the report output root under an artifact
  subdirectory; case artifact paths are unique within the report under exact and
  ASCII case-insensitive matching and cannot use canonical differential
  artifact filenames as their top-level namespace
- `statusCounts` equals the status histogram reconstructed from `cases`
- present `context` string fields and list entries are non-blank, contain no
  leading or trailing whitespace, and context lists do not contain duplicates
  under exact or ASCII case-insensitive matching
- present `context.projectName` is an ASCII token
- present `context.defaultProfile`, `context.defaultMode`,
  `context.primaryOmArtifact`, and `context.primaryOoxmlSource` are ASCII
  identifiers
- present `context.enabledCorpusGroups` and `context.validationModes` entries
  are ASCII identifiers
- present `context.enabledCorpusGroups` and `context.validationModes` are
  non-empty, and `enabledCorpusSourceCount` covers the enabled corpus groups
- when `context` is present, `profile` equals `context.defaultProfile`
- `context` equals the current `SourceRegistrySummary`
- `profile` equals the registry default profile when using the context-aware
  gate path

Relevant APIs:

- `build_differential_report`
- `build_differential_report_with_source_context`
- `load_differential_report_from_json`
- `load_differential_report_from_path`
- `write_differential_report_to_path`

Use `build_differential_report` only for local or legacy fixtures that are not
intended to drive CI gating.

## Gate Summary JSON

The gate summary is the small downstream CI artifact derived from a validated
report.

It contains:

- `passed`
- `blockingCaseCount`
- `blockingCases`
- `incompleteOracleCount`
- `missingRuntimeCount`
- `failedCaseCount`
- `unsupportedCaseCount`
- `skippedCaseCount`

Blocking statuses are:

- `failed`
- `missingOracle`
- `missingRuntime`

Non-blocking statuses are:

- `passed`
- `unsupported`
- `skipped`

`office-codegen` validates that:

- `blockingCaseCount` equals `blockingCases.length`
- `passed` is true exactly when `blockingCaseCount == 0`
- `blockingCaseCount` equals `failedCaseCount + incompleteOracleCount +
  missingRuntimeCount`, with overflow rejected as a contract error
- `blockingCases` contains non-blank, trimmed case names that are unique under
  exact and ASCII case-insensitive matching

Relevant APIs:

- `summarize_differential_gate`
- `try_summarize_differential_gate`
- `summarize_differential_gate_with_source_context`
- `load_differential_gate_from_json`
- `load_differential_gate_from_path`
- `load_differential_gate_from_path_with_source_context`
- `write_differential_gate_to_path`
- `write_differential_gate_from_report_path_with_source_context`

CI should prefer `write_differential_gate_from_report_path_with_source_context`.
That path loads the report, validates stale report counts, validates source
registry context, writes the gate summary JSON, and returns the same summary to
the caller.
Use `try_summarize_differential_gate` when a runner already has an in-memory
report and still needs stale count validation before deriving a gate summary.

If the runner owns the output directory, it should prefer
`write_differential_report_and_gate_to_output_root`. That path writes both
canonical artifacts under the output root and returns the exact paths used. It
preflights report counts and source registry context before creating the output
directory, and rejects output roots that are files or symlinks plus canonical
artifact paths that are already directories or symlinks, so context mismatch or
path conflicts cannot leave a report-only partial artifact or follow artifact
links outside the output root.

Downstream CI can validate a completed output directory with
`load_differential_artifacts_from_output_root`. That path loads both canonical
artifacts, preflights canonical artifact paths, validates report context,
validates the gate summary, and rejects a stored gate summary that differs from
the report-derived gate.
The returned bundle exposes `passed()`, `blocking_case_count()`, and
`blocking_cases()` so CI code can make a pass/fail decision without re-reading
raw JSON fields.

## Artifact Flow

Canonical artifact filenames are exposed by
`differential_artifact_contract()`:

- `differential_report.json`
- `differential_gate_summary.json`

Runners can derive both paths under an output directory with
`differential_artifact_paths(output_root)`. This returns:

- `reportPath = output_root / differential_report.json`
- `gateSummaryPath = output_root / differential_gate_summary.json`

The intended flow is:

1. Load [sources.toml](/home/seorii/dev/hancomac/ootd/specs/sources.toml).
2. Build `SourceRegistrySummary`.
3. Compare Excel Oracle results and runtime results into case records.
4. Build a differential report with source context.
5. Write `differential_report.json`.
6. Load the report artifact through the context-aware gate path.
7. Write `differential_gate_summary.json`.
8. Downstream CI loads the gate summary and validates the gate invariant before
   making a pass/fail decision.

The current `office-codegen` layer intentionally does not run Excel, spawn the
runtime harness, or select corpus files. It fixes the artifact model and
validation gates that those later runners must use.

## Failure Policy

Contract errors mean the artifact is malformed, stale, generated from a
different registry, or unsafe to use for CI gating. They should fail the job
before interpreting case status.

Case status failures mean the artifact is well-formed and a real compatibility
result was observed. These are represented in `blockingCases`.

Unsupported and skipped cases are counted but do not make the gate fail.
