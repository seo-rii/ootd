# Workbook Dirty Domains

`ExcelRuntime::workbook_dirty_domains` exposes the save-relevant state of an open workbook without
collapsing it into Excel's prompt-facing `Workbook.Saved` flag. This is the current OOTD-046
contract.

## Domains

| Domain | Meaning |
|---|---|
| `prompt_dirty` | Closing may require a save decision. `Workbook.Saved` reads and writes only this domain. |
| `semantic_dirty` | A supported workbook, worksheet, cell, name, chart, or drawing mutation changed workbook semantics. |
| `serialization_dirty` | Saving the current runtime would rewrite at least one model, calculation, or package artifact. This is derived from the other runtime and codec states. |
| `formula_cache_dirty` | Calculation changed a cached formula result or dynamic-array materialization. This does not by itself change `Workbook.Saved`. |
| `package_graph_dirty` | A relationship, part, content-type binding, or stale calculation-chain artifact must change. |
| `external_refresh_dirty` | External refresh output is pending serialization. No external refresh command is supported yet, so this remains `false`; OOTD-065 owns activation of this domain. |

## State Transitions

| Operation | Prompt | Semantic | Serialization | Formula cache | Package graph |
|---|---:|---:|---:|---:|---:|
| Load/open a verified baseline | false | false | false | false | false |
| `Workbook.Saved = false` | true | unchanged | unchanged | unchanged | unchanged |
| `Workbook.Saved = true` | false | unchanged | unchanged | unchanged | unchanged |
| Supported content/model mutation | true | true | true | unchanged | only when its package closure changes |
| Calculation with changed cached results | unchanged | unchanged | true | true | true only when a stale calculation chain must be removed |
| Calculation metadata update without cache changes | unchanged | unchanged | true | false | true only when a stale calculation chain must be removed |
| Snapshot-only `save_workbook` | unchanged | unchanged | unchanged | unchanged | unchanged |
| Successful `SaveCopyAs` | unchanged | unchanged | unchanged | unchanged | unchanged |
| Successful durable `Save`, `SaveAs`, `Close(save)`, or host-writer save | false | false | false | false | false |
| Any failure before runtime baseline commit | unchanged | unchanged | unchanged | unchanged | unchanged |

Close-with-discard removes the workbook from the runtime, so no post-close dirty snapshot exists.
`Application.Calculation` can make calculation properties serialization-dirty without changing the
prompt flag; its exact Excel prompt semantics remain Oracle-gated.

## Save Transaction Boundary

Filesystem saves prepare and verify the next package, perform the same-directory durable write and
atomic replace/create-new sequence, then commit the runtime baseline. Failures at temporary-file
creation, write, flush, file sync, replacement, or parent-directory sync preserve the complete
pre-call dirty-domain snapshot. A parent-directory sync failure may leave a valid replaced output,
but the live runtime remains retryable and dirty because it has not committed the new baseline.

Host-writer saves follow the same rule: both `write_all` and `flush` must succeed before the
baseline and dirty domains are cleared.

## Current Command Ownership

- Range/cell, worksheet, defined-name, chart, drawing, and workbook-structure mutation paths use
  the shared semantic marker.
- Recalculation writeback owns `formula_cache_dirty`; it does not set prompt or semantic state.
- Pending relationship graphs, copied chart support parts, and stale calculation-chain removal own
  `package_graph_dirty`.
- Serialization state is derived rather than independently toggled, preventing a prompt setter
  from clearing data that still needs to be written.
- Only successful baseline commit clears all domains and consumes the calculation snapshot.

`RefreshAll` and other external-refresh members remain part of OOTD-013/OOTD-065. Until they have
an observable backend or host callback, they must not claim an external-refresh state transition.
