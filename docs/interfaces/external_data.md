# External Data and Offline-Open Contract

OOTD treats external workbook links, DDE/OLE links, workbook connections, query tables, and data
model parts as package metadata and cached data. Loading and saving these artifacts does not grant
permission to contact their targets, start a provider, execute a query, or refresh a cache.

## Inventory

`XlsxCodec::load` creates a deterministic `ExternalDataInventory` from the package only. Detection
currently covers:

- `xl/externalLinks/*` parts and `externalBook`, `ddeLink`, and `oleLink` elements;
- `xl/connections.xml`, `xl/queryTables/*`, `xl/model/*`, and `xl/customData/*` parts;
- the corresponding known content types, including orphan declarations; and
- external-link/path, connections, query-table, model, and Excel long-path relationship markers,
  including external targets without an internal package part.

The inventory exposes sorted kinds, part URIs, relationship-part URIs, and relationship metadata.
`LoadedXlsxWorkbook::source_or_current_external_data_inventory` merges the immutable load-time
inventory with a fresh scan of the public package, so replacing or adding a package part cannot
erase the original security observation.

External formula-token classification and every vendor-specific Power Query/Data Model extension
are not yet complete. Unknown parts remain subject to the normal lossless-preservation contract.

## Runtime Policy

`ExcelRuntime::open_workbook` uses `ExternalDataPolicy::OfflinePreserve`:

- the input bytes are parsed locally;
- cached external-link values, connection definitions, query-table metadata, and model bytes are
  preserved through unrelated edits;
- no link update, refresh, provider activation, DDE/OLE execution, filesystem target read, or
  network request is attempted; and
- the open workbook exposes `workbook_external_data_access_report`, whose update, refresh, and
  external-access attempt flags remain `false`.

Security-sensitive hosts can call `open_workbook_with_external_data_policy` with `Refuse`. If any
source/current marker is present, open returns stable `ExternalDataPolicyRefused` before allocating
or registering a workbook handle. A marker-free workbook can still be opened under that policy.

The Excel Object Model `Workbooks.Open` surface does not invent a host-only argument. Omitted,
`0`, or `false` `UpdateLinks` selects the default offline behavior. Any nonzero/`true` update request
is `Unsupported` before the source file is read. `Workbook.RefreshAll` is likewise `Unsupported`;
the rejected call does not mark a refresh or external-access attempt and does not set
`external_refresh_dirty`.

## Windows Excel Oracle Boundary

The desktop Excel Oracle uses a stricter `Refuse` boundary. Before constructing an Excel COM
session, it scans both the caller's source package and the copied sandbox input with bounded ZIP
and XML readers. Known external-link, connection, query-table, and Data Model paths, content types,
and relationship types are rejected. This includes Transitional and Strict external-link
relationships plus Excel long-path relationship extensions; an ordinary external hyperlink remains
allowed by this classification.

Each decision is atomically recorded as `manifest/preflight/source_input.json` or
`manifest/preflight/sandbox_input.json`. The audit states the active-content and external-data
policies, whether the package was eligible for Excel activation, size metrics for accepted inputs,
the input SHA-256 when the bounded regular file can be read, and the stable rejection reason for
denied inputs. A rejected source therefore leaves evidence without activating Excel. The network
field records required host isolation rather than claiming that the runner configured it. These
controls supplement, rather than replace, the required offline and disposable Windows host profile.

## Save Contract

Serialization never refreshes external data. Default lossless save preserves cached external-data
parts and relationships byte-for-byte during an unrelated cell edit in the synthetic corpus. The
active-content and digital-signature policies remain independent: external-data inventory neither
strips active content nor bypasses signed-package rewrite refusal.

## Remaining Scope

OOTD-065 remains `Partial`. Before a refresh capability can be enabled, it needs a host callback or
provider boundary with explicit allowlists, credential isolation, timeouts, cancellation, audit
events, and deterministic `external_refresh_dirty` transitions. A pinned Oracle corpus must still
cover linked workbooks, DDE/OLE, ODBC/OLE DB/web/text connections, QueryTables, Power Query, Data
Model, external pivot sources, and locale-specific link paths. Preflight marker coverage must grow
with that corpus; it is not a claim that desktop Excel is safe outside the isolated host profile.
