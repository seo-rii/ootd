# XLSX Support Snapshot Validation Contract

`LoadedXlsxWorkbook` keeps source bytes and parsed summaries for worksheet, drawing, and chart
parts so unrelated edits can preserve package content exactly. These snapshots are public mutable
state today, so load, graph materialization, and save validate their owner graph before using them
as a lossless-rewrite baseline.

## Validation Boundaries

`XlsxCodec::load` validates every collected snapshot's owners, inventories, and source-map shape
before returning the decoded workbook. It deliberately retains a source relationship whose target
part is missing so existing preserve/diagnostic flows can report that malformed package at their
policy boundary. The public chart/drawing materializer repeats validation with declared-target
checks before its no-op fast path and again after it invalidates source snapshots changed by
materialization.

Save performs the check in two ordered phases. Owner validation runs after the live worksheet
collection has been matched to the package but before dirty worksheet recovery or XML rewrite can
look up a snapshot by `SheetId`. Existing byte and parsed-summary comparisons then retain their
part-specific diagnostics. Finally, support-graph validation runs before typed chart/drawing graph
validation and serialization.

Snapshots describe the loaded source graph, so active-content Strip does not validate them again
after removing package parts. The stripped package still passes the independent content-type and
relationship-closure gates.

## Sheet And Part Ownership

Every `worksheet_support_parts` key must resolve to a live worksheet model, and every
`sheet_drawing_support_parts` key must resolve to a live sheet. A worksheet support snapshot can
only belong to `SheetKind::Worksheet`. In both maps, the stored host-part URI must exactly equal the
owning model sheet's part URI.

When a host relationships part is declared, its URI must be the owner-relative `.rels` URI derived
from the host part. A parsed worksheet relationship summary cannot outlive its source bytes, but
runtime copy/move may retain newly rewritten source bytes without manufacturing a parsed source
summary. Relationship source data cannot exist without a declared owner part.

## Drawing And Chart Inventory

The drawing relationship-ID list must exactly match the ordered binding IDs, and binding targets
must exactly match the drawing-part inventory. Every inventory rejects duplicate raw URIs,
malformed OPC names, and case/percent spellings that collapse to one canonical identity. Declared
drawing and chart relationship parts must be owner-relative to an inventoried drawing or chart.

For retained source snapshots:

- drawing source-byte keys and drawing-summary keys are identical subsets of drawing parts;
- drawing relationship source-byte keys are a subset of declared drawing relationship parts;
- drawing opaque source keys exactly match their inventory;
- chart source, summary, relationship, support, and opaque source keys exactly match their
  respective inventories; and
- at materialization and save boundaries, internal summary targets must be declared by the
  corresponding chart, support, or opaque inventory, while relationships with
  `TargetMode="External"` do not require package parts.

Chart summaries account exactly for declared chart relationship parts. Chart support and opaque
inventories, and drawing chart/opaque inventories, may retain additional historical or shared parts
that are not referenced by every retained summary; each retained internal summary edge must still
point into its inventory before mutation or serialization.

## Deliberate Preservation Exceptions

A sheet with no worksheet or drawing support entry is valid; absence means that no preservation
snapshot was collected for that domain. Newly materialized chart/drawing graphs are also not added
to the historical source snapshot. Materialization may invalidate a changed drawing's source bytes
and summary together, or a changed drawing relationship source independently, so those maps may be
proper subsets of their inventories.

Shared chart, style, and opaque package parts may legitimately appear in snapshots owned by more
than one sheet. This validator establishes each snapshot's sheet owner and internal coherence; the
typed chart/drawing and final OPC gates separately validate serialization ownership and package
closure. Drawingless package-bound chart sheets remain preserve-only.

## Follow-up Boundaries

Public snapshot and model fields remain mutable until OOTD-054 moves callers to validated
commands. Orphan `worksheet_data` is owned by OOTD-032, and atomic workbook-ID reassignment by
OOTD-033. Exact host XML ownership beyond the recorded relationship graph remains a later
QName/owner-aware rewrite concern. Current evidence is synthetic and does not replace the pinned
desktop Excel corpus required by OOTD-043 and OOTD-085.
