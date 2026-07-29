# Chart And Drawing Graph Save Validation Contract

OOTD treats decoded and newly materialized chart/drawing state as a typed view over an OPC
relationship graph. Public model fields are still mutable, so every chart graph entry point must
fail closed before it rewrites drawing or chart XML.

## Validation Boundaries

`materialize_state_only_chart_graphs` validates the workbook model and the complete chart/drawing
graph even when no state-only graph needs allocation. `XlsxCodec::save` validates the graph again
against the actual package selected for serialization, after worksheet/package binding and loaded
support-snapshot checks and before dirty drawing or chart parts are rewritten.

The save-time validator receives the serialization package separately from
`LoadedXlsxWorkbook::package`. This matters after a state-only graph has been materialized and the
package has been moved out of the temporary workbook state.

## Model Identity And Ownership

The preflight requires all of the following:

- every drawing and chart map key equals the model object's ID;
- each drawing, chart, and chart frame belongs to the current workbook;
- every drawing host sheet exists, and every chart frame's host matches its drawing host;
- chart-frame IDs are unique, drawing-local non-visual IDs are unambiguous, and each frame refers
  to an existing chart;
- every typed chart has at least one drawing host;
- chart-sheet bindings refer to their chart sheet, drawing, and primary chart consistently; and
- state-only and materialized package bindings are complete rather than partially populated.

A package-bound chart sheet without a typed chart binding remains a valid preserve-only state. It
represents a loaded drawingless or otherwise unmodeled chart sheet and is not treated as a request
to invent a chart. A materialized chart may also have multiple drawing hosts; this preserves a
shared loaded chart part. The materializer still refuses to create a new state-only chart with
multiple hosts because allocating that shared graph is not implemented.

## OPC Relationship Graph

For every materialized drawing and chart, the preflight resolves `raw_part_uri` through
`OpcPackage` and requires the target part to exist. Resolved canonical part identity has exactly one
typed model owner, so case or percent aliases cannot let two drawing/chart records claim the same
package part.

Each materialized drawing must have exactly one internal drawing relationship from its host sheet's
owner-relative `.rels` part to the resolved drawing part. Each materialized chart frame must retain
a `drawing-part#relationship-id` binding whose relationship exists, is an internal chart
relationship, and resolves to the chart model's package part. These checks form the typed
sheet-to-drawing-to-chart ownership chain.

Loaded support snapshots are checked first so a missing or changed preserved part keeps its more
specific diagnostic. Package-wide content-type coherence and internal relationship closure remain
the final OPC gates. Active-content Strip runs only after the preserved graph has passed this
preflight; the stripped package then passes the existing post-cleanup OPC gates.

## Deliberate Follow-up Boundaries

This contract does not require a `sheet_drawing_support_parts` entry for every typed graph. Newly
materialized graphs do not rebuild the loaded preservation snapshot, so snapshot-key and snapshot
owner coherence remain a separate OOTD-031 stage. Exact host-XML `<drawing r:id>` ownership beyond
the relationship edge also remains with that support-snapshot boundary.

Public DTO fields remain open until OOTD-054; this preflight prevents those mutations from being
serialized but does not prevent their temporary construction. Strict OOXML chart/drawing decoding
and mutation remain unsupported, and the current evidence is synthetic rather than desktop Excel
Oracle evidence.
