# Workbook Sheet Record Parse Policy

OOTD treats workbook sheet records as package graph declarations. A sheet is not admitted into the
model unless its identity, display name, and workbook relationship all validate together.

## Implemented Strict Boundary

The default XLSX load and save-validation paths require every typed SpreadsheetML `sheet` record
to provide:

- a `name` attribute that is not blank, has at most 31 characters, and contains no control
  character or `:`, `\`, `/`, `?`, `*`, `[`, or `]`;
- a decimal `sheetId` in the inclusive range `1..=4294967295`;
- an `id` attribute in the loaded dialect's Office Document relationships namespace;
- a matching internal workbook relationship whose type is worksheet, chart sheet, dialog sheet,
  or the supported macro-sheet relationship type.

Within one workbook, sheet IDs, ASCII-case-insensitive sheet names, and relationship IDs must each
be unique. The relationship target becomes the sheet part URI only after these checks succeed.
Save validation reparses the actual workbook relationship part, so it applies the same graph
contract as load.

Malformed records return `OmErrorCode::Parse`. OOTD does not synthesize sequential IDs, `SheetN`
names, default worksheet kinds, missing relationship IDs, or empty part URIs.

An internal save-rewrite phase may encounter duplicate source-XML names after multiple validated
runtime renames or copies while the in-memory model already owns the final unique names. That phase
relaxes only the duplicate-name check long enough to rewrite records by sheet/relationship identity.
The rewritten workbook part is then reparsed under the full strict policy before ZIP serialization;
every other required-field, range, relationship, and duplicate check remains active throughout.

## Repair Boundary

There is no implicit repair mode. Unsupported `Workbooks.Open` repair options already fail before
the source file is read, and the codec exposes no repair policy that could opt into invention. A
malformed workbook therefore produces a deterministic parse error and no runtime workbook handle.

A future explicit repair mode must return a structured report identifying every correction before
it may admit a repaired model. It must not weaken the default path or silently reuse these former
fallbacks.

## Remaining Boundaries

This slice validates the semantic identity and relationship closure of workbook sheet records. It
does not yet claim full SpreadsheetML content-model ordering or parent-path validation; those stay
with OOTD-015/OOTD-049. Defined-name repair and scope fidelity stay with OOTD-078, while malformed
worksheet row/cell attributes and duplicate coordinates are tracked by OOTD-017.

The current evidence is synthetic. Desktop Excel open/save/reopen observations remain required
before this behavior is marked Oracle-verified.
