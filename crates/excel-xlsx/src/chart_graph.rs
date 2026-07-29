use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Write};

use excel_model::{DrawingObjectModel, WorksheetData};
use office_common::{ChartId, OmError, OmErrorCode, OmResult, SheetId, SheetKind, SheetVisibility};
use office_opc::{CompressionMethod, OpcPackage, OpcPart};
use quick_xml::escape::escape;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use super::{
    LoadedXlsxWorkbook, PendingChartRelationshipGraph, PendingDrawingRelationshipGraph,
    PendingPackagePart, PendingPackageRelationship, chart_object_anchor_xml,
    encode_chart_model_xml, normalize_relationship_target, xml_local_name,
};

const CONTENT_TYPES_PART_NAME: &str = "[Content_Types].xml";
const WORKBOOK_PART_NAME: &str = "xl/workbook.xml";
const WORKBOOK_RELS_PART_NAME: &str = "xl/_rels/workbook.xml.rels";
const CHARTSHEET_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const DRAWING_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const CHART_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const CHARTSHEET_PART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
const DRAWING_PART_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CHART_PART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const RELATIONSHIPS_PART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-package.relationships+xml";
const CONTENT_TYPES_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/package/2006/content-types";
const PACKAGE_RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships";
const SPREADSHEET_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const OFFICE_RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SPREADSHEET_DRAWING_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const DRAWING_MAIN_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const CHART_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";

pub(crate) fn has_state_only_chart_graphs(workbook: &LoadedXlsxWorkbook) -> bool {
    !workbook.pending_drawing_relationship_graphs.is_empty()
        || !workbook.pending_chart_relationship_graphs.is_empty()
        || workbook
            .state
            .worksheets
            .iter()
            .any(|sheet| sheet.kind == SheetKind::ChartSheet && sheet.part_uri.is_none())
        || workbook
            .state
            .chart_sheets
            .values()
            .any(|binding| binding.raw_part_uri.is_none())
        || workbook
            .state
            .drawings
            .values()
            .any(|drawing| drawing.raw_part_uri.is_none())
        || workbook.state.drawings.values().any(|drawing| {
            drawing.objects.iter().any(|object| match object {
                DrawingObjectModel::ChartFrame(chart_object) => chart_object.raw_binding.is_none(),
                DrawingObjectModel::UnsupportedRaw { .. } => false,
            })
        })
        || workbook
            .state
            .charts
            .values()
            .any(|chart| chart.raw_part_uri.is_none())
}

/// Allocates package parts and relationships for chart models that exist only in WorkbookState.
///
/// Existing part names and relationship ids are preserved. New identifiers are allocated from
/// the first free sequential value in deterministic model order.
pub fn materialize_state_only_chart_graphs(
    mut workbook: LoadedXlsxWorkbook,
) -> OmResult<LoadedXlsxWorkbook> {
    materialize_state_only_chart_graphs_in_place(&mut workbook)?;
    Ok(workbook)
}

pub(crate) fn materialize_state_only_chart_graphs_in_place(
    workbook: &mut LoadedXlsxWorkbook,
) -> OmResult<()> {
    workbook.state.validate_for_save()?;
    validate_chart_graphs_for_save(workbook, &workbook.package)?;
    if !has_state_only_chart_graphs(workbook) {
        return Ok(());
    }

    normalize_state_only_non_visual_ids(workbook)?;
    let mut content_types_xml = workbook
        .package
        .part(CONTENT_TYPES_PART_NAME)
        .ok_or_else(|| {
            OmError::new(
                OmErrorCode::Parse,
                format!("workbook package is missing {CONTENT_TYPES_PART_NAME}"),
            )
        })?
        .bytes
        .clone();
    let mut used_part_names = workbook
        .package
        .parts()
        .iter()
        .map(|part| part.name.clone())
        .collect::<BTreeSet<_>>();
    used_part_names.extend(content_type_override_part_names(&content_types_xml)?);

    materialize_chart_sheet_shells(workbook, &mut used_part_names, &mut content_types_xml)?;
    materialize_pending_existing_chart_relationship_graphs(
        workbook,
        &mut used_part_names,
        &mut content_types_xml,
    )?;
    materialize_new_drawings(workbook, &mut used_part_names, &mut content_types_xml)?;
    materialize_charts_in_existing_drawings(
        workbook,
        &mut used_part_names,
        &mut content_types_xml,
    )?;

    workbook
        .package
        .replace_part_bytes(CONTENT_TYPES_PART_NAME, content_types_xml)?;
    invalidate_changed_preservation_snapshots(workbook);
    ensure_all_state_only_charts_materialized(workbook)?;
    validate_chart_graphs_for_save(workbook, &workbook.package)
}

pub(crate) fn validate_chart_graphs_for_save(
    workbook: &LoadedXlsxWorkbook,
    package: &OpcPackage,
) -> OmResult<()> {
    let workbook_id = workbook.state.model.id;
    let sheets = workbook
        .state
        .worksheets
        .iter()
        .map(|sheet| (sheet.id, sheet))
        .collect::<BTreeMap<_, _>>();
    let mut chart_hosts = BTreeMap::new();
    let mut chart_object_ids = BTreeSet::new();
    let mut opaque_object_ids = BTreeSet::new();
    let mut materialized_part_owners = BTreeMap::new();

    for sheet in &workbook.state.worksheets {
        if sheet.workbook_id != workbook_id {
            return Err(OmError::invalid_state(format!(
                "sheet {} belongs to a different workbook",
                sheet.name
            )));
        }
        if sheet.kind == SheetKind::ChartSheet {
            let Some(binding) = workbook.state.chart_sheets.get(&sheet.id) else {
                if sheet.part_uri.is_some() && sheet.relationship_id.is_some() {
                    continue;
                }
                return Err(OmError::invalid_state(format!(
                    "chart sheet {} is missing its chart binding",
                    sheet.name
                )));
            };
            if binding.sheet_id != sheet.id {
                return Err(OmError::invalid_state(format!(
                    "chart sheet binding {} does not match host sheet {}",
                    binding.sheet_id.0, sheet.id.0
                )));
            }
            if !workbook.state.charts.contains_key(&binding.chart_id) {
                return Err(OmError::invalid_state(format!(
                    "chart sheet {} references missing chart {}",
                    sheet.name, binding.chart_id.0
                )));
            }
            if let Some(drawing_id) = binding.drawing_id
                && !workbook.state.drawings.contains_key(&drawing_id)
            {
                return Err(OmError::invalid_state(format!(
                    "chart sheet {} references missing drawing {}",
                    sheet.name, drawing_id.0
                )));
            }
            if sheet.part_uri.is_none() {
                if sheet.relationship_id.is_some() || binding.raw_part_uri.is_some() {
                    return Err(OmError::invalid_state(format!(
                        "state-only chart sheet {} has a partial package binding",
                        sheet.name
                    )));
                }
                if binding.drawing_id.is_none() {
                    return Err(OmError::invalid_state(format!(
                        "state-only chart sheet {} has no drawing binding",
                        sheet.name
                    )));
                }
            } else if sheet.relationship_id.is_none() || binding.raw_part_uri.is_none() {
                return Err(OmError::invalid_state(format!(
                    "materialized chart sheet {} has an incomplete package binding",
                    sheet.name
                )));
            } else if binding.raw_part_uri != sheet.part_uri {
                return Err(OmError::invalid_state(format!(
                    "chart sheet {} has mismatched package part bindings",
                    sheet.name
                )));
            }
        }
    }

    for (sheet_id, binding) in &workbook.state.chart_sheets {
        if binding.sheet_id != *sheet_id
            || sheets
                .get(sheet_id)
                .is_none_or(|sheet| sheet.kind != SheetKind::ChartSheet)
        {
            return Err(OmError::invalid_state(format!(
                "chart sheet binding {} has no matching chart sheet",
                sheet_id.0
            )));
        }
    }

    let mut drawing_hosts = BTreeMap::new();
    for (drawing_id, drawing) in &workbook.state.drawings {
        if drawing.id != *drawing_id {
            return Err(OmError::invalid_state(format!(
                "drawing map key {} does not match model id {}",
                drawing_id.0, drawing.id.0
            )));
        }
        if drawing.workbook_id != workbook_id {
            return Err(OmError::invalid_state(format!(
                "drawing {} belongs to a different workbook",
                drawing_id.0
            )));
        }
        let host = sheets.get(&drawing.host_sheet_id).ok_or_else(|| {
            OmError::invalid_state(format!(
                "drawing {} references missing host sheet {}",
                drawing_id.0, drawing.host_sheet_id.0
            ))
        })?;
        drawing_hosts
            .entry(drawing.host_sheet_id)
            .or_insert_with(Vec::new)
            .push((*drawing_id, drawing.raw_part_uri.is_some()));

        if let Some(part_uri) = drawing.raw_part_uri.as_deref() {
            let drawing_part = package.part(part_uri).ok_or_else(|| {
                OmError::invalid_state(format!(
                    "drawing {} package part is missing: {part_uri}",
                    drawing_id.0
                ))
            })?;
            let canonical_part_identity = OpcPackage::canonical_part_identity(&drawing_part.name)?;
            let owner = format!("drawing {}", drawing_id.0);
            if let Some(existing_owner) =
                materialized_part_owners.insert(canonical_part_identity, owner.clone())
            {
                return Err(OmError::invalid_state(format!(
                    "package part {} is owned by both {existing_owner} and {owner}",
                    drawing_part.name
                )));
            }

            let host_part_uri = host.part_uri.as_deref().ok_or_else(|| {
                OmError::invalid_state(format!(
                    "materialized drawing {} has an unbound host sheet {}",
                    drawing_id.0, host.name
                ))
            })?;
            let host_relationships_part_uri = relationships_part_uri_for(host_part_uri);
            let host_relationship_count = package
                .part(&host_relationships_part_uri)
                .map(|part| {
                    super::parse_relationship_entries_for_part(part.bytes.as_slice(), host_part_uri)
                })
                .transpose()?
                .unwrap_or_default()
                .iter()
                .filter(|relationship| {
                    relationship.relationship_type == DRAWING_RELATIONSHIP_TYPE
                        && relationship.target_mode.is_none()
                        && package
                            .part(&relationship.target)
                            .is_some_and(|target_part| target_part.name == drawing_part.name)
                })
                .count();
            if host_relationship_count != 1 {
                return Err(OmError::invalid_state(format!(
                    "drawing {} is not owned by host sheet {} relationships: expected one relationship to {}, found {host_relationship_count}",
                    drawing_id.0, host.name, drawing_part.name
                )));
            }
        }
        let pending_relationship_graph =
            workbook.pending_drawing_relationship_graphs.get(drawing_id);
        if drawing.raw_part_uri.is_some() && pending_relationship_graph.is_some() {
            return Err(OmError::invalid_state(format!(
                "materialized drawing {} must not retain a pending relationship graph",
                drawing_id.0
            )));
        }

        let mut drawing_chart_ids = BTreeSet::new();
        let mut materialized_non_visual_ids = BTreeSet::new();
        let mut opaque_relationship_ids = BTreeSet::new();
        for object in &drawing.objects {
            let chart_object = match object {
                DrawingObjectModel::ChartFrame(chart_object) => chart_object,
                DrawingObjectModel::UnsupportedRaw {
                    id,
                    raw_anchor_xml,
                    relationship_ids,
                    non_visual_id,
                    ..
                } => {
                    if !opaque_object_ids.insert(*id) {
                        return Err(OmError::invalid_state(format!(
                            "opaque drawing object id {} is duplicated",
                            id.0
                        )));
                    }
                    validate_opaque_anchor_xml(raw_anchor_xml)?;
                    for relationship_id in relationship_ids {
                        opaque_relationship_ids.insert(relationship_id.clone());
                    }
                    if let Some(non_visual_id) = non_visual_id
                        && !materialized_non_visual_ids.insert(*non_visual_id)
                    {
                        return Err(OmError::invalid_state(format!(
                            "drawing {} contains duplicate materialized non-visual id {}",
                            drawing_id.0, non_visual_id
                        )));
                    }
                    continue;
                }
            };
            if !chart_object_ids.insert(chart_object.id) {
                return Err(OmError::invalid_state(format!(
                    "chart object id {} is duplicated",
                    chart_object.id.0
                )));
            }
            if chart_object.workbook_id != workbook_id
                || chart_object.host_sheet_id != drawing.host_sheet_id
            {
                return Err(OmError::invalid_state(format!(
                    "chart object {} does not match drawing {} ownership",
                    chart_object.id.0, drawing_id.0
                )));
            }
            if !drawing_chart_ids.insert(chart_object.chart_id) {
                return Err(OmError::invalid_state(format!(
                    "drawing {} references chart {} more than once",
                    drawing_id.0, chart_object.chart_id.0
                )));
            }
            if chart_object.raw_binding.is_some() {
                let non_visual_id = match chart_object.non_visual_id {
                    Some(non_visual_id) => non_visual_id,
                    None => u32::try_from(chart_object.id.0).map_err(|_| {
                        OmError::invalid_state(format!(
                            "materialized chart object id {} cannot be encoded as a DrawingML non-visual id",
                            chart_object.id.0
                        ))
                    })?,
                };
                if !materialized_non_visual_ids.insert(non_visual_id) {
                    return Err(OmError::invalid_state(format!(
                        "drawing {} contains duplicate materialized non-visual id {}",
                        drawing_id.0, non_visual_id
                    )));
                }
            }
            let chart = workbook
                .state
                .charts
                .get(&chart_object.chart_id)
                .ok_or_else(|| {
                    OmError::invalid_state(format!(
                        "chart object {} references missing chart {}",
                        chart_object.id.0, chart_object.chart_id.0
                    ))
                })?;
            if chart.workbook_id != workbook_id {
                return Err(OmError::invalid_state(format!(
                    "chart {} belongs to a different workbook",
                    chart.id.0
                )));
            }
            if drawing.raw_part_uri.is_none() && chart.raw_part_uri.is_some() {
                return Err(OmError::unsupported(format!(
                    "state-only drawing {} cannot bind existing chart {}",
                    drawing_id.0, chart.id.0
                )));
            }
            if chart.raw_part_uri.is_none() && chart_object.raw_binding.is_some() {
                return Err(OmError::invalid_state(format!(
                    "state-only chart object {} has a partial drawing binding",
                    chart_object.id.0
                )));
            }
            if chart.raw_part_uri.is_some()
                && drawing.raw_part_uri.is_some()
                && chart_object.raw_binding.is_none()
            {
                return Err(OmError::invalid_state(format!(
                    "materialized chart object {} is missing its drawing binding",
                    chart_object.id.0
                )));
            }
            if let (Some(chart_part_uri), Some(drawing_part_uri), Some(raw_binding)) = (
                chart.raw_part_uri.as_deref(),
                drawing.raw_part_uri.as_deref(),
                chart_object.raw_binding.as_deref(),
            ) {
                validate_materialized_chart_binding(
                    package,
                    chart_object.id,
                    drawing_part_uri,
                    chart_part_uri,
                    raw_binding,
                )?;
            }
            if chart.raw_part_uri.is_none()
                && matches!(
                    chart_object.anchor,
                    None | Some(office_common::DrawingAnchor::UnsupportedRaw)
                )
            {
                return Err(OmError::invalid_state(format!(
                    "state-only chart object {} has no serializable anchor",
                    chart_object.id.0
                )));
            }
            chart_hosts
                .entry(chart.id)
                .or_insert_with(Vec::new)
                .push(format!("drawing {}", drawing_id.0));
        }

        if drawing.raw_part_uri.is_none() && !opaque_relationship_ids.is_empty() {
            let pending_relationship_graph = pending_relationship_graph.ok_or_else(|| {
                OmError::unsupported(format!(
                    "state-only drawing {} requires relationship copying for opaque objects",
                    drawing_id.0
                ))
            })?;
            validate_pending_drawing_relationship_graph(
                *drawing_id,
                pending_relationship_graph,
                &opaque_relationship_ids,
            )?;
        } else if let Some(pending_relationship_graph) = pending_relationship_graph {
            validate_pending_drawing_relationship_graph(
                *drawing_id,
                pending_relationship_graph,
                &opaque_relationship_ids,
            )?;
        }

        if host.kind == SheetKind::ChartSheet {
            let binding = workbook
                .state
                .chart_sheets
                .get(&host.id)
                .expect("chart sheet binding validated above");
            if binding.drawing_id != Some(*drawing_id) {
                return Err(OmError::invalid_state(format!(
                    "chart sheet {} does not bind drawing {}",
                    host.name, drawing_id.0
                )));
            }
            let chart_frames = drawing
                .objects
                .iter()
                .filter_map(|object| match object {
                    DrawingObjectModel::ChartFrame(chart_object) => Some(chart_object),
                    DrawingObjectModel::UnsupportedRaw { .. } => None,
                })
                .collect::<Vec<_>>();
            let primary_chart_count = chart_frames
                .iter()
                .filter(|chart_object| chart_object.chart_id == binding.chart_id)
                .count();
            if primary_chart_count != 1 {
                return Err(OmError::invalid_state(format!(
                    "chart sheet {} must have one unambiguous primary chart frame",
                    host.name
                )));
            }
        }
    }

    for (sheet_id, drawings) in drawing_hosts {
        if drawings.len() > 1 && drawings.iter().any(|(_, materialized)| !materialized) {
            return Err(OmError::unsupported(format!(
                "sheet {} cannot materialize a second drawing part",
                sheet_id.0
            )));
        }
    }

    for (chart_id, chart) in &workbook.state.charts {
        if chart.id != *chart_id {
            return Err(OmError::invalid_state(format!(
                "chart map key {} does not match model id {}",
                chart_id.0, chart.id.0
            )));
        }
        if chart.workbook_id != workbook_id {
            return Err(OmError::invalid_state(format!(
                "chart {} belongs to a different workbook",
                chart_id.0
            )));
        }
        if let Some(part_uri) = chart.raw_part_uri.as_deref() {
            let chart_part = package.part(part_uri).ok_or_else(|| {
                OmError::invalid_state(format!(
                    "chart {} package part is missing: {part_uri}",
                    chart_id.0
                ))
            })?;
            let canonical_part_identity = OpcPackage::canonical_part_identity(&chart_part.name)?;
            let owner = format!("chart {}", chart_id.0);
            if let Some(existing_owner) =
                materialized_part_owners.insert(canonical_part_identity, owner.clone())
            {
                return Err(OmError::invalid_state(format!(
                    "package part {} is owned by both {existing_owner} and {owner}",
                    chart_part.name
                )));
            }
        }
        let pending_relationship_graph = workbook.pending_chart_relationship_graphs.get(chart_id);
        if let Some(pending_relationship_graph) = pending_relationship_graph {
            validate_pending_chart_relationship_graph(*chart_id, pending_relationship_graph)?;
        }
        let host_count = chart_hosts.get(chart_id).map_or(0, Vec::len);
        if host_count == 0 {
            return Err(OmError::invalid_state(format!(
                "chart {} has no drawing or chart sheet host",
                chart_id.0
            )));
        }
        if chart.raw_part_uri.is_none() && host_count > 1 {
            return Err(OmError::unsupported(format!(
                "chart {} is bound to multiple hosts",
                chart_id.0
            )));
        }
    }

    Ok(())
}

fn validate_pending_drawing_relationship_graph(
    drawing_id: office_common::DrawingId,
    graph: &PendingDrawingRelationshipGraph,
    expected_relationship_ids: &BTreeSet<String>,
) -> OmResult<()> {
    if graph.source_drawing_part_uri.trim().is_empty() {
        return Err(OmError::invalid_state(format!(
            "pending relationship graph for drawing {} has no source drawing part URI",
            drawing_id.0
        )));
    }
    validate_pending_relationship_graph(
        "drawing",
        drawing_id.0,
        &graph.root_relationships_part_source_bytes,
        &graph.root_relationships,
        &graph.parts,
        expected_relationship_ids,
    )
}

fn validate_pending_chart_relationship_graph(
    chart_id: ChartId,
    graph: &PendingChartRelationshipGraph,
) -> OmResult<()> {
    if graph.source_chart_part_uri.trim().is_empty() {
        return Err(OmError::invalid_state(format!(
            "pending relationship graph for chart {} has no source chart part URI",
            chart_id.0
        )));
    }
    if graph.source_chart_part_bytes.is_empty() {
        return Err(OmError::invalid_state(format!(
            "pending relationship graph for chart {} has no source chart XML",
            chart_id.0
        )));
    }
    validate_pending_relationship_graph(
        "chart",
        chart_id.0,
        &graph.root_relationships_part_source_bytes,
        &graph.root_relationships,
        &graph.parts,
        &BTreeSet::new(),
    )
}

fn validate_pending_relationship_graph(
    owner_kind: &str,
    owner_id: u64,
    root_relationships_part_source_bytes: &[u8],
    root_relationships: &[PendingPackageRelationship],
    parts: &BTreeMap<String, PendingPackagePart>,
    expected_relationship_ids: &BTreeSet<String>,
) -> OmResult<()> {
    if root_relationships_part_source_bytes.is_empty() {
        return Err(OmError::invalid_state(format!(
            "pending relationship graph for {owner_kind} {owner_id} has no source relationship XML"
        )));
    }
    let mut root_relationship_ids = BTreeSet::new();
    for relationship in root_relationships {
        validate_pending_package_relationship(relationship)?;
        if !root_relationship_ids.insert(relationship.relationship_id.clone()) {
            return Err(OmError::invalid_state(format!(
                "pending relationship graph for {owner_kind} {owner_id} contains duplicate root relationship {}",
                relationship.relationship_id
            )));
        }
    }
    if !expected_relationship_ids.is_subset(&root_relationship_ids) {
        return Err(OmError::invalid_state(format!(
            "pending relationship graph for {owner_kind} {owner_id} is missing required relationship ids"
        )));
    }

    for (source_part_uri, part) in parts {
        if source_part_uri != &part.source_part_uri || source_part_uri.trim().is_empty() {
            return Err(OmError::invalid_state(format!(
                "pending relationship graph for {owner_kind} {owner_id} contains a mismatched package part key"
            )));
        }
        let mut relationship_ids = BTreeSet::new();
        for relationship in &part.relationships {
            validate_pending_package_relationship(relationship)?;
            if !relationship_ids.insert(relationship.relationship_id.clone()) {
                return Err(OmError::invalid_state(format!(
                    "pending package part {} contains duplicate relationship {}",
                    source_part_uri, relationship.relationship_id
                )));
            }
        }
        if !part.relationships.is_empty() && part.relationships_part_source_bytes.is_none() {
            return Err(OmError::invalid_state(format!(
                "pending package part {} has relationships without source relationship XML",
                source_part_uri
            )));
        }
        if part.relationships_part_source_bytes.is_some()
            != part.relationships_part_compression.is_some()
        {
            return Err(OmError::invalid_state(format!(
                "pending package part {} has incomplete relationship part metadata",
                source_part_uri
            )));
        }
    }

    let mut pending_targets = root_relationships
        .iter()
        .filter(|relationship| !pending_relationship_is_external(relationship))
        .map(|relationship| relationship.target.clone())
        .collect::<Vec<_>>();
    let mut reachable_parts = BTreeSet::new();
    while let Some(source_part_uri) = pending_targets.pop() {
        if !reachable_parts.insert(source_part_uri.clone()) {
            continue;
        }
        let part = parts.get(&source_part_uri).ok_or_else(|| {
            OmError::invalid_state(format!(
                "pending relationship graph for {owner_kind} {owner_id} is missing internal target {source_part_uri}"
            ))
        })?;
        pending_targets.extend(
            part.relationships
                .iter()
                .filter(|relationship| !pending_relationship_is_external(relationship))
                .map(|relationship| relationship.target.clone()),
        );
    }
    if reachable_parts.len() != parts.len() {
        return Err(OmError::invalid_state(format!(
            "pending relationship graph for {owner_kind} {owner_id} contains unreachable package parts"
        )));
    }
    Ok(())
}

fn validate_pending_package_relationship(
    relationship: &PendingPackageRelationship,
) -> OmResult<()> {
    if relationship.relationship_id.trim().is_empty()
        || relationship.relationship_type.trim().is_empty()
        || relationship.target.trim().is_empty()
    {
        return Err(OmError::invalid_state(
            "pending package relationship requires an id, type, and target",
        ));
    }
    if relationship
        .target_mode
        .as_deref()
        .is_some_and(|target_mode| {
            !target_mode.eq_ignore_ascii_case("External")
                && !target_mode.eq_ignore_ascii_case("Internal")
        })
    {
        return Err(OmError::invalid_state(format!(
            "pending package relationship {} has unsupported TargetMode",
            relationship.relationship_id
        )));
    }
    Ok(())
}

fn pending_relationship_is_external(relationship: &PendingPackageRelationship) -> bool {
    relationship
        .target_mode
        .as_deref()
        .is_some_and(|target_mode| target_mode.eq_ignore_ascii_case("External"))
}

fn normalize_state_only_non_visual_ids(workbook: &mut LoadedXlsxWorkbook) -> OmResult<()> {
    for (drawing_id, drawing) in &mut workbook.state.drawings {
        let mut used_ids = drawing
            .objects
            .iter()
            .filter_map(|object| match object {
                DrawingObjectModel::ChartFrame(chart_object)
                    if chart_object.raw_binding.is_some() =>
                {
                    chart_object
                        .non_visual_id
                        .or_else(|| u32::try_from(chart_object.id.0).ok())
                }
                DrawingObjectModel::UnsupportedRaw { non_visual_id, .. } => *non_visual_id,
                _ => None,
            })
            .collect::<BTreeSet<_>>();
        for object in &mut drawing.objects {
            let DrawingObjectModel::ChartFrame(chart_object) = object else {
                continue;
            };
            if chart_object.raw_binding.is_some() {
                continue;
            }
            let requested_id = chart_object
                .non_visual_id
                .or_else(|| u32::try_from(chart_object.id.0).ok());
            let non_visual_id = requested_id
                .filter(|candidate| used_ids.insert(*candidate))
                .or_else(|| (1..=u32::MAX).find(|candidate| used_ids.insert(*candidate)))
                .ok_or_else(|| {
                    OmError::invalid_state(format!(
                        "drawing {} has no available DrawingML non-visual id",
                        drawing_id.0
                    ))
                })?;
            chart_object.non_visual_id = Some(non_visual_id);
        }
    }
    Ok(())
}

fn validate_materialized_chart_binding(
    package: &OpcPackage,
    chart_object_id: office_common::ChartObjectId,
    drawing_part_uri: &str,
    chart_part_uri: &str,
    raw_binding: &str,
) -> OmResult<()> {
    let (bound_drawing_part_uri, relationship_id) =
        raw_binding.rsplit_once('#').ok_or_else(|| {
            OmError::invalid_state(format!(
                "chart object {} has an invalid drawing binding: {raw_binding}",
                chart_object_id.0
            ))
        })?;
    if bound_drawing_part_uri != drawing_part_uri || relationship_id.is_empty() {
        return Err(OmError::invalid_state(format!(
            "chart object {} drawing binding does not match {drawing_part_uri}",
            chart_object_id.0
        )));
    }
    let relationships_part_uri = relationships_part_uri_for(drawing_part_uri);
    let relationships_xml = package
        .part(&relationships_part_uri)
        .ok_or_else(|| {
            OmError::invalid_state(format!(
                "chart object {} drawing relationships part is missing: {relationships_part_uri}",
                chart_object_id.0
            ))
        })?
        .bytes
        .as_slice();
    let relationship = relationship_by_id(relationships_xml, relationship_id)?.ok_or_else(|| {
        OmError::invalid_state(format!(
            "chart object {} relationship {relationship_id} is missing from {relationships_part_uri}",
            chart_object_id.0
        ))
    })?;
    let parent = drawing_part_uri
        .rsplit_once('/')
        .map_or("", |(parent, _)| parent);
    let base_segments = parent
        .split('/')
        .filter(|segment| !segment.is_empty())
        .collect::<Vec<_>>();
    let resolved_target = relationship
        .target
        .as_deref()
        .and_then(|target| normalize_relationship_target(target, &base_segments));
    if relationship.relationship_type.as_deref() != Some(CHART_RELATIONSHIP_TYPE)
        || relationship.target_mode.is_some()
        || resolved_target.as_deref() != Some(chart_part_uri)
    {
        return Err(OmError::invalid_state(format!(
            "chart object {} relationship {relationship_id} does not target {chart_part_uri}",
            chart_object_id.0
        )));
    }
    Ok(())
}

fn materialize_chart_sheet_shells(
    workbook: &mut LoadedXlsxWorkbook,
    used_part_names: &mut BTreeSet<String>,
    content_types_xml: &mut Vec<u8>,
) -> OmResult<()> {
    let sheet_ids = workbook
        .state
        .worksheets
        .iter()
        .filter(|sheet| sheet.kind == SheetKind::ChartSheet && sheet.part_uri.is_none())
        .map(|sheet| sheet.id)
        .collect::<Vec<_>>();
    if sheet_ids.is_empty() {
        return Ok(());
    }

    let mut workbook_xml = workbook
        .package
        .part(WORKBOOK_PART_NAME)
        .ok_or_else(|| OmError::new(OmErrorCode::Parse, "workbook.xml is missing"))?
        .bytes
        .clone();
    let workbook_rels_part = workbook.package.part(WORKBOOK_RELS_PART_NAME);
    let mut workbook_rels_xml = workbook_rels_part
        .map(|part| part.bytes.clone())
        .unwrap_or_else(empty_relationships_xml);
    let workbook_rels_compression = workbook_rels_part
        .map(|part| part.compression)
        .unwrap_or(CompressionMethod::Stored);
    let mut used_relationship_ids = relationship_ids(workbook_rels_xml.as_slice())?;

    for sheet_id in sheet_ids {
        let insertion_index = workbook
            .state
            .worksheets
            .iter()
            .position(|sheet| sheet.id == sheet_id)
            .expect("state-only chart sheet id collected above");
        let relationship_id = next_relationship_id(&mut used_relationship_ids);
        let part_uri =
            next_available_sequential_part_uri(used_part_names, "xl/chartsheets/sheet", ".xml");
        let relationship_target = part_uri
            .strip_prefix("xl/")
            .unwrap_or(part_uri.as_str())
            .to_string();

        let sheet = workbook
            .state
            .worksheets
            .iter_mut()
            .find(|sheet| sheet.id == sheet_id)
            .expect("state-only chart sheet id collected above");
        workbook_xml = insert_sheet_into_workbook_xml(
            workbook_xml.as_slice(),
            insertion_index,
            sheet.name.as_str(),
            sheet.id,
            relationship_id.as_str(),
            sheet.visibility,
        )?;
        workbook_rels_xml = append_relationship(
            workbook_rels_xml.as_slice(),
            relationship_id.as_str(),
            CHARTSHEET_RELATIONSHIP_TYPE,
            relationship_target.as_str(),
        )?;
        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>"#
            .to_vec();
        workbook.package.add_part(OpcPart {
            name: part_uri.clone(),
            content_type: Some(CHARTSHEET_PART_CONTENT_TYPE.to_string()),
            compression: CompressionMethod::Stored,
            bytes: sheet_xml.clone(),
        })?;
        *content_types_xml = append_content_type_override_if_missing(
            content_types_xml.as_slice(),
            part_uri.as_str(),
            CHARTSHEET_PART_CONTENT_TYPE,
        )?;
        sheet.relationship_id = Some(relationship_id);
        sheet.part_uri = Some(part_uri.clone());
        workbook
            .state
            .chart_sheets
            .get_mut(&sheet_id)
            .expect("chart sheet binding validated above")
            .raw_part_uri = Some(part_uri.clone());
        workbook
            .state
            .worksheet_data
            .entry(sheet_id)
            .or_insert_with(WorksheetData::default)
            .source_xml = sheet_xml;
    }

    workbook
        .package
        .replace_part_bytes(WORKBOOK_PART_NAME, workbook_xml)?;
    if workbook.package.contains(WORKBOOK_RELS_PART_NAME) {
        workbook
            .package
            .replace_part_bytes(WORKBOOK_RELS_PART_NAME, workbook_rels_xml)?;
    } else {
        workbook.package.add_part(OpcPart {
            name: WORKBOOK_RELS_PART_NAME.to_string(),
            content_type: Some(RELATIONSHIPS_PART_CONTENT_TYPE.to_string()),
            compression: workbook_rels_compression,
            bytes: workbook_rels_xml,
        })?;
    }
    Ok(())
}

fn materialize_pending_existing_chart_relationship_graphs(
    workbook: &mut LoadedXlsxWorkbook,
    used_part_names: &mut BTreeSet<String>,
    content_types_xml: &mut Vec<u8>,
) -> OmResult<()> {
    let plans = workbook
        .pending_chart_relationship_graphs
        .keys()
        .filter_map(|chart_id| {
            workbook
                .state
                .charts
                .get(chart_id)
                .and_then(|chart| chart.raw_part_uri.clone())
                .map(|part_uri| (*chart_id, part_uri))
        })
        .collect::<Vec<_>>();
    for (chart_id, chart_part_uri) in plans {
        materialize_chart_part(
            workbook,
            chart_id,
            &chart_part_uri,
            used_part_names,
            content_types_xml,
        )?;
    }
    Ok(())
}

fn materialize_chart_part(
    workbook: &mut LoadedXlsxWorkbook,
    chart_id: ChartId,
    chart_part_uri: &str,
    used_part_names: &mut BTreeSet<String>,
    content_types_xml: &mut Vec<u8>,
) -> OmResult<()> {
    let pending_relationship_graph = workbook
        .pending_chart_relationship_graphs
        .get(&chart_id)
        .cloned();
    let copied_part_uris = if let Some(graph) = pending_relationship_graph.as_ref() {
        materialize_pending_package_parts(
            workbook,
            &graph.parts,
            used_part_names,
            content_types_xml,
        )?
    } else {
        BTreeMap::new()
    };
    let relationship_xml = pending_relationship_graph
        .as_ref()
        .map(|graph| {
            let opaque_relationships = graph
                .root_relationships
                .iter()
                .map(|relationship| {
                    materialized_relationship_record(
                        relationship,
                        chart_part_uri,
                        &copied_part_uris,
                    )
                })
                .collect::<OmResult<Vec<_>>>()?;
            rewrite_root_relationships(
                graph.root_relationships_part_source_bytes.as_slice(),
                &opaque_relationships,
                &[],
            )
        })
        .transpose()?;
    let chart = workbook
        .state
        .charts
        .get_mut(&chart_id)
        .expect("chart graph validated above");
    chart.raw_part_uri = Some(chart_part_uri.to_string());
    let chart_xml = encode_chart_model_xml(
        pending_relationship_graph
            .as_ref()
            .map(|graph| graph.source_chart_part_bytes.as_slice()),
        chart,
    )?;
    workbook.package.remove_part(chart_part_uri);
    workbook
        .package
        .remove_part(&relationships_part_uri_for(chart_part_uri));
    workbook.package.add_part(OpcPart {
        name: chart_part_uri.to_string(),
        content_type: Some(CHART_PART_CONTENT_TYPE.to_string()),
        compression: pending_relationship_graph
            .as_ref()
            .map_or(CompressionMethod::Stored, |graph| {
                graph.source_chart_part_compression
            }),
        bytes: chart_xml,
    })?;
    if let (Some(graph), Some(relationship_xml)) =
        (pending_relationship_graph.as_ref(), relationship_xml)
    {
        workbook.package.add_part(OpcPart {
            name: relationships_part_uri_for(chart_part_uri),
            content_type: Some(RELATIONSHIPS_PART_CONTENT_TYPE.to_string()),
            compression: graph.root_relationships_part_compression,
            bytes: relationship_xml,
        })?;
    }
    *content_types_xml = append_content_type_override_if_missing(
        content_types_xml.as_slice(),
        chart_part_uri,
        CHART_PART_CONTENT_TYPE,
    )?;
    workbook.pending_chart_relationship_graphs.remove(&chart_id);
    Ok(())
}

fn materialize_new_drawings(
    workbook: &mut LoadedXlsxWorkbook,
    used_part_names: &mut BTreeSet<String>,
    content_types_xml: &mut Vec<u8>,
) -> OmResult<()> {
    let drawing_ids = workbook
        .state
        .drawings
        .iter()
        .filter_map(|(drawing_id, drawing)| drawing.raw_part_uri.is_none().then_some(*drawing_id))
        .collect::<Vec<_>>();

    for drawing_id in drawing_ids {
        let (host_sheet_id, objects, chart_objects) = {
            let drawing = workbook
                .state
                .drawings
                .get(&drawing_id)
                .expect("state-only drawing id collected above");
            (
                drawing.host_sheet_id,
                drawing.objects.clone(),
                drawing
                    .objects
                    .iter()
                    .filter_map(|object| match object {
                        DrawingObjectModel::ChartFrame(chart_object) => Some(chart_object.clone()),
                        DrawingObjectModel::UnsupportedRaw { .. } => None,
                    })
                    .collect::<Vec<_>>(),
            )
        };
        if objects.is_empty() {
            return Err(OmError::invalid_state(format!(
                "state-only drawing {} has no drawing objects",
                drawing_id.0
            )));
        }
        let host_part_uri = workbook
            .state
            .worksheets
            .iter()
            .find(|sheet| sheet.id == host_sheet_id)
            .and_then(|sheet| sheet.part_uri.clone())
            .ok_or_else(|| {
                OmError::invalid_state(format!(
                    "host sheet {} is missing a package part",
                    host_sheet_id.0
                ))
            })?;
        let drawing_part_uri =
            next_available_sequential_part_uri(used_part_names, "xl/drawings/drawing", ".xml");
        let drawing_rels_part_uri = relationships_part_uri_for(&drawing_part_uri);
        used_part_names.insert(drawing_rels_part_uri.clone());
        let pending_relationship_graph = workbook
            .pending_drawing_relationship_graphs
            .get(&drawing_id)
            .cloned();
        let copied_part_uris = if let Some(graph) = pending_relationship_graph.as_ref() {
            materialize_pending_package_parts(
                workbook,
                &graph.parts,
                used_part_names,
                content_types_xml,
            )?
        } else {
            BTreeMap::new()
        };

        let mut chart_part_uris = BTreeMap::new();
        for chart_object in &chart_objects {
            let chart_part_uri = next_available_chart_part_uri(used_part_names);
            chart_part_uris.insert(chart_object.chart_id, chart_part_uri);
        }
        let mut used_relationship_ids = pending_relationship_graph
            .as_ref()
            .into_iter()
            .flat_map(|graph| graph.root_relationships.iter())
            .map(|relationship| relationship.relationship_id.clone())
            .collect::<BTreeSet<_>>();
        let relationship_ids = chart_objects
            .iter()
            .map(|chart_object| {
                (
                    chart_object.id,
                    next_relationship_id(&mut used_relationship_ids),
                )
            })
            .collect::<BTreeMap<_, _>>();

        let opaque_relationships = pending_relationship_graph
            .as_ref()
            .into_iter()
            .flat_map(|graph| graph.root_relationships.iter())
            .map(|relationship| {
                materialized_relationship_record(relationship, &drawing_part_uri, &copied_part_uris)
            })
            .collect::<OmResult<Vec<_>>>()?;
        let chart_relationships = chart_objects
            .iter()
            .map(|chart_object| {
                let relationship_id = relationship_ids
                    .get(&chart_object.id)
                    .expect("relationship id allocated above");
                let chart_part_uri = chart_part_uris
                    .get(&chart_object.chart_id)
                    .expect("chart part uri allocated above");
                MaterializedRelationshipRecord {
                    relationship_id: relationship_id.clone(),
                    relationship_type: CHART_RELATIONSHIP_TYPE.to_string(),
                    target: relative_relationship_target(&drawing_part_uri, chart_part_uri),
                    target_mode: None,
                }
            })
            .collect::<Vec<_>>();
        let drawing_rels_xml = if let Some(graph) = pending_relationship_graph.as_ref() {
            rewrite_root_relationships(
                graph.root_relationships_part_source_bytes.as_slice(),
                &opaque_relationships,
                &chart_relationships,
            )?
        } else {
            relationships_xml(chart_relationships)
        };
        let drawing_xml = drawing_xml(&objects, &relationship_ids)?;

        for (chart_id, chart_part_uri) in &chart_part_uris {
            materialize_chart_part(
                workbook,
                *chart_id,
                chart_part_uri,
                used_part_names,
                content_types_xml,
            )?;
        }
        workbook.package.add_part(OpcPart {
            name: drawing_part_uri.clone(),
            content_type: Some(DRAWING_PART_CONTENT_TYPE.to_string()),
            compression: CompressionMethod::Stored,
            bytes: drawing_xml,
        })?;
        workbook.package.add_part(OpcPart {
            name: drawing_rels_part_uri,
            content_type: Some(RELATIONSHIPS_PART_CONTENT_TYPE.to_string()),
            compression: pending_relationship_graph
                .as_ref()
                .map_or(CompressionMethod::Stored, |graph| {
                    graph.root_relationships_part_compression
                }),
            bytes: drawing_rels_xml,
        })?;
        *content_types_xml = append_content_type_override_if_missing(
            content_types_xml.as_slice(),
            drawing_part_uri.as_str(),
            DRAWING_PART_CONTENT_TYPE,
        )?;

        attach_drawing_to_host(workbook, &host_part_uri, &drawing_part_uri)?;
        let drawing = workbook
            .state
            .drawings
            .get_mut(&drawing_id)
            .expect("state-only drawing id collected above");
        drawing.raw_part_uri = Some(drawing_part_uri.clone());
        for (object_index, object) in drawing.objects.iter_mut().enumerate() {
            let DrawingObjectModel::ChartFrame(chart_object) = object else {
                continue;
            };
            let relationship_id = relationship_ids
                .get(&chart_object.id)
                .expect("relationship id allocated above");
            chart_object.raw_binding = Some(format!("{drawing_part_uri}#{relationship_id}"));
            chart_object.z_order = u32::try_from(object_index).ok();
        }
        workbook
            .pending_drawing_relationship_graphs
            .remove(&drawing_id);
    }
    Ok(())
}

fn materialize_charts_in_existing_drawings(
    workbook: &mut LoadedXlsxWorkbook,
    used_part_names: &mut BTreeSet<String>,
    content_types_xml: &mut Vec<u8>,
) -> OmResult<()> {
    let plans = workbook
        .state
        .drawings
        .iter()
        .filter_map(|(drawing_id, drawing)| {
            let drawing_part_uri = drawing.raw_part_uri.clone()?;
            let chart_objects = drawing
                .objects
                .iter()
                .filter_map(|object| match object {
                    DrawingObjectModel::ChartFrame(chart_object)
                        if workbook
                            .state
                            .charts
                            .get(&chart_object.chart_id)
                            .is_some_and(|chart| chart.raw_part_uri.is_none()) =>
                    {
                        Some(chart_object.clone())
                    }
                    _ => None,
                })
                .collect::<Vec<_>>();
            (!chart_objects.is_empty()).then_some((*drawing_id, drawing_part_uri, chart_objects))
        })
        .collect::<Vec<_>>();

    for (drawing_id, drawing_part_uri, chart_objects) in plans {
        let drawing_rels_part_uri = relationships_part_uri_for(&drawing_part_uri);
        let existing_rels = workbook
            .package
            .part(&drawing_rels_part_uri)
            .map(|part| (part.bytes.clone(), part.compression));
        let mut used_relationship_ids = existing_rels
            .as_ref()
            .map(|(xml, _)| relationship_ids(xml.as_slice()))
            .transpose()?
            .unwrap_or_default();
        let mut updated_rels = existing_rels
            .as_ref()
            .map(|(xml, _)| xml.clone())
            .unwrap_or_else(empty_relationships_xml);
        let mut relationship_ids_by_object = BTreeMap::new();
        let mut chart_part_uris = BTreeMap::new();

        for chart_object in &chart_objects {
            let relationship_id = next_relationship_id(&mut used_relationship_ids);
            let chart_part_uri = next_available_chart_part_uri(used_part_names);
            updated_rels = append_relationship(
                updated_rels.as_slice(),
                relationship_id.as_str(),
                CHART_RELATIONSHIP_TYPE,
                relative_relationship_target(&drawing_part_uri, &chart_part_uri).as_str(),
            )?;
            relationship_ids_by_object.insert(chart_object.id, relationship_id);
            chart_part_uris.insert(chart_object.chart_id, chart_part_uri);
        }

        if let Some((_, _)) = existing_rels {
            workbook
                .package
                .replace_part_bytes(&drawing_rels_part_uri, updated_rels)?;
        } else {
            workbook.package.add_part(OpcPart {
                name: drawing_rels_part_uri,
                content_type: Some(RELATIONSHIPS_PART_CONTENT_TYPE.to_string()),
                compression: CompressionMethod::Stored,
                bytes: updated_rels,
            })?;
        }

        for (chart_id, chart_part_uri) in &chart_part_uris {
            materialize_chart_part(
                workbook,
                *chart_id,
                chart_part_uri,
                used_part_names,
                content_types_xml,
            )?;
        }

        append_chart_anchors(
            workbook,
            &drawing_part_uri,
            &chart_objects,
            &relationship_ids_by_object,
        )?;
        let drawing = workbook
            .state
            .drawings
            .get_mut(&drawing_id)
            .expect("existing drawing id came from state");
        for object in &mut drawing.objects {
            let DrawingObjectModel::ChartFrame(chart_object) = object else {
                continue;
            };
            if let Some(relationship_id) = relationship_ids_by_object.get(&chart_object.id) {
                chart_object.raw_binding = Some(format!("{drawing_part_uri}#{relationship_id}"));
            }
        }
    }
    Ok(())
}

fn attach_drawing_to_host(
    workbook: &mut LoadedXlsxWorkbook,
    host_part_uri: &str,
    drawing_part_uri: &str,
) -> OmResult<()> {
    let rels_part_uri = relationships_part_uri_for(host_part_uri);
    let existing_rels = workbook
        .package
        .part(&rels_part_uri)
        .map(|part| (part.bytes.clone(), part.compression));
    if existing_rels
        .as_ref()
        .map(|(xml, _)| relationship_type_exists(xml.as_slice(), DRAWING_RELATIONSHIP_TYPE))
        .transpose()?
        .unwrap_or(false)
    {
        return Err(OmError::invalid_state(format!(
            "drawing host already has a live drawing relationship: {host_part_uri}"
        )));
    }
    let mut used_relationship_ids = existing_rels
        .as_ref()
        .map(|(xml, _)| relationship_ids(xml.as_slice()))
        .transpose()?
        .unwrap_or_default();
    let relationship_id = next_relationship_id(&mut used_relationship_ids);
    let rels_xml = append_relationship(
        existing_rels
            .as_ref()
            .map_or_else(empty_relationships_xml, |(xml, _)| xml.clone())
            .as_slice(),
        relationship_id.as_str(),
        DRAWING_RELATIONSHIP_TYPE,
        relative_relationship_target(host_part_uri, drawing_part_uri).as_str(),
    )?;
    if let Some((_, _)) = existing_rels {
        workbook
            .package
            .replace_part_bytes(&rels_part_uri, rels_xml)?;
    } else {
        workbook.package.add_part(OpcPart {
            name: rels_part_uri,
            content_type: Some(RELATIONSHIPS_PART_CONTENT_TYPE.to_string()),
            compression: CompressionMethod::Stored,
            bytes: rels_xml,
        })?;
    }

    let host_xml = workbook
        .package
        .part(host_part_uri)
        .ok_or_else(|| {
            OmError::new(
                OmErrorCode::Parse,
                format!("drawing host part is missing: {host_part_uri}"),
            )
        })?
        .bytes
        .clone();
    let updated_host_xml = attach_drawing_element(host_xml.as_slice(), &relationship_id)?;
    workbook
        .package
        .replace_part_bytes(host_part_uri, updated_host_xml.clone())?;
    if let Some(sheet_id) = workbook
        .state
        .worksheets
        .iter()
        .find(|sheet| sheet.part_uri.as_deref() == Some(host_part_uri))
        .map(|sheet| sheet.id)
        && let Some(sheet_data) = workbook.state.worksheet_data.get_mut(&sheet_id)
    {
        sheet_data.source_xml = updated_host_xml;
    }
    Ok(())
}

fn drawing_xml(
    objects: &[DrawingObjectModel],
    relationship_ids: &BTreeMap<office_common::ChartObjectId, String>,
) -> OmResult<Vec<u8>> {
    let mut root_namespace_attrs = BTreeMap::from([
        (
            "xmlns:xdr".to_string(),
            SPREADSHEET_DRAWING_NAMESPACE.to_string(),
        ),
        ("xmlns:a".to_string(), DRAWING_MAIN_NAMESPACE.to_string()),
        ("xmlns:c".to_string(), CHART_NAMESPACE.to_string()),
        (
            "xmlns:r".to_string(),
            OFFICE_RELATIONSHIPS_NAMESPACE.to_string(),
        ),
    ]);
    for object in objects {
        let DrawingObjectModel::UnsupportedRaw {
            root_namespace_attrs: object_namespace_attrs,
            ..
        } = object
        else {
            continue;
        };
        for (name, value) in object_namespace_attrs {
            if name != "xmlns" && !name.starts_with("xmlns:") {
                return Err(OmError::invalid_state(format!(
                    "opaque drawing root attribute is not a namespace declaration: {name}"
                )));
            }
            if let Some(existing) = root_namespace_attrs.get(name)
                && existing != value
            {
                return Err(OmError::invalid_state(format!(
                    "opaque drawing namespace {name} conflicts with the materialized drawing root"
                )));
            }
            root_namespace_attrs.insert(name.clone(), value.clone());
        }
    }

    let mut xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr"#
        .to_vec();
    for (name, value) in root_namespace_attrs {
        xml.extend_from_slice(format!(r#" {name}="{}""#, escape(&value)).as_bytes());
    }
    xml.extend_from_slice(b">\n");
    for object in objects {
        match object {
            DrawingObjectModel::ChartFrame(chart_object) => {
                let relationship_id = relationship_ids
                    .get(&chart_object.id)
                    .expect("relationship id allocated above");
                xml.extend_from_slice(
                    chart_object_anchor_xml(chart_object, relationship_id)?.as_slice(),
                );
            }
            DrawingObjectModel::UnsupportedRaw { raw_anchor_xml, .. } => {
                validate_opaque_anchor_xml(raw_anchor_xml)?;
                xml.extend_from_slice(raw_anchor_xml.as_bytes());
            }
        }
        xml.push(b'\n');
    }
    xml.extend_from_slice(b"</xdr:wsDr>");
    Ok(xml)
}

fn validate_opaque_anchor_xml(raw_anchor_xml: &str) -> OmResult<()> {
    if raw_anchor_xml.trim().is_empty() {
        return Err(OmError::invalid_state(
            "opaque drawing anchor XML must not be empty",
        ));
    }
    let mut reader = Reader::from_reader(Cursor::new(raw_anchor_xml.as_bytes()));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut root_seen = false;
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) => {
                if depth == 0 {
                    if root_seen
                        || !matches!(
                            xml_local_name(element.name().as_ref()),
                            b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor"
                        )
                    {
                        return Err(OmError::invalid_state(
                            "opaque drawing object must contain exactly one DrawingML anchor",
                        ));
                    }
                    root_seen = true;
                }
                depth += 1;
            }
            Ok(Event::Empty(element)) if depth == 0 => {
                if root_seen
                    || !matches!(
                        xml_local_name(element.name().as_ref()),
                        b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor"
                    )
                {
                    return Err(OmError::invalid_state(
                        "opaque drawing object must contain exactly one DrawingML anchor",
                    ));
                }
                root_seen = true;
            }
            Ok(Event::End(_)) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OmError::invalid_state("opaque drawing anchor XML is unbalanced")
                })?;
            }
            Ok(Event::Text(text))
                if depth == 0 && !text.xml_content().map_err(xml_error)?.trim().is_empty() =>
            {
                return Err(OmError::invalid_state(
                    "opaque drawing anchor XML contains text outside the root anchor",
                ));
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    if !root_seen || depth != 0 {
        return Err(OmError::invalid_state(
            "opaque drawing object must contain one complete DrawingML anchor",
        ));
    }
    Ok(())
}

fn append_chart_anchors(
    workbook: &mut LoadedXlsxWorkbook,
    drawing_part_uri: &str,
    chart_objects: &[excel_model::ChartObjectModel],
    relationship_ids: &BTreeMap<office_common::ChartObjectId, String>,
) -> OmResult<()> {
    let drawing_xml = workbook
        .package
        .part(drawing_part_uri)
        .ok_or_else(|| {
            OmError::new(
                OmErrorCode::Parse,
                format!("drawing part is missing: {drawing_part_uri}"),
            )
        })?
        .bytes
        .clone();
    let mut reader = Reader::from_reader(Cursor::new(drawing_xml.as_slice()));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    let mut root_seen = false;
    let mut depth = 0usize;

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if depth == 0 => {
                if xml_local_name(element.name().as_ref()) != b"wsDr" {
                    return Err(OmError::new(
                        OmErrorCode::Parse,
                        format!("drawing root was not wsDr in {drawing_part_uri}"),
                    ));
                }
                root_seen = true;
                depth = 1;
                writer
                    .write_event(Event::Start(with_drawing_namespaces(
                        &element,
                        reader.decoder(),
                    )?))
                    .map_err(xml_error)?;
            }
            Ok(Event::Empty(element)) if depth == 0 => {
                if xml_local_name(element.name().as_ref()) != b"wsDr" {
                    return Err(OmError::new(
                        OmErrorCode::Parse,
                        format!("drawing root was not wsDr in {drawing_part_uri}"),
                    ));
                }
                root_seen = true;
                let qualified_name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
                writer
                    .write_event(Event::Start(with_drawing_namespaces(
                        &element,
                        reader.decoder(),
                    )?))
                    .map_err(xml_error)?;
                write_chart_anchors(&mut writer, chart_objects, relationship_ids)?;
                writer
                    .write_event(Event::End(BytesEnd::new(qualified_name)))
                    .map_err(xml_error)?;
            }
            Ok(Event::Start(element)) => {
                depth += 1;
                writer
                    .write_event(Event::Start(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::End(element)) if root_seen && depth == 1 => {
                write_chart_anchors(&mut writer, chart_objects, relationship_ids)?;
                writer
                    .write_event(Event::End(element.into_owned()))
                    .map_err(xml_error)?;
                depth = 0;
            }
            Ok(Event::End(element)) => {
                depth = depth.saturating_sub(1);
                writer
                    .write_event(Event::End(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned()).map_err(xml_error)?,
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    if !root_seen {
        return Err(OmError::new(
            OmErrorCode::Parse,
            format!("drawing root was not found in {drawing_part_uri}"),
        ));
    }
    workbook
        .package
        .replace_part_bytes(drawing_part_uri, writer.into_inner().into_inner())
}

fn write_chart_anchors(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    chart_objects: &[excel_model::ChartObjectModel],
    relationship_ids: &BTreeMap<office_common::ChartObjectId, String>,
) -> OmResult<()> {
    for chart_object in chart_objects {
        let relationship_id = relationship_ids
            .get(&chart_object.id)
            .expect("relationship id allocated above");
        writer
            .get_mut()
            .write_all(chart_object_anchor_xml(chart_object, relationship_id)?.as_slice())
            .map_err(|error| OmError::io(error.to_string()))?;
    }
    Ok(())
}

fn with_drawing_namespaces(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> OmResult<BytesStart<'static>> {
    let mut rewritten =
        BytesStart::new(String::from_utf8_lossy(element.name().as_ref()).into_owned());
    let mut namespaces = BTreeSet::new();
    for attr in element.attributes() {
        let attr = attr.map_err(xml_error)?;
        let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
        if let Some(expected_namespace) = expected_drawing_namespace(&key) {
            let actual_namespace = attr.decode_and_unescape_value(decoder).map_err(xml_error)?;
            if actual_namespace.as_ref() != expected_namespace {
                return Err(OmError::invalid_state(format!(
                    "drawing namespace {key} is bound to an unexpected URI"
                )));
            }
            namespaces.insert(key.clone());
        }
        rewritten.push_attribute(attr);
    }
    for (key, value) in [
        ("xmlns:xdr", SPREADSHEET_DRAWING_NAMESPACE),
        ("xmlns:a", DRAWING_MAIN_NAMESPACE),
        ("xmlns:c", CHART_NAMESPACE),
        ("xmlns:r", OFFICE_RELATIONSHIPS_NAMESPACE),
    ] {
        if !namespaces.contains(key) {
            rewritten.push_attribute((key, value));
        }
    }
    Ok(rewritten)
}

fn expected_drawing_namespace(attribute_name: &str) -> Option<&'static str> {
    match attribute_name {
        "xmlns:xdr" => Some(SPREADSHEET_DRAWING_NAMESPACE),
        "xmlns:a" => Some(DRAWING_MAIN_NAMESPACE),
        "xmlns:c" => Some(CHART_NAMESPACE),
        "xmlns:r" => Some(OFFICE_RELATIONSHIPS_NAMESPACE),
        _ => None,
    }
}

fn attach_drawing_element(xml: &[u8], relationship_id: &str) -> OmResult<Vec<u8>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    let mut root_seen = false;
    let mut root_depth = 0usize;
    let mut drawing_seen = false;
    let mut relationship_prefixes = BTreeSet::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if !root_seen => {
                root_seen = true;
                root_depth = 1;
                let (element, prefixes) = with_relationship_namespace(&element, reader.decoder())?;
                relationship_prefixes = prefixes;
                writer
                    .write_event(Event::Start(element))
                    .map_err(xml_error)?;
            }
            Ok(Event::Empty(element)) if !root_seen => {
                root_seen = true;
                let qualified_name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
                let (element, _) = with_relationship_namespace(&element, reader.decoder())?;
                writer
                    .write_event(Event::Start(element))
                    .map_err(xml_error)?;
                writer
                    .write_event(Event::Empty(drawing_element(
                        relationship_id,
                        qualified_name.as_bytes(),
                    )?))
                    .map_err(xml_error)?;
                writer
                    .write_event(Event::End(BytesEnd::new(qualified_name)))
                    .map_err(xml_error)?;
            }
            Ok(Event::Start(element)) => {
                if root_depth == 1 && xml_local_name(element.name().as_ref()) == b"drawing" {
                    if drawing_seen {
                        return Err(OmError::invalid_state(
                            "drawing host contains multiple drawing bindings",
                        ));
                    }
                    drawing_seen = true;
                    writer
                        .write_event(Event::Start(with_relationship_id(
                            &element,
                            relationship_id,
                            &relationship_prefixes,
                        )?))
                        .map_err(xml_error)?;
                } else {
                    writer
                        .write_event(Event::Start(element.into_owned()))
                        .map_err(xml_error)?;
                }
                root_depth += 1;
            }
            Ok(Event::Empty(element)) => {
                if root_depth == 1 && xml_local_name(element.name().as_ref()) == b"drawing" {
                    if drawing_seen {
                        return Err(OmError::invalid_state(
                            "drawing host contains multiple drawing bindings",
                        ));
                    }
                    drawing_seen = true;
                    writer
                        .write_event(Event::Empty(with_relationship_id(
                            &element,
                            relationship_id,
                            &relationship_prefixes,
                        )?))
                        .map_err(xml_error)?;
                } else {
                    writer
                        .write_event(Event::Empty(element.into_owned()))
                        .map_err(xml_error)?;
                }
            }
            Ok(Event::End(element)) if root_depth == 1 => {
                if !drawing_seen {
                    let qualified_name = element.name().as_ref().to_vec();
                    writer
                        .write_event(Event::Empty(drawing_element(
                            relationship_id,
                            qualified_name.as_slice(),
                        )?))
                        .map_err(xml_error)?;
                }
                writer
                    .write_event(Event::End(element.into_owned()))
                    .map_err(xml_error)?;
                root_depth = 0;
            }
            Ok(Event::End(element)) => {
                root_depth = root_depth.saturating_sub(1);
                writer
                    .write_event(Event::End(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned()).map_err(xml_error)?,
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    if !root_seen {
        return Err(OmError::new(
            OmErrorCode::Parse,
            "drawing host XML has no root element",
        ));
    }
    Ok(writer.into_inner().into_inner())
}

fn with_relationship_id(
    element: &BytesStart<'_>,
    relationship_id: &str,
    relationship_prefixes: &BTreeSet<String>,
) -> OmResult<BytesStart<'static>> {
    let mut rewritten =
        BytesStart::new(String::from_utf8_lossy(element.name().as_ref()).into_owned());
    let mut relationship_id_seen = false;
    for attr in element.attributes() {
        let attr = attr.map_err(xml_error)?;
        let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
        let is_relationship_id = key.rsplit_once(':').is_some_and(|(prefix, local_name)| {
            local_name == "id" && relationship_prefixes.contains(prefix)
        });
        if is_relationship_id {
            if relationship_id_seen {
                return Err(OmError::invalid_state(
                    "drawing element contains multiple relationship ids",
                ));
            }
            relationship_id_seen = true;
            push_escaped_attribute(&mut rewritten, key.as_str(), relationship_id);
        } else {
            rewritten.push_attribute(attr);
        }
    }
    if !relationship_id_seen {
        push_escaped_attribute(&mut rewritten, "r:id", relationship_id);
    }
    Ok(rewritten)
}

fn with_relationship_namespace(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> OmResult<(BytesStart<'static>, BTreeSet<String>)> {
    let mut rewritten =
        BytesStart::new(String::from_utf8_lossy(element.name().as_ref()).into_owned());
    let mut has_relationship_namespace = false;
    let mut relationship_prefixes = BTreeSet::new();
    for attr in element.attributes() {
        let attr = attr.map_err(xml_error)?;
        let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
        if let Some(prefix) = key.strip_prefix("xmlns:") {
            let value = attr.decode_and_unescape_value(decoder).map_err(xml_error)?;
            if value.as_ref() == OFFICE_RELATIONSHIPS_NAMESPACE {
                relationship_prefixes.insert(prefix.to_string());
            }
            if prefix == "r" {
                if value.as_ref() != OFFICE_RELATIONSHIPS_NAMESPACE {
                    return Err(OmError::invalid_state(
                        "worksheet relationship namespace prefix r is bound to an unexpected URI",
                    ));
                }
                has_relationship_namespace = true;
            }
        }
        rewritten.push_attribute(attr);
    }
    if !has_relationship_namespace {
        rewritten.push_attribute(("xmlns:r", OFFICE_RELATIONSHIPS_NAMESPACE));
    }
    relationship_prefixes.insert("r".to_string());
    Ok((rewritten, relationship_prefixes))
}

fn drawing_element(
    relationship_id: &str,
    host_root_qualified_name: &[u8],
) -> OmResult<BytesStart<'static>> {
    let mut element = BytesStart::new("drawing");
    push_escaped_attribute(&mut element, "r:id", relationship_id);
    qualified_generated_child(&element, host_root_qualified_name)
}

fn push_escaped_attribute(element: &mut BytesStart<'_>, key: &str, value: &str) {
    element.push_attribute((key, value));
}

fn append_relationship(
    xml: &[u8],
    relationship_id: &str,
    relationship_type: &str,
    target: &str,
) -> OmResult<Vec<u8>> {
    let mut relationship = BytesStart::new("Relationship");
    push_escaped_attribute(&mut relationship, "Id", relationship_id);
    push_escaped_attribute(&mut relationship, "Type", relationship_type);
    push_escaped_attribute(&mut relationship, "Target", target);
    append_empty_xml_child_before_container_end(xml, b"Relationships", relationship)
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct MaterializedRelationshipRecord {
    relationship_id: String,
    relationship_type: String,
    target: String,
    target_mode: Option<String>,
}

fn materialize_pending_package_parts(
    workbook: &mut LoadedXlsxWorkbook,
    parts: &BTreeMap<String, PendingPackagePart>,
    used_part_names: &mut BTreeSet<String>,
    content_types_xml: &mut Vec<u8>,
) -> OmResult<BTreeMap<String, String>> {
    let mut copied_part_uris = BTreeMap::new();
    for part in parts.values() {
        let target_part_uri = next_available_copied_part_uri(
            used_part_names,
            &part.source_part_uri,
            part.relationships_part_source_bytes.is_some(),
        )?;
        copied_part_uris.insert(part.source_part_uri.clone(), target_part_uri);
    }

    for part in parts.values() {
        let target_part_uri = copied_part_uris
            .get(&part.source_part_uri)
            .expect("pending package part URI allocated above");
        workbook.package.add_part(OpcPart {
            name: target_part_uri.clone(),
            content_type: part.content_type.clone(),
            compression: part.compression,
            bytes: part.bytes.clone(),
        })?;
        if let Some(content_type) = part.content_type.as_deref() {
            *content_types_xml = append_content_type_override_if_missing(
                content_types_xml.as_slice(),
                target_part_uri,
                content_type,
            )?;
        }
        if let Some(source_relationships_xml) = part.relationships_part_source_bytes.as_deref() {
            let target_relationships_part_uri = relationships_part_uri_for(target_part_uri);
            let target_relationships_xml = rewrite_pending_relationship_targets(
                source_relationships_xml,
                part,
                target_part_uri,
                &copied_part_uris,
            )?;
            workbook.package.add_part(OpcPart {
                name: target_relationships_part_uri,
                content_type: Some(RELATIONSHIPS_PART_CONTENT_TYPE.to_string()),
                compression: part
                    .relationships_part_compression
                    .unwrap_or(CompressionMethod::Stored),
                bytes: target_relationships_xml,
            })?;
        }
    }
    Ok(copied_part_uris)
}

fn materialized_relationship_record(
    relationship: &PendingPackageRelationship,
    source_part_uri: &str,
    copied_part_uris: &BTreeMap<String, String>,
) -> OmResult<MaterializedRelationshipRecord> {
    let target = if pending_relationship_is_external(relationship) {
        relationship.target.clone()
    } else {
        let target_part_uri = copied_part_uris.get(&relationship.target).ok_or_else(|| {
            OmError::invalid_state(format!(
                "pending package relationship {} is missing copied target {}",
                relationship.relationship_id, relationship.target
            ))
        })?;
        relative_relationship_target(source_part_uri, target_part_uri)
    };
    Ok(MaterializedRelationshipRecord {
        relationship_id: relationship.relationship_id.clone(),
        relationship_type: relationship.relationship_type.clone(),
        target,
        target_mode: relationship.target_mode.clone(),
    })
}

fn rewrite_pending_relationship_targets(
    source_xml: &[u8],
    part: &PendingPackagePart,
    target_part_uri: &str,
    copied_part_uris: &BTreeMap<String, String>,
) -> OmResult<Vec<u8>> {
    let relationships = part
        .relationships
        .iter()
        .map(|relationship| (relationship.relationship_id.as_str(), relationship))
        .collect::<BTreeMap<_, _>>();
    let mut seen_relationship_ids = BTreeSet::new();
    let mut reader = Reader::from_reader(Cursor::new(source_xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element))
                if xml_local_name(element.name().as_ref()) == b"Relationship" =>
            {
                writer
                    .write_event(Event::Start(rewrite_pending_relationship_element(
                        &element,
                        reader.decoder(),
                        &relationships,
                        &mut seen_relationship_ids,
                        target_part_uri,
                        copied_part_uris,
                    )?))
                    .map_err(xml_error)?;
            }
            Ok(Event::Empty(element))
                if xml_local_name(element.name().as_ref()) == b"Relationship" =>
            {
                writer
                    .write_event(Event::Empty(rewrite_pending_relationship_element(
                        &element,
                        reader.decoder(),
                        &relationships,
                        &mut seen_relationship_ids,
                        target_part_uri,
                        copied_part_uris,
                    )?))
                    .map_err(xml_error)?;
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned()).map_err(xml_error)?,
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    if seen_relationship_ids.len() != relationships.len() {
        return Err(OmError::invalid_state(format!(
            "pending package part {} relationship summary does not match source XML",
            part.source_part_uri
        )));
    }
    Ok(writer.into_inner().into_inner())
}

fn rewrite_pending_relationship_element(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    relationships: &BTreeMap<&str, &PendingPackageRelationship>,
    seen_relationship_ids: &mut BTreeSet<String>,
    target_part_uri: &str,
    copied_part_uris: &BTreeMap<String, String>,
) -> OmResult<BytesStart<'static>> {
    let mut relationship_id = None;
    for attr in element.attributes() {
        let attr = attr.map_err(xml_error)?;
        if attr.key.as_ref() == b"Id" {
            relationship_id = Some(
                attr.decode_and_unescape_value(decoder)
                    .map_err(xml_error)?
                    .into_owned(),
            );
        }
    }
    let relationship_id = relationship_id
        .ok_or_else(|| OmError::invalid_state("pending package Relationship element has no Id"))?;
    let relationship = relationships.get(relationship_id.as_str()).ok_or_else(|| {
        OmError::invalid_state(format!(
            "pending package relationship {} is missing from its typed graph",
            relationship_id
        ))
    })?;
    if !seen_relationship_ids.insert(relationship_id.clone()) {
        return Err(OmError::invalid_state(format!(
            "pending package relationship {} occurs more than once in source XML",
            relationship_id
        )));
    }
    let rewritten_target = if pending_relationship_is_external(relationship) {
        None
    } else {
        let copied_target = copied_part_uris.get(&relationship.target).ok_or_else(|| {
            OmError::invalid_state(format!(
                "pending package relationship {} is missing copied target {}",
                relationship_id, relationship.target
            ))
        })?;
        Some(relative_relationship_target(target_part_uri, copied_target))
    };

    let mut rewritten =
        BytesStart::new(String::from_utf8_lossy(element.name().as_ref()).into_owned());
    let mut target_seen = false;
    for attr in element.attributes() {
        let attr = attr.map_err(xml_error)?;
        let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
        let value = attr
            .decode_and_unescape_value(decoder)
            .map_err(xml_error)?
            .into_owned();
        let value = if attr.key.as_ref() == b"Target" {
            target_seen = true;
            rewritten_target.as_deref().unwrap_or(value.as_str())
        } else {
            value.as_str()
        };
        rewritten.push_attribute((key.as_str(), value));
    }
    if !target_seen {
        let target = rewritten_target.as_deref().unwrap_or(&relationship.target);
        rewritten.push_attribute(("Target", target));
    }
    Ok(rewritten)
}

fn rewrite_root_relationships(
    source_xml: &[u8],
    opaque_relationships: &[MaterializedRelationshipRecord],
    chart_relationships: &[MaterializedRelationshipRecord],
) -> OmResult<Vec<u8>> {
    let opaque_relationships = opaque_relationships
        .iter()
        .map(|relationship| (relationship.relationship_id.as_str(), relationship))
        .collect::<BTreeMap<_, _>>();
    let mut seen_relationship_ids = BTreeSet::new();
    let mut reader = Reader::from_reader(Cursor::new(source_xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    let mut root_name = None::<Vec<u8>>;
    let mut depth = 0usize;
    let mut skipped_depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(_)) if skipped_depth > 0 => skipped_depth += 1,
            Ok(Event::End(_)) if skipped_depth > 0 => skipped_depth -= 1,
            Ok(Event::Empty(_)) if skipped_depth > 0 => {}
            Ok(Event::Eof) if skipped_depth > 0 => {
                return Err(OmError::invalid_state(
                    "relationship XML ended inside a skipped relationship",
                ));
            }
            Ok(Event::Start(element)) if depth == 0 => {
                root_name = Some(element.name().as_ref().to_vec());
                depth = 1;
                writer
                    .write_event(Event::Start(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::Empty(element)) if depth == 0 => {
                let qualified_root_name = element.name().as_ref().to_vec();
                let end_name = String::from_utf8_lossy(&qualified_root_name).into_owned();
                writer
                    .write_event(Event::Start(element.into_owned()))
                    .map_err(xml_error)?;
                write_materialized_relationships(
                    &mut writer,
                    &qualified_root_name,
                    opaque_relationships.values().copied(),
                )?;
                write_materialized_relationships(
                    &mut writer,
                    &qualified_root_name,
                    chart_relationships.iter(),
                )?;
                writer
                    .write_event(Event::End(BytesEnd::new(end_name)))
                    .map_err(xml_error)?;
                root_name = Some(qualified_root_name);
            }
            Ok(Event::Start(element))
                if depth == 1 && xml_local_name(element.name().as_ref()) == b"Relationship" =>
            {
                let relationship_id = relationship_element_id(&element, reader.decoder())?;
                if let Some(relationship) = opaque_relationships.get(relationship_id.as_str()) {
                    if !seen_relationship_ids.insert(relationship_id.clone()) {
                        return Err(OmError::invalid_state(format!(
                            "relationship {} occurs more than once",
                            relationship_id
                        )));
                    }
                    writer
                        .write_event(Event::Start(rewrite_materialized_relationship_element(
                            &element,
                            reader.decoder(),
                            relationship,
                        )?))
                        .map_err(xml_error)?;
                    depth += 1;
                } else {
                    skipped_depth = 1;
                }
            }
            Ok(Event::Empty(element))
                if depth == 1 && xml_local_name(element.name().as_ref()) == b"Relationship" =>
            {
                let relationship_id = relationship_element_id(&element, reader.decoder())?;
                if let Some(relationship) = opaque_relationships.get(relationship_id.as_str()) {
                    if !seen_relationship_ids.insert(relationship_id.clone()) {
                        return Err(OmError::invalid_state(format!(
                            "relationship {} occurs more than once",
                            relationship_id
                        )));
                    }
                    writer
                        .write_event(Event::Empty(rewrite_materialized_relationship_element(
                            &element,
                            reader.decoder(),
                            relationship,
                        )?))
                        .map_err(xml_error)?;
                }
            }
            Ok(Event::Start(element)) => {
                depth += 1;
                writer
                    .write_event(Event::Start(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::End(element)) if depth == 1 => {
                let qualified_root_name = root_name.as_deref().ok_or_else(|| {
                    OmError::invalid_state("relationship XML has no root element")
                })?;
                write_materialized_relationships(
                    &mut writer,
                    qualified_root_name,
                    chart_relationships.iter(),
                )?;
                writer
                    .write_event(Event::End(element.into_owned()))
                    .map_err(xml_error)?;
                depth = 0;
            }
            Ok(Event::End(element)) => {
                depth = depth.saturating_sub(1);
                writer
                    .write_event(Event::End(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned()).map_err(xml_error)?,
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    if root_name.is_none() || seen_relationship_ids.len() != opaque_relationships.len() {
        return Err(OmError::invalid_state(
            "relationship summary does not match source XML",
        ));
    }
    Ok(writer.into_inner().into_inner())
}

fn relationship_element_id(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> OmResult<String> {
    for attr in element.attributes() {
        let attr = attr.map_err(xml_error)?;
        if attr.key.as_ref() == b"Id" {
            return attr
                .decode_and_unescape_value(decoder)
                .map(|value| value.into_owned())
                .map_err(xml_error);
        }
    }
    Err(OmError::invalid_state("Relationship element has no Id"))
}

fn rewrite_materialized_relationship_element(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    relationship: &MaterializedRelationshipRecord,
) -> OmResult<BytesStart<'static>> {
    let mut rewritten =
        BytesStart::new(String::from_utf8_lossy(element.name().as_ref()).into_owned());
    let mut target_seen = false;
    for attr in element.attributes() {
        let attr = attr.map_err(xml_error)?;
        let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
        let value = attr
            .decode_and_unescape_value(decoder)
            .map_err(xml_error)?
            .into_owned();
        let value = if attr.key.as_ref() == b"Target" {
            target_seen = true;
            relationship.target.as_str()
        } else {
            value.as_str()
        };
        rewritten.push_attribute((key.as_str(), value));
    }
    if !target_seen {
        rewritten.push_attribute(("Target", relationship.target.as_str()));
    }
    Ok(rewritten)
}

fn write_materialized_relationships<'a>(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    root_name: &[u8],
    relationships: impl IntoIterator<Item = &'a MaterializedRelationshipRecord>,
) -> OmResult<()> {
    let prefix = xml_prefix(root_name)
        .map(|prefix| format!("{}:", String::from_utf8_lossy(prefix)))
        .unwrap_or_default();
    for relationship in relationships {
        let mut element = BytesStart::new(format!("{prefix}Relationship"));
        push_escaped_attribute(&mut element, "Id", &relationship.relationship_id);
        push_escaped_attribute(&mut element, "Type", &relationship.relationship_type);
        push_escaped_attribute(&mut element, "Target", &relationship.target);
        if let Some(target_mode) = relationship.target_mode.as_deref() {
            push_escaped_attribute(&mut element, "TargetMode", target_mode);
        }
        writer
            .write_event(Event::Empty(element))
            .map_err(xml_error)?;
    }
    Ok(())
}

fn relationships_xml(
    relationships: impl IntoIterator<Item = MaterializedRelationshipRecord>,
) -> Vec<u8> {
    let mut xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
"#
    .to_vec();
    for relationship in relationships {
        let target_mode = relationship
            .target_mode
            .as_deref()
            .map_or_else(String::new, |mode| {
                format!(" TargetMode=\"{}\"", escape(mode))
            });
        xml.extend_from_slice(
            format!(
                "  <Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"{target_mode}/>\n",
                escape(&relationship.relationship_id),
                escape(&relationship.relationship_type),
                escape(&relationship.target)
            )
            .as_bytes(),
        );
    }
    xml.extend_from_slice(b"</Relationships>");
    xml
}

fn empty_relationships_xml() -> Vec<u8> {
    br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#
        .to_vec()
}

fn relationship_ids(xml: &[u8]) -> OmResult<BTreeSet<String>> {
    let mut ids = BTreeSet::new();
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if xml_local_name(element.name().as_ref()) == b"Relationship" =>
            {
                for attr in element.attributes() {
                    let attr = attr.map_err(xml_error)?;
                    if attr.key.as_ref() == b"Id" {
                        ids.insert(
                            attr.decode_and_unescape_value(reader.decoder())
                                .map_err(xml_error)?
                                .into_owned(),
                        );
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    Ok(ids)
}

struct RelationshipRecord {
    relationship_type: Option<String>,
    target: Option<String>,
    target_mode: Option<String>,
}

fn relationship_by_id(xml: &[u8], relationship_id: &str) -> OmResult<Option<RelationshipRecord>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if xml_local_name(element.name().as_ref()) == b"Relationship" =>
            {
                let mut id = None;
                let mut relationship_type = None;
                let mut target = None;
                let mut target_mode = None;
                for attr in element.attributes() {
                    let attr = attr.map_err(xml_error)?;
                    let value = attr
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(xml_error)?
                        .into_owned();
                    match attr.key.as_ref() {
                        b"Id" => id = Some(value),
                        b"Type" => relationship_type = Some(value),
                        b"Target" => target = Some(value),
                        b"TargetMode" => target_mode = Some(value),
                        _ => {}
                    }
                }
                if id.as_deref() == Some(relationship_id) {
                    return Ok(Some(RelationshipRecord {
                        relationship_type,
                        target,
                        target_mode,
                    }));
                }
            }
            Ok(Event::Eof) => return Ok(None),
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
}

fn relationship_type_exists(xml: &[u8], relationship_type: &str) -> OmResult<bool> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if xml_local_name(element.name().as_ref()) == b"Relationship" =>
            {
                for attr in element.attributes() {
                    let attr = attr.map_err(xml_error)?;
                    if attr.key.as_ref() == b"Type"
                        && attr
                            .decode_and_unescape_value(reader.decoder())
                            .map_err(xml_error)?
                            .as_ref()
                            == relationship_type
                    {
                        return Ok(true);
                    }
                }
            }
            Ok(Event::Eof) => return Ok(false),
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
}

fn next_relationship_id(used_ids: &mut BTreeSet<String>) -> String {
    let id = (1..)
        .map(|index| format!("rId{index}"))
        .find(|candidate| !used_ids.contains(candidate))
        .expect("relationship id search is unbounded");
    used_ids.insert(id.clone());
    id
}

fn next_available_sequential_part_uri(
    used_part_names: &mut BTreeSet<String>,
    prefix: &str,
    suffix: &str,
) -> String {
    (1..)
        .map(|index| format!("{prefix}{index}{suffix}"))
        .find(|candidate| used_part_names.insert(candidate.clone()))
        .expect("part uri search is unbounded")
}

fn next_available_chart_part_uri(used_part_names: &mut BTreeSet<String>) -> String {
    loop {
        let candidate =
            next_available_sequential_part_uri(used_part_names, "xl/charts/chart", ".xml");
        if used_part_names.insert(relationships_part_uri_for(&candidate)) {
            return candidate;
        }
    }
}

fn next_available_copied_part_uri(
    used_part_names: &mut BTreeSet<String>,
    source_part_uri: &str,
    has_relationships_part: bool,
) -> OmResult<String> {
    let source_part_uri = source_part_uri.trim_start_matches('/');
    if source_part_uri.is_empty() || source_part_uri.ends_with('/') {
        return Err(OmError::invalid_state(
            "pending package part has an invalid source URI",
        ));
    }
    let candidate_is_available = |candidate: &str, used_part_names: &BTreeSet<String>| {
        !used_part_names.contains(candidate)
            && (!has_relationships_part
                || !used_part_names.contains(&relationships_part_uri_for(candidate)))
    };
    if candidate_is_available(source_part_uri, used_part_names) {
        used_part_names.insert(source_part_uri.to_string());
        if has_relationships_part {
            used_part_names.insert(relationships_part_uri_for(source_part_uri));
        }
        return Ok(source_part_uri.to_string());
    }

    let (parent, file_name) = source_part_uri
        .rsplit_once('/')
        .map_or(("", source_part_uri), |(parent, file_name)| {
            (parent, file_name)
        });
    let (stem, suffix) = file_name
        .rfind('.')
        .map_or((file_name, ""), |extension_start| {
            (&file_name[..extension_start], &file_name[extension_start..])
        });
    let digit_start = stem
        .char_indices()
        .rev()
        .find_map(|(index, character)| {
            (!character.is_ascii_digit()).then_some(index + character.len_utf8())
        })
        .unwrap_or(0);
    let base = if digit_start == 0 && stem.chars().all(|character| character.is_ascii_digit()) {
        stem
    } else {
        &stem[..digit_start]
    };
    let first_index = stem[digit_start..]
        .parse::<usize>()
        .ok()
        .and_then(|index| index.checked_add(1))
        .unwrap_or(1);
    for index in first_index.. {
        let file_name = format!("{base}{index}{suffix}");
        let candidate = if parent.is_empty() {
            file_name
        } else {
            format!("{parent}/{file_name}")
        };
        if candidate_is_available(&candidate, used_part_names) {
            used_part_names.insert(candidate.clone());
            if has_relationships_part {
                used_part_names.insert(relationships_part_uri_for(&candidate));
            }
            return Ok(candidate);
        }
    }
    unreachable!("part uri search is unbounded")
}

fn relationships_part_uri_for(part_uri: &str) -> String {
    if let Some((prefix, file_name)) = part_uri.rsplit_once('/') {
        format!("{prefix}/_rels/{file_name}.rels")
    } else {
        format!("_rels/{part_uri}.rels")
    }
}

fn relative_relationship_target(source_part_uri: &str, target_part_uri: &str) -> String {
    let Some((parent, _)) = source_part_uri.rsplit_once('/') else {
        return target_part_uri.to_string();
    };
    let base_segments = parent
        .split('/')
        .filter(|segment| !segment.is_empty())
        .collect::<Vec<_>>();
    let target_segments = target_part_uri
        .split('/')
        .filter(|segment| !segment.is_empty())
        .collect::<Vec<_>>();
    let shared_prefix_len = base_segments
        .iter()
        .zip(target_segments.iter())
        .take_while(|(left, right)| left == right)
        .count();
    let mut relative = vec![".."; base_segments.len().saturating_sub(shared_prefix_len)];
    relative.extend_from_slice(&target_segments[shared_prefix_len..]);
    relative.join("/")
}

fn append_content_type_override_if_missing(
    xml: &[u8],
    part_uri: &str,
    content_type: &str,
) -> OmResult<Vec<u8>> {
    let expected_part_name = format!("/{}", part_uri.trim_start_matches('/'));
    if let Some(existing_content_type) = content_type_overrides(xml)?.get(&expected_part_name) {
        if existing_content_type != content_type {
            return Err(OmError::invalid_state(format!(
                "content type override for {expected_part_name} is {existing_content_type}, expected {content_type}"
            )));
        }
        return Ok(xml.to_vec());
    }
    let mut element = BytesStart::new("Override");
    push_escaped_attribute(&mut element, "PartName", &expected_part_name);
    push_escaped_attribute(&mut element, "ContentType", content_type);
    append_empty_xml_child_before_container_end(xml, b"Types", element)
}

fn content_type_override_part_names(xml: &[u8]) -> OmResult<BTreeSet<String>> {
    Ok(content_type_overrides(xml)?
        .into_keys()
        .map(|part_name| part_name.trim_start_matches('/').to_string())
        .collect())
}

fn content_type_overrides(xml: &[u8]) -> OmResult<BTreeMap<String, String>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut overrides = BTreeMap::new();
    let mut root_seen = false;
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if depth == 0 => {
                validate_root_container_namespace(&element, reader.decoder(), b"Types")?;
                root_seen = true;
                depth = 1;
            }
            Ok(Event::Empty(element)) if depth == 0 => {
                validate_root_container_namespace(&element, reader.decoder(), b"Types")?;
                root_seen = true;
            }
            Ok(Event::Start(element)) if depth == 1 => {
                if xml_local_name(element.name().as_ref()) == b"Override" {
                    collect_content_type_override(&element, reader.decoder(), &mut overrides)?;
                }
                depth += 1;
            }
            Ok(Event::Empty(element)) if depth == 1 => {
                if xml_local_name(element.name().as_ref()) == b"Override" {
                    collect_content_type_override(&element, reader.decoder(), &mut overrides)?;
                }
            }
            Ok(Event::Start(_)) => depth += 1,
            Ok(Event::End(_)) => depth = depth.saturating_sub(1),
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    if !root_seen {
        return Err(OmError::new(
            OmErrorCode::Parse,
            "content types XML has no root element",
        ));
    }
    Ok(overrides)
}

fn collect_content_type_override(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    overrides: &mut BTreeMap<String, String>,
) -> OmResult<()> {
    let mut part_name = None;
    let mut content_type = None;
    for attr in element.attributes() {
        let attr = attr.map_err(xml_error)?;
        let value = attr
            .decode_and_unescape_value(decoder)
            .map_err(xml_error)?
            .into_owned();
        match attr.key.as_ref() {
            b"PartName" => part_name = Some(value),
            b"ContentType" => content_type = Some(value),
            _ => {}
        }
    }
    if let (Some(part_name), Some(content_type)) = (part_name, content_type)
        && let Some(previous) = overrides.insert(part_name.clone(), content_type.clone())
        && previous != content_type
    {
        return Err(OmError::invalid_state(format!(
            "content type override for {part_name} is duplicated with conflicting MIME types"
        )));
    }
    Ok(())
}

fn validate_root_container_namespace(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    expected_local_name: &[u8],
) -> OmResult<()> {
    if xml_local_name(element.name().as_ref()) != expected_local_name {
        return Err(OmError::new(
            OmErrorCode::Parse,
            format!(
                "XML root was not {}",
                String::from_utf8_lossy(expected_local_name)
            ),
        ));
    }
    let expected_namespace = match expected_local_name {
        b"Types" => CONTENT_TYPES_NAMESPACE,
        b"Relationships" => PACKAGE_RELATIONSHIPS_NAMESPACE,
        b"workbook" => SPREADSHEET_NAMESPACE,
        _ => return Ok(()),
    };
    let qualified_name = element.name();
    let qualified_name = qualified_name.as_ref();
    let namespace_attribute = qualified_name
        .iter()
        .position(|byte| *byte == b':')
        .map(|separator| {
            format!(
                "xmlns:{}",
                String::from_utf8_lossy(&qualified_name[..separator])
            )
        })
        .unwrap_or_else(|| "xmlns".to_string());
    for attr in element.attributes() {
        let attr = attr.map_err(xml_error)?;
        if attr.key.as_ref() == namespace_attribute.as_bytes() {
            let actual_namespace = attr.decode_and_unescape_value(decoder).map_err(xml_error)?;
            if actual_namespace.as_ref() == expected_namespace {
                return Ok(());
            }
            break;
        }
    }
    Err(OmError::new(
        OmErrorCode::Parse,
        format!(
            "XML root {} is not in the expected namespace",
            String::from_utf8_lossy(expected_local_name)
        ),
    ))
}

fn append_empty_xml_child_before_container_end(
    xml: &[u8],
    container_name: &[u8],
    child: BytesStart<'static>,
) -> OmResult<Vec<u8>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    let mut container_seen = false;
    let mut inserted = false;
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if depth == 0 => {
                validate_root_container_namespace(&element, reader.decoder(), container_name)?;
                container_seen = true;
                depth = 1;
                writer
                    .write_event(Event::Start(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::Empty(element)) if depth == 0 => {
                validate_root_container_namespace(&element, reader.decoder(), container_name)?;
                container_seen = true;
                inserted = true;
                let qualified_name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
                writer
                    .write_event(Event::Start(element.into_owned()))
                    .map_err(xml_error)?;
                writer
                    .write_event(Event::Empty(qualified_generated_child(
                        &child,
                        qualified_name.as_bytes(),
                    )?))
                    .map_err(xml_error)?;
                writer
                    .write_event(Event::End(BytesEnd::new(qualified_name)))
                    .map_err(xml_error)?;
            }
            Ok(Event::Start(element)) => {
                depth += 1;
                writer
                    .write_event(Event::Start(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::End(element)) if depth == 1 => {
                let qualified_name = element.name().as_ref().to_vec();
                writer
                    .write_event(Event::Empty(qualified_generated_child(
                        &child,
                        qualified_name.as_slice(),
                    )?))
                    .map_err(xml_error)?;
                writer
                    .write_event(Event::End(element.into_owned()))
                    .map_err(xml_error)?;
                inserted = true;
                depth = 0;
            }
            Ok(Event::End(element)) => {
                depth = depth.saturating_sub(1);
                writer
                    .write_event(Event::End(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned()).map_err(xml_error)?,
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    if !container_seen || !inserted {
        return Err(OmError::new(
            OmErrorCode::Parse,
            format!(
                "XML container {} was not found",
                String::from_utf8_lossy(container_name)
            ),
        ));
    }
    Ok(writer.into_inner().into_inner())
}

fn qualified_generated_child(
    child: &BytesStart<'static>,
    container_qualified_name: &[u8],
) -> OmResult<BytesStart<'static>> {
    let child_name = child.name();
    let local_name = String::from_utf8_lossy(xml_local_name(child_name.as_ref())).into_owned();
    let qualified_name = if let Some(separator) = container_qualified_name
        .iter()
        .position(|byte| *byte == b':')
    {
        format!(
            "{}:{local_name}",
            String::from_utf8_lossy(&container_qualified_name[..separator])
        )
    } else {
        local_name
    };
    let mut qualified = BytesStart::new(qualified_name);
    for attr in child.attributes().with_checks(false) {
        let attr = attr.map_err(xml_error)?;
        let key = attr.key.as_ref().to_vec();
        let value = attr.value.into_owned();
        qualified.push_attribute((key.as_slice(), value.as_slice()));
    }
    Ok(qualified)
}

fn insert_sheet_into_workbook_xml(
    xml: &[u8],
    insert_at: usize,
    sheet_name: &str,
    sheet_id: SheetId,
    relationship_id: &str,
    visibility: SheetVisibility,
) -> OmResult<Vec<u8>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    let mut inside_sheets = false;
    let mut sheets_seen = false;
    let mut document_depth = 0usize;
    let mut workbook_prefix = None::<Vec<u8>>;
    let mut sheet_index = 0usize;
    let mut inserted = false;
    let mut sheet = BytesStart::new("sheet");
    push_escaped_attribute(&mut sheet, "name", sheet_name);
    push_escaped_attribute(&mut sheet, "sheetId", &sheet_id.0.to_string());
    push_escaped_attribute(&mut sheet, "r:id", relationship_id);
    if let Some(state) = match visibility {
        SheetVisibility::Visible => None,
        SheetVisibility::Hidden => Some("hidden"),
        SheetVisibility::VeryHidden => Some("veryHidden"),
    } {
        push_escaped_attribute(&mut sheet, "state", state);
    }

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if document_depth == 0 => {
                validate_root_container_namespace(&element, reader.decoder(), b"workbook")?;
                workbook_prefix = xml_prefix(element.name().as_ref()).map(|prefix| prefix.to_vec());
                document_depth = 1;
                writer
                    .write_event(Event::Start(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::Empty(element)) if document_depth == 0 => {
                validate_root_container_namespace(&element, reader.decoder(), b"workbook")?;
                writer
                    .write_event(Event::Empty(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::Start(element))
                if document_depth == 1
                    && xml_local_name(element.name().as_ref()) == b"sheets"
                    && xml_prefix(element.name().as_ref()) == workbook_prefix.as_deref() =>
            {
                if sheets_seen {
                    return Err(OmError::invalid_state(
                        "workbook.xml contains multiple sheets containers",
                    ));
                }
                sheets_seen = true;
                inside_sheets = true;
                document_depth += 1;
                writer
                    .write_event(Event::Start(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::Empty(element))
                if document_depth == 1
                    && xml_local_name(element.name().as_ref()) == b"sheets"
                    && xml_prefix(element.name().as_ref()) == workbook_prefix.as_deref() =>
            {
                if sheets_seen {
                    return Err(OmError::invalid_state(
                        "workbook.xml contains multiple sheets containers",
                    ));
                }
                sheets_seen = true;
                inserted = true;
                let qualified_name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
                writer
                    .write_event(Event::Start(element.into_owned()))
                    .map_err(xml_error)?;
                writer
                    .write_event(Event::Empty(qualified_generated_child(
                        &sheet,
                        qualified_name.as_bytes(),
                    )?))
                    .map_err(xml_error)?;
                writer
                    .write_event(Event::End(BytesEnd::new(qualified_name)))
                    .map_err(xml_error)?;
            }
            Ok(Event::Start(element))
                if inside_sheets
                    && document_depth == 2
                    && xml_local_name(element.name().as_ref()) == b"sheet" =>
            {
                if !inserted && sheet_index == insert_at {
                    let qualified_name = element.name().as_ref().to_vec();
                    writer
                        .write_event(Event::Empty(qualified_generated_child(
                            &sheet,
                            qualified_name.as_slice(),
                        )?))
                        .map_err(xml_error)?;
                    inserted = true;
                }
                sheet_index += 1;
                document_depth += 1;
                writer
                    .write_event(Event::Start(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::Empty(element))
                if inside_sheets
                    && document_depth == 2
                    && xml_local_name(element.name().as_ref()) == b"sheet" =>
            {
                if !inserted && sheet_index == insert_at {
                    let qualified_name = element.name().as_ref().to_vec();
                    writer
                        .write_event(Event::Empty(qualified_generated_child(
                            &sheet,
                            qualified_name.as_slice(),
                        )?))
                        .map_err(xml_error)?;
                    inserted = true;
                }
                sheet_index += 1;
                writer
                    .write_event(Event::Empty(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::Start(element)) => {
                document_depth += 1;
                writer
                    .write_event(Event::Start(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::End(element)) if inside_sheets && document_depth == 2 => {
                if !inserted {
                    let qualified_name = element.name().as_ref().to_vec();
                    writer
                        .write_event(Event::Empty(qualified_generated_child(
                            &sheet,
                            qualified_name.as_slice(),
                        )?))
                        .map_err(xml_error)?;
                    inserted = true;
                }
                inside_sheets = false;
                document_depth -= 1;
                writer
                    .write_event(Event::End(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::End(element)) => {
                document_depth = document_depth.saturating_sub(1);
                writer
                    .write_event(Event::End(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned()).map_err(xml_error)?,
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    if !sheets_seen || !inserted {
        return Err(OmError::new(
            OmErrorCode::Parse,
            "workbook.xml does not contain a sheets container",
        ));
    }
    Ok(writer.into_inner().into_inner())
}

fn xml_prefix(name: &[u8]) -> Option<&[u8]> {
    name.iter()
        .position(|byte| *byte == b':')
        .map(|separator| &name[..separator])
}

fn invalidate_changed_preservation_snapshots(workbook: &mut LoadedXlsxWorkbook) {
    let package = &workbook.package;
    let part_changed = |part_uri: &str, source_bytes: &[u8]| {
        package
            .part(part_uri)
            .is_none_or(|part| part.bytes != source_bytes)
    };
    if workbook
        .support_parts
        .content_types_source_bytes
        .as_deref()
        .is_some_and(|source| part_changed(CONTENT_TYPES_PART_NAME, source))
    {
        workbook.support_parts.content_types_source_bytes = None;
        workbook.support_parts.content_types_summary = None;
    }
    if workbook
        .support_parts
        .workbook_relationships_part_source_bytes
        .as_deref()
        .is_some_and(|source| part_changed(WORKBOOK_RELS_PART_NAME, source))
    {
        workbook
            .support_parts
            .workbook_relationships_part_source_bytes = None;
        workbook.support_parts.workbook_relationships_summary = None;
    }
    for support in workbook.worksheet_support_parts.values_mut() {
        if support
            .relationships_part_uri
            .as_deref()
            .zip(support.relationships_part_source_bytes.as_deref())
            .is_some_and(|(part_uri, source)| part_changed(part_uri, source))
        {
            support.relationships_part_source_bytes = None;
            support.relationships_summary = None;
        }
    }
    for support in workbook.sheet_drawing_support_parts.values_mut() {
        if support
            .sheet_part_uri
            .as_deref()
            .zip(support.sheet_part_source_bytes.as_deref())
            .is_some_and(|(part_uri, source)| part_changed(part_uri, source))
        {
            support.sheet_part_source_bytes = None;
        }
        if support
            .relationships_part_uri
            .as_deref()
            .zip(support.relationships_part_source_bytes.as_deref())
            .is_some_and(|(part_uri, source)| part_changed(part_uri, source))
        {
            support.relationships_part_source_bytes = None;
        }
        let changed_drawings = support
            .drawing_part_source_bytes
            .iter()
            .filter_map(|(part_uri, source)| {
                part_changed(part_uri, source).then(|| part_uri.clone())
            })
            .collect::<Vec<_>>();
        for part_uri in changed_drawings {
            support.drawing_part_source_bytes.remove(&part_uri);
            support.drawing_summaries.remove(&part_uri);
        }
        let changed_rels = support
            .drawing_relationships_part_source_bytes
            .iter()
            .filter_map(|(part_uri, source)| {
                part_changed(part_uri, source).then(|| part_uri.clone())
            })
            .collect::<Vec<_>>();
        for part_uri in changed_rels {
            support
                .drawing_relationships_part_source_bytes
                .remove(&part_uri);
        }
    }
}

fn ensure_all_state_only_charts_materialized(workbook: &LoadedXlsxWorkbook) -> OmResult<()> {
    if let Some(drawing_id) = workbook.pending_drawing_relationship_graphs.keys().next() {
        return Err(OmError::invalid_state(format!(
            "pending relationship graph for drawing {} was not materialized",
            drawing_id.0
        )));
    }
    if let Some(chart_id) = workbook.pending_chart_relationship_graphs.keys().next() {
        return Err(OmError::invalid_state(format!(
            "pending relationship graph for chart {} was not materialized",
            chart_id.0
        )));
    }
    if let Some(chart) = workbook
        .state
        .charts
        .values()
        .find(|chart| chart.raw_part_uri.is_none())
    {
        return Err(OmError::invalid_state(format!(
            "state-only chart {} was not materialized",
            chart.id.0
        )));
    }
    if let Some(drawing) = workbook
        .state
        .drawings
        .values()
        .find(|drawing| drawing.raw_part_uri.is_none())
    {
        return Err(OmError::invalid_state(format!(
            "state-only drawing {} was not materialized",
            drawing.id.0
        )));
    }
    Ok(())
}

fn xml_error(error: impl std::fmt::Display) -> OmError {
    OmError::new(OmErrorCode::Parse, error.to_string())
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use excel_model::{
        ChartModel, ChartObjectModel, ChartSheetBinding, ChartType, DefinedNameTable, DrawingModel,
        DrawingObjectModel, WorkbookState, WorksheetData,
    };
    use office_common::{
        AbsoluteAnchor, ChartId, ChartObjectId, DrawingAnchor, DrawingId, DrawingObjectId,
        FileFormat, LoadOptions, ObjectPlacement, OmErrorCode, PointEmu, SaveOptions, SheetId,
        SheetKind, SheetVisibility, SizeEmu, WorkbookId, WorkbookModel, WorksheetModel,
    };
    use office_opc::{CompressionMethod, OpcPackage, OpcPart};

    use super::super::{
        LoadedXlsxWorkbook, PendingChartRelationshipGraph, PendingDrawingRelationshipGraph,
        PendingPackagePart, PendingPackageRelationship, XlsxCodec,
    };
    use super::{
        CHART_PART_CONTENT_TYPE, CONTENT_TYPES_PART_NAME, DRAWING_PART_CONTENT_TYPE,
        DRAWING_RELATIONSHIP_TYPE, RELATIONSHIPS_PART_CONTENT_TYPE, append_chart_anchors,
        append_content_type_override_if_missing, append_relationship, attach_drawing_element,
        insert_sheet_into_workbook_xml, materialize_state_only_chart_graphs,
    };

    fn base_workbook() -> LoadedXlsxWorkbook {
        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#
            .to_vec();
        LoadedXlsxWorkbook {
            state: WorkbookState {
                model: WorkbookModel {
                    id: WorkbookId(7),
                    display_name: "Book1".to_string(),
                    format: FileFormat::Xlsx,
                    date1904: false,
                    is_addin: false,
                },
                worksheets: vec![WorksheetModel {
                    id: SheetId(1),
                    workbook_id: WorkbookId(7),
                    name: "Sheet1".to_string(),
                    kind: SheetKind::Worksheet,
                    visibility: SheetVisibility::Visible,
                    relationship_id: Some("rId1".to_string()),
                    part_uri: Some("xl/worksheets/sheet1.xml".to_string()),
                }],
                worksheet_data: BTreeMap::from([(
                    SheetId(1),
                    WorksheetData {
                        source_xml: sheet_xml.clone(),
                        ..WorksheetData::default()
                    },
                )]),
                defined_names: DefinedNameTable::default(),
                charts: BTreeMap::new(),
                drawings: BTreeMap::new(),
                chart_sheets: BTreeMap::new(),
                opaque_parts: Vec::new(),
            },
            package: OpcPackage::try_new(vec![
                OpcPart {
                    name: "[Content_Types].xml".to_string(),
                    content_type: None,
                    compression: CompressionMethod::Stored,
                    bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#
                        .to_vec(),
                },
                OpcPart {
                    name: "_rels/.rels".to_string(),
                    content_type: Some(RELATIONSHIPS_PART_CONTENT_TYPE.to_string()),
                    compression: CompressionMethod::Stored,
                    bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#
                        .to_vec(),
                },
                OpcPart {
                    name: "xl/workbook.xml".to_string(),
                    content_type: Some(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                            .to_string(),
                    ),
                    compression: CompressionMethod::Stored,
                    bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#
                        .to_vec(),
                },
                OpcPart {
                    name: "xl/_rels/workbook.xml.rels".to_string(),
                    content_type: Some(RELATIONSHIPS_PART_CONTENT_TYPE.to_string()),
                    compression: CompressionMethod::Stored,
                    bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#
                        .to_vec(),
                },
                OpcPart {
                    name: "xl/worksheets/sheet1.xml".to_string(),
                    content_type: Some(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
                            .to_string(),
                    ),
                    compression: CompressionMethod::Stored,
                    bytes: sheet_xml,
                },
            ])
            .expect("chart-graph test package should have valid part identities"),
            detected_format: FileFormat::Xlsx,
            calculation_properties: Default::default(),
            support_parts: Default::default(),
            worksheet_support_parts: BTreeMap::new(),
            sheet_drawing_support_parts: BTreeMap::new(),
            pending_drawing_relationship_graphs: BTreeMap::new(),
            pending_chart_relationship_graphs: BTreeMap::new(),
            active_content_inventory: Default::default(),
            digital_signature_inventory: Default::default(),
            external_data_inventory: Default::default(),
        }
    }

    fn state_only_chart(chart_id: ChartId) -> ChartModel {
        ChartModel {
            id: chart_id,
            workbook_id: WorkbookId(7),
            chart_type: ChartType::Column,
            style: None,
            series: Vec::new(),
            title: None,
            legend: None,
            axes: Vec::new(),
            groups: Vec::new(),
            vary_by_categories: None,
            gap_width: None,
            gap_depth: None,
            overlap: None,
            bar_shape: None,
            has_series_lines: None,
            has_drop_lines: None,
            has_hi_lo_lines: None,
            has_up_down_bars: None,
            first_slice_angle: None,
            explosion: None,
            bubble_scale: None,
            show_negative_bubbles: None,
            has_3d_shading: None,
            doughnut_hole_size: None,
            second_plot_size: None,
            size_represents: None,
            split_type: None,
            split_value: None,
            data_labels: None,
            data_table: None,
            data_table_dirty: false,
            plot_area_layout: None,
            plot_area_layout_dirty: false,
            show_data_labels_over_maximum: None,
            display_blanks_as: None,
            plot_visible_only: None,
            view_3d: None,
            view_3d_dirty: false,
            rounded_corners: None,
            protection: None,
            protection_dirty: false,
            raw_part_uri: None,
            series_topology_dirty: false,
            content_dirty: true,
            dirty: true,
        }
    }

    fn state_only_chart_object(
        chart_object_id: ChartObjectId,
        chart_id: ChartId,
        host_sheet_id: SheetId,
    ) -> ChartObjectModel {
        ChartObjectModel {
            id: chart_object_id,
            anchor_attrs: BTreeMap::new(),
            position_attrs: BTreeMap::new(),
            extents_attrs: BTreeMap::new(),
            marker_attrs: Default::default(),
            graphic_frame_attrs: BTreeMap::new(),
            graphic_frame_transform_xml: None,
            graphic_data_attrs: BTreeMap::new(),
            graphic_data_child_xmls: Vec::new(),
            chart_reference_attrs: BTreeMap::new(),
            non_visual_frame_attrs: BTreeMap::new(),
            graphic_attrs: BTreeMap::new(),
            non_visual_id: Some(chart_object_id.0 as u32),
            non_visual_attrs: BTreeMap::new(),
            non_visual_child_xml: None,
            non_visual_frame_properties_xml: None,
            client_data_attrs: BTreeMap::new(),
            client_data_xml: None,
            anchor_extension_xmls: Vec::new(),
            workbook_id: WorkbookId(7),
            host_sheet_id,
            chart_id,
            name: format!("Chart {}", chart_object_id.0),
            anchor: Some(DrawingAnchor::Absolute(AbsoluteAnchor {
                position: PointEmu {
                    x: office_common::Emu(0),
                    y: office_common::Emu(0),
                },
                extents: SizeEmu {
                    cx: office_common::Emu(5_486_400),
                    cy: office_common::Emu(3_200_400),
                },
            })),
            placement: ObjectPlacement::FreeFloating,
            z_order: Some(0),
            raw_binding: None,
            dirty: true,
        }
    }

    fn materialized_embedded_chart_workbook() -> LoadedXlsxWorkbook {
        let mut workbook = base_workbook();
        let chart_id = ChartId(1);
        let drawing_id = DrawingId(1);
        workbook
            .state
            .charts
            .insert(chart_id, state_only_chart(chart_id));
        workbook.state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![DrawingObjectModel::ChartFrame(state_only_chart_object(
                    ChartObjectId(1),
                    chart_id,
                    SheetId(1),
                ))],
                raw_part_uri: None,
                dirty: true,
            },
        );
        materialize_state_only_chart_graphs(workbook).expect("materialize chart graph")
    }

    #[test]
    fn direct_save_materializes_state_only_embedded_chart_graph() {
        let mut workbook = base_workbook();
        let chart_id = ChartId(1);
        let drawing_id = DrawingId(1);
        workbook
            .state
            .charts
            .insert(chart_id, state_only_chart(chart_id));
        workbook.state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![DrawingObjectModel::ChartFrame(state_only_chart_object(
                    ChartObjectId(1),
                    chart_id,
                    SheetId(1),
                ))],
                raw_part_uri: None,
                dirty: true,
            },
        );

        let bytes = XlsxCodec
            .save(&workbook, SaveOptions::default())
            .expect("direct codec save should materialize embedded chart graph");
        assert!(workbook.state.charts[&chart_id].raw_part_uri.is_none());
        assert!(workbook.state.drawings[&drawing_id].raw_part_uri.is_none());

        let package = OpcPackage::from_bytes(&bytes).expect("saved package");
        for part_uri in [
            "xl/charts/chart1.xml",
            "xl/drawings/drawing1.xml",
            "xl/drawings/_rels/drawing1.xml.rels",
            "xl/worksheets/_rels/sheet1.xml.rels",
        ] {
            assert!(package.contains(part_uri), "missing {part_uri}");
        }
        let reopened = XlsxCodec
            .load(&bytes, LoadOptions::default())
            .expect("reopen materialized workbook");
        assert_eq!(
            reopened.state.charts[&chart_id].raw_part_uri.as_deref(),
            Some("xl/charts/chart1.xml")
        );
        assert_eq!(
            reopened.state.drawings[&drawing_id].raw_part_uri.as_deref(),
            Some("xl/drawings/drawing1.xml")
        );
    }

    #[test]
    fn direct_save_materializes_mixed_drawing_with_opaque_anchor_in_order() {
        let mut workbook = base_workbook();
        let chart_id = ChartId(1);
        let drawing_id = DrawingId(1);
        workbook
            .state
            .charts
            .insert(chart_id, state_only_chart(chart_id));
        let raw_anchor_xml = r#"<xdr:oneCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="914400" cy="457200"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Preserved Shape"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><foo:meta/></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Keep me</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>"#.to_string();
        workbook.state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![
                    DrawingObjectModel::UnsupportedRaw {
                        id: DrawingObjectId(9),
                        raw_part_uri: None,
                        raw_anchor_xml: raw_anchor_xml.clone(),
                        root_namespace_attrs: BTreeMap::from([(
                            "xmlns:foo".to_string(),
                            "urn:opaque-shape".to_string(),
                        )]),
                        relationship_ids: Vec::new(),
                        non_visual_id: Some(2),
                    },
                    DrawingObjectModel::ChartFrame(state_only_chart_object(
                        ChartObjectId(1),
                        chart_id,
                        SheetId(1),
                    )),
                ],
                raw_part_uri: None,
                dirty: true,
            },
        );

        let bytes = XlsxCodec
            .save(&workbook, SaveOptions::default())
            .expect("mixed state-only drawing should materialize");
        let package = OpcPackage::from_bytes(&bytes).expect("saved package");
        let drawing_xml = String::from_utf8(
            package
                .part("xl/drawings/drawing1.xml")
                .expect("drawing part")
                .bytes
                .clone(),
        )
        .expect("drawing XML");
        assert!(drawing_xml.contains(r#"xmlns:foo="urn:opaque-shape""#));
        assert!(drawing_xml.contains(&raw_anchor_xml), "{drawing_xml}");
        assert!(
            drawing_xml.find("Preserved Shape").expect("raw shape")
                < drawing_xml.find("Chart 1").expect("chart frame")
        );

        let reopened = XlsxCodec
            .load(&bytes, LoadOptions::default())
            .expect("reopen mixed drawing");
        let drawing = reopened
            .state
            .drawings
            .values()
            .next()
            .expect("reopened drawing");
        assert!(matches!(
            drawing.objects.as_slice(),
            [DrawingObjectModel::UnsupportedRaw { relationship_ids, .. }, DrawingObjectModel::ChartFrame(_)]
                if relationship_ids.is_empty()
        ));
        let saved_again = XlsxCodec
            .save(&reopened, SaveOptions::default())
            .expect("save reopened mixed drawing");
        let saved_again_package =
            OpcPackage::from_bytes(&saved_again).expect("second saved package");
        assert_eq!(
            saved_again_package
                .part("xl/drawings/drawing1.xml")
                .expect("second drawing part")
                .bytes,
            package
                .part("xl/drawings/drawing1.xml")
                .expect("first drawing part")
                .bytes
        );
    }

    #[test]
    fn direct_save_materializes_transitive_opaque_drawing_relationship_graph() {
        let mut workbook = base_workbook();
        let chart_id = ChartId(1);
        let drawing_id = DrawingId(1);
        workbook
            .state
            .charts
            .insert(chart_id, state_only_chart(chart_id));
        workbook
            .package
            .add_part(OpcPart {
                name: "xl/media/image1.png".to_string(),
                content_type: Some("image/png".to_string()),
                compression: CompressionMethod::Stored,
                bytes: vec![9, 9, 9],
            })
            .expect("reserve image1 target");
        let content_types = append_content_type_override_if_missing(
            workbook
                .package
                .part(CONTENT_TYPES_PART_NAME)
                .expect("content types manifest")
                .bytes
                .as_slice(),
            "xl/media/image1.png",
            "image/png",
        )
        .expect("declare reserved image content type");
        workbook
            .package
            .replace_part_bytes(CONTENT_TYPES_PART_NAME, content_types)
            .expect("replace content types manifest");
        workbook.state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![
                    DrawingObjectModel::UnsupportedRaw {
                        id: DrawingObjectId(9),
                        raw_part_uri: None,
                        raw_anchor_xml: r#"<xdr:oneCellAnchor><xdr:from/><xdr:ext cx="1" cy="1"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Linked Picture"><a:hlinkClick r:id="rIdLink"/></xdr:cNvPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1" r:link="rIdImageAlias"/></xdr:blipFill></xdr:pic><xdr:oleObject r:id="rIdPackage"/><xdr:clientData/></xdr:oneCellAnchor>"#.to_string(),
                        root_namespace_attrs: BTreeMap::new(),
                        relationship_ids: vec![
                            "rIdLink".to_string(),
                            "rId1".to_string(),
                            "rIdImageAlias".to_string(),
                            "rIdPackage".to_string(),
                        ],
                        non_visual_id: Some(2),
                    },
                    DrawingObjectModel::ChartFrame(state_only_chart_object(
                        ChartObjectId(1),
                        chart_id,
                        SheetId(1),
                    )),
                ],
                raw_part_uri: None,
                dirty: true,
            },
        );
        let root_relationships = vec![
            PendingPackageRelationship {
                relationship_id: "rId1".to_string(),
                relationship_type:
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                        .to_string(),
                target: "xl/media/image1.png".to_string(),
                target_mode: None,
            },
            PendingPackageRelationship {
                relationship_id: "rIdLink".to_string(),
                relationship_type:
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                        .to_string(),
                target: "https://example.com/picture".to_string(),
                target_mode: Some("External".to_string()),
            },
            PendingPackageRelationship {
                relationship_id: "rIdImageAlias".to_string(),
                relationship_type:
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                        .to_string(),
                target: "xl/media/image1.png".to_string(),
                target_mode: None,
            },
            PendingPackageRelationship {
                relationship_id: "rIdPackage".to_string(),
                relationship_type:
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
                        .to_string(),
                target: "xl/embeddings/oleObject1.bin".to_string(),
                target_mode: Some("Internal".to_string()),
            },
        ];
        workbook.pending_drawing_relationship_graphs.insert(
            drawing_id,
            PendingDrawingRelationshipGraph {
                source_drawing_part_uri: "xl/drawings/drawing7.xml".to_string(),
                root_relationships_part_source_bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:keep="urn:keep" keep:root="yes">
  <pr:Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png" keep:edge="image"/>
  <pr:Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/picture" TargetMode="External"/>
  <pr:Relationship Id="rIdImageAlias" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png" keep:edge="alias"/>
  <pr:Relationship Id="rIdPackage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/oleObject1.bin" TargetMode="Internal"/>
  <pr:Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart7.xml"/>
</pr:Relationships>"#.to_vec(),
                root_relationships_part_compression: CompressionMethod::Stored,
                root_relationships,
                parts: BTreeMap::from([
                    (
                        "customXml/item1.xml".to_string(),
                        PendingPackagePart {
                            source_part_uri: "customXml/item1.xml".to_string(),
                            bytes: b"<custom>keep</custom>".to_vec(),
                            content_type: Some("application/xml".to_string()),
                            compression: CompressionMethod::Stored,
                            relationships_part_source_bytes: None,
                            relationships_part_compression: None,
                            relationships: Vec::new(),
                        },
                    ),
                    (
                        "xl/embeddings/oleObject1.bin".to_string(),
                        PendingPackagePart {
                            source_part_uri: "xl/embeddings/oleObject1.bin".to_string(),
                            bytes: vec![4, 5, 6],
                            content_type: Some("application/vnd.openxmlformats-officedocument.oleObject".to_string()),
                            compression: CompressionMethod::Stored,
                            relationships_part_source_bytes: Some(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:keep="urn:keep" keep:root="child"><Relationship Id="rIdChild" Type="urn:child" Target="../../customXml/item1.xml" keep:edge="child"/></Relationships>"#.to_vec()),
                            relationships_part_compression: Some(CompressionMethod::Stored),
                            relationships: vec![PendingPackageRelationship {
                                relationship_id: "rIdChild".to_string(),
                                relationship_type: "urn:child".to_string(),
                                target: "customXml/item1.xml".to_string(),
                                target_mode: None,
                            }],
                        },
                    ),
                    (
                        "xl/media/image1.png".to_string(),
                        PendingPackagePart {
                            source_part_uri: "xl/media/image1.png".to_string(),
                            bytes: vec![1, 2, 3],
                            content_type: Some("image/png".to_string()),
                            compression: CompressionMethod::Stored,
                            relationships_part_source_bytes: None,
                            relationships_part_compression: None,
                            relationships: Vec::new(),
                        },
                    ),
                ]),
            },
        );

        let bytes = XlsxCodec
            .save(&workbook, SaveOptions::default())
            .expect("opaque relationship graph should materialize");
        let package = OpcPackage::from_bytes(&bytes).expect("saved package");
        assert_eq!(
            package
                .part("xl/media/image2.png")
                .expect("copied image")
                .bytes,
            vec![1, 2, 3]
        );
        assert_eq!(
            package
                .part("xl/embeddings/oleObject1.bin")
                .expect("copied package")
                .bytes,
            vec![4, 5, 6]
        );
        assert!(package.contains("customXml/item1.xml"));
        let drawing_relationships = String::from_utf8(
            package
                .part("xl/drawings/_rels/drawing1.xml.rels")
                .expect("drawing relationships")
                .bytes
                .clone(),
        )
        .expect("drawing relationships XML");
        assert!(drawing_relationships.contains(r#"keep:root="yes""#));
        assert!(drawing_relationships.contains(r#"keep:edge="image""#));
        assert!(drawing_relationships.contains(r#"Id="rId1""#));
        assert!(drawing_relationships.contains(r#"Target="../media/image2.png""#));
        assert_eq!(
            drawing_relationships.matches("../media/image2.png").count(),
            2
        );
        assert!(drawing_relationships.contains(r#"Id="rId2""#));
        assert!(!drawing_relationships.contains(r#"Id="rId9""#));
        assert!(drawing_relationships.contains(r#"TargetMode="External""#));
        assert!(drawing_relationships.contains(r#"TargetMode="Internal""#));
        let package_relationships = String::from_utf8(
            package
                .part("xl/embeddings/_rels/oleObject1.bin.rels")
                .expect("copied package relationships")
                .bytes
                .clone(),
        )
        .expect("package relationships XML");
        assert!(package_relationships.contains(r#"keep:root="child""#));
        assert!(package_relationships.contains(r#"keep:edge="child""#));
        assert!(package_relationships.contains(r#"Target="../../customXml/item1.xml""#));

        let reopened = XlsxCodec
            .load(&bytes, LoadOptions::default())
            .expect("reopen relationship-backed drawing");
        let saved_again = XlsxCodec
            .save(&reopened, SaveOptions::default())
            .expect("save relationship-backed drawing twice");
        let saved_again_package = OpcPackage::from_bytes(&saved_again).expect("second package");
        assert_eq!(
            saved_again_package
                .part("xl/media/image2.png")
                .expect("second copied image")
                .bytes,
            vec![1, 2, 3]
        );
        assert_eq!(
            saved_again_package
                .part("xl/embeddings/_rels/oleObject1.bin.rels")
                .expect("second copied package relationships")
                .bytes,
            package
                .part("xl/embeddings/_rels/oleObject1.bin.rels")
                .expect("first copied package relationships")
                .bytes
        );
    }

    #[test]
    fn direct_save_materializes_transitive_opaque_chart_relationship_graph() {
        let mut workbook = base_workbook();
        let chart_id = ChartId(1);
        let drawing_id = DrawingId(1);
        workbook
            .state
            .charts
            .insert(chart_id, state_only_chart(chart_id));
        workbook.state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![DrawingObjectModel::ChartFrame(state_only_chart_object(
                    ChartObjectId(1),
                    chart_id,
                    SheetId(1),
                ))],
                raw_part_uri: None,
                dirty: true,
            },
        );
        workbook.pending_chart_relationship_graphs.insert(
            chart_id,
            PendingChartRelationshipGraph {
                source_chart_part_uri: "xl/charts/chart7.xml".to_string(),
                source_chart_part_bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:keep="urn:keep" keep:root="chart">
  <c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/></c:barChart></c:plotArea></c:chart>
  <c:userShapes r:id="rIdUserShapes" keep:edge="shape"/>
  <c:externalData r:id="rIdExternal" keep:edge="external"/>
  <c:extLst><c:ext uri="urn:opaque"><keep:payload value="preserve"/></c:ext></c:extLst>
</c:chartSpace>"#
                    .to_vec(),
                source_chart_part_compression: CompressionMethod::Stored,
                root_relationships_part_source_bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:keep="urn:keep" keep:root="chart-rels">
  <pr:Relationship Id="rIdUserShapes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes" Target="../drawings/userShapes7.xml" TargetMode="Internal" keep:edge="shape"/>
  <pr:Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="https://example.test/data.xlsx" TargetMode="External" keep:edge="external"/>
  <pr:Relationship Id="rIdStyleIgnored" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style7.xml"/>
</pr:Relationships>"#
                    .to_vec(),
                root_relationships_part_compression: CompressionMethod::Stored,
                root_relationships: vec![
                    PendingPackageRelationship {
                        relationship_id: "rIdUserShapes".to_string(),
                        relationship_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes".to_string(),
                        target: "xl/drawings/userShapes7.xml".to_string(),
                        target_mode: Some("Internal".to_string()),
                    },
                    PendingPackageRelationship {
                        relationship_id: "rIdExternal".to_string(),
                        relationship_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink".to_string(),
                        target: "https://example.test/data.xlsx".to_string(),
                        target_mode: Some("External".to_string()),
                    },
                ],
                parts: BTreeMap::from([
                    (
                        "xl/drawings/userShapes7.xml".to_string(),
                        PendingPackagePart {
                            source_part_uri: "xl/drawings/userShapes7.xml".to_string(),
                            bytes: br#"<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" keep="yes"/>"#.to_vec(),
                            content_type: Some("application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml".to_string()),
                            compression: CompressionMethod::Stored,
                            relationships_part_source_bytes: Some(br#"<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:keep="urn:keep" keep:root="child"><pr:Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image7.png" keep:edge="image"/></pr:Relationships>"#.to_vec()),
                            relationships_part_compression: Some(CompressionMethod::Stored),
                            relationships: vec![PendingPackageRelationship {
                                relationship_id: "rIdImage".to_string(),
                                relationship_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image".to_string(),
                                target: "xl/media/image7.png".to_string(),
                                target_mode: None,
                            }],
                        },
                    ),
                    (
                        "xl/media/image7.png".to_string(),
                        PendingPackagePart {
                            source_part_uri: "xl/media/image7.png".to_string(),
                            bytes: vec![7, 8, 9],
                            content_type: Some("image/png".to_string()),
                            compression: CompressionMethod::Stored,
                            relationships_part_source_bytes: None,
                            relationships_part_compression: None,
                            relationships: Vec::new(),
                        },
                    ),
                ]),
            },
        );

        let bytes = XlsxCodec
            .save(&workbook, SaveOptions::default())
            .expect("opaque chart relationship graph should materialize");
        let package = OpcPackage::from_bytes(&bytes).expect("saved package");
        let chart_xml = String::from_utf8(
            package
                .part("xl/charts/chart1.xml")
                .expect("materialized chart")
                .bytes
                .clone(),
        )
        .expect("chart XML");
        assert!(chart_xml.contains(r#"keep:root="chart""#));
        assert!(chart_xml.contains(r#"r:id="rIdUserShapes""#));
        assert!(chart_xml.contains(r#"r:id="rIdExternal""#));
        assert!(chart_xml.contains(r#"<keep:payload value="preserve""#));

        let chart_relationships = String::from_utf8(
            package
                .part("xl/charts/_rels/chart1.xml.rels")
                .expect("materialized chart relationships")
                .bytes
                .clone(),
        )
        .expect("chart relationships XML");
        assert!(chart_relationships.contains(r#"<pr:Relationships"#));
        assert!(chart_relationships.contains(r#"keep:root="chart-rels""#));
        assert!(chart_relationships.contains(r#"keep:edge="shape""#));
        assert!(chart_relationships.contains(r#"Target="../drawings/userShapes7.xml""#));
        assert!(chart_relationships.contains(r#"Target="https://example.test/data.xlsx""#));
        assert!(chart_relationships.contains(r#"TargetMode="External""#));
        assert!(!chart_relationships.contains("rIdStyleIgnored"));

        assert_eq!(
            package
                .part("xl/media/image7.png")
                .expect("recursive image target")
                .bytes,
            vec![7, 8, 9]
        );
        let child_relationships = String::from_utf8(
            package
                .part("xl/drawings/_rels/userShapes7.xml.rels")
                .expect("user shapes relationships")
                .bytes
                .clone(),
        )
        .expect("user shapes relationships XML");
        assert!(child_relationships.contains(r#"keep:root="child""#));
        assert!(child_relationships.contains(r#"keep:edge="image""#));
        assert!(child_relationships.contains(r#"Target="../media/image7.png""#));

        let reopened = XlsxCodec
            .load(&bytes, LoadOptions::default())
            .expect("reopen chart relationship graph");
        let saved_again = XlsxCodec
            .save(&reopened, SaveOptions::default())
            .expect("save chart relationship graph twice");
        let saved_again_package = OpcPackage::from_bytes(&saved_again).expect("second package");
        assert_eq!(
            saved_again_package
                .part("xl/charts/_rels/chart1.xml.rels")
                .expect("second chart relationships")
                .bytes,
            package
                .part("xl/charts/_rels/chart1.xml.rels")
                .expect("first chart relationships")
                .bytes
        );
        assert_eq!(
            saved_again_package
                .part("xl/drawings/_rels/userShapes7.xml.rels")
                .expect("second user shapes relationships")
                .bytes,
            package
                .part("xl/drawings/_rels/userShapes7.xml.rels")
                .expect("first user shapes relationships")
                .bytes
        );
    }

    #[test]
    fn direct_save_rejects_relationship_backed_opaque_anchor_before_materializing() {
        let mut workbook = base_workbook();
        let drawing_id = DrawingId(1);
        workbook.state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![DrawingObjectModel::UnsupportedRaw {
                    id: DrawingObjectId(9),
                    raw_part_uri: None,
                    raw_anchor_xml: r#"<xdr:oneCellAnchor><xdr:from/><xdr:ext cx="1" cy="1"/><xdr:pic><xdr:blipFill><a:blip r:embed="rIdImage1"/></xdr:blipFill></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>"#.to_string(),
                    root_namespace_attrs: BTreeMap::new(),
                    relationship_ids: vec!["rIdImage1".to_string()],
                    non_visual_id: Some(2),
                }],
                raw_part_uri: None,
                dirty: true,
            },
        );
        let package_part_names = workbook
            .package
            .parts()
            .iter()
            .map(|part| part.name.clone())
            .collect::<Vec<_>>();

        let error = XlsxCodec
            .save(&workbook, SaveOptions::default())
            .expect_err("relationship-backed raw anchor should require graph copying");
        assert_eq!(error.code, OmErrorCode::Unsupported);
        assert!(error.message.contains("requires relationship copying"));
        assert_eq!(
            workbook
                .package
                .parts()
                .iter()
                .map(|part| part.name.clone())
                .collect::<Vec<_>>(),
            package_part_names
        );
        assert!(workbook.state.drawings[&drawing_id].raw_part_uri.is_none());
    }

    #[test]
    fn direct_save_rewrites_stale_drawing_placeholder_without_live_relationship() {
        let mut workbook = base_workbook();
        let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:ext="urn:extension"><sheetData/><drawing r:id="rId9" ext:id="keep"/></worksheet>"#
            .to_vec();
        workbook
            .package
            .replace_part_bytes("xl/worksheets/sheet1.xml", worksheet_xml.clone())
            .expect("replace worksheet XML");
        workbook
            .state
            .worksheet_data
            .get_mut(&SheetId(1))
            .expect("worksheet data")
            .source_xml = worksheet_xml;
        workbook
            .package
            .add_part(OpcPart {
                name: "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-package.relationships+xml".to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test" TargetMode="External"/></Relationships>"#
                    .to_vec(),
            })
            .expect("add worksheet relationships");

        let chart_id = ChartId(1);
        let drawing_id = DrawingId(1);
        workbook
            .state
            .charts
            .insert(chart_id, state_only_chart(chart_id));
        workbook.state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![DrawingObjectModel::ChartFrame(state_only_chart_object(
                    ChartObjectId(1),
                    chart_id,
                    SheetId(1),
                ))],
                raw_part_uri: None,
                dirty: true,
            },
        );

        let bytes = XlsxCodec
            .save(&workbook, SaveOptions::default())
            .expect("stale drawing placeholder should be rebound");
        let package = OpcPackage::from_bytes(&bytes).expect("saved package");
        let worksheet_xml = String::from_utf8(
            package
                .part("xl/worksheets/sheet1.xml")
                .expect("worksheet part")
                .bytes
                .clone(),
        )
        .expect("worksheet utf8");
        assert_eq!(worksheet_xml.matches("<drawing").count(), 1);
        assert!(worksheet_xml.contains("<drawing r:id=\"rId2\""));
        assert!(worksheet_xml.contains("ext:id=\"keep\""));
        assert!(!worksheet_xml.contains("rId9"));
        let relationships = String::from_utf8(
            package
                .part("xl/worksheets/_rels/sheet1.xml.rels")
                .expect("worksheet relationships")
                .bytes
                .clone(),
        )
        .expect("relationships utf8");
        assert!(relationships.contains("Id=\"rId1\""));
        assert!(relationships.contains("Id=\"rId2\""));
        assert!(relationships.contains(DRAWING_RELATIONSHIP_TYPE));
    }

    #[test]
    fn direct_save_rejects_state_only_drawing_when_host_has_live_drawing_relationship() {
        let mut workbook = base_workbook();
        let worksheet_xml = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rId1"/></worksheet>"#
            .to_vec();
        workbook
            .package
            .replace_part_bytes("xl/worksheets/sheet1.xml", worksheet_xml.clone())
            .expect("replace worksheet XML");
        workbook
            .state
            .worksheet_data
            .get_mut(&SheetId(1))
            .expect("worksheet data")
            .source_xml = worksheet_xml;
        workbook
            .package
            .add_part(OpcPart {
                name: "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-package.relationships+xml".to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: format!(
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{DRAWING_RELATIONSHIP_TYPE}\" Target=\"../drawings/drawing9.xml\"/></Relationships>"
                )
                .into_bytes(),
            })
            .expect("add worksheet relationships");

        let chart_id = ChartId(1);
        let drawing_id = DrawingId(1);
        workbook
            .state
            .charts
            .insert(chart_id, state_only_chart(chart_id));
        workbook.state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![DrawingObjectModel::ChartFrame(state_only_chart_object(
                    ChartObjectId(1),
                    chart_id,
                    SheetId(1),
                ))],
                raw_part_uri: None,
                dirty: true,
            },
        );

        let error = XlsxCodec
            .save(&workbook, SaveOptions::default())
            .expect_err("live drawing relationships must not be replaced");
        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("live drawing relationship"));
    }

    #[test]
    fn direct_save_materializes_wholly_state_only_chart_sheet_graph() {
        let mut workbook = base_workbook();
        let sheet_id = SheetId(2);
        let chart_id = ChartId(1);
        let drawing_id = DrawingId(1);
        workbook.state.worksheets.insert(
            0,
            WorksheetModel {
                id: sheet_id,
                workbook_id: WorkbookId(7),
                name: "Chart1".to_string(),
                kind: SheetKind::ChartSheet,
                visibility: SheetVisibility::Visible,
                relationship_id: None,
                part_uri: None,
            },
        );
        workbook
            .state
            .worksheet_data
            .insert(sheet_id, WorksheetData::default());
        workbook
            .state
            .charts
            .insert(chart_id, state_only_chart(chart_id));
        workbook.state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: WorkbookId(7),
                host_sheet_id: sheet_id,
                objects: vec![DrawingObjectModel::ChartFrame(state_only_chart_object(
                    ChartObjectId(1),
                    chart_id,
                    sheet_id,
                ))],
                raw_part_uri: None,
                dirty: true,
            },
        );
        workbook.state.chart_sheets.insert(
            sheet_id,
            ChartSheetBinding {
                sheet_id,
                chart_id,
                drawing_id: Some(drawing_id),
                raw_part_uri: None,
            },
        );

        let bytes = XlsxCodec
            .save(&workbook, SaveOptions::default())
            .expect("direct codec save should materialize chart sheet graph");
        assert!(workbook.state.worksheets[0].part_uri.is_none());
        assert!(
            workbook.state.chart_sheets[&sheet_id]
                .raw_part_uri
                .is_none()
        );

        let package = OpcPackage::from_bytes(&bytes).expect("saved package");
        for part_uri in [
            "xl/chartsheets/sheet1.xml",
            "xl/chartsheets/_rels/sheet1.xml.rels",
            "xl/drawings/drawing1.xml",
            "xl/drawings/_rels/drawing1.xml.rels",
            "xl/charts/chart1.xml",
        ] {
            assert!(package.contains(part_uri), "missing {part_uri}");
        }
        let workbook_xml = String::from_utf8(
            package
                .part("xl/workbook.xml")
                .expect("workbook part")
                .bytes
                .clone(),
        )
        .expect("workbook utf8");
        assert!(workbook_xml.find("name=\"Chart1\"") < workbook_xml.find("name=\"Sheet1\""));
        assert!(workbook_xml.contains("name=\"Chart1\" sheetId=\"2\" r:id=\"rId2\""));
        let workbook_rels = String::from_utf8(
            package
                .part("xl/_rels/workbook.xml.rels")
                .expect("workbook relationships")
                .bytes
                .clone(),
        )
        .expect("workbook relationships utf8");
        assert!(workbook_rels.contains("Id=\"rId2\""));
        assert!(workbook_rels.contains("Target=\"chartsheets/sheet1.xml\""));
        let content_types = String::from_utf8(
            package
                .part("[Content_Types].xml")
                .expect("content types")
                .bytes
                .clone(),
        )
        .expect("content types utf8");
        for part_name in [
            "/xl/chartsheets/sheet1.xml",
            "/xl/drawings/drawing1.xml",
            "/xl/charts/chart1.xml",
        ] {
            assert!(content_types.contains(part_name), "missing {part_name}");
        }
        let reopened = XlsxCodec
            .load(&bytes, LoadOptions::default())
            .expect("reopen materialized chart sheet");
        assert_eq!(reopened.state.worksheets.len(), 2);
        assert_eq!(reopened.state.worksheets[0].kind, SheetKind::ChartSheet);
        assert_eq!(
            reopened.state.chart_sheets[&sheet_id]
                .raw_part_uri
                .as_deref(),
            Some("xl/chartsheets/sheet1.xml")
        );
        assert_eq!(
            reopened.state.charts[&chart_id].raw_part_uri.as_deref(),
            Some("xl/charts/chart1.xml")
        );
        let repeated = XlsxCodec
            .save(&reopened, SaveOptions::default())
            .expect("repeat chart sheet save");
        let repeated_package = OpcPackage::from_bytes(&repeated).expect("repeated package");
        assert_eq!(package.parts(), repeated_package.parts());
    }

    #[test]
    fn direct_save_appends_state_only_chart_to_existing_drawing() {
        let mut initial = base_workbook();
        let first_chart_id = ChartId(1);
        let drawing_id = DrawingId(1);
        initial
            .state
            .charts
            .insert(first_chart_id, state_only_chart(first_chart_id));
        initial.state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![DrawingObjectModel::ChartFrame(state_only_chart_object(
                    ChartObjectId(1),
                    first_chart_id,
                    SheetId(1),
                ))],
                raw_part_uri: None,
                dirty: true,
            },
        );
        let initial_bytes = XlsxCodec
            .save(&initial, SaveOptions::default())
            .expect("materialize initial drawing");
        let mut loaded = XlsxCodec
            .load(&initial_bytes, LoadOptions::default())
            .expect("load initial drawing");

        let second_chart_id = ChartId(2);
        let loaded_workbook_id = loaded.state.model.id;
        let mut second_chart = state_only_chart(second_chart_id);
        second_chart.workbook_id = loaded_workbook_id;
        loaded.state.charts.insert(second_chart_id, second_chart);
        let drawing = loaded
            .state
            .drawings
            .get_mut(&drawing_id)
            .expect("loaded drawing");
        let mut second_chart_object =
            state_only_chart_object(ChartObjectId(2), second_chart_id, SheetId(1));
        second_chart_object.workbook_id = loaded_workbook_id;
        second_chart_object.z_order = Some(1);
        drawing
            .objects
            .push(DrawingObjectModel::ChartFrame(second_chart_object));
        drawing.dirty = true;

        let saved = XlsxCodec
            .save(&loaded, SaveOptions::default())
            .expect("append state-only chart to existing drawing");
        assert!(loaded.state.charts[&second_chart_id].raw_part_uri.is_none());
        let package = OpcPackage::from_bytes(&saved).expect("saved package");
        assert!(package.contains("xl/charts/chart2.xml"));
        let drawing_rels = String::from_utf8(
            package
                .part("xl/drawings/_rels/drawing1.xml.rels")
                .expect("drawing relationships")
                .bytes
                .clone(),
        )
        .expect("drawing relationships utf8");
        assert!(drawing_rels.contains("Id=\"rId2\""));
        assert!(drawing_rels.contains("Target=\"../charts/chart2.xml\""));

        let reopened = XlsxCodec
            .load(&saved, LoadOptions::default())
            .expect("reopen appended chart");
        assert_eq!(reopened.state.charts.len(), 2);
        assert_eq!(reopened.state.drawings[&drawing_id].objects.len(), 2);
    }

    #[test]
    fn direct_save_rejects_partial_state_only_chart_binding() {
        let mut workbook = base_workbook();
        let chart_id = ChartId(1);
        let drawing_id = DrawingId(1);
        workbook
            .state
            .charts
            .insert(chart_id, state_only_chart(chart_id));
        let mut chart_object = state_only_chart_object(ChartObjectId(1), chart_id, SheetId(1));
        chart_object.raw_binding = Some("xl/drawings/drawing1.xml#rId1".to_string());
        workbook.state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![DrawingObjectModel::ChartFrame(chart_object)],
                raw_part_uri: None,
                dirty: true,
            },
        );

        let error = XlsxCodec
            .save(&workbook, SaveOptions::default())
            .expect_err("partial graph must be rejected");
        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("partial drawing binding"));
    }

    #[test]
    fn generated_children_follow_prefixed_container_namespaces() {
        let relationships = append_relationship(
            br#"<p:Relationships xmlns:p="http://schemas.openxmlformats.org/package/2006/relationships"/>"#,
            "rId1",
            "urn:test",
            "target.xml",
        )
        .expect("append prefixed relationship");
        assert!(
            String::from_utf8(relationships)
                .expect("relationships utf8")
                .contains("<p:Relationship ")
        );

        let content_types = append_content_type_override_if_missing(
            br#"<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"/>"#,
            "xl/charts/chart1.xml",
            "application/test",
        )
        .expect("append prefixed content type");
        assert!(
            String::from_utf8(content_types)
                .expect("content types utf8")
                .contains("<ct:Override ")
        );

        let workbook = insert_sheet_into_workbook_xml(
            br#"<s:workbook xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><s:sheets/></s:workbook>"#,
            0,
            "Chart1",
            SheetId(2),
            "rId2",
            SheetVisibility::Visible,
        )
        .expect("append prefixed sheet");
        assert!(
            String::from_utf8(workbook)
                .expect("workbook utf8")
                .contains("<s:sheet ")
        );

        let worksheet = attach_drawing_element(
            br#"<s:worksheet xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
            "rId3",
        )
        .expect("append prefixed drawing");
        assert!(
            String::from_utf8(worksheet)
                .expect("worksheet utf8")
                .contains("<s:drawing ")
        );
    }

    #[test]
    fn generated_xml_escapes_values_and_preserves_raw_attributes() {
        let workbook = insert_sheet_into_workbook_xml(
            br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets/></workbook>"#,
            0,
            "R&D \"Q\"",
            SheetId(2),
            "rId2",
            SheetVisibility::Visible,
        )
        .expect("insert escaped sheet name");
        let workbook = String::from_utf8(workbook).expect("workbook utf8");
        assert!(
            workbook.contains("name=\"R&amp;D &quot;Q&quot;\""),
            "{workbook}"
        );

        let worksheet = attach_drawing_element(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" custom="A&amp;B&quot;C"/>"#,
            "rId1",
        )
        .expect("preserve escaped root attribute");
        let worksheet = String::from_utf8(worksheet).expect("worksheet utf8");
        assert!(worksheet.contains("custom=\"A&amp;B&quot;C\""));
        assert!(worksheet.contains("<drawing r:id=\"rId1\""));
    }

    #[test]
    fn package_insertions_ignore_nested_foreign_containers() {
        let relationships = append_relationship(
            br#"<p:Relationships xmlns:p="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:ext="urn:extension"><ext:wrapper><ext:Relationships><ext:Relationship Id="nested"/></ext:Relationships></ext:wrapper></p:Relationships>"#,
            "rId1",
            "urn:test",
            "target.xml",
        )
        .expect("append root relationship");
        let relationships = String::from_utf8(relationships).expect("relationships utf8");
        assert!(
            relationships.contains(
                "<ext:Relationships><ext:Relationship Id=\"nested\"/></ext:Relationships>"
            )
        );
        assert_eq!(relationships.matches("Target=\"target.xml\"").count(), 1);

        let workbook = insert_sheet_into_workbook_xml(
            br#"<s:workbook xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:ext="urn:extension"><s:extLst><ext:sheets/></s:extLst><s:sheets/></s:workbook>"#,
            0,
            "Chart1",
            SheetId(2),
            "rId2",
            SheetVisibility::Visible,
        )
        .expect("insert into workbook sheets");
        let workbook = String::from_utf8(workbook).expect("workbook utf8");
        assert!(workbook.contains("<ext:sheets/>"));
        assert_eq!(workbook.matches("name=\"Chart1\"").count(), 1);
        assert!(workbook.contains("<s:sheets><s:sheet "));
    }

    #[test]
    fn drawing_namespace_prefixes_must_have_expected_identity() {
        let error = attach_drawing_element(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="urn:not-relationships"/>"#,
            "rId1",
        )
        .expect_err("wrong relationship namespace must be rejected");
        assert_eq!(error.code, OmErrorCode::InvalidState);

        let mut workbook = base_workbook();
        workbook
            .package
            .add_part(OpcPart {
                name: "xl/drawings/drawing1.xml".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: br#"<xdr:wsDr xmlns:xdr="urn:not-spreadsheet-drawing"/>"#.to_vec(),
            })
            .expect("add drawing part");
        let chart_object = state_only_chart_object(ChartObjectId(1), ChartId(1), SheetId(1));
        let error = append_chart_anchors(
            &mut workbook,
            "xl/drawings/drawing1.xml",
            std::slice::from_ref(&chart_object),
            &BTreeMap::from([(ChartObjectId(1), "rId1".to_string())]),
        )
        .expect_err("wrong drawing namespace must be rejected");
        assert_eq!(error.code, OmErrorCode::InvalidState);
    }

    #[test]
    fn content_type_conflicts_are_rejected_and_stale_part_names_are_reserved() {
        let conflict = append_content_type_override_if_missing(
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/charts/chart1.xml" ContentType="application/wrong"/></Types>"#,
            "xl/charts/chart1.xml",
            CHART_PART_CONTENT_TYPE,
        )
        .expect_err("conflicting content type must be rejected");
        assert_eq!(conflict.code, OmErrorCode::InvalidState);

        let mut workbook = base_workbook();
        let content_types = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/stale"/></Types>"#
            .to_vec();
        workbook
            .package
            .replace_part_bytes("[Content_Types].xml", content_types)
            .expect("replace content types");
        let chart_id = ChartId(1);
        let drawing_id = DrawingId(1);
        workbook
            .state
            .charts
            .insert(chart_id, state_only_chart(chart_id));
        workbook.state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![DrawingObjectModel::ChartFrame(state_only_chart_object(
                    ChartObjectId(1),
                    chart_id,
                    SheetId(1),
                ))],
                raw_part_uri: None,
                dirty: true,
            },
        );
        let materialized = materialize_state_only_chart_graphs(workbook)
            .expect("stale override URI should remain reserved during allocation");
        assert!(!materialized.package.contains("xl/charts/chart1.xml"));
        assert!(materialized.package.contains("xl/charts/chart2.xml"));

        let error = XlsxCodec
            .save(&materialized, SaveOptions::default())
            .expect_err("final save must reject the stale content type Override");
        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("xl/charts/chart1.xml"));
    }

    #[test]
    fn materialized_chart_binding_must_resolve_to_the_chart_part() {
        let mut workbook = base_workbook();
        let chart_id = ChartId(1);
        let drawing_id = DrawingId(1);
        workbook
            .state
            .charts
            .insert(chart_id, state_only_chart(chart_id));
        workbook.state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![DrawingObjectModel::ChartFrame(state_only_chart_object(
                    ChartObjectId(1),
                    chart_id,
                    SheetId(1),
                ))],
                raw_part_uri: None,
                dirty: true,
            },
        );
        workbook = materialize_state_only_chart_graphs(workbook).expect("materialize chart graph");
        let DrawingObjectModel::ChartFrame(chart_object) = &mut workbook
            .state
            .drawings
            .get_mut(&drawing_id)
            .expect("drawing")
            .objects[0]
        else {
            panic!("chart frame");
        };
        chart_object.raw_binding = Some("xl/drawings/other.xml#rId999".to_string());
        let second_chart_id = ChartId(2);
        workbook
            .state
            .charts
            .insert(second_chart_id, state_only_chart(second_chart_id));
        let mut second_chart_object =
            state_only_chart_object(ChartObjectId(2), second_chart_id, SheetId(1));
        second_chart_object.z_order = Some(1);
        let drawing = workbook
            .state
            .drawings
            .get_mut(&drawing_id)
            .expect("drawing");
        drawing
            .objects
            .push(DrawingObjectModel::ChartFrame(second_chart_object));
        drawing.dirty = true;

        let error = XlsxCodec
            .save(&workbook, SaveOptions::default())
            .expect_err("mismatched materialized binding must be rejected");
        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("drawing binding does not match"));
    }

    #[test]
    fn state_only_drawing_allocates_unique_non_visual_ids() {
        let mut workbook = base_workbook();
        let first_chart_id = ChartId(1);
        let second_chart_id = ChartId(2);
        workbook
            .state
            .charts
            .insert(first_chart_id, state_only_chart(first_chart_id));
        workbook
            .state
            .charts
            .insert(second_chart_id, state_only_chart(second_chart_id));
        let first = state_only_chart_object(ChartObjectId(1), first_chart_id, SheetId(1));
        let mut second = state_only_chart_object(ChartObjectId(2), second_chart_id, SheetId(1));
        second.non_visual_id = first.non_visual_id;
        workbook.state.drawings.insert(
            DrawingId(1),
            DrawingModel {
                id: DrawingId(1),
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![
                    DrawingObjectModel::ChartFrame(first),
                    DrawingObjectModel::ChartFrame(second),
                ],
                raw_part_uri: None,
                dirty: true,
            },
        );

        let workbook = materialize_state_only_chart_graphs(workbook)
            .expect("duplicate requested DrawingML ids should be remapped");
        let non_visual_ids = workbook.state.drawings[&DrawingId(1)]
            .objects
            .iter()
            .filter_map(|object| match object {
                DrawingObjectModel::ChartFrame(chart_object) => chart_object.non_visual_id,
                DrawingObjectModel::UnsupportedRaw { .. } => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(non_visual_ids.len(), 2);
        assert_ne!(non_visual_ids[0], non_visual_ids[1]);
        XlsxCodec
            .save(&workbook, SaveOptions::default())
            .expect("remapped DrawingML ids should save");
    }

    #[test]
    fn existing_drawing_append_does_not_touch_nested_foreign_wsdr() {
        let mut workbook = base_workbook();
        workbook
            .package
            .add_part(OpcPart {
                name: "xl/drawings/drawing1.xml".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:ext="urn:extension"><xdr:extLst><ext:wsDr><ext:item/></ext:wsDr></xdr:extLst></xdr:wsDr>"#
                    .to_vec(),
            })
            .expect("add drawing part");
        let chart_object = state_only_chart_object(ChartObjectId(1), ChartId(1), SheetId(1));
        append_chart_anchors(
            &mut workbook,
            "xl/drawings/drawing1.xml",
            std::slice::from_ref(&chart_object),
            &BTreeMap::from([(ChartObjectId(1), "rId1".to_string())]),
        )
        .expect("append chart anchor");

        let drawing_xml = String::from_utf8(
            workbook
                .package
                .part("xl/drawings/drawing1.xml")
                .expect("drawing part")
                .bytes
                .clone(),
        )
        .expect("drawing utf8");
        assert_eq!(drawing_xml.matches("<xdr:absoluteAnchor").count(), 1);
        assert!(drawing_xml.contains("<ext:wsDr><ext:item/></ext:wsDr>"));
    }

    #[test]
    fn public_materializer_is_atomic_when_chart_encoding_fails() {
        let mut workbook = base_workbook();
        let chart_id = ChartId(1);
        let drawing_id = DrawingId(1);
        let mut chart = state_only_chart(chart_id);
        chart.chart_type = ChartType::Unsupported("futureChart".to_string());
        workbook.state.charts.insert(chart_id, chart);
        workbook.state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![DrawingObjectModel::ChartFrame(state_only_chart_object(
                    ChartObjectId(1),
                    chart_id,
                    SheetId(1),
                ))],
                raw_part_uri: None,
                dirty: true,
            },
        );
        let state_before = workbook.state.clone();
        let package_before = workbook.package.clone();

        materialize_state_only_chart_graphs(workbook.clone())
            .expect_err("unsupported chart encoding should fail");
        assert_eq!(workbook.state, state_before);
        assert_eq!(workbook.package, package_before);
    }

    #[test]
    fn public_materializer_rejects_invalid_model_without_state_only_graphs() {
        let mut workbook = base_workbook();
        workbook.state.worksheet_data.clear();

        let error = materialize_state_only_chart_graphs(workbook)
            .expect_err("model validation must run before the no-graph fast path");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("has no worksheet data"));
    }

    #[test]
    fn public_materializer_rejects_invalid_model_before_graph_materialization() {
        let mut workbook = base_workbook();
        let chart_id = ChartId(1);
        workbook
            .state
            .charts
            .insert(chart_id, state_only_chart(chart_id));
        workbook.state.worksheets.clear();

        let error = materialize_state_only_chart_graphs(workbook)
            .expect_err("model validation must precede chart graph validation");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("at least one worksheet"));
    }

    #[test]
    fn public_materializer_rejects_materialized_drawing_map_key_drift() {
        let mut workbook = materialized_embedded_chart_workbook();
        workbook
            .state
            .drawings
            .get_mut(&DrawingId(1))
            .expect("drawing")
            .id = DrawingId(2);

        let error = materialize_state_only_chart_graphs(workbook)
            .expect_err("materialized drawing identity drift must be rejected");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(
            error
                .message
                .contains("drawing map key 1 does not match model id 2")
        );
    }

    #[test]
    fn direct_save_rejects_materialized_chart_map_key_drift() {
        let mut workbook = materialized_embedded_chart_workbook();
        workbook
            .state
            .charts
            .get_mut(&ChartId(1))
            .expect("chart")
            .id = ChartId(2);

        let error = XlsxCodec
            .save(&workbook, SaveOptions::default())
            .expect_err("materialized chart identity drift must be rejected");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(
            error
                .message
                .contains("chart map key 1 does not match model id 2")
        );
    }

    #[test]
    fn direct_save_rejects_materialized_chart_object_ownership_drift() {
        let mut workbook = materialized_embedded_chart_workbook();
        let DrawingObjectModel::ChartFrame(chart_object) = &mut workbook
            .state
            .drawings
            .get_mut(&DrawingId(1))
            .expect("drawing")
            .objects[0]
        else {
            panic!("chart frame");
        };
        chart_object.workbook_id = WorkbookId(8);

        let error = XlsxCodec
            .save(&workbook, SaveOptions::default())
            .expect_err("materialized chart object ownership drift must be rejected");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(
            error
                .message
                .contains("chart object 1 does not match drawing 1 ownership")
        );
    }

    #[test]
    fn public_materializer_rejects_duplicate_materialized_drawing_part_ownership() {
        let mut workbook = materialized_embedded_chart_workbook();
        let mut duplicate = workbook.state.drawings[&DrawingId(1)].clone();
        duplicate.id = DrawingId(2);
        duplicate.objects.clear();
        workbook.state.drawings.insert(duplicate.id, duplicate);

        let error = materialize_state_only_chart_graphs(workbook)
            .expect_err("one package part must not have two drawing model owners");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(
            error
                .message
                .contains("xl/drawings/drawing1.xml is owned by both drawing 1 and drawing 2")
        );
    }

    #[test]
    fn public_materializer_rejects_duplicate_materialized_chart_part_ownership() {
        let mut workbook = materialized_embedded_chart_workbook();
        let mut duplicate_chart = workbook.state.charts[&ChartId(1)].clone();
        duplicate_chart.id = ChartId(2);
        workbook
            .state
            .charts
            .insert(duplicate_chart.id, duplicate_chart);
        let DrawingObjectModel::ChartFrame(first_chart_object) = workbook
            .state
            .drawings
            .get(&DrawingId(1))
            .expect("drawing")
            .objects[0]
            .clone()
        else {
            panic!("chart frame");
        };
        let mut duplicate_chart_object = first_chart_object;
        duplicate_chart_object.id = ChartObjectId(2);
        duplicate_chart_object.non_visual_id = Some(2);
        duplicate_chart_object.chart_id = ChartId(2);
        workbook
            .state
            .drawings
            .get_mut(&DrawingId(1))
            .expect("drawing")
            .objects
            .push(DrawingObjectModel::ChartFrame(duplicate_chart_object));

        let error = materialize_state_only_chart_graphs(workbook)
            .expect_err("one package part must not have two chart model owners");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(
            error
                .message
                .contains("xl/charts/chart1.xml is owned by both chart 1 and chart 2")
        );
    }

    #[test]
    fn public_materializer_rejects_materialized_drawing_host_relationship_drift() {
        let mut workbook = materialized_embedded_chart_workbook();
        let worksheet_relationships_part_uri = "xl/worksheets/_rels/sheet1.xml.rels";
        let relationships_xml = String::from_utf8(
            workbook
                .package
                .part(worksheet_relationships_part_uri)
                .expect("worksheet relationships")
                .bytes
                .clone(),
        )
        .expect("worksheet relationships utf8")
        .replace("../drawings/drawing1.xml", "../drawings/drawing2.xml");
        workbook
            .package
            .replace_part_bytes(
                worksheet_relationships_part_uri,
                relationships_xml.into_bytes(),
            )
            .expect("replace worksheet relationships");

        let error = materialize_state_only_chart_graphs(workbook)
            .expect_err("drawing host relationship drift must be rejected");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(
            error
                .message
                .contains("drawing 1 is not owned by host sheet Sheet1 relationships")
        );
    }

    #[test]
    fn unrelated_state_only_materialization_preserves_shared_materialized_chart() {
        let mut workbook = materialized_embedded_chart_workbook();
        let first_drawing_part_uri = "xl/drawings/drawing1.xml";
        let second_drawing_part_uri = "xl/drawings/drawing2.xml";
        let second_drawing_xml = workbook
            .package
            .part(first_drawing_part_uri)
            .expect("first drawing part")
            .bytes
            .clone();
        workbook
            .package
            .add_part(OpcPart {
                name: second_drawing_part_uri.to_string(),
                content_type: Some(DRAWING_PART_CONTENT_TYPE.to_string()),
                compression: CompressionMethod::Stored,
                bytes: second_drawing_xml.clone(),
            })
            .expect("add second drawing part");
        workbook
            .package
            .add_part(OpcPart {
                name: "xl/drawings/_rels/drawing2.xml.rels".to_string(),
                content_type: Some(RELATIONSHIPS_PART_CONTENT_TYPE.to_string()),
                compression: CompressionMethod::Stored,
                bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>"#
                    .to_vec(),
            })
            .expect("add second drawing relationships");
        let content_types_xml = workbook
            .package
            .part(CONTENT_TYPES_PART_NAME)
            .expect("content types")
            .bytes
            .clone();
        workbook
            .package
            .replace_part_bytes(
                CONTENT_TYPES_PART_NAME,
                append_content_type_override_if_missing(
                    &content_types_xml,
                    second_drawing_part_uri,
                    DRAWING_PART_CONTENT_TYPE,
                )
                .expect("add drawing content type"),
            )
            .expect("replace content types");
        let worksheet_relationships_part_uri = "xl/worksheets/_rels/sheet1.xml.rels";
        let worksheet_relationships_xml = workbook
            .package
            .part(worksheet_relationships_part_uri)
            .expect("worksheet relationships")
            .bytes
            .clone();
        workbook
            .package
            .replace_part_bytes(
                worksheet_relationships_part_uri,
                append_relationship(
                    &worksheet_relationships_xml,
                    "rId2",
                    DRAWING_RELATIONSHIP_TYPE,
                    "../drawings/drawing2.xml",
                )
                .expect("append second drawing relationship"),
            )
            .expect("replace worksheet relationships");

        let DrawingObjectModel::ChartFrame(first_chart_object) = workbook
            .state
            .drawings
            .get(&DrawingId(1))
            .expect("first drawing")
            .objects[0]
            .clone()
        else {
            panic!("chart frame");
        };
        let mut shared_chart_object = first_chart_object;
        shared_chart_object.id = ChartObjectId(2);
        shared_chart_object.raw_binding = Some("xl/drawings/drawing2.xml#rId1".to_string());
        workbook.state.drawings.insert(
            DrawingId(2),
            DrawingModel {
                id: DrawingId(2),
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![DrawingObjectModel::ChartFrame(shared_chart_object)],
                raw_part_uri: Some(second_drawing_part_uri.to_string()),
                dirty: false,
            },
        );

        let state_only_chart_id = ChartId(2);
        workbook
            .state
            .charts
            .insert(state_only_chart_id, state_only_chart(state_only_chart_id));
        let mut state_only_object =
            state_only_chart_object(ChartObjectId(3), state_only_chart_id, SheetId(1));
        state_only_object.z_order = Some(1);
        workbook
            .state
            .drawings
            .get_mut(&DrawingId(1))
            .expect("first drawing")
            .objects
            .push(DrawingObjectModel::ChartFrame(state_only_object));

        let materialized = materialize_state_only_chart_graphs(workbook)
            .expect("unrelated chart materialization must preserve a shared loaded chart");

        assert_eq!(
            materialized.state.charts[&ChartId(1)]
                .raw_part_uri
                .as_deref(),
            Some("xl/charts/chart1.xml")
        );
        assert_eq!(
            materialized.state.charts[&state_only_chart_id]
                .raw_part_uri
                .as_deref(),
            Some("xl/charts/chart2.xml")
        );
        assert_eq!(
            materialized
                .package
                .part(second_drawing_part_uri)
                .expect("preserved second drawing")
                .bytes,
            second_drawing_xml
        );
    }

    #[test]
    fn direct_save_rejects_materialized_chart_without_object_binding() {
        let mut workbook = base_workbook();
        let chart_id = ChartId(1);
        let drawing_id = DrawingId(1);
        workbook
            .state
            .charts
            .insert(chart_id, state_only_chart(chart_id));
        workbook.state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: WorkbookId(7),
                host_sheet_id: SheetId(1),
                objects: vec![DrawingObjectModel::ChartFrame(state_only_chart_object(
                    ChartObjectId(1),
                    chart_id,
                    SheetId(1),
                ))],
                raw_part_uri: None,
                dirty: true,
            },
        );
        workbook = materialize_state_only_chart_graphs(workbook).expect("materialize chart graph");
        let DrawingObjectModel::ChartFrame(chart_object) = &mut workbook
            .state
            .drawings
            .get_mut(&drawing_id)
            .expect("drawing")
            .objects[0]
        else {
            panic!("chart frame");
        };
        chart_object.raw_binding = None;

        let error = XlsxCodec
            .save(&workbook, SaveOptions::default())
            .expect_err("missing materialized object binding must be rejected");
        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("missing its drawing binding"));
    }
}
