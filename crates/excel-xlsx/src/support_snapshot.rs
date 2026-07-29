use std::collections::{BTreeMap, BTreeSet};

use office_common::{OmError, OmResult, SheetId, SheetKind, WorksheetModel};

use super::relationships::relationships_part_uri_for_part;
use super::{LoadedXlsxWorkbook, SheetDrawingSupportParts, WorksheetSupportParts};

#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum RetainedTargetPolicy {
    AllowSourceDangling,
    RequireDeclared,
}

pub(crate) fn validate_support_snapshot_owners(workbook: &LoadedXlsxWorkbook) -> OmResult<()> {
    let sheets = workbook
        .state
        .worksheets()
        .iter()
        .map(|sheet| (sheet.id, sheet))
        .collect::<BTreeMap<_, _>>();

    for (sheet_id, support) in &workbook.worksheet_support_parts {
        let sheet = sheets.get(sheet_id).copied().ok_or_else(|| {
            OmError::invalid_state(format!(
                "worksheet support snapshot key {} has no worksheet owner",
                sheet_id.0
            ))
        })?;
        validate_worksheet_snapshot_owner(*sheet_id, sheet, support)?;
    }

    for (sheet_id, support) in &workbook.sheet_drawing_support_parts {
        let sheet = sheets.get(sheet_id).copied().ok_or_else(|| {
            OmError::invalid_state(format!(
                "drawing support snapshot key {} has no sheet owner",
                sheet_id.0
            ))
        })?;
        validate_drawing_snapshot_owner(*sheet_id, sheet, support)?;
    }

    Ok(())
}

pub(crate) fn validate_support_snapshot_graph(
    workbook: &LoadedXlsxWorkbook,
    target_policy: RetainedTargetPolicy,
) -> OmResult<()> {
    for (sheet_id, support) in &workbook.sheet_drawing_support_parts {
        validate_sheet_drawing_snapshot_graph(*sheet_id, support, target_policy)?;
    }
    Ok(())
}

pub(crate) fn validate_support_snapshots(
    workbook: &LoadedXlsxWorkbook,
    target_policy: RetainedTargetPolicy,
) -> OmResult<()> {
    validate_support_snapshot_owners(workbook)?;
    validate_support_snapshot_graph(workbook, target_policy)
}

fn validate_worksheet_snapshot_owner(
    sheet_id: SheetId,
    sheet: &WorksheetModel,
    support: &WorksheetSupportParts,
) -> OmResult<()> {
    if sheet.kind != SheetKind::Worksheet {
        return Err(OmError::invalid_state(format!(
            "worksheet support snapshot key {} owns non-worksheet sheet {}",
            sheet_id.0, sheet.name
        )));
    }
    validate_snapshot_host_part(
        "worksheet support",
        sheet_id,
        sheet.part_uri.as_deref(),
        support.worksheet_part_uri.as_deref(),
    )?;
    validate_snapshot_relationship_owner(
        "worksheet support",
        sheet_id,
        support.worksheet_part_uri.as_deref(),
        support.relationships_part_uri.as_deref(),
    )?;
    if support.relationships_part_uri.is_none()
        && (support.relationships_part_source_bytes.is_some()
            || support.relationships_summary.is_some())
    {
        return Err(OmError::invalid_state(format!(
            "worksheet support snapshot for sheet {} has relationship data without an owner part",
            sheet_id.0
        )));
    }
    if support.relationships_part_source_bytes.is_none() && support.relationships_summary.is_some()
    {
        return Err(OmError::invalid_state(format!(
            "worksheet relationship summary has no source snapshot for sheet {}",
            sheet_id.0
        )));
    }
    Ok(())
}

fn validate_drawing_snapshot_owner(
    sheet_id: SheetId,
    sheet: &WorksheetModel,
    support: &SheetDrawingSupportParts,
) -> OmResult<()> {
    validate_snapshot_host_part(
        "drawing support",
        sheet_id,
        sheet.part_uri.as_deref(),
        support.sheet_part_uri.as_deref(),
    )?;
    validate_snapshot_relationship_owner(
        "drawing support",
        sheet_id,
        support.sheet_part_uri.as_deref(),
        support.relationships_part_uri.as_deref(),
    )?;
    if support.relationships_part_uri.is_none() && support.relationships_part_source_bytes.is_some()
    {
        return Err(OmError::invalid_state(format!(
            "drawing support snapshot for sheet {} has drawing graph data without a relationships owner part",
            sheet_id.0
        )));
    }
    Ok(())
}

fn validate_snapshot_host_part(
    description: &str,
    sheet_id: SheetId,
    model_part_uri: Option<&str>,
    snapshot_part_uri: Option<&str>,
) -> OmResult<()> {
    let model_part_uri = model_part_uri.ok_or_else(|| {
        OmError::invalid_state(format!(
            "{description} snapshot key {} has an unbound model sheet",
            sheet_id.0
        ))
    })?;
    let snapshot_part_uri = snapshot_part_uri.ok_or_else(|| {
        OmError::invalid_state(format!(
            "{description} snapshot for sheet {} has no owner part URI",
            sheet_id.0
        ))
    })?;
    if snapshot_part_uri != model_part_uri {
        return Err(OmError::invalid_state(format!(
            "{description} snapshot for sheet {} owns {snapshot_part_uri} instead of {model_part_uri}",
            sheet_id.0
        )));
    }
    Ok(())
}

fn validate_snapshot_relationship_owner(
    description: &str,
    sheet_id: SheetId,
    owner_part_uri: Option<&str>,
    relationships_part_uri: Option<&str>,
) -> OmResult<()> {
    let Some(relationships_part_uri) = relationships_part_uri else {
        return Ok(());
    };
    let owner_part_uri = owner_part_uri.ok_or_else(|| {
        OmError::invalid_state(format!(
            "{description} snapshot for sheet {} has a relationships part without an owner part",
            sheet_id.0
        ))
    })?;
    let expected = relationships_part_uri_for_part(owner_part_uri).ok_or_else(|| {
        OmError::invalid_state(format!(
            "{description} snapshot relationship owner cannot be derived for sheet {} part {owner_part_uri}",
            sheet_id.0
        ))
    })?;
    if relationships_part_uri != expected {
        return Err(OmError::invalid_state(format!(
            "{description} snapshot for sheet {} owns relationships part {relationships_part_uri} instead of {expected}",
            sheet_id.0
        )));
    }
    Ok(())
}

fn validate_sheet_drawing_snapshot_graph(
    sheet_id: SheetId,
    support: &SheetDrawingSupportParts,
    target_policy: RetainedTargetPolicy,
) -> OmResult<()> {
    let binding_ids = support
        .drawing_relationships
        .iter()
        .map(|binding| binding.relationship_id.clone())
        .collect::<Vec<_>>();
    if binding_ids != support.drawing_relationship_ids {
        return Err(OmError::invalid_state(format!(
            "drawing relationship ids and bindings differ for sheet {}",
            sheet_id.0
        )));
    }
    let binding_targets = support
        .drawing_relationships
        .iter()
        .map(|binding| binding.target.clone())
        .collect::<Vec<_>>();
    if binding_targets != support.drawing_part_uris {
        return Err(OmError::invalid_state(format!(
            "drawing relationship targets and part inventory differ for sheet {}",
            sheet_id.0
        )));
    }
    if support.relationships_part_uri.is_none()
        && (!support.drawing_relationship_ids.is_empty()
            || !support.drawing_relationships.is_empty()
            || !support.drawing_part_uris.is_empty())
    {
        return Err(OmError::invalid_state(format!(
            "drawing support snapshot for sheet {} has drawing graph data without a relationships owner part",
            sheet_id.0
        )));
    }

    let drawing_opaque_parts = validate_drawing_snapshot_inventory(support, sheet_id)?;
    let (chart_parts, chart_relationship_parts, chart_support_parts, chart_opaque_parts) =
        validate_chart_snapshot_inventory(support, sheet_id)?;

    if target_policy == RetainedTargetPolicy::RequireDeclared {
        validate_drawing_summary_targets(support, &chart_parts, &drawing_opaque_parts, sheet_id)?;
        validate_chart_summary_targets(
            support,
            &chart_relationship_parts,
            &chart_support_parts,
            &chart_opaque_parts,
            sheet_id,
        )?;
    }
    Ok(())
}

fn validate_drawing_snapshot_inventory(
    support: &SheetDrawingSupportParts,
    sheet_id: SheetId,
) -> OmResult<BTreeSet<String>> {
    let drawing_parts = declared_part_uris(
        &support.drawing_part_uris,
        "drawing part inventory",
        sheet_id,
    )?;
    let drawing_relationship_parts = declared_part_uris(
        &support.drawing_relationships_part_uris,
        "drawing relationships part inventory",
        sheet_id,
    )?;
    validate_relationship_part_owners(
        &drawing_relationship_parts,
        &drawing_parts,
        "drawing relationships part",
        sheet_id,
    )?;
    ensure_subset_map_keys(
        &support.drawing_part_source_bytes,
        &drawing_parts,
        "drawing part source snapshot",
        sheet_id,
    )?;
    let source_keys = support
        .drawing_part_source_bytes
        .keys()
        .cloned()
        .collect::<BTreeSet<_>>();
    let summary_keys = support
        .drawing_summaries
        .keys()
        .cloned()
        .collect::<BTreeSet<_>>();
    if source_keys != summary_keys {
        return Err(OmError::invalid_state(format!(
            "drawing part source and summary snapshot keys differ for sheet {}",
            sheet_id.0
        )));
    }
    ensure_subset_map_keys(
        &support.drawing_relationships_part_source_bytes,
        &drawing_relationship_parts,
        "drawing relationships source snapshot",
        sheet_id,
    )?;
    let drawing_opaque_parts = declared_part_uris(
        &support.drawing_opaque_relationship_part_uris,
        "drawing opaque part inventory",
        sheet_id,
    )?;
    ensure_exact_map_keys(
        &support.drawing_opaque_relationship_part_source_bytes,
        &drawing_opaque_parts,
        "drawing opaque source snapshot",
        sheet_id,
    )?;
    Ok(drawing_opaque_parts)
}

fn validate_chart_snapshot_inventory(
    support: &SheetDrawingSupportParts,
    sheet_id: SheetId,
) -> OmResult<(
    BTreeSet<String>,
    BTreeSet<String>,
    BTreeSet<String>,
    BTreeSet<String>,
)> {
    let chart_parts =
        declared_part_uris(&support.chart_part_uris, "chart part inventory", sheet_id)?;
    let chart_relationship_parts = declared_part_uris(
        &support.chart_relationships_part_uris,
        "chart relationships part inventory",
        sheet_id,
    )?;
    validate_relationship_part_owners(
        &chart_relationship_parts,
        &chart_parts,
        "chart relationships part",
        sheet_id,
    )?;
    ensure_exact_map_keys(
        &support.chart_part_source_bytes,
        &chart_parts,
        "chart part source snapshot",
        sheet_id,
    )?;
    ensure_exact_map_keys(
        &support.chart_summaries,
        &chart_parts,
        "chart summary snapshot",
        sheet_id,
    )?;
    ensure_exact_map_keys(
        &support.chart_relationships_part_source_bytes,
        &chart_relationship_parts,
        "chart relationships source snapshot",
        sheet_id,
    )?;
    let chart_support_parts = declared_part_uris(
        &support.chart_support_part_uris,
        "chart support part inventory",
        sheet_id,
    )?;
    ensure_exact_map_keys(
        &support.chart_support_part_source_bytes,
        &chart_support_parts,
        "chart support source snapshot",
        sheet_id,
    )?;
    let chart_opaque_parts = declared_part_uris(
        &support.chart_opaque_relationship_part_uris,
        "chart opaque part inventory",
        sheet_id,
    )?;
    ensure_exact_map_keys(
        &support.chart_opaque_relationship_part_source_bytes,
        &chart_opaque_parts,
        "chart opaque source snapshot",
        sheet_id,
    )?;
    Ok((
        chart_parts,
        chart_relationship_parts,
        chart_support_parts,
        chart_opaque_parts,
    ))
}

fn validate_drawing_summary_targets(
    support: &SheetDrawingSupportParts,
    chart_parts: &BTreeSet<String>,
    opaque_parts: &BTreeSet<String>,
    sheet_id: SheetId,
) -> OmResult<()> {
    for summary in support.drawing_summaries.values() {
        for relationship in &summary.chart_relationships {
            if !chart_parts.contains(&relationship.target) {
                return Err(OmError::invalid_state(format!(
                    "drawing summary chart target {} is not declared by sheet {}",
                    relationship.target, sheet_id.0
                )));
            }
        }
        for relationship in &summary.opaque_relationships {
            if is_external(relationship.target_mode.as_deref()) {
                continue;
            }
            if !opaque_parts.contains(&relationship.target) {
                return Err(OmError::invalid_state(format!(
                    "drawing summary opaque target {} is not declared by sheet {}",
                    relationship.target, sheet_id.0
                )));
            }
        }
    }
    Ok(())
}

fn validate_chart_summary_targets(
    support: &SheetDrawingSupportParts,
    relationship_parts: &BTreeSet<String>,
    support_parts: &BTreeSet<String>,
    opaque_parts: &BTreeSet<String>,
    sheet_id: SheetId,
) -> OmResult<()> {
    let mut summary_relationship_parts = BTreeSet::new();
    for (chart_part_uri, summary) in &support.chart_summaries {
        if let Some(relationships_part_uri) = summary.relationships_part_uri.as_deref() {
            let expected = relationships_part_uri_for_part(chart_part_uri).ok_or_else(|| {
                OmError::invalid_state(format!(
                    "chart summary relationship owner cannot be derived for sheet {} part {chart_part_uri}",
                    sheet_id.0
                ))
            })?;
            if relationships_part_uri != expected
                || !relationship_parts.contains(relationships_part_uri)
            {
                return Err(OmError::invalid_state(format!(
                    "chart summary relationships part {relationships_part_uri} is not owned by {chart_part_uri} for sheet {}",
                    sheet_id.0
                )));
            }
            summary_relationship_parts.insert(relationships_part_uri.to_string());
        }
        for relationship in &summary.support_relationships {
            if is_external(relationship.target_mode.as_deref()) {
                continue;
            }
            if !support_parts.contains(&relationship.target) {
                return Err(OmError::invalid_state(format!(
                    "chart summary support target {} is not declared by sheet {}",
                    relationship.target, sheet_id.0
                )));
            }
        }
        for relationship in &summary.opaque_relationships {
            if is_external(relationship.target_mode.as_deref()) {
                continue;
            }
            if !opaque_parts.contains(&relationship.target) {
                return Err(OmError::invalid_state(format!(
                    "chart summary opaque target {} is not declared by sheet {}",
                    relationship.target, sheet_id.0
                )));
            }
        }
    }
    if summary_relationship_parts != *relationship_parts {
        return Err(OmError::invalid_state(format!(
            "chart summary relationships parts and relationship inventory differ for sheet {}",
            sheet_id.0
        )));
    }
    Ok(())
}

fn declared_part_uris(
    part_uris: &[String],
    description: &str,
    sheet_id: SheetId,
) -> OmResult<BTreeSet<String>> {
    let mut raw = BTreeSet::new();
    let mut canonical = BTreeSet::new();
    for part_uri in part_uris {
        if !raw.insert(part_uri.clone()) {
            return Err(OmError::invalid_state(format!(
                "{description} repeats {part_uri} for sheet {}",
                sheet_id.0
            )));
        }
        let identity =
            office_opc::OpcPackage::canonical_part_identity(part_uri).map_err(|error| {
                OmError::invalid_state(format!(
                    "{description} contains invalid part URI {part_uri} for sheet {}: {error}",
                    sheet_id.0
                ))
            })?;
        if !canonical.insert(identity) {
            return Err(OmError::invalid_state(format!(
                "{description} contains a canonical duplicate at {part_uri} for sheet {}",
                sheet_id.0
            )));
        }
    }
    Ok(raw)
}

fn ensure_exact_map_keys<T>(
    map: &BTreeMap<String, T>,
    declared: &BTreeSet<String>,
    description: &str,
    sheet_id: SheetId,
) -> OmResult<()> {
    ensure_subset_map_keys(map, declared, description, sheet_id)?;
    let keys = map.keys().cloned().collect::<BTreeSet<_>>();
    if keys != *declared {
        return Err(OmError::invalid_state(format!(
            "{description} keys and declared inventory differ for sheet {}",
            sheet_id.0
        )));
    }
    Ok(())
}

fn ensure_subset_map_keys<T>(
    map: &BTreeMap<String, T>,
    declared: &BTreeSet<String>,
    description: &str,
    sheet_id: SheetId,
) -> OmResult<()> {
    for part_uri in map.keys() {
        if !declared.contains(part_uri) {
            return Err(OmError::invalid_state(format!(
                "{description} {part_uri} is not declared by sheet {}",
                sheet_id.0
            )));
        }
    }
    Ok(())
}

fn validate_relationship_part_owners(
    relationship_parts: &BTreeSet<String>,
    owner_parts: &BTreeSet<String>,
    description: &str,
    sheet_id: SheetId,
) -> OmResult<()> {
    let expected = owner_parts
        .iter()
        .filter_map(|part_uri| relationships_part_uri_for_part(part_uri))
        .collect::<BTreeSet<_>>();
    for relationships_part_uri in relationship_parts {
        if !expected.contains(relationships_part_uri) {
            return Err(OmError::invalid_state(format!(
                "{description} {relationships_part_uri} has no declared owner for sheet {}",
                sheet_id.0
            )));
        }
    }
    Ok(())
}

fn is_external(target_mode: Option<&str>) -> bool {
    target_mode.is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
}
