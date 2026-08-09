use std::collections::BTreeSet;
use std::io::Cursor;

use excel_model::{TableStructuralOwner, WorksheetStructuralOwners};
use office_common::{OmError, OmResult};
use office_opc::OpcPackage;
use quick_xml::NsReader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::NamespaceResolver;

use super::cells::parse_bounded_a1_rect;
use crate::dialect::{
    OoxmlDialect, OoxmlRelationshipKind, relationship_type_is, validate_known_relationship_dialect,
};
use crate::relationships::{
    parse_relationship_entries_with_options, relationship_base_segments_for_part,
    relationships_part_uri_for_part,
};
use crate::xml::{decode_general_reference, resolved_element_is, unqualified_attribute_is};
use crate::xml_error;

const TABLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

pub(crate) fn resolve_table_structural_owners(
    structural_owners: &mut WorksheetStructuralOwners,
    worksheet_part_uri: &str,
    package: &OpcPackage,
    dialect: OoxmlDialect,
) -> OmResult<()> {
    let relationship_part_uri =
        relationships_part_uri_for_part(worksheet_part_uri).ok_or_else(|| {
            OmError::parse(format!("invalid worksheet part URI: {worksheet_part_uri}"))
        })?;
    let Some(relationship_part) = package.part(&relationship_part_uri) else {
        if structural_owners.table_relationship_ids.is_empty() {
            structural_owners.table_owners.clear();
            return Ok(());
        }
        return Err(OmError::parse(format!(
            "{worksheet_part_uri}: tablePart markers require worksheet relationships part {relationship_part_uri}"
        )));
    };
    let base_segments = relationship_base_segments_for_part(worksheet_part_uri);
    let relationship_entries = parse_relationship_entries_with_options(
        relationship_part.bytes.as_slice(),
        &base_segments,
        true,
    )
    .map_err(|error| {
        OmError::new(
            error.code,
            format!("{relationship_part_uri}: {}", error.message),
        )
    })?;
    validate_known_relationship_dialect(
        dialect,
        relationship_entries
            .iter()
            .map(|relationship| relationship.relationship_type.as_str()),
        &relationship_part_uri,
    )?;

    let declared_relationship_ids = structural_owners
        .table_relationship_ids
        .iter()
        .map(String::as_str)
        .collect::<BTreeSet<_>>();
    if declared_relationship_ids.len() != structural_owners.table_relationship_ids.len() {
        return Err(OmError::parse(format!(
            "{worksheet_part_uri}: duplicate tablePart relationship marker"
        )));
    }
    for relationship in relationship_entries.iter().filter(|relationship| {
        relationship_type_is(
            dialect,
            &relationship.relationship_type,
            OoxmlRelationshipKind::Table,
        )
    }) {
        if !declared_relationship_ids.contains(relationship.id.as_str()) {
            return Err(OmError::parse(format!(
                "{relationship_part_uri}: table relationship {} is not declared by worksheet tableParts",
                relationship.id,
            )));
        }
    }

    let mut canonical_table_targets = BTreeSet::new();
    let mut table_owners = Vec::with_capacity(structural_owners.table_relationship_ids.len());
    for relationship_id in &structural_owners.table_relationship_ids {
        let relationship = relationship_entries
            .iter()
            .find(|relationship| relationship.id == *relationship_id)
            .ok_or_else(|| {
                OmError::parse(format!(
                    "{worksheet_part_uri}: tablePart relationship {relationship_id} is missing from {relationship_part_uri}"
                ))
            })?;
        if !relationship_type_is(
            dialect,
            &relationship.relationship_type,
            OoxmlRelationshipKind::Table,
        ) {
            return Err(OmError::parse(format!(
                "{relationship_part_uri}: worksheet tablePart relationship {relationship_id} has non-table type {}",
                relationship.relationship_type,
            )));
        }
        if relationship
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
        {
            return Err(OmError::parse(format!(
                "{relationship_part_uri}: worksheet tablePart relationship {relationship_id} must target an internal part"
            )));
        }
        let canonical_target =
            OpcPackage::canonical_part_identity(&relationship.target).map_err(|error| {
                OmError::new(
                    error.code,
                    format!(
                        "{relationship_part_uri}: table relationship {relationship_id}: {}",
                        error.message
                    ),
                )
            })?;
        if !canonical_table_targets.insert(canonical_target) {
            return Err(OmError::parse(format!(
                "{relationship_part_uri}: multiple worksheet tablePart relationships target {}",
                relationship.target,
            )));
        }
        let table_part = package.part(&relationship.target).ok_or_else(|| {
            OmError::parse(format!(
                "{relationship_part_uri}: table relationship {relationship_id} targets missing part {}",
                relationship.target,
            ))
        })?;
        if table_part.content_type.as_deref() != Some(TABLE_CONTENT_TYPE) {
            return Err(OmError::parse(format!(
                "{}: worksheet table relationship {relationship_id} has unsupported content type {:?}",
                table_part.name, table_part.content_type,
            )));
        }
        table_owners.push(parse_table_structural_owner(
            table_part.bytes.as_slice(),
            dialect.spreadsheetml_namespace(),
            relationship_id,
            &table_part.name,
        )?);
    }
    structural_owners.table_owners = table_owners;
    Ok(())
}

pub(crate) fn parse_table_structural_owner(
    table_xml: &[u8],
    spreadsheet_namespace: &str,
    relationship_id: &str,
    table_part_uri: &str,
) -> OmResult<TableStructuralOwner> {
    let parse_table_range = |element: &BytesStart<'_>,
                             resolver: &NamespaceResolver,
                             decoder: quick_xml::encoding::Decoder|
     -> OmResult<_> {
        let mut reference = None;
        for attr in element.attributes() {
            let attr = attr.map_err(|error| {
                OmError::parse(format!(
                    "{table_part_uri}: invalid table attribute: {error}"
                ))
            })?;
            if unqualified_attribute_is(resolver, attr.key, b"ref") {
                if reference.is_some() {
                    return Err(OmError::parse(format!(
                        "{table_part_uri}: table has duplicate ref attributes"
                    )));
                }
                reference = Some(
                    attr.decode_and_unescape_value(decoder)
                        .map_err(|error| {
                            OmError::parse(format!(
                                "{table_part_uri}: invalid table ref attribute: {error}"
                            ))
                        })?
                        .into_owned(),
                );
            }
        }
        let reference = reference.ok_or_else(|| {
            OmError::parse(format!(
                "{table_part_uri}: table is missing an A1 range reference"
            ))
        })?;
        parse_bounded_a1_rect(&reference, table_part_uri, "table")
    };

    let mut reader = NsReader::from_reader(Cursor::new(table_xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut element_depth = 0usize;
    let mut root_seen = false;
    let mut table_range = None;
    let mut table_columns_depth = None;
    let mut table_column_depth = None;
    let mut formula_depth = None;
    let mut current_formula = None::<String>;
    let mut formulas = Vec::new();

    loop {
        match reader.read_resolved_event_into(&mut buffer) {
            Ok((namespace, Event::Start(element))) => {
                let is_table = resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"table",
                );
                if element_depth == 0 {
                    if root_seen || !is_table {
                        return Err(OmError::parse(format!(
                            "{table_part_uri}: expected one table root in the active SpreadsheetML namespace"
                        )));
                    }
                    root_seen = true;
                    table_range = Some(parse_table_range(
                        &element,
                        reader.resolver(),
                        reader.decoder(),
                    )?);
                } else if element_depth == 1
                    && resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"tableColumns",
                    )
                {
                    table_columns_depth = Some(element_depth + 1);
                } else if table_columns_depth == Some(element_depth)
                    && resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"tableColumn",
                    )
                {
                    table_column_depth = Some(element_depth + 1);
                } else if table_column_depth == Some(element_depth)
                    && (resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"calculatedColumnFormula",
                    ) || resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"totalsRowFormula",
                    ))
                {
                    formula_depth = Some(element_depth + 1);
                    current_formula = Some(String::new());
                } else if formula_depth == Some(element_depth) {
                    return Err(OmError::parse(format!(
                        "{table_part_uri}: table formula contains nested XML"
                    )));
                }
                element_depth += 1;
            }
            Ok((namespace, Event::Empty(element))) => {
                if element_depth == 0 {
                    if root_seen
                        || !resolved_element_is(
                            &namespace,
                            element.local_name(),
                            spreadsheet_namespace.as_bytes(),
                            b"table",
                        )
                    {
                        return Err(OmError::parse(format!(
                            "{table_part_uri}: expected one table root in the active SpreadsheetML namespace"
                        )));
                    }
                    root_seen = true;
                    table_range = Some(parse_table_range(
                        &element,
                        reader.resolver(),
                        reader.decoder(),
                    )?);
                } else if table_column_depth == Some(element_depth)
                    && (resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"calculatedColumnFormula",
                    ) || resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"totalsRowFormula",
                    ))
                {
                    formulas.push(String::new());
                } else if formula_depth == Some(element_depth) {
                    return Err(OmError::parse(format!(
                        "{table_part_uri}: table formula contains nested XML"
                    )));
                }
            }
            Ok((_, Event::Text(text))) => {
                if formula_depth == Some(element_depth)
                    && let Some(formula) = current_formula.as_mut()
                {
                    formula.push_str(&text.xml_content().map_err(xml_error)?);
                } else if element_depth == 0
                    && !text.xml_content().map_err(xml_error)?.trim().is_empty()
                {
                    return Err(OmError::parse(format!(
                        "{table_part_uri}: text is not allowed outside the table root"
                    )));
                }
            }
            Ok((_, Event::CData(text))) => {
                if formula_depth == Some(element_depth)
                    && let Some(formula) = current_formula.as_mut()
                {
                    formula.push_str(&text.xml_content().map_err(xml_error)?);
                } else if element_depth == 0 {
                    return Err(OmError::parse(format!(
                        "{table_part_uri}: CDATA is not allowed outside the table root"
                    )));
                }
            }
            Ok((_, Event::GeneralRef(reference))) => {
                if formula_depth == Some(element_depth)
                    && let Some(formula) = current_formula.as_mut()
                {
                    formula.push_str(&decode_general_reference(&reference, table_part_uri)?);
                } else if element_depth == 0 {
                    return Err(OmError::parse(format!(
                        "{table_part_uri}: entity reference is not allowed outside the table root"
                    )));
                }
            }
            Ok((namespace, Event::End(element))) => {
                if formula_depth == Some(element_depth)
                    && (resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"calculatedColumnFormula",
                    ) || resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"totalsRowFormula",
                    ))
                {
                    formulas.push(current_formula.take().unwrap_or_default());
                    formula_depth = None;
                }
                if table_column_depth == Some(element_depth)
                    && resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"tableColumn",
                    )
                {
                    table_column_depth = None;
                }
                if table_columns_depth == Some(element_depth)
                    && resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"tableColumns",
                    )
                {
                    table_columns_depth = None;
                }
                element_depth = element_depth.saturating_sub(1);
            }
            Ok((_, Event::Eof)) => {
                if element_depth != 0 {
                    return Err(OmError::parse(format!(
                        "{table_part_uri}: table XML ended with an unclosed element"
                    )));
                }
                break;
            }
            Ok(_) => {}
            Err(error) => {
                return Err(OmError::parse(format!(
                    "{table_part_uri}: invalid table XML: {error}"
                )));
            }
        }
        buffer.clear();
    }

    let range = table_range.ok_or_else(|| {
        OmError::parse(format!(
            "{table_part_uri}: table root is missing an A1 range reference"
        ))
    })?;
    Ok(TableStructuralOwner {
        relationship_id: relationship_id.to_string(),
        part_uri: table_part_uri.to_string(),
        range,
        formulas,
    })
}
