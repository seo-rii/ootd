use super::relationships::parse_relationship_entries_with_options;
use super::{xml_error, xml_local_name};
use office_common::{
    ExternalDataInventory, ExternalDataKind, ExternalDataRelationship, OmError, OmErrorCode,
    OmResult,
};
use office_opc::{OpcPackage, OpcPart};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::BTreeSet;
use std::io::Cursor;

const CONTENT_TYPES_PART_NAME: &str = "[Content_Types].xml";
const RELATIONSHIPS_CONTENT_TYPE: &str = "application/vnd.openxmlformats-package.relationships+xml";

pub(super) fn collect_external_data_inventory(
    package: &OpcPackage,
) -> OmResult<ExternalDataInventory> {
    let mut kinds = BTreeSet::new();
    let mut part_uris = BTreeSet::new();
    let mut relationships = Vec::new();
    let mut has_content_type_markers = false;

    if let Some(content_types) = package.part(CONTENT_TYPES_PART_NAME) {
        let manifest_kinds = external_data_kinds_from_content_types(&content_types.bytes)?;
        has_content_type_markers = !manifest_kinds.is_empty();
        kinds.extend(manifest_kinds);
    }

    for part in package.parts() {
        if is_relationships_part(part) {
            let marker_kinds = external_data_kinds_from_relationships(&part.bytes)?;
            let Some(base_segments) = relationship_base_segments(&part.name) else {
                if marker_kinds.is_empty() {
                    continue;
                }
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    format!(
                        "external-data inventory cannot resolve relationship owner for {}",
                        part.name
                    ),
                ));
            };
            let base_segments = base_segments.iter().map(String::as_str).collect::<Vec<_>>();
            for relationship in
                parse_relationship_entries_with_options(&part.bytes, &base_segments, true)?
            {
                let relationship_kinds =
                    external_data_kinds_from_relationship_type(&relationship.relationship_type);
                if relationship_kinds.is_empty() {
                    continue;
                }
                kinds.extend(relationship_kinds.iter().copied());
                relationships.push(ExternalDataRelationship {
                    relationship_part_uri: part.name.clone(),
                    id: relationship.id,
                    relationship_type: relationship.relationship_type,
                    target: relationship.target,
                    target_mode: relationship.target_mode,
                    kinds: relationship_kinds,
                });
            }
            continue;
        }

        let part_kinds = external_data_kinds_for_part(part)?;
        if !part_kinds.is_empty() {
            kinds.extend(part_kinds);
            part_uris.insert(part.name.clone());
        }
    }

    Ok(ExternalDataInventory::new(
        kinds.into_iter().collect(),
        part_uris.into_iter().collect(),
        relationships,
        has_content_type_markers,
    ))
}

fn external_data_kinds_for_part(part: &OpcPart) -> OmResult<BTreeSet<ExternalDataKind>> {
    let mut kinds = BTreeSet::new();
    let name = part.name.trim_start_matches('/').to_ascii_lowercase();

    if name.starts_with("xl/externallinks/") {
        kinds.insert(ExternalDataKind::ExternalLink);
    }
    if name == "xl/connections.xml" {
        kinds.insert(ExternalDataKind::Connection);
    }
    if name.starts_with("xl/querytables/") {
        kinds.insert(ExternalDataKind::QueryTable);
    }
    if name.starts_with("xl/model/") || name.starts_with("xl/customdata/") {
        kinds.insert(ExternalDataKind::DataModel);
    }
    if let Some(content_type) = part.content_type.as_deref()
        && let Some(kind) = external_data_kind_from_content_type(content_type)
    {
        kinds.insert(kind);
    }

    if kinds.contains(&ExternalDataKind::ExternalLink) && part_is_xml(part) {
        kinds.extend(external_link_subtypes(&part.bytes)?);
    }
    Ok(kinds)
}

fn external_link_subtypes(xml: &[u8]) -> OmResult<BTreeSet<ExternalDataKind>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut kinds = BTreeSet::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                match xml_local_name(element.name().as_ref()) {
                    b"externalBook" => {
                        kinds.insert(ExternalDataKind::ExternalWorkbook);
                    }
                    b"ddeLink" => {
                        kinds.insert(ExternalDataKind::DdeLink);
                    }
                    b"oleLink" => {
                        kinds.insert(ExternalDataKind::OleLink);
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    Ok(kinds)
}

fn external_data_kinds_from_content_types(
    content_types_xml: &[u8],
) -> OmResult<BTreeSet<ExternalDataKind>> {
    let mut reader = Reader::from_reader(Cursor::new(content_types_xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut kinds = BTreeSet::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if matches!(
                    xml_local_name(element.name().as_ref()),
                    b"Default" | b"Override"
                ) =>
            {
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(xml_error)?;
                    if attribute.key.as_ref() != b"ContentType" {
                        continue;
                    }
                    let content_type = attribute
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(xml_error)?;
                    if let Some(kind) = external_data_kind_from_content_type(&content_type) {
                        kinds.insert(kind);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    Ok(kinds)
}

fn external_data_kind_from_content_type(content_type: &str) -> Option<ExternalDataKind> {
    let media_type = content_type
        .split_once(';')
        .map_or(content_type, |(media_type, _)| media_type)
        .trim()
        .to_ascii_lowercase();
    match media_type.as_str() {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.externallink+xml"
        | "application/vnd.ms-excel.externallink" => Some(ExternalDataKind::ExternalLink),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"
        | "application/vnd.ms-excel.connections" => Some(ExternalDataKind::Connection),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.querytable+xml"
        | "application/vnd.ms-excel.querytable" => Some(ExternalDataKind::QueryTable),
        "application/vnd.ms-excel.model" | "application/vnd.ms-excel.datamodel" => {
            Some(ExternalDataKind::DataModel)
        }
        _ => None,
    }
}

fn external_data_kinds_from_relationships(
    relationships_xml: &[u8],
) -> OmResult<BTreeSet<ExternalDataKind>> {
    let mut reader = Reader::from_reader(Cursor::new(relationships_xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut kinds = BTreeSet::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if xml_local_name(element.name().as_ref()) == b"Relationship" =>
            {
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(xml_error)?;
                    if attribute.key.as_ref() != b"Type" {
                        continue;
                    }
                    let relationship_type = attribute
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(xml_error)?;
                    kinds.extend(external_data_kinds_from_relationship_type(
                        &relationship_type,
                    ));
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    Ok(kinds)
}

fn external_data_kinds_from_relationship_type(relationship_type: &str) -> Vec<ExternalDataKind> {
    let relationship_type = relationship_type.trim().to_ascii_lowercase();
    let mut kinds = BTreeSet::new();
    if relationship_type.ends_with("/externallink") {
        kinds.insert(ExternalDataKind::ExternalLink);
    }
    if relationship_type.ends_with("/externallinkpath")
        || relationship_type.contains("/xlexternallinkpath/")
        || relationship_type.ends_with("/externallinklongpath")
        || relationship_type.contains("/xlexternallinklongpath/")
    {
        kinds.insert(ExternalDataKind::ExternalLink);
        kinds.insert(ExternalDataKind::ExternalWorkbook);
    }
    if relationship_type.ends_with("/oleobjectlinklongpath") {
        kinds.insert(ExternalDataKind::ExternalLink);
        kinds.insert(ExternalDataKind::OleLink);
    }
    if relationship_type.ends_with("/connections") {
        kinds.insert(ExternalDataKind::Connection);
    }
    if relationship_type.ends_with("/querytable") {
        kinds.insert(ExternalDataKind::QueryTable);
    }
    if relationship_type.ends_with("/model") || relationship_type.ends_with("/modelconnection") {
        kinds.insert(ExternalDataKind::DataModel);
    }
    kinds.into_iter().collect()
}

fn relationship_base_segments(part_uri: &str) -> Option<Vec<String>> {
    if part_uri.eq_ignore_ascii_case("_rels/.rels") {
        return Some(Vec::new());
    }
    let (parent, relationship_file_name) = part_uri.rsplit_once("/_rels/")?;
    let owner_file_name = relationship_file_name.strip_suffix(".rels")?;
    if parent.is_empty() || owner_file_name.is_empty() || owner_file_name.contains('/') {
        return None;
    }
    Some(parent.split('/').map(str::to_owned).collect())
}

fn part_is_xml(part: &OpcPart) -> bool {
    part.content_type.as_deref().is_some_and(|content_type| {
        let media_type = content_type
            .split_once(';')
            .map_or(content_type, |(media_type, _)| media_type)
            .trim();
        media_type.eq_ignore_ascii_case("application/xml")
            || media_type.eq_ignore_ascii_case("text/xml")
            || media_type.to_ascii_lowercase().ends_with("+xml")
    }) || part
        .name
        .rsplit_once('.')
        .is_some_and(|(_, extension)| extension.eq_ignore_ascii_case("xml"))
}

fn is_relationships_part(part: &OpcPart) -> bool {
    part.content_type.as_deref().is_some_and(|content_type| {
        content_type
            .split_once(';')
            .map_or(content_type, |(media_type, _)| media_type)
            .trim()
            .eq_ignore_ascii_case(RELATIONSHIPS_CONTENT_TYPE)
    }) || part.name.eq_ignore_ascii_case("_rels/.rels")
        || part
            .name
            .rsplit_once('/')
            .is_some_and(|(parent, file_name)| {
                parent
                    .rsplit_once('/')
                    .is_some_and(|(_, directory)| directory.eq_ignore_ascii_case("_rels"))
                    && file_name
                        .rsplit_once('.')
                        .is_some_and(|(_, extension)| extension.eq_ignore_ascii_case("rels"))
            })
}

#[cfg(test)]
mod tests {
    use super::{
        ExternalDataKind, OpcPackage, OpcPart, collect_external_data_inventory,
        external_data_kind_from_content_type, external_data_kinds_from_relationship_type,
        relationship_base_segments,
    };
    use office_opc::CompressionMethod;

    #[test]
    fn external_data_classifiers_cover_standard_and_excel_extension_markers() {
        assert_eq!(
            external_data_kind_from_content_type(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"
            ),
            Some(ExternalDataKind::Connection)
        );
        assert_eq!(
            external_data_kinds_from_relationship_type(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"
            ),
            vec![
                ExternalDataKind::ExternalLink,
                ExternalDataKind::ExternalWorkbook,
            ]
        );
        assert_eq!(
            external_data_kinds_from_relationship_type(
                "http://schemas.microsoft.com/office/2019/04/relationships/oleObjectLinkLongPath"
            ),
            vec![ExternalDataKind::ExternalLink, ExternalDataKind::OleLink]
        );
        assert_eq!(
            external_data_kinds_from_relationship_type(
                "http://purl.oclc.org/ooxml/officeDocument/relationships/externalLink"
            ),
            vec![ExternalDataKind::ExternalLink]
        );
        assert!(
            external_data_kinds_from_relationship_type(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            )
            .is_empty()
        );
        assert_eq!(
            relationship_base_segments("xl/externalLinks/_rels/externalLink1.xml.rels"),
            Some(vec!["xl".to_string(), "externalLinks".to_string()])
        );
    }

    #[test]
    fn inventory_keeps_orphan_content_type_and_external_relationship_markers_visible() {
        let package = OpcPackage::try_new(vec![
            OpcPart {
                name: "[Content_Types].xml".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/missing/queryTable.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/></Types>"#.to_vec(),
            },
            OpcPart {
                name: "_rels/.rels".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-package.relationships+xml".to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="https://example.invalid/external.xlsx" TargetMode="External"/></Relationships>"#.to_vec(),
            },
        ])
        .expect("external-data test package should have valid part identities");

        let inventory = collect_external_data_inventory(&package).expect("external-data inventory");
        assert_eq!(
            inventory.kinds(),
            &[ExternalDataKind::ExternalLink, ExternalDataKind::QueryTable,]
        );
        assert!(inventory.part_uris().is_empty());
        assert_eq!(
            inventory.relationship_part_uris(),
            &["_rels/.rels".to_string()]
        );
        assert!(inventory.has_content_type_markers());
        assert!(inventory.has_artifacts());
    }
}
