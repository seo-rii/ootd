use super::{xml_error, xml_local_name};
use office_common::OmResult;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::io::Cursor;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct RelationshipEntry {
    pub(super) id: String,
    pub(super) relationship_type: String,
    pub(super) target: String,
    pub(super) target_mode: Option<String>,
}

pub(super) fn normalize_relationship_target(value: &str, base_segments: &[&str]) -> Option<String> {
    let mut segments = if value.starts_with('/') {
        Vec::new()
    } else {
        base_segments
            .iter()
            .map(|segment| (*segment).to_string())
            .collect()
    };
    for segment in value.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            _ => segments.push(segment.to_string()),
        }
    }
    (!segments.is_empty()).then(|| segments.join("/"))
}

pub(super) fn parse_workbook_relationship_entries(
    rels_xml: &[u8],
) -> OmResult<Vec<RelationshipEntry>> {
    parse_relationship_entries(rels_xml, &["xl"])
}

pub(super) fn parse_relationship_entries(
    rels_xml: &[u8],
    base_segments: &[&str],
) -> OmResult<Vec<RelationshipEntry>> {
    parse_relationship_entries_with_options(rels_xml, base_segments, false)
}

pub(super) fn parse_relationship_entries_with_options(
    rels_xml: &[u8],
    base_segments: &[&str],
    include_external: bool,
) -> OmResult<Vec<RelationshipEntry>> {
    let mut reader = Reader::from_reader(Cursor::new(rels_xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut relationships = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if xml_local_name(element.name().as_ref()) == b"Relationship" =>
            {
                let mut id = None;
                let mut relationship_type = None;
                let mut target = None;
                let mut external = false;
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
                        b"TargetMode" if value.eq_ignore_ascii_case("External") => {
                            target_mode = Some(value);
                            external = true;
                        }
                        b"TargetMode" => target_mode = Some(value),
                        b"Target" => target = Some(value),
                        _ => {}
                    }
                }
                if let (Some(id), Some(relationship_type), Some(target)) =
                    (id, relationship_type, target)
                {
                    if external && !include_external {
                        buffer.clear();
                        continue;
                    }
                    let target = if external {
                        target
                    } else if let Some(target) =
                        normalize_relationship_target(&target, base_segments)
                    {
                        target
                    } else {
                        buffer.clear();
                        continue;
                    };
                    relationships.push(RelationshipEntry {
                        id,
                        relationship_type,
                        target,
                        target_mode,
                    });
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }

    Ok(relationships)
}

pub(super) fn worksheet_relationships_part_uri(worksheet_part_uri: &str) -> Option<String> {
    relationships_part_uri_for_part(worksheet_part_uri)
}

pub(super) fn relationships_part_uri_for_part(part_uri: &str) -> Option<String> {
    let (parent, file_name) = part_uri.rsplit_once('/')?;
    Some(format!("{parent}/_rels/{file_name}.rels"))
}
