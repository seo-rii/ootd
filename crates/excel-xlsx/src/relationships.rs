use super::{xml_error, xml_local_name};
use office_common::{OmError, OmResult};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::BTreeSet;
use std::io::Cursor;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct RelationshipEntry {
    pub(super) id: String,
    pub(super) relationship_type: String,
    pub(super) target: String,
    pub(super) target_mode: Option<String>,
}

pub(super) fn normalize_relationship_target(value: &str, base_segments: &[&str]) -> Option<String> {
    if value.is_empty()
        || value
            .chars()
            .any(|ch| ch.is_control() || matches!(ch, '\\' | '?' | '#' | ':'))
    {
        return None;
    }
    let absolute = value.starts_with('/');
    let value = value.strip_prefix('/').unwrap_or(value);
    if value.is_empty() || value.starts_with('/') || value.ends_with('/') {
        return None;
    }
    let mut segments = if absolute {
        Vec::new()
    } else {
        base_segments
            .iter()
            .map(|segment| (*segment).to_string())
            .collect()
    };
    for segment in value.split('/') {
        if segment.is_empty() {
            return None;
        }
        let bytes = segment.as_bytes();
        let mut decoded_segment = Vec::with_capacity(bytes.len());
        let mut index = 0;
        while index < bytes.len() {
            if bytes[index] != b'%' {
                decoded_segment.push(bytes[index]);
                index += 1;
                continue;
            }
            let high = bytes.get(index + 1).copied().and_then(|byte| match byte {
                b'0'..=b'9' => Some(byte - b'0'),
                b'a'..=b'f' => Some(byte - b'a' + 10),
                b'A'..=b'F' => Some(byte - b'A' + 10),
                _ => None,
            })?;
            let low = bytes.get(index + 2).copied().and_then(|byte| match byte {
                b'0'..=b'9' => Some(byte - b'0'),
                b'a'..=b'f' => Some(byte - b'a' + 10),
                b'A'..=b'F' => Some(byte - b'A' + 10),
                _ => None,
            })?;
            let decoded = (high << 4) | low;
            if decoded == 0 || decoded.is_ascii_control() || matches!(decoded, b'/' | b'\\') {
                return None;
            }
            decoded_segment.push(decoded);
            index += 3;
        }
        let decoded_segment = String::from_utf8(decoded_segment).ok()?;
        match decoded_segment.as_str() {
            "." => {}
            ".." => {
                segments.pop()?;
            }
            _ if decoded_segment.ends_with('.') => return None,
            _ => segments.push(segment.to_string()),
        }
    }
    (!segments.is_empty()).then(|| segments.join("/"))
}

#[cfg(test)]
pub(super) fn parse_workbook_relationship_entries(
    rels_xml: &[u8],
) -> OmResult<Vec<RelationshipEntry>> {
    parse_relationship_entries_for_part(rels_xml, "xl/workbook.xml")
}

pub(super) fn parse_relationship_entries_for_part(
    rels_xml: &[u8],
    owner_part_uri: &str,
) -> OmResult<Vec<RelationshipEntry>> {
    let base_segments = relationship_base_segments_for_part(owner_part_uri);
    parse_relationship_entries(rels_xml, &base_segments)
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
    let mut relationship_ids = BTreeSet::new();

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
                        b"TargetMode" if value.eq_ignore_ascii_case("Internal") => {
                            target_mode = Some(value);
                        }
                        b"TargetMode" => {
                            return Err(OmError::parse(format!(
                                "unsupported relationship TargetMode {value:?}"
                            )));
                        }
                        b"Target" => target = Some(value),
                        _ => {}
                    }
                }
                let id = id
                    .filter(|value| !value.is_empty())
                    .ok_or_else(|| OmError::parse("relationship is missing a non-empty Id"))?;
                let relationship_type = relationship_type
                    .filter(|value| !value.is_empty())
                    .ok_or_else(|| OmError::parse("relationship is missing a non-empty Type"))?;
                let target = target
                    .filter(|value| !value.is_empty())
                    .ok_or_else(|| OmError::parse("relationship is missing a non-empty Target"))?;
                if !relationship_ids.insert(id.clone()) {
                    return Err(OmError::parse(format!("duplicate relationship Id {id:?}")));
                }
                let target = if external {
                    target
                } else {
                    normalize_relationship_target(&target, base_segments).ok_or_else(|| {
                        OmError::parse(format!("invalid internal relationship target {target:?}"))
                    })?
                };
                if external && !include_external {
                    buffer.clear();
                    continue;
                }
                relationships.push(RelationshipEntry {
                    id,
                    relationship_type,
                    target,
                    target_mode,
                });
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
    let part_uri = part_uri.strip_prefix('/').unwrap_or(part_uri);
    if part_uri.is_empty() || part_uri.ends_with('/') {
        return None;
    }
    Some(match part_uri.rsplit_once('/') {
        Some((parent, file_name)) => format!("{parent}/_rels/{file_name}.rels"),
        None => format!("_rels/{part_uri}.rels"),
    })
}

pub(super) fn relationship_base_segments_for_part(part_uri: &str) -> Vec<&str> {
    let part_uri = part_uri.strip_prefix('/').unwrap_or(part_uri);
    part_uri
        .rsplit_once('/')
        .map(|(parent, _)| {
            parent
                .split('/')
                .filter(|segment| !segment.is_empty())
                .collect()
        })
        .unwrap_or_default()
}
