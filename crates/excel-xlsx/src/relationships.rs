use super::{
    xml::{resolved_element_is, unqualified_attribute_is},
    xml_error,
};
use office_common::{OmError, OmResult};
use office_opc::OpcPackage;
use quick_xml::NsReader;
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

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum RelationshipPartOwner {
    PackageRoot,
    Part {
        canonical_owner_part_uri: String,
        base_segments: Vec<String>,
    },
}

pub(super) fn content_type_is_package_relationships(content_type: Option<&str>) -> bool {
    const PACKAGE_RELATIONSHIPS_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-package.relationships+xml";

    content_type.is_some_and(|content_type| {
        content_type
            .split_once(';')
            .map_or(content_type, |(media_type, _)| media_type)
            .trim()
            .eq_ignore_ascii_case(PACKAGE_RELATIONSHIPS_CONTENT_TYPE)
    })
}

pub(super) fn relationship_part_owner(part_uri: &str) -> OmResult<Option<RelationshipPartOwner>> {
    let identity = OpcPackage::canonical_part_identity(part_uri)?;
    if identity == "_rels/.rels" {
        return Ok(Some(RelationshipPartOwner::PackageRoot));
    }

    let identity_segments = identity.split('/').collect::<Vec<_>>();
    let raw_part_uri = part_uri.strip_prefix('/').unwrap_or(part_uri);
    let raw_segments = raw_part_uri.split('/').collect::<Vec<_>>();
    if identity_segments.len() != raw_segments.len() {
        return Err(OmError::invalid_state(format!(
            "OPC canonical identity changed the segment count for {part_uri}"
        )));
    }
    let relationship_file_name = identity_segments
        .last()
        .copied()
        .expect("a canonical OPC part identity has a final path segment");
    let has_relationship_suffix = relationship_file_name.ends_with(".rels");
    let expected_relationship_directory_index = identity_segments.len().checked_sub(2);
    let relationship_directory_indexes = identity_segments
        .iter()
        .enumerate()
        .filter_map(|(index, segment)| (*segment == "_rels").then_some(index))
        .collect::<Vec<_>>();
    if !has_relationship_suffix && relationship_directory_indexes.is_empty() {
        return Ok(None);
    }
    if !has_relationship_suffix
        || expected_relationship_directory_index.is_none()
        || relationship_directory_indexes.as_slice()
            != expected_relationship_directory_index.as_slice()
    {
        let nested_owner_detail = (has_relationship_suffix
            && expected_relationship_directory_index.is_some_and(|expected_index| {
                relationship_directory_indexes.contains(&expected_index)
                    && relationship_directory_indexes.len() > 1
            }))
        .then_some(" and cannot be owned by another relationship part")
        .unwrap_or_default();
        return Err(OmError::invalid_state(format!(
            "package relationship part has invalid OPC placement{nested_owner_detail}: {part_uri}"
        )));
    }

    let owner_file_name = relationship_file_name
        .strip_suffix(".rels")
        .expect("relationship suffix was checked");
    if owner_file_name.is_empty() {
        return Err(OmError::invalid_state(format!(
            "non-root package relationship part has an empty owner name: {part_uri}"
        )));
    }
    let canonical_owner_part_uri = if identity_segments.len() == 2 {
        owner_file_name.to_string()
    } else {
        format!(
            "{}/{}",
            identity_segments[..identity_segments.len() - 2].join("/"),
            owner_file_name
        )
    };
    if canonical_owner_part_uri == "[content_types].xml" {
        return Err(OmError::invalid_state(format!(
            "package relationship part cannot use [Content_Types].xml as an owner: {part_uri}"
        )));
    }
    let owner_segments = canonical_owner_part_uri.split('/').collect::<Vec<_>>();
    let owner_is_relationship_payload = owner_segments
        .last()
        .is_some_and(|file_name| file_name.ends_with(".rels"))
        || (owner_segments.len() >= 2 && owner_segments[owner_segments.len() - 2] == "_rels");
    if owner_is_relationship_payload {
        return Err(OmError::invalid_state(format!(
            "package relationship part cannot be owned by another relationship part: {part_uri}"
        )));
    }

    Ok(Some(RelationshipPartOwner::Part {
        canonical_owner_part_uri,
        base_segments: raw_segments[..raw_segments.len() - 2]
            .iter()
            .map(|segment| (*segment).to_string())
            .collect(),
    }))
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
    const PACKAGE_RELATIONSHIPS_NAMESPACE: &[u8] =
        b"http://schemas.openxmlformats.org/package/2006/relationships";

    let mut reader = NsReader::from_reader(Cursor::new(rels_xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut relationships = Vec::new();
    let mut relationship_ids = BTreeSet::new();
    let mut root_seen = false;
    let mut depth = 0_usize;

    loop {
        match reader.read_resolved_event_into(&mut buffer) {
            Ok((namespace, event @ (Event::Start(_) | Event::Empty(_)))) => {
                let (element, remains_open) = match event {
                    Event::Start(element) => (element, true),
                    Event::Empty(element) => (element, false),
                    _ => unreachable!("matched package relationship element event"),
                };
                let local_name = element.local_name();
                let is_package_relationships_root = resolved_element_is(
                    &namespace,
                    local_name,
                    PACKAGE_RELATIONSHIPS_NAMESPACE,
                    b"Relationships",
                );
                let is_package_relationship = resolved_element_is(
                    &namespace,
                    local_name,
                    PACKAGE_RELATIONSHIPS_NAMESPACE,
                    b"Relationship",
                );

                if depth == 0 {
                    if root_seen {
                        return Err(OmError::parse(
                            "invalid package relationships root: document has more than one root",
                        ));
                    }
                    if !is_package_relationships_root {
                        return Err(OmError::parse(
                            "invalid package relationships root: expected Relationships in the package relationships namespace",
                        ));
                    }
                    root_seen = true;
                } else if local_name.as_ref() == b"Relationship" {
                    if depth != 1 || !is_package_relationship {
                        return Err(OmError::parse(
                            "package Relationship element must be a direct child in the package relationships namespace",
                        ));
                    }

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
                        if unqualified_attribute_is(reader.resolver(), attr.key, b"Id") {
                            id = Some(value);
                        } else if unqualified_attribute_is(reader.resolver(), attr.key, b"Type") {
                            relationship_type = Some(value);
                        } else if unqualified_attribute_is(
                            reader.resolver(),
                            attr.key,
                            b"TargetMode",
                        ) {
                            if value.eq_ignore_ascii_case("External") {
                                target_mode = Some(value);
                                external = true;
                            } else if value.eq_ignore_ascii_case("Internal") {
                                target_mode = Some(value);
                            } else {
                                return Err(OmError::parse(format!(
                                    "unsupported relationship TargetMode {value:?}"
                                )));
                            }
                        } else if unqualified_attribute_is(reader.resolver(), attr.key, b"Target") {
                            target = Some(value);
                        }
                    }

                    let id = id
                        .filter(|value| !value.is_empty())
                        .ok_or_else(|| OmError::parse("relationship is missing a non-empty Id"))?;
                    let relationship_type = relationship_type
                        .filter(|value| !value.is_empty())
                        .ok_or_else(|| {
                            OmError::parse("relationship is missing a non-empty Type")
                        })?;
                    let target = target.filter(|value| !value.is_empty()).ok_or_else(|| {
                        OmError::parse("relationship is missing a non-empty Target")
                    })?;
                    if !relationship_ids.insert(id.clone()) {
                        return Err(OmError::parse(format!("duplicate relationship Id {id:?}")));
                    }
                    let target = if external {
                        target
                    } else {
                        normalize_relationship_target(&target, base_segments).ok_or_else(|| {
                            OmError::parse(format!(
                                "invalid internal relationship target {target:?}"
                            ))
                        })?
                    };
                    if !external || include_external {
                        relationships.push(RelationshipEntry {
                            id,
                            relationship_type,
                            target,
                            target_mode,
                        });
                    }
                }

                if remains_open {
                    depth += 1;
                }
            }
            Ok((_, Event::End(_))) => {
                if depth == 0 {
                    return Err(OmError::parse(
                        "invalid package relationships structure: unmatched closing element",
                    ));
                }
                depth -= 1;
            }
            Ok((_, Event::Text(text))) => {
                let text = text.xml_content().map_err(xml_error)?;
                if depth <= 1 && !text.trim().is_empty() {
                    return Err(OmError::parse(
                        "invalid text in package relationships root structure",
                    ));
                }
            }
            Ok((_, Event::CData(_))) if depth <= 1 => {
                return Err(OmError::parse(
                    "invalid CDATA in package relationships root structure",
                ));
            }
            Ok((_, Event::Eof)) => {
                if !root_seen {
                    return Err(OmError::parse(
                        "invalid package relationships root: Relationships element is missing",
                    ));
                }
                break;
            }
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }

    Ok(relationships)
}

pub(super) fn ensure_package_relationship_closure(package: &OpcPackage) -> OmResult<()> {
    for relationship_part in package.parts() {
        let Some(owner) = relationship_part_owner(&relationship_part.name)? else {
            if content_type_is_package_relationships(relationship_part.content_type.as_deref()) {
                return Err(OmError::invalid_state(format!(
                    "package relationships content type has invalid OPC part placement: {}",
                    relationship_part.name
                )));
            }
            continue;
        };
        let base_segments = match owner {
            RelationshipPartOwner::PackageRoot => Vec::new(),
            RelationshipPartOwner::Part {
                canonical_owner_part_uri,
                base_segments,
            } => {
                let owner_part = package.part(&canonical_owner_part_uri).ok_or_else(|| {
                    OmError::invalid_state(format!(
                        "package relationship part {} has no canonical owner part {}",
                        relationship_part.name, canonical_owner_part_uri
                    ))
                })?;
                if content_type_is_package_relationships(owner_part.content_type.as_deref()) {
                    return Err(OmError::invalid_state(format!(
                        "package relationship part {} cannot be owned by relationship payload {}",
                        relationship_part.name, owner_part.name
                    )));
                }
                base_segments
            }
        };
        let base_segments = base_segments.iter().map(String::as_str).collect::<Vec<_>>();
        let entries = parse_relationship_entries_with_options(
            relationship_part.bytes.as_slice(),
            &base_segments,
            true,
        )?;
        for entry in entries {
            if entry
                .target_mode
                .as_deref()
                .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
            {
                continue;
            }
            if !package.contains(&entry.target) {
                return Err(OmError::invalid_state(format!(
                    "package relationship {} in {} targets missing internal part {}",
                    entry.id, relationship_part.name, entry.target
                )));
            }
        }
    }
    Ok(())
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
