use super::relationships::{
    RelationshipEntry, parse_relationship_entries_with_options, relationships_part_uri_for_part,
};
use super::{xml_error, xml_local_name};
use office_common::{
    ActiveContentAuditManifest, ActiveContentContentTypeEntryKind, ActiveContentKind,
    ActiveContentPolicy, ActiveContentRemovedContentTypeEntry, ActiveContentRemovedPart,
    ActiveContentRemovedRelationship, OmError, OmErrorCode, OmResult,
};
use office_opc::{OpcPackage, OpcPart};
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, Reader, Writer};
use std::collections::{BTreeMap, BTreeSet, VecDeque};
use std::io::Cursor;

const RELATIONSHIPS_CONTENT_TYPE: &str = "application/vnd.openxmlformats-package.relationships+xml";
const CONTENT_TYPES_PART_NAME: &str = "[Content_Types].xml";

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ActiveContentInventory {
    kinds: Vec<ActiveContentKind>,
    part_uris: Vec<String>,
    relationship_part_uris: Vec<String>,
    has_content_type_markers: bool,
}

#[derive(Debug, Clone)]
struct PackageRelationshipEntry {
    relationship_part_uri: String,
    owner_part_uri: Option<String>,
    entry: RelationshipEntry,
}

pub(super) fn strip_active_content_from_package(
    mut package: OpcPackage,
    workbook_part_uri: &str,
) -> OmResult<(OpcPackage, ActiveContentAuditManifest)> {
    let inventory = collect_active_content_inventory(&package)?;
    let mut audit = ActiveContentAuditManifest::observed(
        ActiveContentPolicy::Strip,
        inventory.kinds().to_vec(),
    );
    if !inventory.has_artifacts() {
        return Ok((package, audit));
    }

    let relationships = collect_package_relationship_entries(&package)?;
    let mut roots = package
        .parts()
        .iter()
        .filter(|part| {
            !is_relationships_part(part) && !active_content_kinds_for_part(part).is_empty()
        })
        .map(|part| part.name.clone())
        .collect::<BTreeSet<_>>();
    roots.extend(active_content_part_uris_from_manifest(&package)?);
    for relationship in &relationships {
        if active_content_kind_from_relationship_type(&relationship.entry.relationship_type)
            .is_some()
            && !relationship_is_external(&relationship.entry)
            && let Some(target_part) = package.part(&relationship.entry.target)
        {
            roots.insert(target_part.name.clone());
        }
    }

    for protected_part_uri in [CONTENT_TYPES_PART_NAME, workbook_part_uri] {
        if roots
            .iter()
            .any(|part_uri| part_uri.eq_ignore_ascii_case(protected_part_uri))
        {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!(
                    "active-content strip cannot remove required package part {protected_part_uri}"
                ),
            ));
        }
    }

    let mut outgoing = BTreeMap::<String, BTreeSet<String>>::new();
    let mut incoming = BTreeMap::<String, BTreeSet<Option<String>>>::new();
    for relationship in &relationships {
        if relationship_is_external(&relationship.entry) {
            continue;
        }
        let Some(target_part) = package.part(&relationship.entry.target) else {
            continue;
        };
        let target_part_uri = target_part.name.clone();
        incoming
            .entry(target_part_uri.clone())
            .or_default()
            .insert(relationship.owner_part_uri.clone());
        if let Some(owner_part_uri) = relationship
            .owner_part_uri
            .as_deref()
            .and_then(|owner_part_uri| package.part(owner_part_uri))
            .map(|part| part.name.clone())
        {
            outgoing
                .entry(owner_part_uri)
                .or_default()
                .insert(target_part_uri);
        }
    }

    let mut candidate_parts = roots.clone();
    let mut queue = roots.iter().cloned().collect::<VecDeque<_>>();
    while let Some(owner_part_uri) = queue.pop_front() {
        if let Some(target_part_uris) = outgoing.get(&owner_part_uri) {
            for target_part_uri in target_part_uris {
                if candidate_parts.insert(target_part_uri.clone()) {
                    queue.push_back(target_part_uri.clone());
                }
            }
        }
    }

    let mut removed_part_uris = candidate_parts.clone();
    loop {
        let shared_parts = removed_part_uris
            .iter()
            .filter(|part_uri| !roots.contains(*part_uri))
            .filter(|part_uri| {
                incoming.get(*part_uri).is_some_and(|owners| {
                    owners.iter().any(|owner| {
                        owner
                            .as_ref()
                            .is_none_or(|owner| !removed_part_uris.contains(owner))
                    })
                })
            })
            .cloned()
            .collect::<Vec<_>>();
        if shared_parts.is_empty() {
            break;
        }
        for part_uri in shared_parts {
            removed_part_uris.remove(&part_uri);
        }
    }
    audit.retained_shared_part_uris = candidate_parts
        .difference(&removed_part_uris)
        .cloned()
        .collect();

    let mut removed_relationships = Vec::<PackageRelationshipEntry>::new();
    let mut removed_relationship_ids_by_part = BTreeMap::<String, BTreeSet<String>>::new();
    let mut removed_relationship_ids_by_owner = BTreeMap::<String, BTreeSet<String>>::new();
    let mut relationship_parts_to_remove = removed_part_uris
        .iter()
        .filter_map(|part_uri| relationships_part_uri_for_part(part_uri))
        .filter(|relationship_part_uri| package.contains(relationship_part_uri))
        .collect::<BTreeSet<_>>();
    for relationship in relationships {
        let owner_is_removed = relationship
            .owner_part_uri
            .as_ref()
            .is_some_and(|owner| removed_part_uris.contains(owner));
        let target_is_removed = !relationship_is_external(&relationship.entry)
            && package
                .part(&relationship.entry.target)
                .is_some_and(|part| removed_part_uris.contains(&part.name));
        let relationship_is_active =
            active_content_kind_from_relationship_type(&relationship.entry.relationship_type)
                .is_some();
        if !(owner_is_removed || target_is_removed || relationship_is_active) {
            continue;
        }
        removed_relationship_ids_by_part
            .entry(relationship.relationship_part_uri.clone())
            .or_default()
            .insert(relationship.entry.id.clone());
        if let Some(owner_part_uri) = relationship.owner_part_uri.as_ref()
            && !owner_is_removed
        {
            removed_relationship_ids_by_owner
                .entry(owner_part_uri.clone())
                .or_default()
                .insert(relationship.entry.id.clone());
        }
        if owner_is_removed {
            relationship_parts_to_remove.insert(relationship.relationship_part_uri.clone());
        }
        removed_relationships.push(relationship);
    }

    for (owner_part_uri, relationship_ids) in &removed_relationship_ids_by_owner {
        let Some(owner_part) = package.part(owner_part_uri) else {
            continue;
        };
        if !part_is_xml(owner_part) {
            continue;
        }
        let rewritten = strip_relationship_anchors_from_owner_xml(
            owner_part.bytes.as_slice(),
            relationship_ids,
        )?;
        if rewritten != owner_part.bytes {
            package.replace_part_bytes(owner_part_uri, rewritten)?;
            audit.rewritten_owner_part_uris.push(owner_part_uri.clone());
        }
    }

    for (relationship_part_uri, relationship_ids) in &removed_relationship_ids_by_part {
        if relationship_parts_to_remove.contains(relationship_part_uri) {
            continue;
        }
        let Some(relationship_part) = package.part(relationship_part_uri) else {
            continue;
        };
        package.replace_part_bytes(
            relationship_part_uri,
            strip_relationship_entries_by_id(relationship_part.bytes.as_slice(), relationship_ids)?,
        )?;
    }

    removed_part_uris.extend(relationship_parts_to_remove);
    for part_uri in &removed_part_uris {
        let Some(part) = package.part(part_uri).cloned() else {
            continue;
        };
        audit.removed_parts.push(ActiveContentRemovedPart {
            part_uri: part.name.clone(),
            content_type: part.content_type.clone(),
            byte_len: u64::try_from(part.bytes.len()).unwrap_or(u64::MAX),
        });
        package.remove_part(&part.name);
    }

    let content_types_xml = package
        .part(CONTENT_TYPES_PART_NAME)
        .ok_or_else(|| OmError::parse("workbook package is missing [Content_Types].xml"))?
        .bytes
        .clone();
    let (content_types_xml, mut removed_content_type_entries) =
        strip_active_content_type_entries(content_types_xml.as_slice(), &removed_part_uris)?;
    package.replace_part_bytes(CONTENT_TYPES_PART_NAME, content_types_xml)?;

    audit.removed_relationships = removed_relationships
        .into_iter()
        .map(|relationship| ActiveContentRemovedRelationship {
            relationship_part_uri: relationship.relationship_part_uri,
            owner_part_uri: relationship.owner_part_uri,
            id: relationship.entry.id,
            relationship_type: relationship.entry.relationship_type,
            target: relationship.entry.target,
            target_mode: relationship.entry.target_mode,
        })
        .collect();
    audit.removed_relationships.sort_by(|left, right| {
        (
            &left.relationship_part_uri,
            &left.id,
            &left.relationship_type,
            &left.target,
        )
            .cmp(&(
                &right.relationship_part_uri,
                &right.id,
                &right.relationship_type,
                &right.target,
            ))
    });
    audit.removed_parts.sort_by(|left, right| {
        left.part_uri
            .to_ascii_lowercase()
            .cmp(&right.part_uri.to_ascii_lowercase())
    });
    removed_content_type_entries.sort_by(|left, right| {
        (left.entry_kind, &left.selector, &left.content_type).cmp(&(
            right.entry_kind,
            &right.selector,
            &right.content_type,
        ))
    });
    audit.removed_content_type_entries = removed_content_type_entries;
    audit.rewritten_owner_part_uris.sort();
    audit.rewritten_owner_part_uris.dedup();

    let remaining = collect_active_content_inventory(&package)?;
    if remaining.has_artifacts() {
        return Err(OmError::new(
            OmErrorCode::InvalidState,
            format!(
                "active-content strip left package markers for {:?}",
                remaining.kinds()
            ),
        ));
    }
    let workbook_xml = package.part(workbook_part_uri).ok_or_else(|| {
        OmError::parse(format!(
            "workbook package is missing discovered main part {workbook_part_uri}"
        ))
    })?;
    let (sheet_count, visible_sheet_count) = workbook_sheet_counts(&workbook_xml.bytes)?;
    if sheet_count == 0 {
        return Err(OmError::new(
            OmErrorCode::InvalidState,
            "active-content strip would leave the workbook without a sheet",
        ));
    }
    if visible_sheet_count == 0 {
        return Err(OmError::new(
            OmErrorCode::InvalidState,
            "active-content strip would leave the workbook without a visible sheet",
        ));
    }
    Ok((package, audit))
}

impl ActiveContentInventory {
    pub fn has_artifacts(&self) -> bool {
        !self.kinds.is_empty()
    }

    pub fn kinds(&self) -> &[ActiveContentKind] {
        &self.kinds
    }

    pub fn part_uris(&self) -> &[String] {
        &self.part_uris
    }

    pub fn relationship_part_uris(&self) -> &[String] {
        &self.relationship_part_uris
    }

    pub fn has_content_type_markers(&self) -> bool {
        self.has_content_type_markers
    }
}

pub(super) fn collect_active_content_inventory(
    package: &OpcPackage,
) -> OmResult<ActiveContentInventory> {
    let mut kinds = BTreeSet::new();
    let mut part_uris = BTreeSet::new();
    let mut relationship_part_uris = BTreeSet::new();
    let mut has_content_type_markers = false;

    for part in package.parts() {
        let part_kinds = active_content_kinds_for_part(part);
        if !part_kinds.is_empty() {
            kinds.extend(part_kinds);
            part_uris.insert(part.name.clone());
        }

        if part.name.eq_ignore_ascii_case(CONTENT_TYPES_PART_NAME) {
            let content_type_kinds =
                active_content_kinds_from_content_types(part.bytes.as_slice())?;
            if !content_type_kinds.is_empty() {
                has_content_type_markers = true;
                kinds.extend(content_type_kinds);
            }
        }

        if is_relationships_part(part) {
            let relationship_kinds =
                active_content_kinds_from_relationships(part.bytes.as_slice())?;
            if !relationship_kinds.is_empty() {
                kinds.extend(relationship_kinds);
                relationship_part_uris.insert(part.name.clone());
            }
        }
    }

    Ok(ActiveContentInventory {
        kinds: kinds.into_iter().collect(),
        part_uris: part_uris.into_iter().collect(),
        relationship_part_uris: relationship_part_uris.into_iter().collect(),
        has_content_type_markers,
    })
}

fn collect_package_relationship_entries(
    package: &OpcPackage,
) -> OmResult<Vec<PackageRelationshipEntry>> {
    let mut relationships = Vec::new();
    for part in package
        .parts()
        .iter()
        .filter(|part| is_relationships_part(part))
    {
        let marker_kinds = active_content_kinds_from_relationships(part.bytes.as_slice())?;
        let Some((owner_part_uri, base_segments)) = relationship_owner_and_base(&part.name) else {
            if marker_kinds.is_empty() {
                continue;
            }
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!(
                    "active-content strip cannot resolve relationship owner for {}",
                    part.name
                ),
            ));
        };
        let base_segments = base_segments.iter().map(String::as_str).collect::<Vec<_>>();
        for entry in
            parse_relationship_entries_with_options(part.bytes.as_slice(), &base_segments, true)?
        {
            relationships.push(PackageRelationshipEntry {
                relationship_part_uri: part.name.clone(),
                owner_part_uri: owner_part_uri.clone(),
                entry,
            });
        }
    }
    relationships.sort_by(|left, right| {
        (&left.relationship_part_uri, &left.entry.id)
            .cmp(&(&right.relationship_part_uri, &right.entry.id))
    });
    Ok(relationships)
}

fn active_content_part_uris_from_manifest(package: &OpcPackage) -> OmResult<BTreeSet<String>> {
    let content_types_xml = package
        .part(CONTENT_TYPES_PART_NAME)
        .ok_or_else(|| OmError::parse("workbook package is missing [Content_Types].xml"))?
        .bytes
        .as_slice();
    let mut reader = Reader::from_reader(Cursor::new(content_types_xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut part_uris = BTreeSet::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                let Some(entry) = content_type_entry(&element, reader.decoder())? else {
                    buffer.clear();
                    continue;
                };
                if active_content_kind_from_content_type(&entry.content_type).is_none() {
                    buffer.clear();
                    continue;
                }
                match entry.entry_kind {
                    ActiveContentContentTypeEntryKind::Override => {
                        if let Some(part) = package.part(entry.selector.trim_start_matches('/')) {
                            part_uris.insert(part.name.clone());
                        }
                    }
                    ActiveContentContentTypeEntryKind::Default => {
                        part_uris.extend(package.parts().iter().filter_map(|part| {
                            part.name.rsplit_once('.').and_then(|(_, extension)| {
                                extension
                                    .eq_ignore_ascii_case(&entry.selector)
                                    .then(|| part.name.clone())
                            })
                        }));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    Ok(part_uris)
}

fn relationship_owner_and_base(part_uri: &str) -> Option<(Option<String>, Vec<String>)> {
    if part_uri.eq_ignore_ascii_case("_rels/.rels") {
        return Some((None, Vec::new()));
    }
    let (parent, relationship_file_name) = part_uri.rsplit_once("/_rels/")?;
    let owner_file_name = relationship_file_name.strip_suffix(".rels")?;
    if parent.is_empty() || owner_file_name.is_empty() || owner_file_name.contains('/') {
        return None;
    }
    let owner_part_uri = format!("{parent}/{owner_file_name}");
    let base_segments = parent.split('/').map(str::to_owned).collect();
    Some((Some(owner_part_uri), base_segments))
}

fn relationship_is_external(relationship: &RelationshipEntry) -> bool {
    relationship
        .target_mode
        .as_deref()
        .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
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
    }) || part.name.rsplit_once('.').is_some_and(|(_, extension)| {
        extension.eq_ignore_ascii_case("xml") || extension.eq_ignore_ascii_case("vml")
    })
}

fn strip_relationship_entries_by_id(
    relationships_xml: &[u8],
    relationship_ids: &BTreeSet<String>,
) -> OmResult<Vec<u8>> {
    let mut reader = Reader::from_reader(Cursor::new(relationships_xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    let mut skip_depth = 0usize;

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(_)) if skip_depth > 0 => skip_depth += 1,
            Ok(Event::End(_)) if skip_depth > 0 => skip_depth -= 1,
            Ok(_) if skip_depth > 0 => {}
            Ok(Event::Start(element))
                if xml_local_name(element.name().as_ref()) == b"Relationship"
                    && relationship_element_id(&element, reader.decoder())?
                        .is_some_and(|id| relationship_ids.contains(&id)) =>
            {
                skip_depth = 1;
            }
            Ok(Event::Empty(element))
                if xml_local_name(element.name().as_ref()) == b"Relationship"
                    && relationship_element_id(&element, reader.decoder())?
                        .is_some_and(|id| relationship_ids.contains(&id)) => {}
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned()).map_err(xml_error)?,
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    Ok(writer.into_inner().into_inner())
}

fn relationship_element_id(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> OmResult<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == b"Id" {
            return Ok(Some(
                attribute
                    .decode_and_unescape_value(decoder)
                    .map_err(xml_error)?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn strip_relationship_anchors_from_owner_xml(
    owner_xml: &[u8],
    relationship_ids: &BTreeSet<String>,
) -> OmResult<Vec<u8>> {
    const TRANSITIONAL_RELATIONSHIP_NAMESPACE: &[u8] =
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_RELATIONSHIP_NAMESPACE: &[u8] =
        b"http://purl.oclc.org/ooxml/officeDocument/relationships";

    let mut reader = NsReader::from_reader(Cursor::new(owner_xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut skip_depth = 0usize;
    let mut removed_anchor = false;

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(_)) if skip_depth > 0 => {
                skip_depth += 1;
                document_depth += 1;
            }
            Ok(Event::End(_)) if skip_depth > 0 => {
                skip_depth -= 1;
                document_depth = document_depth.saturating_sub(1);
            }
            Ok(Event::Start(element)) => {
                if element_has_removed_relationship_anchor(
                    &reader,
                    &element,
                    relationship_ids,
                    TRANSITIONAL_RELATIONSHIP_NAMESPACE,
                    STRICT_RELATIONSHIP_NAMESPACE,
                )? {
                    if document_depth == 0 {
                        return Err(OmError::new(
                            OmErrorCode::InvalidState,
                            "active-content strip cannot remove an XML document root relationship anchor",
                        ));
                    }
                    skip_depth = 1;
                    removed_anchor = true;
                } else {
                    writer
                        .write_event(Event::Start(element.into_owned()))
                        .map_err(xml_error)?;
                }
                document_depth += 1;
            }
            Ok(Event::Empty(element)) if skip_depth == 0 => {
                if element_has_removed_relationship_anchor(
                    &reader,
                    &element,
                    relationship_ids,
                    TRANSITIONAL_RELATIONSHIP_NAMESPACE,
                    STRICT_RELATIONSHIP_NAMESPACE,
                )? {
                    if document_depth == 0 {
                        return Err(OmError::new(
                            OmErrorCode::InvalidState,
                            "active-content strip cannot remove an XML document root relationship anchor",
                        ));
                    }
                    removed_anchor = true;
                } else {
                    writer
                        .write_event(Event::Empty(element.into_owned()))
                        .map_err(xml_error)?;
                }
            }
            Ok(Event::End(element)) => {
                document_depth = document_depth.saturating_sub(1);
                writer
                    .write_event(Event::End(element.into_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::Eof) => break,
            Ok(event) if skip_depth == 0 => {
                writer.write_event(event.into_owned()).map_err(xml_error)?
            }
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    let rewritten = writer.into_inner().into_inner();
    if removed_anchor {
        strip_empty_active_content_containers(rewritten.as_slice())
    } else {
        Ok(rewritten)
    }
}

fn strip_empty_active_content_containers(owner_xml: &[u8]) -> OmResult<Vec<u8>> {
    struct PendingContainer {
        depth: usize,
        has_element_child: bool,
        events: Vec<Event<'static>>,
    }

    let mut reader = Reader::from_reader(Cursor::new(owner_xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    let mut pending = None::<PendingContainer>;

    loop {
        let event = match reader.read_event_into(&mut buffer) {
            Ok(event) => event,
            Err(error) => return Err(xml_error(error)),
        };
        if matches!(event, Event::Eof) {
            break;
        }
        if let Some(container) = pending.as_mut() {
            match &event {
                Event::Start(_) => {
                    if container.depth == 1 {
                        container.has_element_child = true;
                    }
                    container.depth += 1;
                }
                Event::Empty(_) if container.depth == 1 => {
                    container.has_element_child = true;
                }
                Event::End(_) => container.depth = container.depth.saturating_sub(1),
                _ => {}
            }
            container.events.push(event.into_owned());
            if container.depth == 0 {
                let container = pending.take().ok_or_else(|| {
                    OmError::new(
                        OmErrorCode::InvalidState,
                        "active-content XML container state was lost during strip",
                    )
                })?;
                if container.has_element_child {
                    for event in container.events {
                        writer.write_event(event).map_err(xml_error)?;
                    }
                }
            }
            buffer.clear();
            continue;
        }

        match &event {
            Event::Start(element)
                if matches!(
                    xml_local_name(element.name().as_ref()),
                    b"controls" | b"oleObjects"
                ) =>
            {
                pending = Some(PendingContainer {
                    depth: 1,
                    has_element_child: false,
                    events: vec![event.into_owned()],
                });
            }
            _ => writer.write_event(event.into_owned()).map_err(xml_error)?,
        }
        buffer.clear();
    }
    if pending.is_some() {
        return Err(OmError::parse(
            "unterminated active-content XML container during strip",
        ));
    }
    Ok(writer.into_inner().into_inner())
}

fn workbook_sheet_counts(workbook_xml: &[u8]) -> OmResult<(usize, usize)> {
    let mut reader = Reader::from_reader(Cursor::new(workbook_xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut sheet_count = 0usize;
    let mut visible_sheet_count = 0usize;

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if xml_local_name(element.name().as_ref()) == b"sheet" =>
            {
                sheet_count += 1;
                let mut hidden = false;
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(xml_error)?;
                    if attribute.key.as_ref() != b"state" {
                        continue;
                    }
                    let state = attribute
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(xml_error)?;
                    hidden = state.eq_ignore_ascii_case("hidden")
                        || state.eq_ignore_ascii_case("veryHidden");
                }
                if !hidden {
                    visible_sheet_count += 1;
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    Ok((sheet_count, visible_sheet_count))
}

fn element_has_removed_relationship_anchor(
    reader: &NsReader<Cursor<&[u8]>>,
    element: &quick_xml::events::BytesStart<'_>,
    relationship_ids: &BTreeSet<String>,
    transitional_namespace: &[u8],
    strict_namespace: &[u8],
) -> OmResult<bool> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if local_name.as_ref() != b"id" {
            continue;
        }
        let raw_name = attribute.key.as_ref();
        let relationship_attribute = match namespace {
            ResolveResult::Bound(namespace) => {
                namespace.as_ref() == transitional_namespace
                    || namespace.as_ref() == strict_namespace
            }
            ResolveResult::Unknown(_) => raw_name.contains(&b':'),
            ResolveResult::Unbound => false,
        };
        if !relationship_attribute {
            continue;
        }
        let value = attribute
            .decode_and_unescape_value(reader.decoder())
            .map_err(xml_error)?;
        if relationship_ids.contains(value.as_ref()) {
            return Ok(true);
        }
    }
    Ok(false)
}

fn strip_active_content_type_entries(
    content_types_xml: &[u8],
    removed_part_uris: &BTreeSet<String>,
) -> OmResult<(Vec<u8>, Vec<ActiveContentRemovedContentTypeEntry>)> {
    let mut reader = Reader::from_reader(Cursor::new(content_types_xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    let mut skip_depth = 0usize;
    let mut removed_entries = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(_)) if skip_depth > 0 => skip_depth += 1,
            Ok(Event::End(_)) if skip_depth > 0 => skip_depth -= 1,
            Ok(_) if skip_depth > 0 => {}
            Ok(Event::Start(element)) => {
                if let Some(entry) = content_type_entry(&element, reader.decoder())?
                    && should_strip_content_type_entry(&entry, removed_part_uris)?
                {
                    removed_entries.push(entry);
                    skip_depth = 1;
                } else {
                    writer
                        .write_event(Event::Start(element.into_owned()))
                        .map_err(xml_error)?;
                }
            }
            Ok(Event::Empty(element)) => {
                if let Some(entry) = content_type_entry(&element, reader.decoder())?
                    && should_strip_content_type_entry(&entry, removed_part_uris)?
                {
                    removed_entries.push(entry);
                } else {
                    writer
                        .write_event(Event::Empty(element.into_owned()))
                        .map_err(xml_error)?;
                }
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned()).map_err(xml_error)?,
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
    Ok((writer.into_inner().into_inner(), removed_entries))
}

fn content_type_entry(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> OmResult<Option<ActiveContentRemovedContentTypeEntry>> {
    let entry_kind = match xml_local_name(element.name().as_ref()) {
        b"Default" => ActiveContentContentTypeEntryKind::Default,
        b"Override" => ActiveContentContentTypeEntryKind::Override,
        _ => return Ok(None),
    };
    let mut selector = None;
    let mut content_type = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        let value = attribute
            .decode_and_unescape_value(decoder)
            .map_err(xml_error)?
            .into_owned();
        match attribute.key.as_ref() {
            b"Extension" if entry_kind == ActiveContentContentTypeEntryKind::Default => {
                selector = Some(value);
            }
            b"PartName" if entry_kind == ActiveContentContentTypeEntryKind::Override => {
                selector = Some(value);
            }
            b"ContentType" => content_type = Some(value),
            _ => {}
        }
    }
    match (selector, content_type) {
        (Some(selector), Some(content_type)) => Ok(Some(ActiveContentRemovedContentTypeEntry {
            entry_kind,
            selector,
            content_type,
        })),
        _ => Ok(None),
    }
}

fn should_strip_content_type_entry(
    entry: &ActiveContentRemovedContentTypeEntry,
    removed_part_uris: &BTreeSet<String>,
) -> OmResult<bool> {
    if active_content_kind_from_content_type(&entry.content_type).is_some() {
        return Ok(true);
    }
    if entry.entry_kind != ActiveContentContentTypeEntryKind::Override {
        return Ok(false);
    }

    let selector_identity =
        OpcPackage::canonical_part_identity(&entry.selector).map_err(|error| {
            OmError::invalid_state(format!(
                "active-content strip found an invalid content type Override PartName {:?}: {}",
                entry.selector, error.message
            ))
        })?;
    for part_uri in removed_part_uris {
        if OpcPackage::canonical_part_identity(part_uri)? == selector_identity {
            return Ok(true);
        }
    }
    Ok(false)
}

fn active_content_kinds_for_part(part: &OpcPart) -> BTreeSet<ActiveContentKind> {
    let mut kinds = BTreeSet::new();
    let name = part.name.trim_start_matches('/').to_ascii_lowercase();

    if name == "xl/vbaproject.bin" {
        kinds.insert(ActiveContentKind::VbaProject);
    }
    if name.rsplit('/').next().is_some_and(|file_name| {
        file_name.starts_with("vbaprojectsignature") && file_name.ends_with(".bin")
    }) {
        kinds.insert(ActiveContentKind::VbaProjectSignature);
    }
    if name == "xl/vbadata.xml" {
        kinds.insert(ActiveContentKind::VbaData);
    }
    if name.starts_with("xl/macrosheets/") {
        kinds.insert(ActiveContentKind::XlmMacroSheet);
    }
    if name.starts_with("xl/dialogsheets/") {
        kinds.insert(ActiveContentKind::DialogSheet);
    }
    if name.starts_with("xl/activex/") || name.starts_with("xl/ctrlprops/") {
        kinds.insert(ActiveContentKind::ActiveXControl);
    }
    if name.starts_with("xl/embeddings/") {
        kinds.insert(ActiveContentKind::EmbeddedObject);
    }
    if name.starts_with("customui/") {
        kinds.insert(ActiveContentKind::CustomUi);
    }

    if let Some(content_type) = part.content_type.as_deref() {
        if let Some(kind) = active_content_kind_from_content_type(content_type) {
            kinds.insert(kind);
        }
    }

    kinds
}

fn active_content_kinds_from_content_types(xml: &[u8]) -> OmResult<BTreeSet<ActiveContentKind>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
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
                    if let Some(kind) = active_content_kind_from_content_type(content_type.as_ref())
                    {
                        kinds.insert(kind);
                    }
                }
            }
            Ok(Event::Eof) => return Ok(kinds),
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
}

fn active_content_kind_from_content_type(content_type: &str) -> Option<ActiveContentKind> {
    let media_type = content_type
        .split_once(';')
        .map_or(content_type, |(media_type, _)| media_type)
        .trim()
        .to_ascii_lowercase();
    match media_type.as_str() {
        "application/vnd.ms-office.vbaproject" => Some(ActiveContentKind::VbaProject),
        "application/vnd.ms-office.vbaprojectsignature"
        | "application/vnd.ms-office.vbaprojectsignatureagile"
        | "application/vnd.ms-office.vbaprojectsignaturev3" => {
            Some(ActiveContentKind::VbaProjectSignature)
        }
        "application/vnd.ms-office.vbadata+xml" => Some(ActiveContentKind::VbaData),
        "application/vnd.ms-excel.macrosheet"
        | "application/vnd.ms-excel.macrosheet+xml"
        | "application/vnd.ms-excel.intlmacrosheet"
        | "application/vnd.ms-excel.intlmacrosheet+xml" => Some(ActiveContentKind::XlmMacroSheet),
        "application/vnd.ms-excel.dialogsheet"
        | "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml" => {
            Some(ActiveContentKind::DialogSheet)
        }
        "application/vnd.ms-office.activex"
        | "application/vnd.ms-office.activex+xml"
        | "application/vnd.ms-excel.controlproperties+xml" => {
            Some(ActiveContentKind::ActiveXControl)
        }
        "application/vnd.openxmlformats-officedocument.oleobject" => {
            Some(ActiveContentKind::EmbeddedObject)
        }
        _ => None,
    }
}

fn is_relationships_part(part: &OpcPart) -> bool {
    part.name
        .rsplit_once('.')
        .is_some_and(|(_, extension)| extension.eq_ignore_ascii_case("rels"))
        || part.content_type.as_deref().is_some_and(|content_type| {
            content_type
                .split_once(';')
                .map_or(content_type, |(media_type, _)| media_type)
                .trim()
                .eq_ignore_ascii_case(RELATIONSHIPS_CONTENT_TYPE)
        })
}

fn active_content_kinds_from_relationships(xml: &[u8]) -> OmResult<BTreeSet<ActiveContentKind>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
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
                    if let Some(kind) =
                        active_content_kind_from_relationship_type(relationship_type.as_ref())
                    {
                        kinds.insert(kind);
                    }
                }
            }
            Ok(Event::Eof) => return Ok(kinds),
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }
}

fn active_content_kind_from_relationship_type(
    relationship_type: &str,
) -> Option<ActiveContentKind> {
    let relationship_type = relationship_type.trim().to_ascii_lowercase();
    match relationship_type.as_str() {
        "http://schemas.microsoft.com/office/2006/relationships/vbaproject" => {
            Some(ActiveContentKind::VbaProject)
        }
        "http://schemas.microsoft.com/office/2006/relationships/vbaprojectsignature"
        | "http://schemas.microsoft.com/office/2006/relationships/vbaprojectsignatureagile"
        | "http://schemas.microsoft.com/office/2006/relationships/vbaprojectsignaturev3" => {
            Some(ActiveContentKind::VbaProjectSignature)
        }
        "http://schemas.microsoft.com/office/2006/relationships/vbadata" => {
            Some(ActiveContentKind::VbaData)
        }
        "http://schemas.microsoft.com/office/2006/relationships/xlmacrosheet"
        | "http://schemas.microsoft.com/office/2006/relationships/xlintlmacrosheet" => {
            Some(ActiveContentKind::XlmMacroSheet)
        }
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/dialogsheet"
        | "http://purl.oclc.org/ooxml/officedocument/relationships/dialogsheet" => {
            Some(ActiveContentKind::DialogSheet)
        }
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/control"
        | "http://purl.oclc.org/ooxml/officedocument/relationships/control"
        | "http://schemas.microsoft.com/office/2006/relationships/activexcontrolbinary"
        | "http://schemas.microsoft.com/office/2006/relationships/ctrlprop" => {
            Some(ActiveContentKind::ActiveXControl)
        }
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/oleobject"
        | "http://schemas.openxmlformats.org/officedocument/2006/relationships/package"
        | "http://purl.oclc.org/ooxml/officedocument/relationships/oleobject"
        | "http://purl.oclc.org/ooxml/officedocument/relationships/package" => {
            Some(ActiveContentKind::EmbeddedObject)
        }
        "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
        | "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
        | "http://schemas.microsoft.com/office/2006/relationships/ui/usercustomization" => {
            Some(ActiveContentKind::CustomUi)
        }
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::{
        ActiveContentContentTypeEntryKind, ActiveContentKind, ActiveContentRemovedContentTypeEntry,
        BTreeSet, CONTENT_TYPES_PART_NAME, Cursor, Event, NsReader, OpcPackage, OpcPart, Reader,
        RelationshipEntry, active_content_kind_from_content_type,
        active_content_kind_from_relationship_type, active_content_part_uris_from_manifest,
        collect_package_relationship_entries, element_has_removed_relationship_anchor, part_is_xml,
        relationship_element_id, relationship_is_external, relationship_owner_and_base,
        should_strip_content_type_entry, strip_active_content_from_package,
        strip_active_content_type_entries, strip_empty_active_content_containers,
        strip_relationship_anchors_from_owner_xml, strip_relationship_entries_by_id,
        workbook_sheet_counts,
    };
    use office_opc::CompressionMethod;

    #[test]
    fn strip_helpers_cover_raw_manifest_relationship_and_owner_xml_boundaries() {
        let package = OpcPackage::try_new(vec![
            OpcPart {
                name: CONTENT_TYPES_PART_NAME.to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/custom/payload.bin" ContentType="application/vnd.ms-office.activeX"/></Types>"#.to_vec(),
            },
            OpcPart {
                name: "_rels/.rels".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-package.relationships+xml".to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.to_vec(),
            },
            OpcPart {
                name: "xl/workbook.xml".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                        .to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec(),
            },
            OpcPart {
                name: "xl/_rels/workbook.xml.rels".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-package.relationships+xml".to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#.to_vec(),
            },
            OpcPart {
                name: "xl/worksheets/sheet1.xml".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
                        .to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#.to_vec(),
            },
            OpcPart {
                name: "custom/payload.bin".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: vec![1, 2, 3],
            },
        ])
        .expect("active-content test package should have valid part identities");

        assert_eq!(
            active_content_part_uris_from_manifest(&package).expect("manifest roots"),
            BTreeSet::from(["custom/payload.bin".to_string()])
        );
        let relationships =
            collect_package_relationship_entries(&package).expect("package relationships");
        assert_eq!(relationships.len(), 2);
        assert_eq!(
            relationship_owner_and_base("xl/activeX/_rels/activeX1.xml.rels"),
            Some((
                Some("xl/activeX/activeX1.xml".to_string()),
                vec!["xl".to_string(), "activeX".to_string()]
            ))
        );

        let external = RelationshipEntry {
            id: "rIdExternal".to_string(),
            relationship_type: "urn:fixture".to_string(),
            target: "https://example.invalid".to_string(),
            target_mode: Some("External".to_string()),
        };
        assert!(relationship_is_external(&external));

        let relationship_xml = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdRemove" Type="urn:remove" Target="remove.bin"/><Relationship Id="rIdKeep" Type="urn:keep" Target="keep.bin"/></Relationships>"#;
        let stripped_relationships = strip_relationship_entries_by_id(
            relationship_xml,
            &BTreeSet::from(["rIdRemove".to_string()]),
        )
        .expect("strip relationship by id");
        let stripped_relationships =
            std::str::from_utf8(&stripped_relationships).expect("relationships utf8");
        assert!(!stripped_relationships.contains("rIdRemove"));
        assert!(stripped_relationships.contains("rIdKeep"));

        let mut relationship_reader = Reader::from_reader(Cursor::new(relationship_xml));
        let mut relationship_buffer = Vec::new();
        loop {
            match relationship_reader
                .read_event_into(&mut relationship_buffer)
                .expect("relationship event")
            {
                Event::Empty(element) => {
                    assert_eq!(
                        relationship_element_id(&element, relationship_reader.decoder())
                            .expect("relationship id"),
                        Some("rIdRemove".to_string())
                    );
                    break;
                }
                Event::Eof => panic!("relationship element was not found"),
                _ => relationship_buffer.clear(),
            }
        }

        let owner_xml = br#"<worksheet xmlns="urn:sheet" xmlns:linked="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><controls><control linked:id="rIdRemove"/></controls><keep linked:id="rIdKeep"/></worksheet>"#;
        let rewritten_owner = strip_relationship_anchors_from_owner_xml(
            owner_xml,
            &BTreeSet::from(["rIdRemove".to_string()]),
        )
        .expect("strip owner anchor");
        let rewritten_owner = std::str::from_utf8(&rewritten_owner).expect("owner utf8");
        assert!(!rewritten_owner.contains("controls"));
        assert!(!rewritten_owner.contains("rIdRemove"));
        assert!(rewritten_owner.contains("rIdKeep"));
        assert_eq!(
            strip_empty_active_content_containers(
                br#"<worksheet><controls>  </controls><sheetData/></worksheet>"#
            )
            .expect("strip empty container"),
            b"<worksheet><sheetData/></worksheet>"
        );

        let mut owner_reader = NsReader::from_reader(Cursor::new(owner_xml.as_slice()));
        let mut owner_buffer = Vec::new();
        loop {
            match owner_reader
                .read_event_into(&mut owner_buffer)
                .expect("owner event")
            {
                Event::Empty(element) if element.name().as_ref() == b"control" => {
                    assert!(
                        element_has_removed_relationship_anchor(
                            &owner_reader,
                            &element,
                            &BTreeSet::from(["rIdRemove".to_string()]),
                            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                            b"http://purl.oclc.org/ooxml/officeDocument/relationships",
                        )
                        .expect("removed relationship anchor")
                    );
                    break;
                }
                Event::Eof => panic!("control element was not found"),
                _ => owner_buffer.clear(),
            }
        }

        assert!(part_is_xml(
            package.part("xl/workbook.xml").expect("workbook")
        ));
        assert_eq!(
            workbook_sheet_counts(&package.part("xl/workbook.xml").expect("workbook").bytes)
                .expect("sheet counts"),
            (1, 1)
        );
        let removed_entry = ActiveContentRemovedContentTypeEntry {
            entry_kind: ActiveContentContentTypeEntryKind::Override,
            selector: "/custom/payload.bin".to_string(),
            content_type: "application/octet-stream".to_string(),
        };
        assert!(
            should_strip_content_type_entry(
                &removed_entry,
                &BTreeSet::from(["custom/payload.bin".to_string()])
            )
            .expect("canonical content type Override comparison")
        );
        let (stripped_content_types, stripped_entries) = strip_active_content_type_entries(
            &package
                .part(CONTENT_TYPES_PART_NAME)
                .expect("content types")
                .bytes,
            &BTreeSet::from(["custom/payload.bin".to_string()]),
        )
        .expect("strip content types");
        assert_eq!(stripped_entries.len(), 1);
        assert!(
            !std::str::from_utf8(&stripped_content_types)
                .expect("content types utf8")
                .contains("payload.bin")
        );

        let (stripped_package, audit) =
            strip_active_content_from_package(package, "xl/workbook.xml")
                .expect("strip raw manifest package");
        assert!(!stripped_package.contains("custom/payload.bin"));
        assert_eq!(audit.removed_parts.len(), 1);
    }

    #[test]
    fn content_type_classifier_covers_binary_and_international_macro_variants() {
        for (content_type, expected) in [
            (
                "application/vnd.ms-office.activeX",
                ActiveContentKind::ActiveXControl,
            ),
            (
                "application/vnd.ms-office.activeX+xml",
                ActiveContentKind::ActiveXControl,
            ),
            (
                "application/vnd.ms-excel.intlmacrosheet",
                ActiveContentKind::XlmMacroSheet,
            ),
            (
                "application/vnd.ms-excel.intlmacrosheet+xml; charset=binary",
                ActiveContentKind::XlmMacroSheet,
            ),
        ] {
            assert_eq!(
                active_content_kind_from_content_type(content_type),
                Some(expected),
                "{content_type}"
            );
        }
    }

    #[test]
    fn relationship_classifier_accepts_known_transitional_strict_and_office_markers_only() {
        for (relationship_type, expected) in [
            (
                "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
                ActiveContentKind::VbaProject,
            ),
            (
                "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignatureAgile",
                ActiveContentKind::VbaProjectSignature,
            ),
            (
                "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
                ActiveContentKind::XlmMacroSheet,
            ),
            (
                "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet",
                ActiveContentKind::XlmMacroSheet,
            ),
            (
                "http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet",
                ActiveContentKind::DialogSheet,
            ),
            (
                "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary",
                ActiveContentKind::ActiveXControl,
            ),
            (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
                ActiveContentKind::EmbeddedObject,
            ),
            (
                "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility",
                ActiveContentKind::CustomUi,
            ),
        ] {
            assert_eq!(
                active_content_kind_from_relationship_type(relationship_type),
                Some(expected),
                "{relationship_type}"
            );
        }

        for relationship_type in [
            "http://example.com/relationships/vbaProject",
            "http://schemas.microsoft.com/office/2006/relationships/vbaProject/extra",
            "activeXControlBinary",
        ] {
            assert_eq!(
                active_content_kind_from_relationship_type(relationship_type),
                None,
                "{relationship_type}"
            );
        }
    }
}
