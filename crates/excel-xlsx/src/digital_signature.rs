use super::{xml_error, xml_local_name};
use office_common::OmResult;
use office_opc::{OpcPackage, OpcPart};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::BTreeSet;
use std::io::Cursor;

const DIGITAL_SIGNATURE_RELATIONSHIP_PREFIXES: [&str; 2] = [
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/",
    "http://schemas.microsoft.com/package/2006/relationships/digital-signature/",
];
const DIGITAL_SIGNATURE_RELATIONSHIP_SUFFIXES: [&str; 3] = ["origin", "signature", "certificate"];
const DIGITAL_SIGNATURE_CONTENT_TYPES: [&str; 3] = [
    "application/vnd.openxmlformats-package.digital-signature-origin",
    "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
    "application/vnd.openxmlformats-package.digital-signature-certificate",
];
const RELATIONSHIPS_CONTENT_TYPE: &str = "application/vnd.openxmlformats-package.relationships+xml";

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DigitalSignatureInventory {
    part_uris: Vec<String>,
    relationship_part_uris: Vec<String>,
}

impl DigitalSignatureInventory {
    pub fn has_artifacts(&self) -> bool {
        !self.part_uris.is_empty() || !self.relationship_part_uris.is_empty()
    }

    pub fn part_uris(&self) -> &[String] {
        &self.part_uris
    }

    pub fn relationship_part_uris(&self) -> &[String] {
        &self.relationship_part_uris
    }
}

pub(super) fn collect_digital_signature_inventory(
    package: &OpcPackage,
) -> OmResult<DigitalSignatureInventory> {
    let mut part_uris = BTreeSet::new();
    let mut relationship_part_uris = BTreeSet::new();

    for part in package.parts() {
        if is_signature_part(part) {
            part_uris.insert(part.name.clone());
        }
        if is_relationships_part(part)
            && relationships_part_has_signature_type(part.bytes.as_slice())?
        {
            relationship_part_uris.insert(part.name.clone());
        }
    }

    Ok(DigitalSignatureInventory {
        part_uris: part_uris.into_iter().collect(),
        relationship_part_uris: relationship_part_uris.into_iter().collect(),
    })
}

fn is_signature_part(part: &OpcPart) -> bool {
    part.name
        .split('/')
        .next()
        .is_some_and(|segment| segment.eq_ignore_ascii_case("_xmlsignatures"))
        || part
            .content_type
            .as_deref()
            .is_some_and(is_digital_signature_content_type)
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

fn is_digital_signature_content_type(content_type: &str) -> bool {
    let media_type = content_type
        .split_once(';')
        .map_or(content_type, |(media_type, _)| media_type)
        .trim();
    DIGITAL_SIGNATURE_CONTENT_TYPES
        .iter()
        .any(|candidate| media_type.eq_ignore_ascii_case(candidate))
}

fn relationships_part_has_signature_type(xml: &[u8]) -> OmResult<bool> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();

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
                    if is_digital_signature_relationship_type(relationship_type.as_ref()) {
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

fn is_digital_signature_relationship_type(relationship_type: &str) -> bool {
    DIGITAL_SIGNATURE_RELATIONSHIP_PREFIXES
        .iter()
        .any(|prefix| {
            DIGITAL_SIGNATURE_RELATIONSHIP_SUFFIXES
                .iter()
                .any(|suffix| {
                    relationship_type.len() == prefix.len() + suffix.len()
                        && relationship_type
                            .get(..prefix.len())
                            .is_some_and(|value| value.eq_ignore_ascii_case(prefix))
                        && relationship_type
                            .get(prefix.len()..)
                            .is_some_and(|value| value.eq_ignore_ascii_case(suffix))
                })
        })
}

#[cfg(test)]
mod tests {
    use super::is_digital_signature_relationship_type;

    #[test]
    fn signature_relationship_types_accept_standard_and_legacy_package_namespaces_only() {
        for relationship_type in [
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin",
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature",
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate",
            "http://schemas.microsoft.com/package/2006/relationships/digital-signature/origin",
        ] {
            assert!(
                is_digital_signature_relationship_type(relationship_type),
                "{relationship_type}"
            );
        }
        for relationship_type in [
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin/extra",
            "http://example.com/package/2006/relationships/digital-signature/origin",
            "서명",
        ] {
            assert!(
                !is_digital_signature_relationship_type(relationship_type),
                "{relationship_type}"
            );
        }
    }
}
