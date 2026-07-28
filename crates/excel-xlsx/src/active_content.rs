use super::{xml_error, xml_local_name};
use office_common::OmResult;
use office_opc::{OpcPackage, OpcPart};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::BTreeSet;
use std::io::Cursor;

const RELATIONSHIPS_CONTENT_TYPE: &str = "application/vnd.openxmlformats-package.relationships+xml";
const CONTENT_TYPES_PART_NAME: &str = "[Content_Types].xml";

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum ActiveContentKind {
    VbaProject,
    VbaProjectSignature,
    VbaData,
    XlmMacroSheet,
    DialogSheet,
    ActiveXControl,
    EmbeddedObject,
    CustomUi,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ActiveContentInventory {
    kinds: Vec<ActiveContentKind>,
    part_uris: Vec<String>,
    relationship_part_uris: Vec<String>,
    has_content_type_markers: bool,
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
        ActiveContentKind, active_content_kind_from_content_type,
        active_content_kind_from_relationship_type,
    };

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
