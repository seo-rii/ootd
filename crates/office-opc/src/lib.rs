use office_common::{OmError, OmErrorCode, OmResult};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};
pub use zip::CompressionMethod;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpcPart {
    pub name: String,
    pub content_type: Option<String>,
    pub compression: CompressionMethod,
    pub bytes: Vec<u8>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct OpcPackage {
    parts: Vec<OpcPart>,
}

impl OpcPackage {
    pub fn new(parts: Vec<OpcPart>) -> Self {
        Self { parts }
    }

    pub fn from_bytes(bytes: &[u8]) -> OmResult<Self> {
        let mut archive = ZipArchive::new(Cursor::new(bytes)).map_err(zip_error)?;
        let mut parts = Vec::with_capacity(archive.len());
        let mut content_types_xml = None;

        for index in 0..archive.len() {
            let mut file = archive.by_index(index).map_err(zip_error)?;
            if file.is_dir() {
                continue;
            }
            let name = normalize_part_name(file.name())?;
            let compression = file.compression();
            let mut part_bytes = Vec::new();
            file.read_to_end(&mut part_bytes).map_err(io_error)?;
            if name == "[Content_Types].xml" {
                content_types_xml = Some(part_bytes.clone());
            }
            parts.push(OpcPart {
                name,
                content_type: None,
                compression,
                bytes: part_bytes,
            });
        }

        if let Some(content_types_xml) = content_types_xml {
            let manifest = parse_content_types(&content_types_xml)?;
            for part in &mut parts {
                part.content_type = manifest.resolve(&part.name);
            }
        }

        Ok(Self { parts })
    }

    pub fn to_bytes(&self) -> OmResult<Vec<u8>> {
        let cursor = Cursor::new(Vec::new());
        let mut writer = ZipWriter::new(cursor);

        for part in &self.parts {
            let options = SimpleFileOptions::default().compression_method(part.compression);
            writer.start_file(&part.name, options).map_err(zip_error)?;
            writer.write_all(&part.bytes).map_err(io_error)?;
        }

        let cursor = writer.finish().map_err(zip_error)?;
        Ok(cursor.into_inner())
    }

    pub fn contains(&self, name: &str) -> bool {
        let normalized = name.trim_start_matches('/');
        self.parts.iter().any(|part| part.name == normalized)
    }

    pub fn part(&self, name: &str) -> Option<&OpcPart> {
        let normalized = name.trim_start_matches('/');
        self.parts.iter().find(|part| part.name == normalized)
    }

    pub fn parts(&self) -> &[OpcPart] {
        &self.parts
    }

    pub fn replace_part_bytes(&mut self, name: &str, bytes: Vec<u8>) -> OmResult<()> {
        let normalized = name.trim_start_matches('/');
        let part = self
            .parts
            .iter_mut()
            .find(|part| part.name == normalized)
            .ok_or_else(|| {
                OmError::new(
                    OmErrorCode::NotFound,
                    format!("OPC part not found: {normalized}"),
                )
            })?;
        part.bytes = bytes;
        Ok(())
    }

    pub fn remove_part(&mut self, name: &str) -> bool {
        let normalized = name.trim_start_matches('/');
        let original_len = self.parts.len();
        self.parts.retain(|part| part.name != normalized);
        self.parts.len() != original_len
    }
}

#[derive(Debug, Default)]
struct ContentTypesManifest {
    defaults: BTreeMap<String, String>,
    overrides: BTreeMap<String, String>,
}

impl ContentTypesManifest {
    fn resolve(&self, part_name: &str) -> Option<String> {
        let override_key = format!("/{}", part_name.trim_start_matches('/'));
        if let Some(content_type) = self.overrides.get(&override_key) {
            return Some(content_type.clone());
        }

        let extension = part_name
            .rsplit_once('.')
            .map(|(_, ext)| ext.to_ascii_lowercase())?;
        self.defaults.get(&extension).cloned()
    }
}

fn normalize_part_name(name: &str) -> OmResult<String> {
    let normalized = name.trim_start_matches('/').to_string();
    if normalized.is_empty() || normalized.ends_with('/') {
        return Err(OmError::new(
            OmErrorCode::Parse,
            format!("invalid OPC part name: {name}"),
        ));
    }
    Ok(normalized)
}

fn parse_content_types(xml: &[u8]) -> OmResult<ContentTypesManifest> {
    let mut manifest = ContentTypesManifest::default();
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                match element.name().as_ref() {
                    b"Default" => {
                        let mut extension = None;
                        let mut content_type = None;
                        for attr in element.attributes() {
                            let attr = attr.map_err(xml_error)?;
                            let value = attr
                                .decode_and_unescape_value(reader.decoder())
                                .map_err(xml_error)?
                                .into_owned();
                            match attr.key.as_ref() {
                                b"Extension" => extension = Some(value.to_ascii_lowercase()),
                                b"ContentType" => content_type = Some(value),
                                _ => {}
                            }
                        }
                        if let (Some(extension), Some(content_type)) = (extension, content_type) {
                            manifest.defaults.insert(extension, content_type);
                        }
                    }
                    b"Override" => {
                        let mut part_name = None;
                        let mut content_type = None;
                        for attr in element.attributes() {
                            let attr = attr.map_err(xml_error)?;
                            let value = attr
                                .decode_and_unescape_value(reader.decoder())
                                .map_err(xml_error)?
                                .into_owned();
                            match attr.key.as_ref() {
                                b"PartName" => part_name = Some(value),
                                b"ContentType" => content_type = Some(value),
                                _ => {}
                            }
                        }
                        if let (Some(part_name), Some(content_type)) = (part_name, content_type) {
                            manifest.overrides.insert(part_name, content_type);
                        }
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

    Ok(manifest)
}

fn io_error(error: std::io::Error) -> OmError {
    OmError::new(OmErrorCode::Io, error.to_string())
}

fn zip_error(error: zip::result::ZipError) -> OmError {
    OmError::new(OmErrorCode::Io, error.to_string())
}

fn xml_error(error: impl std::fmt::Display) -> OmError {
    OmError::new(OmErrorCode::Parse, error.to_string())
}

#[cfg(test)]
mod tests {
    use super::{CompressionMethod, OpcPackage, OpcPart};
    use office_common::OmErrorCode;
    use std::io::{Cursor, Write};
    use zip::ZipWriter;
    use zip::write::SimpleFileOptions;

    #[test]
    fn resolves_content_types_from_manifest() {
        let package = OpcPackage::new(vec![
            OpcPart {
                name: "[Content_Types].xml".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#
                    .to_vec(),
            },
            OpcPart {
                name: "_rels/.rels".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: b"<Relationships/>".to_vec(),
            },
            OpcPart {
                name: "xl/workbook.xml".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: b"<workbook/>".to_vec(),
            },
        ]);

        let bytes = package.to_bytes().expect("package bytes");
        let reparsed = OpcPackage::from_bytes(&bytes).expect("package parse");

        assert_eq!(
            reparsed
                .part("_rels/.rels")
                .and_then(|part| part.content_type.clone()),
            Some("application/vnd.openxmlformats-package.relationships+xml".to_string())
        );
        assert_eq!(
            reparsed
                .part("xl/workbook.xml")
                .and_then(|part| part.content_type.clone()),
            Some(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                    .to_string()
            )
        );
    }

    #[test]
    fn replaces_part_bytes_without_touching_other_parts() {
        let mut package = OpcPackage::new(vec![
            OpcPart {
                name: "[Content_Types].xml".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-package.content-types+xml".to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: b"<Types/>".to_vec(),
            },
            OpcPart {
                name: "xl/workbook.xml".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                        .to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: b"<old/>".to_vec(),
            },
            OpcPart {
                name: "customXml/item1.xml".to_string(),
                content_type: Some("application/xml".to_string()),
                compression: CompressionMethod::Stored,
                bytes: b"<custom/>".to_vec(),
            },
        ]);

        package
            .replace_part_bytes("xl/workbook.xml", b"<new/>".to_vec())
            .expect("replace workbook part");

        assert_eq!(
            package
                .part("xl/workbook.xml")
                .expect("workbook part")
                .bytes
                .as_slice(),
            b"<new/>"
        );
        assert_eq!(
            package
                .part("customXml/item1.xml")
                .expect("custom part")
                .bytes
                .as_slice(),
            b"<custom/>"
        );
    }

    #[test]
    fn replaces_part_bytes_with_leading_slash_lookup() {
        let mut package = OpcPackage::new(vec![OpcPart {
            name: "xl/worksheets/sheet1.xml".to_string(),
            content_type: Some(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
                    .to_string(),
            ),
            compression: CompressionMethod::Stored,
            bytes: b"<old/>".to_vec(),
        }]);

        package
            .replace_part_bytes("/xl/worksheets/sheet1.xml", b"<new/>".to_vec())
            .expect("replace part");

        assert_eq!(
            package
                .part("xl/worksheets/sheet1.xml")
                .expect("sheet part")
                .bytes
                .as_slice(),
            b"<new/>"
        );
    }

    #[test]
    fn removes_part_bytes_with_normalized_lookup() {
        let mut package = OpcPackage::new(vec![
            OpcPart {
                name: "xl/workbook.xml".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                        .to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: b"<workbook/>".to_vec(),
            },
            OpcPart {
                name: "xl/calcChain.xml".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
                        .to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: b"<calcChain/>".to_vec(),
            },
        ]);

        assert!(package.remove_part("/xl/calcChain.xml"));
        assert!(!package.contains("xl/calcChain.xml"));
        assert!(!package.remove_part("xl/calcChain.xml"));
        assert!(package.contains("xl/workbook.xml"));
    }

    #[test]
    fn ignores_directory_entries_when_loading_zip() {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        writer
            .add_directory("xl/", SimpleFileOptions::default())
            .expect("directory");
        writer
            .start_file(
                "[Content_Types].xml",
                SimpleFileOptions::default().compression_method(CompressionMethod::Stored),
            )
            .expect("content types");
        writer
            .write_all(br#"<Types/>"#)
            .expect("write content types");
        writer
            .start_file(
                "xl/workbook.xml",
                SimpleFileOptions::default().compression_method(CompressionMethod::Stored),
            )
            .expect("workbook");
        writer.write_all(br#"<workbook/>"#).expect("write workbook");
        let bytes = writer.finish().expect("finish zip").into_inner();

        let package = OpcPackage::from_bytes(&bytes).expect("package parse");

        assert_eq!(package.parts().len(), 2);
        assert!(package.contains("[Content_Types].xml"));
        assert!(package.contains("xl/workbook.xml"));
    }

    #[test]
    fn removes_part_with_leading_slash_lookup() {
        let mut package = OpcPackage::new(vec![
            OpcPart {
                name: "xl/workbook.xml".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                        .to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: b"<workbook/>".to_vec(),
            },
            OpcPart {
                name: "xl/calcChain.xml".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
                        .to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: b"<calcChain/>".to_vec(),
            },
        ]);

        assert!(package.remove_part("/xl/calcChain.xml"));
        assert!(!package.contains("xl/calcChain.xml"));
        assert!(package.contains("xl/workbook.xml"));
        assert!(!package.remove_part("xl/missing.xml"));
    }

    #[test]
    fn override_content_type_takes_precedence_over_default_extension_mapping() {
        let package = OpcPackage::new(vec![
            OpcPart {
                name: "[Content_Types].xml".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#
                    .to_vec(),
            },
            OpcPart {
                name: "xl/workbook.xml".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: b"<workbook/>".to_vec(),
            },
        ]);

        let reparsed = OpcPackage::from_bytes(&package.to_bytes().expect("package bytes"))
            .expect("package parse");

        assert_eq!(
            reparsed
                .part("xl/workbook.xml")
                .and_then(|part| part.content_type.clone()),
            Some(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                    .to_string()
            )
        );
    }

    #[test]
    fn resolves_default_content_type_case_insensitively_for_extensions() {
        let package = OpcPackage::new(vec![
            OpcPart {
                name: "[Content_Types].xml".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="XML" ContentType="application/xml"/>
</Types>"#
                    .to_vec(),
            },
            OpcPart {
                name: "customXml/item1.XML".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: b"<custom/>".to_vec(),
            },
        ]);

        let reparsed = OpcPackage::from_bytes(&package.to_bytes().expect("package bytes"))
            .expect("package parse");

        assert_eq!(
            reparsed
                .part("customXml/item1.XML")
                .and_then(|part| part.content_type.clone()),
            Some("application/xml".to_string())
        );
    }

    #[test]
    fn preserves_part_order_and_compression_across_round_trip() {
        let package = OpcPackage::new(vec![
            OpcPart {
                name: "docProps/core.xml".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-package.core-properties+xml".to_string(),
                ),
                compression: CompressionMethod::Deflated,
                bytes: b"<core/>".to_vec(),
            },
            OpcPart {
                name: "xl/workbook.xml".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                        .to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: b"<workbook/>".to_vec(),
            },
        ]);

        let reparsed = OpcPackage::from_bytes(&package.to_bytes().expect("package bytes"))
            .expect("package parse");

        assert_eq!(
            reparsed
                .parts()
                .iter()
                .map(|part| part.name.as_str())
                .collect::<Vec<_>>(),
            vec!["docProps/core.xml", "xl/workbook.xml"]
        );
        assert_eq!(reparsed.parts()[0].compression, CompressionMethod::Deflated);
        assert_eq!(reparsed.parts()[1].compression, CompressionMethod::Stored);
    }

    #[test]
    fn replace_part_bytes_returns_not_found_for_missing_part() {
        let mut package = OpcPackage::new(vec![OpcPart {
            name: "xl/workbook.xml".to_string(),
            content_type: Some(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                    .to_string(),
            ),
            compression: CompressionMethod::Stored,
            bytes: b"<workbook/>".to_vec(),
        }]);

        let error = package
            .replace_part_bytes("xl/missing.xml", b"<missing/>".to_vec())
            .expect_err("missing part should fail");

        assert_eq!(error.code, OmErrorCode::NotFound);
        assert!(error.message.contains("xl/missing.xml"));
    }

    #[test]
    fn malformed_content_types_manifest_returns_parse_error() {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        writer
            .start_file(
                "[Content_Types].xml",
                SimpleFileOptions::default().compression_method(CompressionMethod::Stored),
            )
            .expect("content types");
        writer
            .write_all(br#"<Types><Default Extension="xml""#)
            .expect("write malformed content types");
        writer
            .start_file(
                "xl/workbook.xml",
                SimpleFileOptions::default().compression_method(CompressionMethod::Stored),
            )
            .expect("workbook");
        writer.write_all(br#"<workbook/>"#).expect("write workbook");
        let bytes = writer.finish().expect("finish zip").into_inner();

        let error = OpcPackage::from_bytes(&bytes).expect_err("malformed manifest should fail");

        assert_eq!(error.code, OmErrorCode::Parse);
    }
}
