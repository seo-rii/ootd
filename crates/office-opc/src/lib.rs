use office_common::{OmError, OmErrorCode, OmResult};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Read, Write};
pub use zip::CompressionMethod;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LoadLimits {
    pub max_archive_bytes: u64,
    pub max_entries: usize,
    pub max_part_name_bytes: usize,
    pub max_entry_uncompressed_bytes: u64,
    pub max_total_uncompressed_bytes: u64,
    pub max_compression_ratio: u64,
}

impl Default for LoadLimits {
    fn default() -> Self {
        Self {
            max_archive_bytes: 128 * 1024 * 1024,
            max_entries: 10_000,
            max_part_name_bytes: 1_024,
            max_entry_uncompressed_bytes: 64 * 1024 * 1024,
            max_total_uncompressed_bytes: 256 * 1024 * 1024,
            max_compression_ratio: 1_000,
        }
    }
}

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
        Self::from_bytes_with_limits(bytes, LoadLimits::default())
    }

    pub fn from_bytes_with_limits(bytes: &[u8], limits: LoadLimits) -> OmResult<Self> {
        let archive_bytes = u64::try_from(bytes.len()).unwrap_or(u64::MAX);
        ensure_within_limit("archive bytes", archive_bytes, limits.max_archive_bytes)?;
        let declared_entry_count = preflight_zip_entry_count(bytes)?;
        ensure_within_limit(
            "entry count",
            declared_entry_count,
            u64::try_from(limits.max_entries).unwrap_or(u64::MAX),
        )?;

        let mut archive = ZipArchive::new(Cursor::new(bytes)).map_err(zip_error)?;
        let indexed_entry_count = u64::try_from(archive.len()).unwrap_or(u64::MAX);
        if indexed_entry_count != declared_entry_count {
            return Err(OmError::parse(format!(
                "OPC ZIP contains a duplicate ZIP entry name or inconsistent central directory: declared {declared_entry_count}, indexed {indexed_entry_count}"
            )));
        }
        ensure_within_limit(
            "entry count",
            indexed_entry_count,
            u64::try_from(limits.max_entries).unwrap_or(u64::MAX),
        )?;

        let mut declared_total = 0_u64;
        let mut part_identities = BTreeSet::new();
        for index in 0..archive.len() {
            let file = archive.by_index(index).map_err(zip_error)?;
            ensure_within_limit(
                "part name bytes",
                u64::try_from(file.name().len()).unwrap_or(u64::MAX),
                u64::try_from(limits.max_part_name_bytes).unwrap_or(u64::MAX),
            )?;
            if file.is_dir() {
                continue;
            }
            let canonical_name = canonical_part_name(file.name(), OmErrorCode::Parse)?;
            if !part_identities.insert(canonical_name.identity.clone()) {
                return Err(OmError::parse(format!(
                    "duplicate OPC part identity: {}",
                    canonical_name.spelling
                )));
            }

            let uncompressed_bytes = file.size();
            ensure_within_limit(
                "entry uncompressed bytes",
                uncompressed_bytes,
                limits.max_entry_uncompressed_bytes,
            )?;
            declared_total = declared_total
                .checked_add(uncompressed_bytes)
                .ok_or_else(|| {
                    resource_limit_error(
                        "total uncompressed bytes",
                        u64::MAX,
                        limits.max_total_uncompressed_bytes,
                    )
                })?;
            ensure_within_limit(
                "total uncompressed bytes",
                declared_total,
                limits.max_total_uncompressed_bytes,
            )?;
            ensure_compression_ratio_within_limit(
                uncompressed_bytes,
                file.compressed_size(),
                limits.max_compression_ratio,
            )?;
        }

        let mut parts = Vec::with_capacity(archive.len());
        let mut content_types_xml = None;
        let mut actual_total = 0_u64;

        for index in 0..archive.len() {
            let mut file = archive.by_index(index).map_err(zip_error)?;
            if file.is_dir() {
                continue;
            }
            let name = normalize_part_name(file.name())?;
            let compression = file.compression();
            let mut part_bytes = Vec::new();
            let remaining_total = limits
                .max_total_uncompressed_bytes
                .saturating_sub(actual_total);
            let read_limit = limits
                .max_entry_uncompressed_bytes
                .min(remaining_total)
                .saturating_add(1);
            file.by_ref()
                .take(read_limit)
                .read_to_end(&mut part_bytes)
                .map_err(io_error)?;
            ensure_within_limit(
                "entry uncompressed bytes",
                u64::try_from(part_bytes.len()).unwrap_or(u64::MAX),
                limits.max_entry_uncompressed_bytes,
            )?;
            actual_total = actual_total
                .checked_add(u64::try_from(part_bytes.len()).unwrap_or(u64::MAX))
                .ok_or_else(|| {
                    resource_limit_error(
                        "total uncompressed bytes",
                        u64::MAX,
                        limits.max_total_uncompressed_bytes,
                    )
                })?;
            ensure_within_limit(
                "total uncompressed bytes",
                actual_total,
                limits.max_total_uncompressed_bytes,
            )?;
            if part_identity_key(&name).as_deref() == Some("[content_types].xml") {
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
        let mut identities = BTreeSet::new();

        for part in &self.parts {
            let canonical_name = canonical_part_name(&part.name, OmErrorCode::InvalidArgument)?;
            if !identities.insert(canonical_name.identity) {
                return Err(OmError::invalid_argument(format!(
                    "duplicate OPC part identity: {}",
                    canonical_name.spelling
                )));
            }
            let options = SimpleFileOptions::default().compression_method(part.compression);
            writer
                .start_file(&canonical_name.spelling, options)
                .map_err(zip_error)?;
            writer.write_all(&part.bytes).map_err(io_error)?;
        }

        let cursor = writer.finish().map_err(zip_error)?;
        Ok(cursor.into_inner())
    }

    pub fn contains(&self, name: &str) -> bool {
        let Some(identity) = part_identity_key(name) else {
            return false;
        };
        self.parts
            .iter()
            .any(|part| part_identity_key(&part.name).as_deref() == Some(identity.as_str()))
    }

    pub fn part(&self, name: &str) -> Option<&OpcPart> {
        let identity = part_identity_key(name)?;
        self.parts
            .iter()
            .find(|part| part_identity_key(&part.name).as_deref() == Some(identity.as_str()))
    }

    pub fn parts(&self) -> &[OpcPart] {
        &self.parts
    }

    pub fn replace_part_bytes(&mut self, name: &str, bytes: Vec<u8>) -> OmResult<()> {
        let canonical_name = canonical_part_name(name, OmErrorCode::InvalidArgument)?;
        let part = self
            .parts
            .iter_mut()
            .find(|part| {
                part_identity_key(&part.name).as_deref() == Some(canonical_name.identity.as_str())
            })
            .ok_or_else(|| {
                OmError::new(
                    OmErrorCode::NotFound,
                    format!("OPC part not found: {}", canonical_name.spelling),
                )
            })?;
        part.bytes = bytes;
        Ok(())
    }

    pub fn add_part(&mut self, mut part: OpcPart) -> OmResult<()> {
        let canonical_name = canonical_part_name(&part.name, OmErrorCode::InvalidArgument)?;
        for existing in &self.parts {
            let existing_name = canonical_part_name(&existing.name, OmErrorCode::InvalidArgument)?;
            if existing_name.identity == canonical_name.identity {
                return Err(OmError::new(
                    OmErrorCode::InvalidArgument,
                    format!("OPC part already exists: {}", canonical_name.spelling),
                ));
            }
        }
        part.name = canonical_name.spelling;
        self.parts.push(part);
        Ok(())
    }

    pub fn remove_part(&mut self, name: &str) -> bool {
        let Some(identity) = part_identity_key(name) else {
            return false;
        };
        let original_len = self.parts.len();
        self.parts
            .retain(|part| part_identity_key(&part.name).as_deref() != Some(identity.as_str()));
        self.parts.len() != original_len
    }
}

fn ensure_within_limit(label: &str, actual: u64, maximum: u64) -> OmResult<()> {
    if actual > maximum {
        return Err(resource_limit_error(label, actual, maximum));
    }
    Ok(())
}

fn ensure_compression_ratio_within_limit(
    uncompressed_bytes: u64,
    compressed_bytes: u64,
    maximum_ratio: u64,
) -> OmResult<()> {
    if uncompressed_bytes == 0 {
        return Ok(());
    }
    let allowed_uncompressed =
        u128::from(compressed_bytes).saturating_mul(u128::from(maximum_ratio));
    if compressed_bytes == 0 || u128::from(uncompressed_bytes) > allowed_uncompressed {
        return Err(OmError::resource_limit(format!(
            "OPC ZIP compression ratio limit exceeded: {uncompressed_bytes} uncompressed bytes, {compressed_bytes} compressed bytes, maximum ratio {maximum_ratio}:1"
        )));
    }
    Ok(())
}

fn resource_limit_error(label: &str, actual: u64, maximum: u64) -> OmError {
    OmError::resource_limit(format!(
        "OPC ZIP {label} limit exceeded: {actual} > {maximum}"
    ))
}

fn preflight_zip_entry_count(bytes: &[u8]) -> OmResult<u64> {
    try_preflight_zip_entry_count(bytes).ok_or_else(|| {
        OmError::parse("OPC ZIP end-of-central-directory entry count is missing or malformed")
    })
}

fn try_preflight_zip_entry_count(bytes: &[u8]) -> Option<u64> {
    const EOCD_SIGNATURE: [u8; 4] = [0x50, 0x4b, 0x05, 0x06];
    const ZIP64_EOCD_SIGNATURE: [u8; 4] = [0x50, 0x4b, 0x06, 0x06];
    const ZIP64_LOCATOR_SIGNATURE: [u8; 4] = [0x50, 0x4b, 0x06, 0x07];
    const EOCD_BYTES: usize = 22;
    const MAX_COMMENT_BYTES: usize = u16::MAX as usize;

    if bytes.len() < EOCD_BYTES {
        return None;
    }
    let search_start = bytes.len().saturating_sub(EOCD_BYTES + MAX_COMMENT_BYTES);
    let latest_start = bytes.len() - EOCD_BYTES;
    for offset in (search_start..=latest_start).rev() {
        if bytes.get(offset..offset.checked_add(4)?)? != EOCD_SIGNATURE {
            continue;
        }
        let comment_bytes = read_u16_le(bytes, offset + 20)? as usize;
        if offset.checked_add(EOCD_BYTES)?.checked_add(comment_bytes)? != bytes.len() {
            continue;
        }
        let entries = read_u16_le(bytes, offset + 10)?;
        if entries != u16::MAX {
            return Some(u64::from(entries));
        }

        let locator_offset = offset.checked_sub(20)?;
        if bytes.get(locator_offset..locator_offset.checked_add(4)?)? != ZIP64_LOCATOR_SIGNATURE {
            return None;
        }
        let zip64_offset = usize::try_from(read_u64_le(bytes, locator_offset + 8)?).ok()?;
        if bytes.get(zip64_offset..zip64_offset.checked_add(4)?)? != ZIP64_EOCD_SIGNATURE {
            return None;
        }
        return read_u64_le(bytes, zip64_offset + 32);
    }
    None
}

fn read_u16_le(bytes: &[u8], offset: usize) -> Option<u16> {
    let end = offset.checked_add(2)?;
    Some(u16::from_le_bytes(bytes.get(offset..end)?.try_into().ok()?))
}

fn read_u64_le(bytes: &[u8], offset: usize) -> Option<u64> {
    let end = offset.checked_add(8)?;
    Some(u64::from_le_bytes(bytes.get(offset..end)?.try_into().ok()?))
}

#[derive(Debug, Default)]
struct ContentTypesManifest {
    defaults: BTreeMap<String, String>,
    overrides: BTreeMap<String, String>,
}

impl ContentTypesManifest {
    fn resolve(&self, part_name: &str) -> Option<String> {
        let override_key = part_identity_key(part_name)?;
        if let Some(content_type) = self.overrides.get(&override_key) {
            return Some(content_type.clone());
        }

        let extension = part_name
            .rsplit_once('.')
            .map(|(_, ext)| ext.to_ascii_lowercase())?;
        self.defaults.get(&extension).cloned()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CanonicalPartName {
    spelling: String,
    identity: String,
}

fn normalize_part_name(name: &str) -> OmResult<String> {
    Ok(canonical_part_name(name, OmErrorCode::Parse)?.spelling)
}

fn part_identity_key(name: &str) -> Option<String> {
    canonical_part_name(name, OmErrorCode::InvalidArgument)
        .ok()
        .map(|name| name.identity)
}

fn canonical_part_name(name: &str, error_code: OmErrorCode) -> OmResult<CanonicalPartName> {
    let spelling = name.strip_prefix('/').unwrap_or(name);
    let invalid = |reason: &str| {
        OmError::new(
            error_code,
            format!("invalid OPC part name {name:?}: {reason}"),
        )
    };
    if spelling.is_empty() {
        return Err(invalid("name is empty"));
    }
    if spelling.starts_with('/') || spelling.ends_with('/') {
        return Err(invalid("leading or trailing slash"));
    }
    if spelling
        .chars()
        .any(|ch| ch.is_control() || matches!(ch, '\\' | '?' | '#' | ':'))
    {
        return Err(invalid("forbidden URI character"));
    }

    let mut identity_segments = Vec::new();
    for segment in spelling.split('/') {
        if segment.is_empty() {
            return Err(invalid("empty path segment"));
        }
        let identity_segment = canonical_part_segment_identity(segment)
            .ok_or_else(|| invalid("malformed or forbidden percent encoding"))?;
        if matches!(identity_segment.as_str(), "." | "..") {
            return Err(invalid("dot path segment"));
        }
        if identity_segment.ends_with('.') {
            return Err(invalid("path segment ends with a dot"));
        }
        identity_segments.push(identity_segment);
    }

    Ok(CanonicalPartName {
        spelling: spelling.to_string(),
        identity: identity_segments.join("/"),
    })
}

fn canonical_part_segment_identity(segment: &str) -> Option<String> {
    let bytes = segment.as_bytes();
    let mut identity = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] != b'%' {
            let byte = bytes[index];
            identity.push(if byte.is_ascii_uppercase() {
                byte.to_ascii_lowercase()
            } else {
                byte
            });
            index += 1;
            continue;
        }

        let high = decode_hex(*bytes.get(index + 1)?)?;
        let low = decode_hex(*bytes.get(index + 2)?)?;
        let decoded = (high << 4) | low;
        if decoded == 0 || decoded.is_ascii_control() || matches!(decoded, b'/' | b'\\') {
            return None;
        }
        if decoded.is_ascii_alphanumeric() || matches!(decoded, b'-' | b'.' | b'_' | b'~') {
            identity.push(decoded.to_ascii_lowercase());
        } else {
            identity.push(b'%');
            identity.push(encode_hex(high));
            identity.push(encode_hex(low));
        }
        index += 3;
    }
    String::from_utf8(identity).ok()
}

fn decode_hex(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn encode_hex(nibble: u8) -> u8 {
    match nibble {
        0..=9 => b'0' + nibble,
        10..=15 => b'A' + nibble - 10,
        _ => unreachable!("hex nibble"),
    }
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
                            let canonical_name =
                                canonical_part_name(&part_name, OmErrorCode::Parse)?;
                            if manifest
                                .overrides
                                .insert(canonical_name.identity, content_type)
                                .is_some()
                            {
                                return Err(OmError::parse(format!(
                                    "duplicate OPC content type override: {part_name}"
                                )));
                            }
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
    use super::{CompressionMethod, LoadLimits, OpcPackage, OpcPart};
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
    fn adds_new_part_with_normalized_name() {
        let mut package = OpcPackage::new(vec![OpcPart {
            name: "xl/workbook.xml".to_string(),
            content_type: Some(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                    .to_string(),
            ),
            compression: CompressionMethod::Stored,
            bytes: b"<workbook/>".to_vec(),
        }]);

        package
            .add_part(OpcPart {
                name: "/xl/worksheets/sheet2.xml".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
                        .to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: b"<worksheet/>".to_vec(),
            })
            .expect("add worksheet part");

        assert!(package.contains("xl/worksheets/sheet2.xml"));
        assert_eq!(
            package
                .part("xl/worksheets/sheet2.xml")
                .expect("worksheet part")
                .bytes
                .as_slice(),
            b"<worksheet/>"
        );
    }

    #[test]
    fn add_part_rejects_duplicate_part_names() {
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
            .add_part(OpcPart {
                name: "/xl/workbook.xml".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                        .to_string(),
                ),
                compression: CompressionMethod::Stored,
                bytes: b"<other/>".to_vec(),
            })
            .expect_err("duplicate part should be rejected");

        assert_eq!(error.code, OmErrorCode::InvalidArgument);
        assert!(error.message.contains("xl/workbook.xml"));
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

    #[test]
    fn bounded_load_rejects_archive_bytes_over_limit() {
        let bytes = zip_bytes(&[("xl/workbook.xml", CompressionMethod::Stored, b"<workbook/>")]);
        let limits = LoadLimits {
            max_archive_bytes: bytes.len() as u64 - 1,
            ..LoadLimits::default()
        };

        let error = OpcPackage::from_bytes_with_limits(&bytes, limits)
            .expect_err("oversized archive should fail");

        assert_eq!(error.code, OmErrorCode::ResourceLimit);
        assert!(error.message.contains("archive bytes"));
    }

    #[test]
    fn bounded_load_rejects_entry_count_over_limit() {
        let bytes = zip_bytes(&[
            ("xl/workbook.xml", CompressionMethod::Stored, b"<workbook/>"),
            (
                "xl/worksheets/sheet1.xml",
                CompressionMethod::Stored,
                b"<worksheet/>",
            ),
        ]);
        let limits = LoadLimits {
            max_entries: 1,
            ..LoadLimits::default()
        };

        let error = OpcPackage::from_bytes_with_limits(&bytes, limits)
            .expect_err("entry flood should fail");

        assert_eq!(error.code, OmErrorCode::ResourceLimit);
        assert!(error.message.contains("entry count"));
    }

    #[test]
    fn bounded_load_rejects_part_name_over_limit() {
        let bytes = zip_bytes(&[("xl/workbook.xml", CompressionMethod::Stored, b"<workbook/>")]);
        let limits = LoadLimits {
            max_part_name_bytes: 8,
            ..LoadLimits::default()
        };

        let error = OpcPackage::from_bytes_with_limits(&bytes, limits)
            .expect_err("long part name should fail");

        assert_eq!(error.code, OmErrorCode::ResourceLimit);
        assert!(error.message.contains("part name bytes"));
    }

    #[test]
    fn bounded_load_rejects_single_and_total_uncompressed_bytes_over_limit() {
        let bytes = zip_bytes(&[
            ("one.bin", CompressionMethod::Stored, b"12345"),
            ("two.bin", CompressionMethod::Stored, b"67890"),
        ]);
        let entry_error = OpcPackage::from_bytes_with_limits(
            &bytes,
            LoadLimits {
                max_entry_uncompressed_bytes: 4,
                ..LoadLimits::default()
            },
        )
        .expect_err("oversized entry should fail");
        let total_error = OpcPackage::from_bytes_with_limits(
            &bytes,
            LoadLimits {
                max_total_uncompressed_bytes: 9,
                ..LoadLimits::default()
            },
        )
        .expect_err("oversized total should fail");

        assert_eq!(entry_error.code, OmErrorCode::ResourceLimit);
        assert!(entry_error.message.contains("entry uncompressed bytes"));
        assert_eq!(total_error.code, OmErrorCode::ResourceLimit);
        assert!(total_error.message.contains("total uncompressed bytes"));
    }

    #[test]
    fn bounded_load_rejects_compression_ratio_over_limit() {
        let repeated = vec![0_u8; 16 * 1024];
        let bytes = zip_bytes(&[(
            "xl/worksheets/sheet1.xml",
            CompressionMethod::Deflated,
            repeated.as_slice(),
        )]);
        let limits = LoadLimits {
            max_compression_ratio: 2,
            ..LoadLimits::default()
        };

        let error = OpcPackage::from_bytes_with_limits(&bytes, limits)
            .expect_err("compression bomb should fail");

        assert_eq!(error.code, OmErrorCode::ResourceLimit);
        assert!(error.message.contains("compression ratio"));
    }

    #[test]
    fn bounded_load_accepts_values_exactly_at_each_limit() {
        let bytes = zip_bytes(&[("one.bin", CompressionMethod::Stored, b"12345")]);
        let limits = LoadLimits {
            max_archive_bytes: bytes.len() as u64,
            max_entries: 1,
            max_part_name_bytes: "one.bin".len(),
            max_entry_uncompressed_bytes: 5,
            max_total_uncompressed_bytes: 5,
            max_compression_ratio: 1,
        };

        let package = OpcPackage::from_bytes_with_limits(&bytes, limits)
            .expect("exact limits should be accepted");

        assert_eq!(package.parts().len(), 1);
        assert_eq!(package.parts()[0].bytes, b"12345");
    }

    #[test]
    fn bounded_load_rejects_duplicate_raw_zip_entry_names() {
        let mut bytes = zip_bytes(&[
            ("one.bin", CompressionMethod::Stored, b"one"),
            ("two.bin", CompressionMethod::Stored, b"two"),
        ]);
        for offset in 0..=bytes.len() - b"two.bin".len() {
            if bytes[offset..].starts_with(b"two.bin") {
                bytes[offset..offset + b"two.bin".len()].copy_from_slice(b"one.bin");
            }
        }

        let error = OpcPackage::from_bytes(&bytes).expect_err("duplicate names should fail");

        assert_eq!(error.code, OmErrorCode::Parse);
        assert!(error.message.contains("duplicate ZIP entry name"));
    }

    #[test]
    fn load_rejects_noncanonical_opc_part_names() {
        for name in [
            "xl/../evil.xml",
            "xl//evil.xml",
            "xl\\evil.xml",
            "xl/workbook.xml.",
            "xl/%GG.xml",
            "xl/%2F.xml",
            "xl/item.xml#fragment",
            "xl/item.xml?query",
            "https:evil.xml",
        ] {
            let bytes = zip_bytes(&[(name, CompressionMethod::Stored, b"opaque")]);
            let error =
                OpcPackage::from_bytes(&bytes).expect_err("noncanonical OPC part name should fail");

            assert_eq!(error.code, OmErrorCode::Parse, "name: {name}");
            assert!(error.message.contains("OPC part name"), "name: {name}");
        }
    }

    #[test]
    fn load_rejects_case_or_percent_equivalent_part_identities() {
        for second_name in ["XL/WORKBOOK.XML", "xl/%77orkbook.xml"] {
            let bytes = zip_bytes(&[
                ("xl/workbook.xml", CompressionMethod::Stored, b"first"),
                (second_name, CompressionMethod::Stored, b"second"),
            ]);
            let error = OpcPackage::from_bytes(&bytes)
                .expect_err("equivalent OPC part identities should fail");

            assert_eq!(error.code, OmErrorCode::Parse, "name: {second_name}");
            assert!(error.message.contains("duplicate OPC part identity"));
        }
    }

    #[test]
    fn part_lookup_and_add_are_case_insensitive() {
        let mut package = OpcPackage::new(vec![OpcPart {
            name: "xl/workbook.xml".to_string(),
            content_type: Some("application/xml".to_string()),
            compression: CompressionMethod::Stored,
            bytes: b"<workbook/>".to_vec(),
        }]);

        assert!(package.contains("/XL/WORKBOOK.XML"));
        assert!(package.part("xl/%77orkbook.xml").is_some());

        let error = package
            .add_part(OpcPart {
                name: "XL/WORKBOOK.XML".to_string(),
                content_type: Some("application/xml".to_string()),
                compression: CompressionMethod::Stored,
                bytes: b"<other/>".to_vec(),
            })
            .expect_err("case-equivalent duplicate should fail");
        assert_eq!(error.code, OmErrorCode::InvalidArgument);
    }

    #[test]
    fn content_type_override_lookup_uses_canonical_part_identity() {
        let package = OpcPackage::new(vec![
            OpcPart {
                name: "[Content_Types].xml".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/XL/WORKBOOK.XML" ContentType="application/workbook"/></Types>"#.to_vec(),
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
                .and_then(|part| part.content_type.as_deref()),
            Some("application/workbook")
        );
    }

    #[test]
    fn serialization_revalidates_infallibly_constructed_packages() {
        let package = OpcPackage::new(vec![OpcPart {
            name: "xl/../evil.xml".to_string(),
            content_type: None,
            compression: CompressionMethod::Stored,
            bytes: b"opaque".to_vec(),
        }]);

        let error = package
            .to_bytes()
            .expect_err("invalid in-memory part name should not serialize");

        assert_eq!(error.code, OmErrorCode::InvalidArgument);
        assert!(error.message.contains("OPC part name"));
    }

    fn zip_bytes(entries: &[(&str, CompressionMethod, &[u8])]) -> Vec<u8> {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        for (name, compression, bytes) in entries {
            writer
                .start_file(
                    *name,
                    SimpleFileOptions::default().compression_method(*compression),
                )
                .expect("start ZIP entry");
            writer.write_all(bytes).expect("write ZIP entry");
        }
        writer.finish().expect("finish ZIP").into_inner()
    }
}
