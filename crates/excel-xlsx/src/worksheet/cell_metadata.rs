//! Workbook cell metadata (`xl/metadata.xml`): which worksheet `c@cm` indices mark a formula as
//! a dynamic array (`XLDAPR` with `dynamicArrayProperties fDynamic="1"`).

use super::super::{xml::resolved_element_is, xml_error};
use office_common::{OmError, OmResult};
use quick_xml::NsReader;
use quick_xml::events::Event;
use std::collections::{BTreeMap, BTreeSet};
use std::io::Cursor;

pub(crate) const SHEET_METADATA_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
const DYNAMIC_ARRAY_NAMESPACE: &[u8] =
    b"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";
const DYNAMIC_ARRAY_METADATA_TYPE: &str = "XLDAPR";

/// The workbook cell-metadata part and the one-based `cm` indices that mark dynamic arrays.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct WorkbookCellMetadata {
    pub part_uri: Option<String>,
    pub dynamic_array_indices: BTreeSet<u32>,
}

impl WorkbookCellMetadata {
    /// The `cm` index new dynamic-array anchors reference, if the part already defines one.
    pub fn dynamic_array_index(&self) -> Option<u32> {
        self.dynamic_array_indices.first().copied()
    }
}

/// Resolves every `cellMetadata` block whose record points at an `XLDAPR` future-metadata block
/// with `fDynamic` set. Malformed record indices fail closed with the part URI.
pub(crate) fn parse_dynamic_array_cell_metadata(
    xml: &[u8],
    spreadsheet_namespace: &str,
    part_uri: &str,
) -> OmResult<BTreeSet<u32>> {
    let mut reader = NsReader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let namespace = spreadsheet_namespace.as_bytes();
    let mut metadata_types = Vec::<String>::new();
    let mut future_blocks = BTreeMap::<String, Vec<bool>>::new();
    let mut current_future = None::<String>;
    let mut in_future_block = false;
    let mut in_cell_metadata = false;
    let mut in_cell_block = false;
    let mut cell_block_records = Vec::<Vec<(u32, u32)>>::new();
    let parse_index = |value: &str, name: &str| {
        value.parse::<u32>().map_err(|_| {
            OmError::parse(format!(
                "{part_uri}: cell metadata record has invalid {name} index: {value}"
            ))
        })
    };
    loop {
        let event = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        match event {
            (resolved, Event::Start(element)) | (resolved, Event::Empty(element))
                if resolved_element_is(
                    &resolved,
                    element.local_name(),
                    namespace,
                    b"metadataType",
                ) =>
            {
                let name = element
                    .try_get_attribute("name")
                    .map_err(xml_error)?
                    .map(|attribute| {
                        attribute
                            .decode_and_unescape_value(reader.decoder())
                            .map(|value| value.into_owned())
                    })
                    .transpose()
                    .map_err(xml_error)?
                    .unwrap_or_default();
                metadata_types.push(name);
            }
            (resolved, Event::Start(element))
                if resolved_element_is(
                    &resolved,
                    element.local_name(),
                    namespace,
                    b"futureMetadata",
                ) =>
            {
                let name = element
                    .try_get_attribute("name")
                    .map_err(xml_error)?
                    .map(|attribute| {
                        attribute
                            .decode_and_unescape_value(reader.decoder())
                            .map(|value| value.into_owned())
                    })
                    .transpose()
                    .map_err(xml_error)?
                    .unwrap_or_default();
                future_blocks.entry(name.clone()).or_default();
                current_future = Some(name);
            }
            (resolved, Event::End(element))
                if resolved_element_is(
                    &resolved,
                    element.local_name(),
                    namespace,
                    b"futureMetadata",
                ) =>
            {
                current_future = None;
            }
            (resolved, Event::Start(element))
                if resolved_element_is(&resolved, element.local_name(), namespace, b"bk") =>
            {
                if let Some(name) = current_future.as_ref() {
                    future_blocks.entry(name.clone()).or_default().push(false);
                    in_future_block = true;
                } else if in_cell_metadata {
                    cell_block_records.push(Vec::new());
                    in_cell_block = true;
                }
            }
            (resolved, Event::Empty(element))
                if resolved_element_is(&resolved, element.local_name(), namespace, b"bk") =>
            {
                if let Some(name) = current_future.as_ref() {
                    future_blocks.entry(name.clone()).or_default().push(false);
                } else if in_cell_metadata {
                    cell_block_records.push(Vec::new());
                }
            }
            (resolved, Event::End(element))
                if resolved_element_is(&resolved, element.local_name(), namespace, b"bk") =>
            {
                in_future_block = false;
                in_cell_block = false;
            }
            (resolved, Event::Start(element)) | (resolved, Event::Empty(element))
                if in_future_block
                    && resolved_element_is(
                        &resolved,
                        element.local_name(),
                        DYNAMIC_ARRAY_NAMESPACE,
                        b"dynamicArrayProperties",
                    ) =>
            {
                let dynamic = element
                    .try_get_attribute("fDynamic")
                    .map_err(xml_error)?
                    .map(|attribute| {
                        attribute
                            .decode_and_unescape_value(reader.decoder())
                            .map(|value| matches!(value.as_ref(), "1" | "true"))
                    })
                    .transpose()
                    .map_err(xml_error)?
                    .unwrap_or(false);
                if let Some(name) = current_future.as_ref()
                    && let Some(last) = future_blocks.entry(name.clone()).or_default().last_mut()
                {
                    *last |= dynamic;
                }
            }
            (resolved, Event::Start(element))
                if resolved_element_is(
                    &resolved,
                    element.local_name(),
                    namespace,
                    b"cellMetadata",
                ) =>
            {
                in_cell_metadata = true;
            }
            (resolved, Event::End(element))
                if resolved_element_is(
                    &resolved,
                    element.local_name(),
                    namespace,
                    b"cellMetadata",
                ) =>
            {
                in_cell_metadata = false;
            }
            (resolved, Event::Start(element)) | (resolved, Event::Empty(element))
                if in_cell_block
                    && resolved_element_is(&resolved, element.local_name(), namespace, b"rc") =>
            {
                let mut record_type = None;
                let mut record_value = None;
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(xml_error)?;
                    let value = attribute
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(xml_error)?
                        .into_owned();
                    match attribute.key.as_ref() {
                        b"t" => record_type = Some(parse_index(&value, "type")?),
                        b"v" => record_value = Some(parse_index(&value, "value")?),
                        _ => {}
                    }
                }
                let (Some(record_type), Some(record_value)) = (record_type, record_value) else {
                    return Err(OmError::parse(format!(
                        "{part_uri}: cell metadata record is missing its t or v index"
                    )));
                };
                if record_type == 0 || record_type as usize > metadata_types.len() {
                    return Err(OmError::parse(format!(
                        "{part_uri}: cell metadata record references unknown metadata type {record_type}"
                    )));
                }
                if let Some(records) = cell_block_records.last_mut() {
                    records.push((record_type, record_value));
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
        buffer.clear();
    }

    let mut dynamic_array_indices = BTreeSet::new();
    for (index, records) in cell_block_records.iter().enumerate() {
        let dynamic = records.iter().any(|&(record_type, record_value)| {
            let type_name = &metadata_types[record_type as usize - 1];
            type_name == DYNAMIC_ARRAY_METADATA_TYPE
                && future_blocks
                    .get(type_name)
                    .and_then(|blocks| blocks.get(record_value as usize))
                    .copied()
                    .unwrap_or(false)
        });
        if dynamic {
            dynamic_array_indices.insert(u32::try_from(index + 1).map_err(|_| {
                OmError::parse(format!("{part_uri}: too many cell metadata blocks"))
            })?);
        }
    }
    Ok(dynamic_array_indices)
}

/// A cell-metadata part with one dynamic-array block, referenced by `cm="1"`, in the form
/// desktop Excel writes.
pub(crate) fn dynamic_array_cell_metadata_xml(spreadsheet_namespace: &str) -> Vec<u8> {
    format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            "\r\n",
            r#"<metadata xmlns="{namespace}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">"#,
            r#"<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>"#,
            r#"<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>"#,
            r#"<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>"#,
            "</metadata>"
        ),
        namespace = spreadsheet_namespace
    )
    .into_bytes()
}
