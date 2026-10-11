//! Values Excel cached for external workbooks in `externalLink` parts.
//!
//! A workbook lists its links in `externalReferences`, and its formulas name them by position:
//! `[1]Sheet1!A1` reads the first link. Each link part caches the linked workbook's sheet names
//! and the values of the cells this workbook references, which formulas read while the link is
//! not updated. Reading the cache never opens the linked file.

use super::relationships::{
    RelationshipEntry, parse_relationship_entries_with_options,
    relationship_base_segments_for_part, relationships_part_uri_for_part,
};
use super::worksheet::parse_cell_reference;
use super::xml::decode_general_reference;
use excel_model::{CellData, ExternalBookCache, WorksheetData};
use office_common::{CellError, CellValue};
use office_opc::OpcPackage;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::collections::BTreeMap;

/// The cache of every `externalReference` in workbook order. A link whose part is missing or
/// unreadable, or that is not a workbook link (DDE, OLE), keeps its place as `None`.
pub(super) fn parse_external_book_caches(
    workbook_xml: &[u8],
    relationships: &BTreeMap<String, RelationshipEntry>,
    package: &OpcPackage,
) -> Vec<Option<ExternalBookCache>> {
    external_reference_ids(workbook_xml)
        .into_iter()
        .map(|id| {
            let relationship = relationships.get(&id)?;
            let part = package.part(&relationship.target)?;
            parse_external_book_cache(&part.bytes, &relationship.target, package)
        })
        .collect()
}

fn relationship_id(element: &BytesStart<'_>, reader: &Reader<&[u8]>) -> Option<String> {
    element.attributes().flatten().find_map(|attr| {
        (attr.key.prefix().is_some() && attr.key.local_name().as_ref() == b"id")
            .then(|| attr.decode_and_unescape_value(reader.decoder()).ok())
            .flatten()
            .map(|value| value.into_owned())
    })
}

fn attribute(element: &BytesStart<'_>, reader: &Reader<&[u8]>, name: &[u8]) -> Option<String> {
    element.attributes().flatten().find_map(|attr| {
        (attr.key.as_ref() == name)
            .then(|| attr.decode_and_unescape_value(reader.decoder()).ok())
            .flatten()
            .map(|value| value.into_owned())
    })
}

fn external_reference_ids(workbook_xml: &[u8]) -> Vec<String> {
    let mut reader = Reader::from_reader(workbook_xml);
    let mut ids = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element))
                if element.local_name().as_ref() == b"externalReference" =>
            {
                if let Some(id) = relationship_id(&element, &reader) {
                    ids.push(id);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    ids
}

/// One cached cell while its `v` text is read.
struct PendingCell {
    sheet: usize,
    coordinates: (u32, u32),
    kind: String,
    text: Option<String>,
}

fn parse_external_book_cache(
    xml: &[u8],
    part_uri: &str,
    package: &OpcPackage,
) -> Option<ExternalBookCache> {
    let mut reader = Reader::from_reader(xml);
    let mut cache = None::<ExternalBookCache>;
    let mut cells = BTreeMap::<usize, BTreeMap<(u32, u32), CellData>>::new();
    let mut sheet = None::<usize>;
    let mut cell = None::<PendingCell>;
    let mut in_value = false;
    loop {
        match reader.read_event().ok()? {
            // An empty value leaves the cell without one.
            Event::Empty(element) if element.local_name().as_ref() == b"v" => {}
            Event::Start(element) | Event::Empty(element) => match element.local_name().as_ref() {
                b"externalBook" => {
                    let target = relationship_id(&element, &reader).and_then(|id| {
                        let rels = package.part(&relationships_part_uri_for_part(part_uri)?)?;
                        // The linked file is an external target.
                        let base_segments = relationship_base_segments_for_part(part_uri);
                        parse_relationship_entries_with_options(&rels.bytes, &base_segments, true)
                            .ok()?
                            .into_iter()
                            .find(|relationship| relationship.id == id)
                            .map(|relationship| relationship.target)
                    });
                    cache = Some(ExternalBookCache {
                        target,
                        ..ExternalBookCache::default()
                    });
                }
                b"sheetName" => {
                    if let Some(cache) = cache.as_mut() {
                        cache
                            .sheet_names
                            .push(attribute(&element, &reader, b"val").unwrap_or_default());
                    }
                }
                b"sheetData" => {
                    sheet = attribute(&element, &reader, b"sheetId")
                        .and_then(|id| id.parse::<usize>().ok());
                }
                b"cell" => {
                    cell = sheet.and_then(|sheet| {
                        let reference = attribute(&element, &reader, b"r")?;
                        Some(PendingCell {
                            sheet,
                            coordinates: parse_cell_reference(&reference, None).ok()?,
                            kind: attribute(&element, &reader, b"t")
                                .unwrap_or_else(|| "n".to_string()),
                            text: None,
                        })
                    });
                }
                b"v" => in_value = cell.is_some(),
                _ => {}
            },
            Event::Text(text) if in_value => {
                if let Some(cell) = cell.as_mut() {
                    let text = text.xml_content().ok()?;
                    cell.text.get_or_insert_with(String::new).push_str(&text);
                }
            }
            Event::GeneralRef(reference) if in_value => {
                if let Some(cell) = cell.as_mut() {
                    let text = decode_general_reference(&reference, part_uri).ok()?;
                    cell.text.get_or_insert_with(String::new).push_str(&text);
                }
            }
            Event::End(element) => match element.local_name().as_ref() {
                b"v" => in_value = false,
                b"cell" => {
                    if let Some(PendingCell {
                        sheet,
                        coordinates,
                        kind,
                        text: Some(text),
                    }) = cell.take()
                    {
                        let value = match kind.as_str() {
                            "b" => CellValue::Bool(text.trim() == "1"),
                            "e" => CellValue::Error(CellError::from_lexical(text.trim())),
                            "n" => CellValue::Number(text.trim().parse().ok()?),
                            _ => CellValue::Text(text),
                        };
                        cells.entry(sheet).or_default().insert(
                            coordinates,
                            CellData {
                                value,
                                formula: None,
                                style_id: None,
                            },
                        );
                    }
                }
                b"sheetData" => sheet = None,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
    }
    let mut cache = cache?;
    cache.sheets = (0..cache.sheet_names.len())
        .map(|index| WorksheetData {
            cells: cells.remove(&index).unwrap_or_default(),
            ..WorksheetData::default()
        })
        .collect();
    Some(cache)
}
