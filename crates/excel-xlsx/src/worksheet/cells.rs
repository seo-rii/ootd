use super::super::{WorksheetSupportParts, io_error, xml_error};

use excel_model::{CellData, WorksheetData};
use office_common::{CellError, CellValue, FormulaSource, OmError, OmErrorCode, OmResult, StyleId};
use quick_xml::escape::partial_escape;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Write};

#[derive(Debug, Clone)]
enum RowContentSegment {
    Opaque(Vec<u8>),
    Cell((u32, u32)),
}

#[derive(Debug, Clone, Copy)]
enum CellContentSegment {
    Opaque,
    Formula,
    Value,
    InlineString,
}

pub(crate) fn collect_support_part_dimension_coords(
    support_parts: Option<&WorksheetSupportParts>,
) -> BTreeSet<(u32, u32)> {
    let mut coordinates = BTreeSet::new();
    let Some(support_parts) = support_parts else {
        return coordinates;
    };

    if support_parts.hyperlink_summaries.is_empty() {
        for hyperlink_ref in &support_parts.hyperlink_refs {
            extend_dimension_coords_from_reference(&mut coordinates, hyperlink_ref);
        }
    } else {
        for hyperlink_summary in &support_parts.hyperlink_summaries {
            extend_dimension_coords_from_reference(&mut coordinates, &hyperlink_summary.reference);
        }
    }

    if support_parts.comment_summaries.is_empty() {
        for anchor_refs in support_parts.comment_anchor_refs.values() {
            for anchor_ref in anchor_refs {
                extend_dimension_coords_from_reference(&mut coordinates, anchor_ref);
            }
        }
    } else {
        for comment_summary in support_parts.comment_summaries.values() {
            for comment in &comment_summary.comments {
                extend_dimension_coords_from_reference(&mut coordinates, &comment.reference);
            }
        }
    }

    coordinates
}

pub(crate) fn extend_dimension_coords_from_reference(
    coordinates: &mut BTreeSet<(u32, u32)>,
    reference: &str,
) {
    for component in reference.split(',') {
        let normalized = component.trim().replace('$', "");
        let trimmed = normalized.trim();
        if trimmed.is_empty() {
            continue;
        }
        if let Some((start, end)) = trimmed.split_once(':') {
            if let Ok(start_coordinates) = parse_cell_reference(start.trim(), None) {
                coordinates.insert(start_coordinates);
            }
            if let Ok(end_coordinates) = parse_cell_reference(end.trim(), None) {
                coordinates.insert(end_coordinates);
            }
            continue;
        }
        if let Ok(single_coordinates) = parse_cell_reference(trimmed, None) {
            coordinates.insert(single_coordinates);
        }
    }
}

pub(crate) fn parse_cell_reference(
    reference: &str,
    current_row: Option<u32>,
) -> OmResult<(u32, u32)> {
    let mut col = 0u32;
    let mut row = current_row.unwrap_or(0);

    for ch in reference.chars() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else if ch.is_ascii_digit() {
            row = reference
                .chars()
                .skip_while(|value| value.is_ascii_alphabetic())
                .collect::<String>()
                .parse::<u32>()
                .map_err(xml_error)?;
            break;
        }
    }

    if row == 0 || col == 0 {
        return Err(OmError::new(
            OmErrorCode::Parse,
            format!("invalid worksheet cell reference: {reference}"),
        ));
    }

    Ok((row, col))
}

pub(crate) fn parse_worksheet_cells(
    worksheet_xml: &[u8],
    shared_strings: &[String],
) -> OmResult<BTreeMap<(u32, u32), CellData>> {
    let mut reader = Reader::from_reader(Cursor::new(worksheet_xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut cells = BTreeMap::new();
    let mut current_row = None;
    let mut current_field = None;
    let mut current_cell: Option<(
        u32,
        u32,
        Option<String>,
        Option<StyleId>,
        String,
        String,
        String,
    )> = None;

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.name().as_ref() == b"row" => {
                let mut row_index = None;
                for attr in element.attributes() {
                    let attr = attr.map_err(xml_error)?;
                    let value = attr
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(xml_error)?
                        .into_owned();
                    if attr.key.as_ref() == b"r" {
                        row_index = value.parse::<u32>().ok();
                    }
                }
                current_row = row_index;
            }
            Ok(Event::End(element)) if element.name().as_ref() == b"row" => current_row = None,
            Ok(Event::Start(element)) if element.name().as_ref() == b"c" => {
                let mut reference = None;
                let mut cell_type = None;
                let mut style_id = None;
                for attr in element.attributes() {
                    let attr = attr.map_err(xml_error)?;
                    let value = attr
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(xml_error)?
                        .into_owned();
                    match attr.key.as_ref() {
                        b"r" => reference = Some(value),
                        b"t" => cell_type = Some(value),
                        b"s" => style_id = value.parse::<u64>().ok().map(StyleId),
                        _ => {}
                    }
                }
                let reference = reference.ok_or_else(|| {
                    OmError::new(
                        OmErrorCode::Parse,
                        "worksheet cell is missing an A1 reference",
                    )
                })?;
                let (row, col) = parse_cell_reference(reference.as_str(), current_row)?;
                current_cell = Some((
                    row,
                    col,
                    cell_type,
                    style_id,
                    String::new(),
                    String::new(),
                    String::new(),
                ));
            }
            Ok(Event::Empty(element)) if element.name().as_ref() == b"c" => {
                let mut reference = None;
                let mut style_id = None;
                for attr in element.attributes() {
                    let attr = attr.map_err(xml_error)?;
                    let value = attr
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(xml_error)?
                        .into_owned();
                    match attr.key.as_ref() {
                        b"r" => reference = Some(value),
                        b"s" => style_id = value.parse::<u64>().ok().map(StyleId),
                        _ => {}
                    }
                }
                let reference = reference.ok_or_else(|| {
                    OmError::new(
                        OmErrorCode::Parse,
                        "worksheet cell is missing an A1 reference",
                    )
                })?;
                let (row, col) = parse_cell_reference(reference.as_str(), current_row)?;
                if style_id.is_some() {
                    cells.insert(
                        (row, col),
                        CellData {
                            value: CellValue::Blank,
                            formula: None,
                            style_id,
                        },
                    );
                }
            }
            Ok(Event::Start(element)) if element.name().as_ref() == b"f" => {
                current_field = Some("formula");
            }
            Ok(Event::Start(element)) if element.name().as_ref() == b"v" => {
                current_field = Some("value");
            }
            Ok(Event::Start(element)) if element.name().as_ref() == b"t" => {
                if let Some((_, _, cell_type, _, _, _, _)) = &current_cell {
                    if cell_type.as_deref() == Some("inlineStr") {
                        current_field = Some("inline");
                    }
                }
            }
            Ok(Event::End(element)) if matches!(element.name().as_ref(), b"f" | b"v" | b"t") => {
                current_field = None;
            }
            Ok(Event::Text(text)) => {
                if let Some(field) = current_field {
                    if let Some(cell) = current_cell.as_mut() {
                        match field {
                            "formula" => cell.4.push_str(&text.xml_content().map_err(xml_error)?),
                            "value" => cell.5.push_str(&text.xml_content().map_err(xml_error)?),
                            "inline" => cell.6.push_str(&text.xml_content().map_err(xml_error)?),
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::CData(text)) => {
                if let Some(field) = current_field {
                    if let Some(cell) = current_cell.as_mut() {
                        match field {
                            "formula" => cell.4.push_str(&text.xml_content().map_err(xml_error)?),
                            "value" => cell.5.push_str(&text.xml_content().map_err(xml_error)?),
                            "inline" => cell.6.push_str(&text.xml_content().map_err(xml_error)?),
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::End(element)) if element.name().as_ref() == b"c" => {
                if let Some((row, col, cell_type, style_id, formula, value, inline)) =
                    current_cell.take()
                {
                    let cell_value = match cell_type.as_deref() {
                        Some("b") => CellValue::Bool(value == "1"),
                        Some("e") => CellValue::Error(parse_cell_error(value.as_str())),
                        Some("s") => {
                            let index = value.parse::<usize>().map_err(xml_error)?;
                            CellValue::Text(shared_strings.get(index).cloned().ok_or_else(
                                || {
                                    OmError::new(
                                        OmErrorCode::Parse,
                                        format!("shared string index out of range: {index}"),
                                    )
                                },
                            )?)
                        }
                        Some("str") => CellValue::Text(value),
                        Some("inlineStr") => CellValue::Text(inline),
                        _ if value.is_empty() => CellValue::Blank,
                        _ => CellValue::Number(value.parse::<f64>().map_err(xml_error)?),
                    };
                    if !matches!(cell_value, CellValue::Blank)
                        || !formula.is_empty()
                        || style_id.is_some()
                    {
                        cells.insert(
                            (row, col),
                            CellData {
                                value: cell_value,
                                formula: if formula.is_empty() {
                                    None
                                } else {
                                    Some(FormulaSource {
                                        text: formula,
                                        is_r1c1: false,
                                    })
                                },
                                style_id,
                            },
                        );
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }

    Ok(cells)
}

pub(crate) fn rewrite_worksheet_xml(
    worksheet: &WorksheetData,
    support_parts: Option<&WorksheetSupportParts>,
) -> OmResult<Vec<u8>> {
    if worksheet.source_xml.is_empty() {
        return Err(OmError::new(
            OmErrorCode::InvalidState,
            "worksheet source xml is missing",
        ));
    }

    let mut row_templates = BTreeMap::<u32, Vec<(String, String)>>::new();
    let mut cell_templates = BTreeMap::<(u32, u32), Vec<(String, String)>>::new();
    let mut formula_templates = BTreeMap::<(u32, u32), Vec<(String, String)>>::new();
    let mut raw_row_fragments = BTreeMap::<u32, Vec<u8>>::new();
    let mut raw_cell_fragments = BTreeMap::<(u32, u32), Vec<u8>>::new();
    let mut row_content_segments = BTreeMap::<u32, Vec<RowContentSegment>>::new();
    let mut cell_content_segments =
        BTreeMap::<(u32, u32), Vec<(CellContentSegment, Vec<u8>)>>::new();
    let mut template_reader = Reader::from_reader(Cursor::new(worksheet.source_xml.as_slice()));
    template_reader.config_mut().trim_text(false);
    let mut template_buffer = Vec::new();
    let mut current_template_row = None;
    let mut current_template_cell = None;
    let mut current_raw_row = None;
    let mut current_raw_cell = None;
    let mut current_row_content_cursor = None;
    let mut current_row_opaque = None::<(u32, usize, usize)>;
    let mut current_row_segment_cell = None::<(u32, u32)>;
    let mut current_cell_content_cursor = None;
    let mut current_cell_segment = None::<((u32, u32), CellContentSegment, usize, usize)>;
    let collect_attributes =
        |element: &BytesStart<'_>, decoder| -> OmResult<Vec<(String, String)>> {
            let mut attributes = Vec::new();
            for attr in element.attributes() {
                let attr = attr.map_err(xml_error)?;
                let value = attr
                    .decode_and_unescape_value(decoder)
                    .map_err(xml_error)?
                    .into_owned();
                attributes.push((
                    String::from_utf8_lossy(attr.key.as_ref()).into_owned(),
                    value,
                ));
            }
            Ok(attributes)
        };

    loop {
        let start = template_reader.buffer_position() as usize;
        let event = template_reader.read_event_into(&mut template_buffer);
        let end = template_reader.buffer_position() as usize;
        let raw_event = &worksheet.source_xml[start..end];

        match event {
            Ok(Event::Start(element)) => match element.name().as_ref() {
                b"row" => {
                    let attributes = collect_attributes(&element, template_reader.decoder())?;
                    current_template_row = attributes
                        .iter()
                        .find(|(key, _)| key == "r")
                        .and_then(|(_, value)| value.parse::<u32>().ok());
                    if let Some(row_index) = current_template_row {
                        current_raw_row = Some((row_index, raw_event.to_vec()));
                        current_row_content_cursor = Some((row_index, end));
                        row_content_segments.entry(row_index).or_default();
                        row_templates.insert(row_index, attributes);
                    }
                }
                b"c" => {
                    let attributes = collect_attributes(&element, template_reader.decoder())?;
                    let reference = attributes
                        .iter()
                        .find(|(key, _)| key == "r")
                        .map(|(_, value)| value.clone())
                        .ok_or_else(|| {
                            OmError::new(
                                OmErrorCode::Parse,
                                "worksheet cell is missing an A1 reference",
                            )
                        })?;
                    let coordinates =
                        parse_cell_reference(reference.as_str(), current_template_row)?;
                    current_template_cell = Some(coordinates);
                    current_raw_cell = Some((coordinates, raw_event.to_vec()));
                    current_cell_content_cursor = Some((coordinates, end));
                    cell_content_segments.entry(coordinates).or_default();
                    if let Some((row_index, cursor)) = current_row_content_cursor {
                        if current_row_opaque.is_none() && current_row_segment_cell.is_none() {
                            if cursor < start {
                                row_content_segments.entry(row_index).or_default().push(
                                    RowContentSegment::Opaque(
                                        worksheet.source_xml[cursor..start].to_vec(),
                                    ),
                                );
                            }
                            current_row_segment_cell = Some(coordinates);
                        }
                    }
                    if let Some((_, raw_row)) = current_raw_row.as_mut() {
                        raw_row.extend_from_slice(raw_event);
                    }
                    cell_templates.insert(coordinates, attributes);
                }
                b"f" => {
                    if let Some(coordinates) = current_template_cell {
                        formula_templates.insert(
                            coordinates,
                            collect_attributes(&element, template_reader.decoder())?,
                        );
                        if current_cell_segment.is_none()
                            && let Some((segment_coordinates, cursor)) = current_cell_content_cursor
                            && segment_coordinates == coordinates
                        {
                            if cursor < start {
                                cell_content_segments.entry(coordinates).or_default().push((
                                    CellContentSegment::Opaque,
                                    worksheet.source_xml[cursor..start].to_vec(),
                                ));
                            }
                            current_cell_segment =
                                Some((coordinates, CellContentSegment::Formula, start, 1));
                        } else if let Some((_, _, _, depth)) = current_cell_segment.as_mut() {
                            *depth += 1;
                        }
                    } else if let Some((row_index, _, depth)) = current_row_opaque.as_mut() {
                        if Some(*row_index) == current_template_row {
                            *depth += 1;
                        }
                    }
                    if let Some((_, raw_row)) = current_raw_row.as_mut() {
                        raw_row.extend_from_slice(raw_event);
                    }
                    if let Some((_, raw_cell)) = current_raw_cell.as_mut() {
                        raw_cell.extend_from_slice(raw_event);
                    }
                }
                b"v" => {
                    if let Some(coordinates) = current_template_cell {
                        if current_cell_segment.is_none()
                            && let Some((segment_coordinates, cursor)) = current_cell_content_cursor
                            && segment_coordinates == coordinates
                        {
                            if cursor < start {
                                cell_content_segments.entry(coordinates).or_default().push((
                                    CellContentSegment::Opaque,
                                    worksheet.source_xml[cursor..start].to_vec(),
                                ));
                            }
                            current_cell_segment =
                                Some((coordinates, CellContentSegment::Value, start, 1));
                        } else if let Some((_, _, _, depth)) = current_cell_segment.as_mut() {
                            *depth += 1;
                        }
                    }
                    if let Some((_, raw_row)) = current_raw_row.as_mut() {
                        raw_row.extend_from_slice(raw_event);
                    }
                    if let Some((_, raw_cell)) = current_raw_cell.as_mut() {
                        raw_cell.extend_from_slice(raw_event);
                    }
                }
                b"is" => {
                    if let Some(coordinates) = current_template_cell {
                        if current_cell_segment.is_none()
                            && let Some((segment_coordinates, cursor)) = current_cell_content_cursor
                            && segment_coordinates == coordinates
                        {
                            if cursor < start {
                                cell_content_segments.entry(coordinates).or_default().push((
                                    CellContentSegment::Opaque,
                                    worksheet.source_xml[cursor..start].to_vec(),
                                ));
                            }
                            current_cell_segment =
                                Some((coordinates, CellContentSegment::InlineString, start, 1));
                        } else if let Some((_, _, _, depth)) = current_cell_segment.as_mut() {
                            *depth += 1;
                        }
                    }
                    if let Some((_, raw_row)) = current_raw_row.as_mut() {
                        raw_row.extend_from_slice(raw_event);
                    }
                    if let Some((_, raw_cell)) = current_raw_cell.as_mut() {
                        raw_cell.extend_from_slice(raw_event);
                    }
                }
                _ => {
                    if let Some((coordinates, _, _, depth)) = current_cell_segment.as_mut() {
                        if Some(*coordinates) == current_template_cell {
                            *depth += 1;
                        }
                    } else if let Some(coordinates) = current_template_cell {
                        if let Some((segment_coordinates, cursor)) = current_cell_content_cursor
                            && segment_coordinates == coordinates
                        {
                            current_cell_segment =
                                Some((coordinates, CellContentSegment::Opaque, cursor, 1));
                        }
                    } else if let Some(row_index) = current_template_row {
                        if let Some((opaque_row_index, _, depth)) = current_row_opaque.as_mut() {
                            if *opaque_row_index == row_index {
                                *depth += 1;
                            }
                        } else if let Some((cursor_row_index, cursor)) = current_row_content_cursor
                            && cursor_row_index == row_index
                        {
                            current_row_opaque = Some((row_index, cursor, 1));
                        }
                    }
                    if let Some((_, raw_row)) = current_raw_row.as_mut() {
                        raw_row.extend_from_slice(raw_event);
                    }
                    if let Some((_, raw_cell)) = current_raw_cell.as_mut() {
                        raw_cell.extend_from_slice(raw_event);
                    }
                }
            },
            Ok(Event::Empty(element)) => match element.name().as_ref() {
                b"row" => {
                    let attributes = collect_attributes(&element, template_reader.decoder())?;
                    if let Some(row_index) = attributes
                        .iter()
                        .find(|(key, _)| key == "r")
                        .and_then(|(_, value)| value.parse::<u32>().ok())
                    {
                        raw_row_fragments.insert(row_index, raw_event.to_vec());
                        row_content_segments.entry(row_index).or_default();
                        row_templates.insert(row_index, attributes);
                    }
                }
                b"c" => {
                    let attributes = collect_attributes(&element, template_reader.decoder())?;
                    let reference = attributes
                        .iter()
                        .find(|(key, _)| key == "r")
                        .map(|(_, value)| value.clone())
                        .ok_or_else(|| {
                            OmError::new(
                                OmErrorCode::Parse,
                                "worksheet cell is missing an A1 reference",
                            )
                        })?;
                    let coordinates =
                        parse_cell_reference(reference.as_str(), current_template_row)?;
                    if let Some((_, raw_row)) = current_raw_row.as_mut() {
                        raw_row.extend_from_slice(raw_event);
                    }
                    if let Some((row_index, cursor)) = current_row_content_cursor
                        && current_row_opaque.is_none()
                        && current_row_segment_cell.is_none()
                    {
                        if cursor < start {
                            row_content_segments.entry(row_index).or_default().push(
                                RowContentSegment::Opaque(
                                    worksheet.source_xml[cursor..start].to_vec(),
                                ),
                            );
                        }
                        row_content_segments
                            .entry(row_index)
                            .or_default()
                            .push(RowContentSegment::Cell(coordinates));
                        current_row_content_cursor = Some((row_index, end));
                    }
                    cell_templates.insert(coordinates, attributes);
                    raw_cell_fragments.insert(coordinates, raw_event.to_vec());
                    cell_content_segments.entry(coordinates).or_default();
                    current_template_cell = None;
                }
                b"f" => {
                    if let Some(coordinates) = current_template_cell {
                        if current_cell_segment.is_none() {
                            formula_templates.insert(
                                coordinates,
                                collect_attributes(&element, template_reader.decoder())?,
                            );
                            if let Some((segment_coordinates, cursor)) = current_cell_content_cursor
                                && segment_coordinates == coordinates
                            {
                                if cursor < start {
                                    cell_content_segments.entry(coordinates).or_default().push((
                                        CellContentSegment::Opaque,
                                        worksheet.source_xml[cursor..start].to_vec(),
                                    ));
                                }
                                cell_content_segments
                                    .entry(coordinates)
                                    .or_default()
                                    .push((CellContentSegment::Formula, raw_event.to_vec()));
                                current_cell_content_cursor = Some((coordinates, end));
                            }
                        }
                    }
                    if let Some((_, raw_row)) = current_raw_row.as_mut() {
                        raw_row.extend_from_slice(raw_event);
                    }
                    if let Some((_, raw_cell)) = current_raw_cell.as_mut() {
                        raw_cell.extend_from_slice(raw_event);
                    }
                }
                b"v" => {
                    if current_cell_segment.is_none()
                        && let Some(coordinates) = current_template_cell
                        && let Some((segment_coordinates, cursor)) = current_cell_content_cursor
                        && segment_coordinates == coordinates
                    {
                        if cursor < start {
                            cell_content_segments.entry(coordinates).or_default().push((
                                CellContentSegment::Opaque,
                                worksheet.source_xml[cursor..start].to_vec(),
                            ));
                        }
                        cell_content_segments
                            .entry(coordinates)
                            .or_default()
                            .push((CellContentSegment::Value, raw_event.to_vec()));
                        current_cell_content_cursor = Some((coordinates, end));
                    }
                    if let Some((_, raw_row)) = current_raw_row.as_mut() {
                        raw_row.extend_from_slice(raw_event);
                    }
                    if let Some((_, raw_cell)) = current_raw_cell.as_mut() {
                        raw_cell.extend_from_slice(raw_event);
                    }
                }
                b"is" => {
                    if current_cell_segment.is_none()
                        && let Some(coordinates) = current_template_cell
                        && let Some((segment_coordinates, cursor)) = current_cell_content_cursor
                        && segment_coordinates == coordinates
                    {
                        if cursor < start {
                            cell_content_segments.entry(coordinates).or_default().push((
                                CellContentSegment::Opaque,
                                worksheet.source_xml[cursor..start].to_vec(),
                            ));
                        }
                        cell_content_segments
                            .entry(coordinates)
                            .or_default()
                            .push((CellContentSegment::InlineString, raw_event.to_vec()));
                        current_cell_content_cursor = Some((coordinates, end));
                    }
                    if let Some((_, raw_row)) = current_raw_row.as_mut() {
                        raw_row.extend_from_slice(raw_event);
                    }
                    if let Some((_, raw_cell)) = current_raw_cell.as_mut() {
                        raw_cell.extend_from_slice(raw_event);
                    }
                }
                _ => {
                    if current_cell_segment.is_none()
                        && let Some((coordinates, cursor)) = current_cell_content_cursor
                        && Some(coordinates) == current_template_cell
                    {
                        cell_content_segments.entry(coordinates).or_default().push((
                            CellContentSegment::Opaque,
                            worksheet.source_xml[cursor..end].to_vec(),
                        ));
                        current_cell_content_cursor = Some((coordinates, end));
                    } else if current_template_cell.is_none()
                        && current_row_opaque.is_none()
                        && let Some((row_index, cursor)) = current_row_content_cursor
                        && Some(row_index) == current_template_row
                    {
                        row_content_segments.entry(row_index).or_default().push(
                            RowContentSegment::Opaque(worksheet.source_xml[cursor..end].to_vec()),
                        );
                        current_row_content_cursor = Some((row_index, end));
                    }
                    if let Some((_, raw_row)) = current_raw_row.as_mut() {
                        raw_row.extend_from_slice(raw_event);
                    }
                    if let Some((_, raw_cell)) = current_raw_cell.as_mut() {
                        raw_cell.extend_from_slice(raw_event);
                    }
                }
            },
            Ok(Event::End(element)) => match element.name().as_ref() {
                b"row" => {
                    if let Some((row_index, cursor)) = current_row_content_cursor
                        && cursor < start
                    {
                        row_content_segments.entry(row_index).or_default().push(
                            RowContentSegment::Opaque(worksheet.source_xml[cursor..start].to_vec()),
                        );
                    }
                    if let Some((row_index, mut raw_row)) = current_raw_row.take() {
                        raw_row.extend_from_slice(raw_event);
                        raw_row_fragments.insert(row_index, raw_row);
                    }
                    current_row_content_cursor = None;
                    current_row_opaque = None;
                    current_row_segment_cell = None;
                    current_template_row = None;
                }
                b"c" => {
                    if let Some((coordinates, cursor)) = current_cell_content_cursor
                        && cursor < start
                    {
                        cell_content_segments.entry(coordinates).or_default().push((
                            CellContentSegment::Opaque,
                            worksheet.source_xml[cursor..start].to_vec(),
                        ));
                    }
                    if let Some((_, raw_row)) = current_raw_row.as_mut() {
                        raw_row.extend_from_slice(raw_event);
                    }
                    if let Some((coordinates, mut raw_cell)) = current_raw_cell.take() {
                        raw_cell.extend_from_slice(raw_event);
                        raw_cell_fragments.insert(coordinates, raw_cell);
                    }
                    if let Some(coordinates) = current_row_segment_cell.take()
                        && let Some((row_index, _)) = current_row_content_cursor
                    {
                        row_content_segments
                            .entry(row_index)
                            .or_default()
                            .push(RowContentSegment::Cell(coordinates));
                        current_row_content_cursor = Some((row_index, end));
                    }
                    current_cell_content_cursor = None;
                    current_cell_segment = None;
                    current_template_cell = None;
                }
                _ => {
                    if let Some((coordinates, segment_kind, segment_start, depth)) =
                        current_cell_segment.as_mut()
                    {
                        *depth -= 1;
                        if *depth == 0 {
                            cell_content_segments
                                .entry(*coordinates)
                                .or_default()
                                .push((
                                    *segment_kind,
                                    worksheet.source_xml[*segment_start..end].to_vec(),
                                ));
                            current_cell_content_cursor = Some((*coordinates, end));
                            current_cell_segment = None;
                        }
                    } else if let Some((row_index, opaque_start, depth)) =
                        current_row_opaque.as_mut()
                    {
                        *depth -= 1;
                        if *depth == 0 {
                            row_content_segments.entry(*row_index).or_default().push(
                                RowContentSegment::Opaque(
                                    worksheet.source_xml[*opaque_start..end].to_vec(),
                                ),
                            );
                            current_row_content_cursor = Some((*row_index, end));
                            current_row_opaque = None;
                        }
                    }
                    if let Some((_, raw_row)) = current_raw_row.as_mut() {
                        raw_row.extend_from_slice(raw_event);
                    }
                    if let Some((_, raw_cell)) = current_raw_cell.as_mut() {
                        raw_cell.extend_from_slice(raw_event);
                    }
                }
            },
            Ok(Event::Eof) => break,
            Ok(_) => {
                if let Some((_, raw_row)) = current_raw_row.as_mut() {
                    raw_row.extend_from_slice(raw_event);
                }
                if let Some((_, raw_cell)) = current_raw_cell.as_mut() {
                    raw_cell.extend_from_slice(raw_event);
                }
            }
            Err(error) => return Err(xml_error(error)),
        }
        template_buffer.clear();
    }

    let mut reader = Reader::from_reader(Cursor::new(worksheet.source_xml.as_slice()));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    let mut skipping_sheet_data = 0usize;
    let mut skipping_dimension = false;
    let support_part_dimension_coords = collect_support_part_dimension_coords(support_parts);
    let anchored_dirty_cells = support_part_dimension_coords
        .iter()
        .copied()
        .filter(|coordinates| {
            worksheet.dirty_cells.contains(coordinates)
                && !worksheet.cells.contains_key(coordinates)
        })
        .collect::<BTreeSet<_>>();
    let dimension_ref = compute_dimension_ref_with_preserved(
        &worksheet.cells,
        raw_cell_fragments
            .keys()
            .copied()
            .filter(|coordinates| !worksheet.dirty_cells.contains(coordinates))
            .chain(
                support_part_dimension_coords
                    .into_iter()
                    .filter(|coordinates| {
                        !worksheet.dirty_cells.contains(coordinates)
                            || anchored_dirty_cells.contains(coordinates)
                    }),
            ),
    );
    let mut depth = 0usize;
    let mut has_dimension = false;
    let mut has_sheet_data = false;
    let mut inserted_dimension = false;
    let mut inserted_sheet_data = false;
    let mut rows = BTreeMap::<u32, BTreeMap<u32, Option<&CellData>>>::new();
    let dirty_rows = worksheet
        .dirty_cells
        .iter()
        .map(|(row, _)| *row)
        .collect::<std::collections::BTreeSet<_>>();

    for (&(row, col), cell) in &worksheet.cells {
        rows.entry(row).or_default().insert(col, Some(cell));
    }
    for (&(row, col), _) in &raw_cell_fragments {
        if worksheet.dirty_cells.contains(&(row, col)) {
            continue;
        }
        rows.entry(row).or_default().entry(col).or_insert(None);
    }
    for &(row, col) in &anchored_dirty_cells {
        rows.entry(row).or_default().entry(col).or_insert(None);
    }
    for (&row_index, attributes) in &row_templates {
        if attributes.iter().any(|(key, _)| key != "r") {
            rows.entry(row_index).or_default();
        }
    }

    let write_placeholder_cell = |writer: &mut Writer<Cursor<Vec<u8>>>,
                                  row_index: u32,
                                  col_index: u32|
     -> OmResult<()> {
        let reference = cell_reference(row_index, col_index);
        let mut cell_tag = BytesStart::new("c");
        let mut wrote_reference = false;
        if let Some(attributes) = cell_templates.get(&(row_index, col_index)) {
            for (key, value) in attributes {
                match key.as_str() {
                    "r" => {
                        cell_tag.push_attribute(("r", reference.as_str()));
                        wrote_reference = true;
                    }
                    "t" => {}
                    _ => cell_tag.push_attribute((key.as_str(), value.as_str())),
                }
            }
        }
        if !wrote_reference {
            cell_tag.push_attribute(("r", reference.as_str()));
        }
        let opaque_segments = cell_content_segments
            .get(&(row_index, col_index))
            .map(|segments| {
                segments
                    .iter()
                    .filter_map(|(segment_kind, raw_bytes)| {
                        matches!(segment_kind, CellContentSegment::Opaque)
                            .then_some(raw_bytes.as_slice())
                    })
                    .filter(|raw_bytes| raw_bytes.iter().any(|byte| !byte.is_ascii_whitespace()))
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        if opaque_segments.is_empty() {
            writer
                .write_event(Event::Empty(cell_tag))
                .map_err(xml_error)?;
            return Ok(());
        }
        writer
            .write_event(Event::Start(cell_tag))
            .map_err(xml_error)?;
        for raw_bytes in opaque_segments {
            writer.get_mut().write_all(raw_bytes).map_err(io_error)?;
        }
        writer
            .write_event(Event::End(BytesEnd::new("c")))
            .map_err(xml_error)?;
        Ok(())
    };

    let write_cell = |writer: &mut Writer<Cursor<Vec<u8>>>,
                      row_index: u32,
                      col_index: u32,
                      cell: &CellData|
     -> OmResult<()> {
        let reference = cell_reference(row_index, col_index);
        let style = cell.style_id.map(|style_id| style_id.0.to_string());
        let cell_type = if cell.formula.is_some() {
            match cell.value {
                CellValue::Bool(_) => Some("b"),
                CellValue::Text(_) => Some("str"),
                CellValue::Error(_) => Some("e"),
                _ => None,
            }
        } else {
            match cell.value {
                CellValue::Bool(_) => Some("b"),
                CellValue::Text(_) => Some("inlineStr"),
                CellValue::Error(_) => Some("e"),
                _ => None,
            }
        };
        let mut cell_tag = BytesStart::new("c");
        let mut wrote_reference = false;
        let mut wrote_style = false;
        let mut wrote_type = false;
        if let Some(attributes) = cell_templates.get(&(row_index, col_index)) {
            for (key, value) in attributes {
                match key.as_str() {
                    "r" => {
                        cell_tag.push_attribute(("r", reference.as_str()));
                        wrote_reference = true;
                    }
                    "s" => {
                        if let Some(style) = style.as_deref() {
                            cell_tag.push_attribute(("s", style));
                            wrote_style = true;
                        }
                    }
                    "t" => {
                        if let Some(cell_type) = cell_type {
                            cell_tag.push_attribute(("t", cell_type));
                            wrote_type = true;
                        }
                    }
                    _ => cell_tag.push_attribute((key.as_str(), value.as_str())),
                }
            }
        }
        if !wrote_reference {
            cell_tag.push_attribute(("r", reference.as_str()));
        }
        if !wrote_style && let Some(style) = style.as_deref() {
            cell_tag.push_attribute(("s", style));
        }
        if !wrote_type && let Some(cell_type) = cell_type {
            cell_tag.push_attribute(("t", cell_type));
        }

        let is_empty_cell = cell.formula.is_none()
            && matches!(cell.value, CellValue::Blank)
            && cell.style_id.is_some();
        if is_empty_cell {
            writer
                .write_event(Event::Empty(cell_tag))
                .map_err(xml_error)?;
            return Ok(());
        }

        writer
            .write_event(Event::Start(cell_tag))
            .map_err(xml_error)?;

        let mut wrote_formula = false;
        let mut wrote_value = false;
        let mut wrote_inline_string = false;
        if let Some(segments) = cell_content_segments.get(&(row_index, col_index)) {
            for (segment_kind, raw_bytes) in segments {
                match segment_kind {
                    CellContentSegment::Opaque => {
                        writer.get_mut().write_all(raw_bytes).map_err(io_error)?;
                    }
                    CellContentSegment::Formula => {
                        if let Some(formula) = &cell.formula {
                            let mut formula_tag = BytesStart::new("f");
                            if let Some(attributes) = formula_templates.get(&(row_index, col_index))
                            {
                                for (key, value) in attributes {
                                    formula_tag.push_attribute((key.as_str(), value.as_str()));
                                }
                            }
                            writer
                                .write_event(Event::Start(formula_tag))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                    formula.text.as_str(),
                                ))))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new("f")))
                                .map_err(xml_error)?;
                            wrote_formula = true;
                        }
                    }
                    CellContentSegment::Value => match &cell.value {
                        CellValue::Blank => {}
                        CellValue::Bool(value) => {
                            writer
                                .write_event(Event::Start(BytesStart::new("v")))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                    if *value { "1" } else { "0" },
                                ))))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new("v")))
                                .map_err(xml_error)?;
                            wrote_value = true;
                        }
                        CellValue::Number(value) => {
                            let value_string = value.to_string();
                            writer
                                .write_event(Event::Start(BytesStart::new("v")))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                    value_string.as_str(),
                                ))))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new("v")))
                                .map_err(xml_error)?;
                            wrote_value = true;
                        }
                        CellValue::Text(value) if cell.formula.is_some() => {
                            writer
                                .write_event(Event::Start(BytesStart::new("v")))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                    value.as_str(),
                                ))))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new("v")))
                                .map_err(xml_error)?;
                            wrote_value = true;
                        }
                        CellValue::Text(value) => {
                            writer
                                .write_event(Event::Start(BytesStart::new("is")))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Start(BytesStart::new("t")))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                    value.as_str(),
                                ))))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new("t")))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new("is")))
                                .map_err(xml_error)?;
                            wrote_inline_string = true;
                        }
                        CellValue::Error(error) => {
                            let value = format_cell_error(*error);
                            writer
                                .write_event(Event::Start(BytesStart::new("v")))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                    value,
                                ))))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new("v")))
                                .map_err(xml_error)?;
                            wrote_value = true;
                        }
                    },
                    CellContentSegment::InlineString => {
                        if let CellValue::Text(value) = &cell.value
                            && cell.formula.is_none()
                        {
                            writer
                                .write_event(Event::Start(BytesStart::new("is")))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Start(BytesStart::new("t")))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                    value.as_str(),
                                ))))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new("t")))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new("is")))
                                .map_err(xml_error)?;
                            wrote_inline_string = true;
                        }
                    }
                }
            }
        }

        if !wrote_formula && let Some(formula) = &cell.formula {
            let mut formula_tag = BytesStart::new("f");
            if let Some(attributes) = formula_templates.get(&(row_index, col_index)) {
                for (key, value) in attributes {
                    formula_tag.push_attribute((key.as_str(), value.as_str()));
                }
            }
            writer
                .write_event(Event::Start(formula_tag))
                .map_err(xml_error)?;
            writer
                .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                    formula.text.as_str(),
                ))))
                .map_err(xml_error)?;
            writer
                .write_event(Event::End(BytesEnd::new("f")))
                .map_err(xml_error)?;
        }
        if !wrote_value && !wrote_inline_string {
            match &cell.value {
                CellValue::Blank => {}
                CellValue::Bool(value) => {
                    writer
                        .write_event(Event::Start(BytesStart::new("v")))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                            if *value { "1" } else { "0" },
                        ))))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::End(BytesEnd::new("v")))
                        .map_err(xml_error)?;
                }
                CellValue::Number(value) => {
                    let value_string = value.to_string();
                    writer
                        .write_event(Event::Start(BytesStart::new("v")))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                            value_string.as_str(),
                        ))))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::End(BytesEnd::new("v")))
                        .map_err(xml_error)?;
                }
                CellValue::Text(value) if cell.formula.is_some() => {
                    writer
                        .write_event(Event::Start(BytesStart::new("v")))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                            value.as_str(),
                        ))))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::End(BytesEnd::new("v")))
                        .map_err(xml_error)?;
                }
                CellValue::Text(value) => {
                    writer
                        .write_event(Event::Start(BytesStart::new("is")))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::Start(BytesStart::new("t")))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                            value.as_str(),
                        ))))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::End(BytesEnd::new("t")))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::End(BytesEnd::new("is")))
                        .map_err(xml_error)?;
                }
                CellValue::Error(error) => {
                    let value = format_cell_error(*error);
                    writer
                        .write_event(Event::Start(BytesStart::new("v")))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::Text(BytesText::from_escaped(partial_escape(value))))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::End(BytesEnd::new("v")))
                        .map_err(xml_error)?;
                }
            }
        }

        writer
            .write_event(Event::End(BytesEnd::new("c")))
            .map_err(xml_error)?;
        Ok(())
    };

    let write_sheet_data = |writer: &mut Writer<Cursor<Vec<u8>>>,
                            element: BytesStart<'static>|
     -> OmResult<()> {
        writer
            .write_event(Event::Start(element))
            .map_err(xml_error)?;

        for (&row_index, cells_in_row) in &rows {
            if !dirty_rows.contains(&row_index) {
                if let Some(raw_row) = raw_row_fragments.get(&row_index) {
                    writer.get_mut().write_all(raw_row).map_err(io_error)?;
                    continue;
                }
            }

            let row_index_string = row_index.to_string();
            let mut row = BytesStart::new("row");
            let mut wrote_row_reference = false;
            if let Some(attributes) = row_templates.get(&row_index) {
                for (key, value) in attributes {
                    if key == "r" {
                        row.push_attribute(("r", row_index_string.as_str()));
                        wrote_row_reference = true;
                        continue;
                    }
                    row.push_attribute((key.as_str(), value.as_str()));
                }
            }
            if !wrote_row_reference {
                row.push_attribute(("r", row_index_string.as_str()));
            }
            writer.write_event(Event::Start(row)).map_err(xml_error)?;

            let mut wrote_existing_cells = std::collections::BTreeSet::new();
            let row_has_original_cells =
                row_content_segments
                    .get(&row_index)
                    .is_some_and(|segments| {
                        segments
                            .iter()
                            .any(|segment| matches!(segment, RowContentSegment::Cell(_)))
                    });
            let original_segment_columns = row_content_segments
                .get(&row_index)
                .map(|segments| {
                    segments
                        .iter()
                        .filter_map(|segment| match segment {
                            RowContentSegment::Opaque(_) => None,
                            RowContentSegment::Cell((_, col_index)) => Some(*col_index),
                        })
                        .collect::<std::collections::BTreeSet<_>>()
                })
                .unwrap_or_default();
            let pending_new_cells = cells_in_row
                .iter()
                .filter_map(|(&col_index, cell)| {
                    if original_segment_columns.contains(&col_index) {
                        return None;
                    }
                    match cell {
                        Some(cell) => Some((col_index, Some(*cell))),
                        None if anchored_dirty_cells.contains(&(row_index, col_index)) => {
                            Some((col_index, None))
                        }
                        None => None,
                    }
                })
                .collect::<Vec<_>>();
            let mut wrote_pending_cells = std::collections::BTreeSet::new();
            let mut next_pending_cell = 0usize;
            let mut write_pending_cells =
                |writer: &mut Writer<Cursor<Vec<u8>>>, before_col: Option<u32>| -> OmResult<()> {
                    while let Some((col_index, cell)) =
                        pending_new_cells.get(next_pending_cell).copied()
                    {
                        if let Some(before_col) = before_col
                            && col_index >= before_col
                        {
                            break;
                        }
                        wrote_pending_cells.insert(col_index);
                        if let Some(cell) = cell {
                            write_cell(writer, row_index, col_index, cell)?;
                        } else {
                            write_placeholder_cell(writer, row_index, col_index)?;
                        }
                        next_pending_cell += 1;
                    }
                    Ok(())
                };
            let mut saw_original_row_cell = false;
            if let Some(segments) = row_content_segments.get(&row_index) {
                for (segment_index, segment) in segments.iter().enumerate() {
                    match segment {
                        RowContentSegment::Opaque(raw_bytes) => {
                            if saw_original_row_cell || !row_has_original_cells {
                                let next_original_col = segments[segment_index + 1..]
                                    .iter()
                                    .find_map(|segment| match segment {
                                        RowContentSegment::Opaque(_) => None,
                                        RowContentSegment::Cell((_, col_index)) => Some(*col_index),
                                    });
                                write_pending_cells(writer, next_original_col)?;
                            }
                            writer.get_mut().write_all(raw_bytes).map_err(io_error)?;
                        }
                        RowContentSegment::Cell((_, col_index)) => {
                            write_pending_cells(writer, Some(*col_index))?;
                            let key = (row_index, *col_index);
                            wrote_existing_cells.insert(key);
                            if !worksheet.dirty_cells.contains(&key)
                                && let Some(raw_cell) = raw_cell_fragments.get(&key)
                            {
                                writer.get_mut().write_all(raw_cell).map_err(io_error)?;
                                saw_original_row_cell = true;
                                continue;
                            }
                            let Some(cell) = cells_in_row.get(col_index).and_then(|cell| *cell)
                            else {
                                if anchored_dirty_cells.contains(&key) {
                                    write_placeholder_cell(writer, row_index, *col_index)?;
                                }
                                saw_original_row_cell = true;
                                continue;
                            };
                            write_cell(writer, row_index, *col_index, cell)?;
                            saw_original_row_cell = true;
                        }
                    }
                }
            }
            write_pending_cells(writer, None)?;
            drop(write_pending_cells);

            for (&col_index, cell) in cells_in_row {
                let key = (row_index, col_index);
                if wrote_existing_cells.contains(&key) || wrote_pending_cells.contains(&col_index) {
                    continue;
                }
                if !worksheet.dirty_cells.contains(&key)
                    && let Some(raw_cell) = raw_cell_fragments.get(&key)
                {
                    writer.get_mut().write_all(raw_cell).map_err(io_error)?;
                    continue;
                }
                let Some(cell) = *cell else {
                    if anchored_dirty_cells.contains(&key) {
                        write_placeholder_cell(writer, row_index, col_index)?;
                    }
                    continue;
                };
                write_cell(writer, row_index, col_index, cell)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("row")))
                .map_err(xml_error)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("sheetData")))
            .map_err(xml_error)?;
        Ok(())
    };

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element))
                if skipping_sheet_data > 0
                    || (skipping_dimension && element.name().as_ref() == b"dimension") =>
            {
                if skipping_sheet_data > 0 {
                    skipping_sheet_data += 1;
                }
                depth += 1;
            }
            Ok(Event::Empty(_)) if skipping_sheet_data > 0 || skipping_dimension => {}
            Ok(Event::End(element)) if skipping_sheet_data > 0 => {
                if skipping_sheet_data == 1 && element.name().as_ref() == b"sheetData" {
                    skipping_sheet_data = 0;
                } else {
                    skipping_sheet_data -= 1;
                }
                depth -= 1;
            }
            Ok(Event::End(element))
                if skipping_dimension && element.name().as_ref() == b"dimension" =>
            {
                skipping_dimension = false;
                depth -= 1;
            }
            Ok(_) if skipping_sheet_data > 0 || skipping_dimension => {}
            Ok(Event::Empty(element)) if element.name().as_ref() == b"dimension" => {
                has_dimension = true;
                inserted_dimension = true;
                let mut dimension = BytesStart::new("dimension");
                dimension.push_attribute(("ref", dimension_ref.as_str()));
                writer
                    .write_event(Event::Empty(dimension))
                    .map_err(xml_error)?;
            }
            Ok(Event::Start(element)) if element.name().as_ref() == b"dimension" => {
                has_dimension = true;
                inserted_dimension = true;
                let mut dimension = BytesStart::new("dimension");
                dimension.push_attribute(("ref", dimension_ref.as_str()));
                writer
                    .write_event(Event::Empty(dimension))
                    .map_err(xml_error)?;
                skipping_dimension = true;
                depth += 1;
            }
            Ok(Event::Empty(element)) if element.name().as_ref() == b"sheetData" => {
                has_sheet_data = true;
                inserted_sheet_data = true;
                write_sheet_data(&mut writer, BytesStart::new("sheetData"))?;
            }
            Ok(Event::Start(element)) if element.name().as_ref() == b"sheetData" => {
                has_sheet_data = true;
                inserted_sheet_data = true;
                write_sheet_data(&mut writer, element.to_owned())?;
                skipping_sheet_data = 1;
                depth += 1;
            }
            Ok(Event::Empty(element)) => {
                if depth == 1 && !inserted_dimension && element.name().as_ref() != b"dimension" {
                    let mut dimension = BytesStart::new("dimension");
                    dimension.push_attribute(("ref", dimension_ref.as_str()));
                    writer
                        .write_event(Event::Empty(dimension))
                        .map_err(xml_error)?;
                    inserted_dimension = true;
                }
                if depth == 1
                    && !inserted_sheet_data
                    && element.name().as_ref() != b"sheetData"
                    && !matches!(
                        element.name().as_ref(),
                        b"sheetPr" | b"dimension" | b"sheetViews" | b"sheetFormatPr" | b"cols"
                    )
                {
                    write_sheet_data(&mut writer, BytesStart::new("sheetData"))?;
                    inserted_sheet_data = true;
                }
                writer
                    .write_event(Event::Empty(element.to_owned()))
                    .map_err(xml_error)?;
            }
            Ok(Event::Start(element)) => {
                if depth == 1 && !inserted_dimension && element.name().as_ref() != b"dimension" {
                    let mut dimension = BytesStart::new("dimension");
                    dimension.push_attribute(("ref", dimension_ref.as_str()));
                    writer
                        .write_event(Event::Empty(dimension))
                        .map_err(xml_error)?;
                    inserted_dimension = true;
                }
                if depth == 1
                    && !inserted_sheet_data
                    && element.name().as_ref() != b"sheetData"
                    && !matches!(
                        element.name().as_ref(),
                        b"sheetPr" | b"dimension" | b"sheetViews" | b"sheetFormatPr" | b"cols"
                    )
                {
                    write_sheet_data(&mut writer, BytesStart::new("sheetData"))?;
                    inserted_sheet_data = true;
                }
                writer
                    .write_event(Event::Start(element.to_owned()))
                    .map_err(xml_error)?;
                depth += 1;
            }
            Ok(Event::End(element)) => {
                if depth == 1 && element.name().as_ref() == b"worksheet" {
                    if !has_dimension && !inserted_dimension {
                        let mut dimension = BytesStart::new("dimension");
                        dimension.push_attribute(("ref", dimension_ref.as_str()));
                        writer
                            .write_event(Event::Empty(dimension))
                            .map_err(xml_error)?;
                        inserted_dimension = true;
                    }
                    if !has_sheet_data && !inserted_sheet_data {
                        write_sheet_data(&mut writer, BytesStart::new("sheetData"))?;
                        inserted_sheet_data = true;
                    }
                }
                writer
                    .write_event(Event::End(element.to_owned()))
                    .map_err(xml_error)?;
                depth -= 1;
            }
            Ok(Event::Eof) => {
                writer.write_event(Event::Eof).map_err(xml_error)?;
                break;
            }
            Err(error) => return Err(xml_error(error)),
            Ok(event) => writer.write_event(event.into_owned()).map_err(xml_error)?,
        }
        buffer.clear();
    }

    Ok(writer.into_inner().into_inner())
}

#[cfg_attr(not(test), allow(dead_code))]
pub(crate) fn compute_dimension_ref(cells: &BTreeMap<(u32, u32), CellData>) -> String {
    compute_dimension_ref_with_preserved(cells, std::iter::empty())
}

pub(crate) fn compute_dimension_ref_with_preserved(
    cells: &BTreeMap<(u32, u32), CellData>,
    preserved_coords: impl IntoIterator<Item = (u32, u32)>,
) -> String {
    let mut coordinates = cells.keys().copied().collect::<Vec<_>>();
    coordinates.extend(preserved_coords);

    let Some(&(mut min_row, mut min_col)) = coordinates.first() else {
        return "A1".to_string();
    };
    let (mut max_row, mut max_col) = (min_row, min_col);

    for (row, col) in coordinates {
        min_row = min_row.min(row);
        min_col = min_col.min(col);
        max_row = max_row.max(row);
        max_col = max_col.max(col);
    }

    if min_row == max_row && min_col == max_col {
        return cell_reference(min_row, min_col);
    }

    format!(
        "{}:{}",
        cell_reference(min_row, min_col),
        cell_reference(max_row, max_col)
    )
}

pub(crate) fn cell_reference(row: u32, col: u32) -> String {
    let mut col_index = col;
    let mut label = String::new();
    while col_index > 0 {
        let remainder = ((col_index - 1) % 26) as u8;
        label.insert(0, (b'A' + remainder) as char);
        col_index = (col_index - 1) / 26;
    }
    format!("{label}{row}")
}

pub(crate) fn parse_cell_error(value: &str) -> CellError {
    match value {
        "#NULL!" => CellError::Null,
        "#DIV/0!" => CellError::Div0,
        "#VALUE!" => CellError::Value,
        "#REF!" => CellError::Ref,
        "#NAME?" => CellError::Name,
        "#NUM!" => CellError::Num,
        "#N/A" => CellError::NA,
        "#GETTING_DATA" => CellError::GettingData,
        "#SPILL!" => CellError::Spill,
        "#CALC!" => CellError::Calc,
        "#FIELD!" => CellError::Field,
        "#BLOCKED!" => CellError::Blocked,
        "#BUSY!" => CellError::Busy,
        "#CONNECT!" => CellError::Connect,
        "#PYTHON!" => CellError::Python,
        "#TIMEOUT!" => CellError::Timeout,
        _ => CellError::Unknown,
    }
}

pub(crate) fn format_cell_error(value: CellError) -> &'static str {
    match value {
        CellError::Null => "#NULL!",
        CellError::Div0 => "#DIV/0!",
        CellError::Value => "#VALUE!",
        CellError::Ref => "#REF!",
        CellError::Name => "#NAME?",
        CellError::Num => "#NUM!",
        CellError::NA => "#N/A",
        CellError::GettingData => "#GETTING_DATA",
        CellError::Spill => "#SPILL!",
        CellError::Calc => "#CALC!",
        CellError::Field => "#FIELD!",
        CellError::Blocked => "#BLOCKED!",
        CellError::Busy => "#BUSY!",
        CellError::Connect => "#CONNECT!",
        CellError::Python => "#PYTHON!",
        CellError::Timeout => "#TIMEOUT!",
        CellError::Unknown => "#UNKNOWN!",
    }
}
