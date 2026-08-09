use super::super::{
    WorksheetSupportParts, io_error,
    xml::{expanded_name_is, qualified_name_like, resolved_element_is, unqualified_attribute_is},
    xml_error,
};

use excel_model::{CellData, WorksheetData, WorksheetStructuralOwners};
use office_common::{
    CellError, CellValue, ExcelLimits, FormulaSource, OmError, OmErrorCode, OmResult, Rect, StyleId,
};
use quick_xml::escape::partial_escape;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{NsReader, Writer};
use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Write};

const EXCEL_MAX_ROW_INDEX: u32 = ExcelLimits::MAX_ROW_INDEX;
const EXCEL_MAX_COLUMN_INDEX: u32 = ExcelLimits::MAX_COLUMN_INDEX;

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
    let bytes = reference.as_bytes();
    let mut index = 0usize;
    if bytes.get(index) == Some(&b'$') {
        index += 1;
    }

    let column_start = index;
    let mut col = 0u32;
    while let Some(byte) = bytes.get(index).copied()
        && byte.is_ascii_alphabetic()
    {
        let value = u32::from(byte.to_ascii_uppercase() - b'A' + 1);
        col = col
            .checked_mul(26)
            .and_then(|current| current.checked_add(value))
            .ok_or_else(|| {
                OmError::parse(format!(
                    "worksheet cell reference exceeds Excel grid XFD1048576: {reference}"
                ))
            })?;
        if col > EXCEL_MAX_COLUMN_INDEX {
            return Err(OmError::parse(format!(
                "worksheet cell reference exceeds Excel grid XFD1048576: {reference}"
            )));
        }
        index += 1;
    }
    if index == column_start {
        return Err(OmError::new(
            OmErrorCode::Parse,
            format!("invalid worksheet cell reference: {reference}"),
        ));
    }

    if index == bytes.len() {
        let row = current_row.ok_or_else(|| {
            OmError::parse(format!("invalid worksheet cell reference: {reference}"))
        })?;
        if row == 0 || row > EXCEL_MAX_ROW_INDEX {
            return Err(OmError::parse(format!(
                "worksheet cell reference exceeds Excel grid XFD1048576: {reference}"
            )));
        }
        return Ok((row, col));
    }

    if bytes.get(index) == Some(&b'$') {
        index += 1;
    }
    let row_start = index;
    let mut row = 0u32;
    while let Some(byte) = bytes.get(index).copied()
        && byte.is_ascii_digit()
    {
        row = row
            .checked_mul(10)
            .and_then(|current| current.checked_add(u32::from(byte - b'0')))
            .ok_or_else(|| {
                OmError::parse(format!(
                    "worksheet cell reference exceeds Excel grid XFD1048576: {reference}"
                ))
            })?;
        if row > EXCEL_MAX_ROW_INDEX {
            return Err(OmError::parse(format!(
                "worksheet cell reference exceeds Excel grid XFD1048576: {reference}"
            )));
        }
        index += 1;
    }
    if index == row_start || index != bytes.len() || row == 0 {
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
    spreadsheet_namespace: &str,
    worksheet_part_uri: &str,
) -> OmResult<ParsedWorksheetCells> {
    let mut reader = NsReader::from_reader(Cursor::new(worksheet_xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut cells = BTreeMap::new();
    let mut dynamic_array_formulas = BTreeSet::new();
    let mut spill_ranges = BTreeMap::new();
    let mut seen_cells = BTreeSet::new();
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
        Option<Rect>,
    )> = None;

    loop {
        match reader.read_resolved_event_into(&mut buffer) {
            Ok((namespace, Event::Start(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"row",
                ) =>
            {
                let mut row_index = None;
                for attr in element.attributes() {
                    let attr = attr.map_err(xml_error)?;
                    let value = attr
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(xml_error)?
                        .into_owned();
                    if unqualified_attribute_is(reader.resolver(), attr.key, b"r") {
                        let parsed = value.parse::<u32>().map_err(|_| {
                            OmError::parse(format!(
                                "{worksheet_part_uri}: invalid worksheet row index: {value}"
                            ))
                        })?;
                        if parsed == 0 || parsed > EXCEL_MAX_ROW_INDEX {
                            return Err(OmError::parse(format!(
                                "{worksheet_part_uri}: invalid worksheet row index: {value}"
                            )));
                        }
                        row_index = Some(parsed);
                    }
                }
                current_row = row_index;
            }
            Ok((namespace, Event::End(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"row",
                ) =>
            {
                current_row = None;
            }
            Ok((namespace, Event::Start(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"c",
                ) =>
            {
                let mut reference = None;
                let mut cell_type = None;
                let mut style_index = None;
                for attr in element.attributes() {
                    let attr = attr.map_err(xml_error)?;
                    let value = attr
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(xml_error)?
                        .into_owned();
                    if unqualified_attribute_is(reader.resolver(), attr.key, b"r") {
                        reference = Some(value);
                    } else if unqualified_attribute_is(reader.resolver(), attr.key, b"t") {
                        cell_type = Some(value);
                    } else if unqualified_attribute_is(reader.resolver(), attr.key, b"s") {
                        style_index = Some(value);
                    }
                }
                let reference = reference.ok_or_else(|| {
                    OmError::new(
                        OmErrorCode::Parse,
                        format!("{worksheet_part_uri}: worksheet cell is missing an A1 reference"),
                    )
                })?;
                let (row, col) =
                    parse_cell_reference(reference.as_str(), current_row).map_err(|error| {
                        OmError::new(
                            error.code,
                            format!("{worksheet_part_uri}: {}", error.message),
                        )
                    })?;
                if let Some(row_index) = current_row
                    && row != row_index
                {
                    return Err(OmError::parse(format!(
                        "{worksheet_part_uri}: cell {reference} is contained by row {row_index}"
                    )));
                }
                let style_id = style_index
                    .map(|value| {
                        value.parse::<u64>().map(StyleId).map_err(|_| {
                            OmError::parse(format!(
                                "{worksheet_part_uri}: cell {reference} has invalid style index: {value}"
                            ))
                        })
                    })
                    .transpose()?;
                if !seen_cells.insert((row, col)) {
                    return Err(OmError::parse(format!(
                        "{worksheet_part_uri}: duplicate worksheet cell coordinate: {reference}"
                    )));
                }
                current_cell = Some((
                    row,
                    col,
                    cell_type,
                    style_id,
                    String::new(),
                    String::new(),
                    String::new(),
                    None,
                ));
            }
            Ok((namespace, Event::Empty(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"c",
                ) =>
            {
                let mut reference = None;
                let mut style_index = None;
                for attr in element.attributes() {
                    let attr = attr.map_err(xml_error)?;
                    let value = attr
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(xml_error)?
                        .into_owned();
                    if unqualified_attribute_is(reader.resolver(), attr.key, b"r") {
                        reference = Some(value);
                    } else if unqualified_attribute_is(reader.resolver(), attr.key, b"s") {
                        style_index = Some(value);
                    }
                }
                let reference = reference.ok_or_else(|| {
                    OmError::new(
                        OmErrorCode::Parse,
                        format!("{worksheet_part_uri}: worksheet cell is missing an A1 reference"),
                    )
                })?;
                let (row, col) =
                    parse_cell_reference(reference.as_str(), current_row).map_err(|error| {
                        OmError::new(
                            error.code,
                            format!("{worksheet_part_uri}: {}", error.message),
                        )
                    })?;
                if let Some(row_index) = current_row
                    && row != row_index
                {
                    return Err(OmError::parse(format!(
                        "{worksheet_part_uri}: cell {reference} is contained by row {row_index}"
                    )));
                }
                let style_id = style_index
                    .map(|value| {
                        value.parse::<u64>().map(StyleId).map_err(|_| {
                            OmError::parse(format!(
                                "{worksheet_part_uri}: cell {reference} has invalid style index: {value}"
                            ))
                        })
                    })
                    .transpose()?;
                if !seen_cells.insert((row, col)) {
                    return Err(OmError::parse(format!(
                        "{worksheet_part_uri}: duplicate worksheet cell coordinate: {reference}"
                    )));
                }
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
            Ok((namespace, Event::Start(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"f",
                ) =>
            {
                let mut formula_type = None;
                let mut formula_reference = None;
                for attr in element.attributes() {
                    let attr = attr.map_err(xml_error)?;
                    let value = attr
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(xml_error)?
                        .into_owned();
                    if unqualified_attribute_is(reader.resolver(), attr.key, b"t") {
                        formula_type = Some(value);
                    } else if unqualified_attribute_is(reader.resolver(), attr.key, b"ref") {
                        formula_reference = Some(value);
                    }
                }
                if formula_type.as_deref() == Some("array") {
                    let formula_reference = formula_reference.ok_or_else(|| {
                        OmError::new(
                            OmErrorCode::Parse,
                            "array formula is missing its spill range reference",
                        )
                    })?;
                    let normalized = formula_reference.replace('$', "");
                    let (first, last) = normalized.split_once(':').map_or(
                        (normalized.as_str(), normalized.as_str()),
                        |(first, last)| (first, last),
                    );
                    let (row_first, col_first) =
                        parse_cell_reference(first, None).map_err(|error| {
                            OmError::new(
                                error.code,
                                format!("{worksheet_part_uri}: {}", error.message),
                            )
                        })?;
                    let (row_last, col_last) =
                        parse_cell_reference(last, None).map_err(|error| {
                            OmError::new(
                                error.code,
                                format!("{worksheet_part_uri}: {}", error.message),
                            )
                        })?;
                    if row_first > row_last || col_first > col_last {
                        return Err(OmError::new(
                            OmErrorCode::Parse,
                            format!("invalid array formula spill range: {formula_reference}"),
                        ));
                    }
                    let spill_range = Rect {
                        row_first,
                        row_last,
                        col_first,
                        col_last,
                    };
                    let Some((row, col, _, _, _, _, _, current_spill_range)) =
                        current_cell.as_mut()
                    else {
                        return Err(OmError::new(
                            OmErrorCode::Parse,
                            "array formula is not contained by a worksheet cell",
                        ));
                    };
                    if (*row, *col) != (row_first, col_first) {
                        return Err(OmError::new(
                            OmErrorCode::Parse,
                            format!(
                                "array formula anchor {} is not the top-left of {formula_reference}",
                                cell_reference(*row, *col),
                            ),
                        ));
                    }
                    *current_spill_range = Some(spill_range);
                }
                current_field = Some("formula");
            }
            Ok((namespace, Event::Start(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"v",
                ) =>
            {
                current_field = Some("value");
            }
            Ok((namespace, Event::Start(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"t",
                ) =>
            {
                if let Some((_, _, cell_type, _, _, _, _, _)) = &current_cell {
                    if cell_type.as_deref() == Some("inlineStr") {
                        current_field = Some("inline");
                    }
                }
            }
            Ok((namespace, Event::End(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"f",
                ) || resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"v",
                ) || resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"t",
                ) =>
            {
                current_field = None;
            }
            Ok((_, Event::Text(text))) => {
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
            Ok((_, Event::CData(text))) => {
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
            Ok((namespace, Event::End(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"c",
                ) =>
            {
                if let Some((row, col, cell_type, style_id, formula, value, inline, spill_range)) =
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
                        _ => {
                            let number = value.parse::<f64>().map_err(xml_error)?;
                            CellValue::try_number(number).map_err(|_| {
                                OmError::parse(format!(
                                    "{worksheet_part_uri}: cell {} numeric value must be finite",
                                    cell_reference(row, col)
                                ))
                            })?
                        }
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
                    if let Some(spill_range) = spill_range {
                        let anchor = (row, col);
                        if spill_ranges.values().any(|existing: &Rect| {
                            existing.row_first <= spill_range.row_last
                                && spill_range.row_first <= existing.row_last
                                && existing.col_first <= spill_range.col_last
                                && spill_range.col_first <= existing.col_last
                        }) {
                            return Err(OmError::new(
                                OmErrorCode::Parse,
                                format!(
                                    "overlapping array formula spill range at {}",
                                    cell_reference(row, col),
                                ),
                            ));
                        }
                        dynamic_array_formulas.insert(anchor);
                        spill_ranges.insert(anchor, spill_range);
                    }
                }
            }
            Ok((_, Event::Eof)) => break,
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }

    let mut spill_owners = BTreeMap::new();
    for &cell in cells.keys() {
        for (&anchor, spill_range) in &spill_ranges {
            if cell != anchor
                && cell.0 >= spill_range.row_first
                && cell.0 <= spill_range.row_last
                && cell.1 >= spill_range.col_first
                && cell.1 <= spill_range.col_last
            {
                spill_owners.insert(cell, anchor);
                break;
            }
        }
    }

    let parse_structural_range = |reference: &str, owner: &str| -> OmResult<Rect> {
        let (first, last) = reference
            .split_once(':')
            .map_or((reference, reference), |(first, last)| (first, last));
        let (row_first, col_first) = parse_cell_reference(first, None).map_err(|error| {
            OmError::new(
                error.code,
                format!("{worksheet_part_uri}: {}", error.message),
            )
        })?;
        let (row_last, col_last) = parse_cell_reference(last, None).map_err(|error| {
            OmError::new(
                error.code,
                format!("{worksheet_part_uri}: {}", error.message),
            )
        })?;
        if row_first > row_last || col_first > col_last {
            return Err(OmError::parse(format!(
                "{worksheet_part_uri}: invalid {owner} range: {reference}"
            )));
        }
        Ok(Rect {
            row_first,
            row_last,
            col_first,
            col_last,
        })
    };
    let parse_merged_range =
        |element: &BytesStart<'_>, decoder: quick_xml::encoding::Decoder| -> OmResult<Rect> {
            let mut reference = None;
            for attr in element.attributes() {
                let attr = attr.map_err(xml_error)?;
                if attr.key.as_ref() == b"ref" {
                    reference = Some(
                        attr.decode_and_unescape_value(decoder)
                            .map_err(xml_error)?
                            .into_owned(),
                    );
                }
            }
            let reference = reference.ok_or_else(|| {
                OmError::parse(format!(
                    "{worksheet_part_uri}: worksheet mergeCell is missing an A1 range reference"
                ))
            })?;
            parse_structural_range(&reference, "merged-cell")
        };
    let parse_data_validation_ranges =
        |element: &BytesStart<'_>, decoder: quick_xml::encoding::Decoder| -> OmResult<Vec<Rect>> {
            let mut sqref = None;
            for attr in element.attributes() {
                let attr = attr.map_err(xml_error)?;
                if attr.key.as_ref() == b"sqref" {
                    sqref = Some(
                        attr.decode_and_unescape_value(decoder)
                            .map_err(xml_error)?
                            .into_owned(),
                    );
                }
            }
            let sqref = sqref.ok_or_else(|| {
                OmError::parse(format!(
                    "{worksheet_part_uri}: worksheet dataValidation is missing an sqref range list"
                ))
            })?;
            let ranges = sqref
                .split_ascii_whitespace()
                .map(|reference| parse_structural_range(reference, "data-validation"))
                .collect::<OmResult<Vec<_>>>()?;
            if ranges.is_empty() {
                return Err(OmError::parse(format!(
                    "{worksheet_part_uri}: worksheet dataValidation has an empty sqref range list"
                )));
            }
            Ok(ranges)
        };
    let mut metadata_reader = NsReader::from_reader(Cursor::new(worksheet_xml));
    metadata_reader.config_mut().trim_text(false);
    let mut metadata_buffer = Vec::new();
    let mut element_depth = 0usize;
    let mut merge_cells_depth = None;
    let mut data_validations_depth = None;
    let mut data_validation_depth = None;
    let mut data_validation_formula_depth = None;
    let mut current_data_validation_formula = None;
    let mut merged_ranges = Vec::new();
    let mut data_validation_ranges = Vec::new();
    let mut data_validation_formulas = Vec::new();
    loop {
        match metadata_reader.read_resolved_event_into(&mut metadata_buffer) {
            Ok((namespace, Event::Start(element))) => {
                let is_merge_cells = resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"mergeCells",
                );
                let is_data_validations = resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"dataValidations",
                );
                if element_depth == 1 && is_merge_cells {
                    merge_cells_depth = Some(element_depth + 1);
                } else if element_depth == 1 && is_data_validations {
                    data_validations_depth = Some(element_depth + 1);
                } else if merge_cells_depth == Some(element_depth)
                    && resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"mergeCell",
                    )
                {
                    merged_ranges.push(parse_merged_range(&element, metadata_reader.decoder())?);
                } else if data_validations_depth == Some(element_depth)
                    && resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"dataValidation",
                    )
                {
                    data_validation_ranges.extend(parse_data_validation_ranges(
                        &element,
                        metadata_reader.decoder(),
                    )?);
                    data_validation_depth = Some(element_depth + 1);
                } else if data_validation_depth == Some(element_depth)
                    && (resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"formula1",
                    ) || resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"formula2",
                    ))
                {
                    data_validation_formula_depth = Some(element_depth + 1);
                    current_data_validation_formula = Some(String::new());
                } else if data_validation_formula_depth == Some(element_depth) {
                    return Err(OmError::parse(format!(
                        "{worksheet_part_uri}: worksheet dataValidation formula contains nested XML"
                    )));
                }
                element_depth += 1;
            }
            Ok((namespace, Event::Empty(element))) => {
                if merge_cells_depth == Some(element_depth)
                    && resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"mergeCell",
                    )
                {
                    merged_ranges.push(parse_merged_range(&element, metadata_reader.decoder())?);
                } else if data_validations_depth == Some(element_depth)
                    && resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"dataValidation",
                    )
                {
                    data_validation_ranges.extend(parse_data_validation_ranges(
                        &element,
                        metadata_reader.decoder(),
                    )?);
                } else if data_validation_depth == Some(element_depth)
                    && (resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"formula1",
                    ) || resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"formula2",
                    ))
                {
                    data_validation_formulas.push(String::new());
                } else if data_validation_formula_depth == Some(element_depth) {
                    return Err(OmError::parse(format!(
                        "{worksheet_part_uri}: worksheet dataValidation formula contains nested XML"
                    )));
                }
            }
            Ok((_, Event::Text(text))) => {
                if data_validation_formula_depth == Some(element_depth)
                    && let Some(formula) = current_data_validation_formula.as_mut()
                {
                    formula.push_str(&text.xml_content().map_err(xml_error)?);
                }
            }
            Ok((_, Event::CData(text))) => {
                if data_validation_formula_depth == Some(element_depth)
                    && let Some(formula) = current_data_validation_formula.as_mut()
                {
                    formula.push_str(&text.xml_content().map_err(xml_error)?);
                }
            }
            Ok((namespace, Event::End(element))) => {
                if data_validation_formula_depth == Some(element_depth)
                    && (resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"formula1",
                    ) || resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"formula2",
                    ))
                {
                    data_validation_formulas
                        .push(current_data_validation_formula.take().unwrap_or_default());
                    data_validation_formula_depth = None;
                }
                if data_validation_depth == Some(element_depth)
                    && resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"dataValidation",
                    )
                {
                    data_validation_depth = None;
                }
                if merge_cells_depth == Some(element_depth)
                    && resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"mergeCells",
                    )
                {
                    merge_cells_depth = None;
                }
                if data_validations_depth == Some(element_depth)
                    && resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"dataValidations",
                    )
                {
                    data_validations_depth = None;
                }
                element_depth = element_depth.saturating_sub(1);
            }
            Ok((_, Event::Eof)) => break,
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        metadata_buffer.clear();
    }

    Ok(ParsedWorksheetCells {
        cells,
        dynamic_array_formulas,
        spill_ranges,
        spill_owners,
        structural_owners: WorksheetStructuralOwners {
            merged_ranges,
            data_validation_ranges,
            data_validation_formulas,
        },
    })
}

#[derive(Debug)]
pub(crate) struct ParsedWorksheetCells {
    pub(crate) cells: BTreeMap<(u32, u32), CellData>,
    pub(crate) dynamic_array_formulas: BTreeSet<(u32, u32)>,
    pub(crate) spill_ranges: BTreeMap<(u32, u32), Rect>,
    pub(crate) spill_owners: BTreeMap<(u32, u32), (u32, u32)>,
    pub(crate) structural_owners: WorksheetStructuralOwners,
}

pub(crate) fn rewrite_worksheet_xml(
    worksheet: &WorksheetData,
    support_parts: Option<&WorksheetSupportParts>,
    spreadsheet_namespace: &str,
) -> OmResult<Vec<u8>> {
    if worksheet.source_xml.is_empty() {
        return Err(OmError::new(
            OmErrorCode::InvalidState,
            "worksheet source xml is missing",
        ));
    }
    for (&(row, col), cell) in &worksheet.cells {
        if cell.value.validate().is_err() {
            return Err(OmError::invalid_state(format!(
                "worksheet cell {} numeric value must be finite",
                cell_reference(row, col)
            )));
        }
    }

    let mut row_templates = BTreeMap::<u32, Vec<(String, String)>>::new();
    let mut cell_templates = BTreeMap::<(u32, u32), Vec<(String, String)>>::new();
    let mut formula_templates = BTreeMap::<(u32, u32), Vec<(String, String)>>::new();
    let mut row_element_names = BTreeMap::<u32, Vec<u8>>::new();
    let mut cell_element_names = BTreeMap::<(u32, u32), Vec<u8>>::new();
    let mut worksheet_element_name = None::<Vec<u8>>;
    let mut raw_row_fragments = BTreeMap::<u32, Vec<u8>>::new();
    let mut raw_cell_fragments = BTreeMap::<(u32, u32), Vec<u8>>::new();
    let mut row_content_segments = BTreeMap::<u32, Vec<RowContentSegment>>::new();
    let mut cell_content_segments =
        BTreeMap::<(u32, u32), Vec<(CellContentSegment, Vec<u8>)>>::new();
    let mut template_reader = NsReader::from_reader(Cursor::new(worksheet.source_xml.as_slice()));
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
        let decoder = template_reader.decoder();
        let event = template_reader
            .read_resolved_event_into(&mut template_buffer)
            .map(|(namespace, event)| {
                let namespace = match namespace {
                    quick_xml::name::ResolveResult::Bound(namespace) => {
                        Some(namespace.as_ref().to_vec())
                    }
                    quick_xml::name::ResolveResult::Unbound
                    | quick_xml::name::ResolveResult::Unknown(_) => None,
                };
                (namespace, event)
            });
        let end = template_reader.buffer_position() as usize;
        let raw_event = &worksheet.source_xml[start..end];

        match event {
            Ok((namespace, Event::Start(element))) => {
                let local_name = element.local_name();
                match if expanded_name_is(
                    namespace.as_deref(),
                    local_name,
                    spreadsheet_namespace.as_bytes(),
                    local_name.as_ref(),
                ) {
                    local_name.as_ref()
                } else {
                    b""
                } {
                    b"worksheet" => {
                        worksheet_element_name = Some(element.name().as_ref().to_vec());
                    }
                    b"row" => {
                        let attributes = collect_attributes(&element, decoder)?;
                        current_template_row = attributes
                            .iter()
                            .find(|(key, _)| key == "r")
                            .and_then(|(_, value)| value.parse::<u32>().ok());
                        if let Some(row_index) = current_template_row {
                            row_element_names.insert(row_index, element.name().as_ref().to_vec());
                            current_raw_row = Some((row_index, raw_event.to_vec()));
                            current_row_content_cursor = Some((row_index, end));
                            row_content_segments.entry(row_index).or_default();
                            row_templates.insert(row_index, attributes);
                        }
                    }
                    b"c" => {
                        let attributes = collect_attributes(&element, decoder)?;
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
                        cell_element_names.insert(coordinates, element.name().as_ref().to_vec());
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
                            formula_templates
                                .insert(coordinates, collect_attributes(&element, decoder)?);
                            if current_cell_segment.is_none()
                                && let Some((segment_coordinates, cursor)) =
                                    current_cell_content_cursor
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
                                && let Some((segment_coordinates, cursor)) =
                                    current_cell_content_cursor
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
                                && let Some((segment_coordinates, cursor)) =
                                    current_cell_content_cursor
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
                            if let Some((opaque_row_index, _, depth)) = current_row_opaque.as_mut()
                            {
                                if *opaque_row_index == row_index {
                                    *depth += 1;
                                }
                            } else if let Some((cursor_row_index, cursor)) =
                                current_row_content_cursor
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
                }
            }
            Ok((namespace, Event::Empty(element))) => {
                let local_name = element.local_name();
                match if expanded_name_is(
                    namespace.as_deref(),
                    local_name,
                    spreadsheet_namespace.as_bytes(),
                    local_name.as_ref(),
                ) {
                    local_name.as_ref()
                } else {
                    b""
                } {
                    b"row" => {
                        let attributes = collect_attributes(&element, decoder)?;
                        if let Some(row_index) = attributes
                            .iter()
                            .find(|(key, _)| key == "r")
                            .and_then(|(_, value)| value.parse::<u32>().ok())
                        {
                            row_element_names.insert(row_index, element.name().as_ref().to_vec());
                            raw_row_fragments.insert(row_index, raw_event.to_vec());
                            row_content_segments.entry(row_index).or_default();
                            row_templates.insert(row_index, attributes);
                        }
                    }
                    b"c" => {
                        let attributes = collect_attributes(&element, decoder)?;
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
                        cell_element_names.insert(coordinates, element.name().as_ref().to_vec());
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
                                formula_templates
                                    .insert(coordinates, collect_attributes(&element, decoder)?);
                                if let Some((segment_coordinates, cursor)) =
                                    current_cell_content_cursor
                                    && segment_coordinates == coordinates
                                {
                                    if cursor < start {
                                        cell_content_segments.entry(coordinates).or_default().push(
                                            (
                                                CellContentSegment::Opaque,
                                                worksheet.source_xml[cursor..start].to_vec(),
                                            ),
                                        );
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
                                RowContentSegment::Opaque(
                                    worksheet.source_xml[cursor..end].to_vec(),
                                ),
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
                }
            }
            Ok((namespace, Event::End(element))) => {
                let local_name = element.local_name();
                match if expanded_name_is(
                    namespace.as_deref(),
                    local_name,
                    spreadsheet_namespace.as_bytes(),
                    local_name.as_ref(),
                ) {
                    local_name.as_ref()
                } else {
                    b""
                } {
                    b"row" => {
                        if let Some((row_index, cursor)) = current_row_content_cursor
                            && cursor < start
                        {
                            row_content_segments.entry(row_index).or_default().push(
                                RowContentSegment::Opaque(
                                    worksheet.source_xml[cursor..start].to_vec(),
                                ),
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
                }
            }
            Ok((_, Event::Eof)) => break,
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

    let worksheet_element_name = worksheet_element_name.ok_or_else(|| {
        OmError::new(
            OmErrorCode::Parse,
            "worksheet source XML does not contain a SpreadsheetML worksheet root",
        )
    })?;
    let mut reader = NsReader::from_reader(Cursor::new(worksheet.source_xml.as_slice()));
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
    let worksheet_dimension_name = qualified_name_like(&worksheet_element_name, "dimension");
    let worksheet_sheet_data_name = qualified_name_like(&worksheet_element_name, "sheetData");
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
                                  col_index: u32,
                                  element_name_reference: &[u8]|
     -> OmResult<()> {
        let reference = cell_reference(row_index, col_index);
        let cell_name = qualified_name_like(element_name_reference, "c");
        let mut cell_tag = BytesStart::new(cell_name.as_str());
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
            .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
            .map_err(xml_error)?;
        Ok(())
    };

    let write_cell = |writer: &mut Writer<Cursor<Vec<u8>>>,
                      row_index: u32,
                      col_index: u32,
                      cell: &CellData,
                      element_name_reference: &[u8]|
     -> OmResult<()> {
        let reference = cell_reference(row_index, col_index);
        let cell_name = qualified_name_like(element_name_reference, "c");
        let formula_name = qualified_name_like(element_name_reference, "f");
        let value_name = qualified_name_like(element_name_reference, "v");
        let inline_string_name = qualified_name_like(element_name_reference, "is");
        let text_name = qualified_name_like(element_name_reference, "t");
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
        let mut cell_tag = BytesStart::new(cell_name.as_str());
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

        let coordinates = (row_index, col_index);
        let is_dynamic_array_formula = worksheet.dynamic_array_formulas.contains(&coordinates);
        let spill_reference = is_dynamic_array_formula.then(|| {
            let spill_range = worksheet
                .spill_ranges
                .get(&coordinates)
                .copied()
                .unwrap_or(Rect::single_cell(row_index, col_index));
            let first = cell_reference(spill_range.row_first, spill_range.col_first);
            let last = cell_reference(spill_range.row_last, spill_range.col_last);
            if first == last {
                first
            } else {
                format!("{first}:{last}")
            }
        });
        let write_formula_xml =
            |writer: &mut Writer<Cursor<Vec<u8>>>, formula: &FormulaSource| -> OmResult<()> {
                let mut formula_tag = BytesStart::new(formula_name.as_str());
                if let Some(attributes) = formula_templates.get(&coordinates) {
                    for (key, value) in attributes {
                        if is_dynamic_array_formula && matches!(key.as_str(), "t" | "ref" | "si") {
                            continue;
                        }
                        if !is_dynamic_array_formula
                            && worksheet.dirty_cells.contains(&coordinates)
                            && matches!(key.as_str(), "t" | "ref" | "si" | "aca")
                        {
                            continue;
                        }
                        formula_tag.push_attribute((key.as_str(), value.as_str()));
                    }
                }
                if let Some(spill_reference) = spill_reference.as_deref() {
                    formula_tag.push_attribute(("t", "array"));
                    formula_tag.push_attribute(("ref", spill_reference));
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
                    .write_event(Event::End(BytesEnd::new(formula_name.as_str())))
                    .map_err(xml_error)?;
                Ok(())
            };

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
                            write_formula_xml(writer, formula)?;
                            wrote_formula = true;
                        }
                    }
                    CellContentSegment::Value => match &cell.value {
                        CellValue::Blank => {}
                        CellValue::Bool(value) => {
                            writer
                                .write_event(Event::Start(BytesStart::new(value_name.as_str())))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                    if *value { "1" } else { "0" },
                                ))))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new(value_name.as_str())))
                                .map_err(xml_error)?;
                            wrote_value = true;
                        }
                        CellValue::Number(value) => {
                            let value_string = value.to_string();
                            writer
                                .write_event(Event::Start(BytesStart::new(value_name.as_str())))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                    value_string.as_str(),
                                ))))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new(value_name.as_str())))
                                .map_err(xml_error)?;
                            wrote_value = true;
                        }
                        CellValue::Text(value) if cell.formula.is_some() => {
                            writer
                                .write_event(Event::Start(BytesStart::new(value_name.as_str())))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                    value.as_str(),
                                ))))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new(value_name.as_str())))
                                .map_err(xml_error)?;
                            wrote_value = true;
                        }
                        CellValue::Text(value) => {
                            writer
                                .write_event(Event::Start(BytesStart::new(
                                    inline_string_name.as_str(),
                                )))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Start(BytesStart::new(text_name.as_str())))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                    value.as_str(),
                                ))))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new(text_name.as_str())))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new(inline_string_name.as_str())))
                                .map_err(xml_error)?;
                            wrote_inline_string = true;
                        }
                        CellValue::Error(error) => {
                            let value = format_cell_error(*error);
                            writer
                                .write_event(Event::Start(BytesStart::new(value_name.as_str())))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                    value,
                                ))))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new(value_name.as_str())))
                                .map_err(xml_error)?;
                            wrote_value = true;
                        }
                    },
                    CellContentSegment::InlineString => {
                        if let CellValue::Text(value) = &cell.value
                            && cell.formula.is_none()
                        {
                            writer
                                .write_event(Event::Start(BytesStart::new(
                                    inline_string_name.as_str(),
                                )))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Start(BytesStart::new(text_name.as_str())))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                    value.as_str(),
                                ))))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new(text_name.as_str())))
                                .map_err(xml_error)?;
                            writer
                                .write_event(Event::End(BytesEnd::new(inline_string_name.as_str())))
                                .map_err(xml_error)?;
                            wrote_inline_string = true;
                        }
                    }
                }
            }
        }

        if !wrote_formula && let Some(formula) = &cell.formula {
            write_formula_xml(writer, formula)?;
        }
        if !wrote_value && !wrote_inline_string {
            match &cell.value {
                CellValue::Blank => {}
                CellValue::Bool(value) => {
                    writer
                        .write_event(Event::Start(BytesStart::new(value_name.as_str())))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                            if *value { "1" } else { "0" },
                        ))))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::End(BytesEnd::new(value_name.as_str())))
                        .map_err(xml_error)?;
                }
                CellValue::Number(value) => {
                    let value_string = value.to_string();
                    writer
                        .write_event(Event::Start(BytesStart::new(value_name.as_str())))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                            value_string.as_str(),
                        ))))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::End(BytesEnd::new(value_name.as_str())))
                        .map_err(xml_error)?;
                }
                CellValue::Text(value) if cell.formula.is_some() => {
                    writer
                        .write_event(Event::Start(BytesStart::new(value_name.as_str())))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                            value.as_str(),
                        ))))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::End(BytesEnd::new(value_name.as_str())))
                        .map_err(xml_error)?;
                }
                CellValue::Text(value) => {
                    writer
                        .write_event(Event::Start(BytesStart::new(inline_string_name.as_str())))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::Start(BytesStart::new(text_name.as_str())))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                            value.as_str(),
                        ))))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::End(BytesEnd::new(text_name.as_str())))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::End(BytesEnd::new(inline_string_name.as_str())))
                        .map_err(xml_error)?;
                }
                CellValue::Error(error) => {
                    let value = format_cell_error(*error);
                    writer
                        .write_event(Event::Start(BytesStart::new(value_name.as_str())))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::Text(BytesText::from_escaped(partial_escape(value))))
                        .map_err(xml_error)?;
                    writer
                        .write_event(Event::End(BytesEnd::new(value_name.as_str())))
                        .map_err(xml_error)?;
                }
            }
        }

        writer
            .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
            .map_err(xml_error)?;
        Ok(())
    };

    let write_sheet_data = |writer: &mut Writer<Cursor<Vec<u8>>>,
                            element: BytesStart<'static>|
     -> OmResult<()> {
        let sheet_data_element_name = element.name().as_ref().to_vec();
        let sheet_data_name = qualified_name_like(&sheet_data_element_name, "sheetData");
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
            let row_element_name = row_element_names
                .get(&row_index)
                .map(Vec::as_slice)
                .unwrap_or(sheet_data_element_name.as_slice());
            let row_name = qualified_name_like(row_element_name, "row");
            let mut row = BytesStart::new(row_name.as_str());
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
            let mut write_pending_cells = |writer: &mut Writer<Cursor<Vec<u8>>>,
                                           before_col: Option<u32>|
             -> OmResult<()> {
                while let Some((col_index, cell)) =
                    pending_new_cells.get(next_pending_cell).copied()
                {
                    if let Some(before_col) = before_col
                        && col_index >= before_col
                    {
                        break;
                    }
                    wrote_pending_cells.insert(col_index);
                    let cell_element_name = cell_element_names
                        .get(&(row_index, col_index))
                        .map(Vec::as_slice)
                        .unwrap_or(row_element_name);
                    if let Some(cell) = cell {
                        write_cell(writer, row_index, col_index, cell, cell_element_name)?;
                    } else {
                        write_placeholder_cell(writer, row_index, col_index, cell_element_name)?;
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
                                    let cell_element_name = cell_element_names
                                        .get(&key)
                                        .map(Vec::as_slice)
                                        .unwrap_or(row_element_name);
                                    write_placeholder_cell(
                                        writer,
                                        row_index,
                                        *col_index,
                                        cell_element_name,
                                    )?;
                                }
                                saw_original_row_cell = true;
                                continue;
                            };
                            let cell_element_name = cell_element_names
                                .get(&key)
                                .map(Vec::as_slice)
                                .unwrap_or(row_element_name);
                            write_cell(writer, row_index, *col_index, cell, cell_element_name)?;
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
                        let cell_element_name = cell_element_names
                            .get(&key)
                            .map(Vec::as_slice)
                            .unwrap_or(row_element_name);
                        write_placeholder_cell(writer, row_index, col_index, cell_element_name)?;
                    }
                    continue;
                };
                let cell_element_name = cell_element_names
                    .get(&key)
                    .map(Vec::as_slice)
                    .unwrap_or(row_element_name);
                write_cell(writer, row_index, col_index, cell, cell_element_name)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new(row_name.as_str())))
                .map_err(xml_error)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new(sheet_data_name.as_str())))
            .map_err(xml_error)?;
        Ok(())
    };

    loop {
        match reader.read_resolved_event_into(&mut buffer) {
            Ok((namespace, Event::Start(element)))
                if skipping_sheet_data > 0
                    || (skipping_dimension
                        && resolved_element_is(
                            &namespace,
                            element.local_name(),
                            spreadsheet_namespace.as_bytes(),
                            b"dimension",
                        )) =>
            {
                if skipping_sheet_data > 0 {
                    skipping_sheet_data += 1;
                }
                depth += 1;
            }
            Ok((_, Event::Empty(_))) if skipping_sheet_data > 0 || skipping_dimension => {}
            Ok((namespace, Event::End(element))) if skipping_sheet_data > 0 => {
                if skipping_sheet_data == 1
                    && resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"sheetData",
                    )
                {
                    skipping_sheet_data = 0;
                } else {
                    skipping_sheet_data -= 1;
                }
                depth -= 1;
            }
            Ok((namespace, Event::End(element)))
                if skipping_dimension
                    && resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"dimension",
                    ) =>
            {
                skipping_dimension = false;
                depth -= 1;
            }
            Ok(_) if skipping_sheet_data > 0 || skipping_dimension => {}
            Ok((namespace, Event::Empty(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"dimension",
                ) =>
            {
                has_dimension = true;
                inserted_dimension = true;
                let dimension_name = qualified_name_like(element.name().as_ref(), "dimension");
                let mut dimension = BytesStart::new(dimension_name.as_str());
                dimension.push_attribute(("ref", dimension_ref.as_str()));
                writer
                    .write_event(Event::Empty(dimension))
                    .map_err(xml_error)?;
            }
            Ok((namespace, Event::Start(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"dimension",
                ) =>
            {
                has_dimension = true;
                inserted_dimension = true;
                let dimension_name = qualified_name_like(element.name().as_ref(), "dimension");
                let mut dimension = BytesStart::new(dimension_name.as_str());
                dimension.push_attribute(("ref", dimension_ref.as_str()));
                writer
                    .write_event(Event::Empty(dimension))
                    .map_err(xml_error)?;
                skipping_dimension = true;
                depth += 1;
            }
            Ok((namespace, Event::Empty(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"sheetData",
                ) =>
            {
                has_sheet_data = true;
                inserted_sheet_data = true;
                write_sheet_data(&mut writer, element.into_owned())?;
            }
            Ok((namespace, Event::Start(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"sheetData",
                ) =>
            {
                has_sheet_data = true;
                inserted_sheet_data = true;
                write_sheet_data(&mut writer, element.to_owned())?;
                skipping_sheet_data = 1;
                depth += 1;
            }
            Ok((namespace, Event::Empty(element))) => {
                let element_namespace_is_spreadsheet = matches!(
                    namespace,
                    quick_xml::name::ResolveResult::Bound(namespace)
                        if namespace.as_ref() == spreadsheet_namespace.as_bytes()
                );
                let local_name = element.local_name();
                let local_name = element_namespace_is_spreadsheet.then_some(local_name.as_ref());
                if depth == 1 && !inserted_dimension && local_name != Some(b"dimension".as_slice())
                {
                    let mut dimension = BytesStart::new(worksheet_dimension_name.as_str());
                    dimension.push_attribute(("ref", dimension_ref.as_str()));
                    writer
                        .write_event(Event::Empty(dimension))
                        .map_err(xml_error)?;
                    inserted_dimension = true;
                }
                if depth == 1
                    && !inserted_sheet_data
                    && local_name != Some(b"sheetData".as_slice())
                    && !matches!(
                        local_name,
                        Some(
                            b"sheetPr" | b"dimension" | b"sheetViews" | b"sheetFormatPr" | b"cols"
                        )
                    )
                {
                    write_sheet_data(
                        &mut writer,
                        BytesStart::new(worksheet_sheet_data_name.clone()),
                    )?;
                    inserted_sheet_data = true;
                }
                writer
                    .write_event(Event::Empty(element.to_owned()))
                    .map_err(xml_error)?;
            }
            Ok((namespace, Event::Start(element))) => {
                let element_namespace_is_spreadsheet = matches!(
                    namespace,
                    quick_xml::name::ResolveResult::Bound(namespace)
                        if namespace.as_ref() == spreadsheet_namespace.as_bytes()
                );
                let local_name = element.local_name();
                let local_name = element_namespace_is_spreadsheet.then_some(local_name.as_ref());
                if depth == 1 && !inserted_dimension && local_name != Some(b"dimension".as_slice())
                {
                    let mut dimension = BytesStart::new(worksheet_dimension_name.as_str());
                    dimension.push_attribute(("ref", dimension_ref.as_str()));
                    writer
                        .write_event(Event::Empty(dimension))
                        .map_err(xml_error)?;
                    inserted_dimension = true;
                }
                if depth == 1
                    && !inserted_sheet_data
                    && local_name != Some(b"sheetData".as_slice())
                    && !matches!(
                        local_name,
                        Some(
                            b"sheetPr" | b"dimension" | b"sheetViews" | b"sheetFormatPr" | b"cols"
                        )
                    )
                {
                    write_sheet_data(
                        &mut writer,
                        BytesStart::new(worksheet_sheet_data_name.clone()),
                    )?;
                    inserted_sheet_data = true;
                }
                writer
                    .write_event(Event::Start(element.to_owned()))
                    .map_err(xml_error)?;
                depth += 1;
            }
            Ok((namespace, Event::End(element))) => {
                if depth == 1
                    && resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"worksheet",
                    )
                {
                    if !has_dimension && !inserted_dimension {
                        let mut dimension = BytesStart::new(worksheet_dimension_name.as_str());
                        dimension.push_attribute(("ref", dimension_ref.as_str()));
                        writer
                            .write_event(Event::Empty(dimension))
                            .map_err(xml_error)?;
                        inserted_dimension = true;
                    }
                    if !has_sheet_data && !inserted_sheet_data {
                        write_sheet_data(
                            &mut writer,
                            BytesStart::new(worksheet_sheet_data_name.clone()),
                        )?;
                        inserted_sheet_data = true;
                    }
                }
                writer
                    .write_event(Event::End(element.to_owned()))
                    .map_err(xml_error)?;
                depth -= 1;
            }
            Ok((_, Event::Eof)) => {
                writer.write_event(Event::Eof).map_err(xml_error)?;
                break;
            }
            Err(error) => return Err(xml_error(error)),
            Ok((_, event)) => writer.write_event(event.into_owned()).map_err(xml_error)?,
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
