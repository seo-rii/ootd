//! Replaying recorded structural shifts onto the parts a worksheet owns: comment anchors, legacy
//! VML shape anchors, hyperlink snapshots, table ranges, and drawing anchors.

use super::{
    WorksheetSupportParts, parse_comment_part_summary, parse_vml_drawing_part_summary,
    replay_shifts_on_ref, xml_error, xml_local_name,
};
use excel_model::{WorkbookState, WorksheetData};
use office_common::{OmError, OmResult, StructuralAxis, StructuralShift};
use office_opc::OpcPackage;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::collections::BTreeSet;
use std::io::Cursor;

/// Rewrites every `comment@ref` in a comments part. A comment whose cell is deleted cannot be
/// removed yet, so the shift is refused.
pub(crate) fn shift_comment_part(xml: &[u8], shift: StructuralShift) -> OmResult<Vec<u8>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        let event = match event {
            Event::Start(element) if xml_local_name(element.name().as_ref()) == b"comment" => {
                Event::Start(shift_comment_element(&element, shift)?)
            }
            Event::Empty(element) if xml_local_name(element.name().as_ref()) == b"comment" => {
                Event::Empty(shift_comment_element(&element, shift)?)
            }
            Event::Eof => break,
            event => event.into_owned(),
        };
        writer.write_event(event).map_err(xml_error)?;
        buffer.clear();
    }
    Ok(writer.into_inner().into_inner())
}

fn shift_comment_element(
    element: &BytesStart<'_>,
    shift: StructuralShift,
) -> OmResult<BytesStart<'static>> {
    let mut shifted =
        BytesStart::new(String::from_utf8_lossy(element.name().as_ref()).into_owned());
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == b"ref" {
            let reference = attribute.unescape_value().map_err(xml_error)?;
            let moved = replay_shifts_on_ref(&[shift], &reference)?.ok_or_else(|| {
                OmError::unsupported(format!(
                    "structural comment removal is not implemented for comment {reference}"
                ))
            })?;
            shifted.push_attribute(("ref", moved.as_str()));
        } else {
            shifted.push_attribute(attribute);
        }
    }
    Ok(shifted)
}

/// Replays structural shifts onto the cell anchors of every picture, shape, and other non-chart
/// object in a drawing part. Chart frames are skipped because the model moves their anchors.
/// Anchors are edited in place, so all other drawing bytes are preserved; `editAs="absolute"` and
/// `absoluteAnchor` objects do not move with cells.
pub(crate) fn shift_drawing_part(xml: &[u8], shifts: &[StructuralShift]) -> OmResult<Vec<u8>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut output = Vec::with_capacity(xml.len());
    let mut copied = 0usize;
    let mut depth = 0usize;
    let mut anchor_start = None::<usize>;
    loop {
        let before = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        match reader.read_event().map_err(xml_error)? {
            Event::Start(element) => {
                depth += 1;
                if depth == 2
                    && matches!(
                        xml_local_name(element.name().as_ref()),
                        b"twoCellAnchor" | b"oneCellAnchor"
                    )
                {
                    anchor_start = Some(before);
                }
            }
            Event::End(_) => {
                if depth == 2
                    && let Some(start) = anchor_start.take()
                {
                    let end = usize::try_from(reader.buffer_position()).unwrap_or(xml.len());
                    let anchor = std::str::from_utf8(&xml[start..end])
                        .map_err(|_| OmError::parse("drawing anchor is not UTF-8"))?;
                    output.extend_from_slice(&xml[copied..start]);
                    output.extend_from_slice(shift_drawing_anchor(anchor, shifts)?.as_bytes());
                    copied = end;
                }
                depth = depth.saturating_sub(1);
            }
            Event::Eof => break,
            _ => {}
        }
    }
    output.extend_from_slice(&xml[copied..]);
    Ok(output)
}

fn shift_drawing_anchor(anchor: &str, shifts: &[StructuralShift]) -> OmResult<String> {
    if anchor.contains("http://schemas.openxmlformats.org/drawingml/2006/chart\"") {
        return Ok(anchor.to_string());
    }
    let open_end = anchor.find('>').unwrap_or(anchor.len());
    let open_tag = &anchor[..open_end];
    let name_end = open_tag[1..]
        .find(|ch: char| ch.is_whitespace() || ch == '>' || ch == '/')
        .map_or(open_tag.len(), |offset| offset + 1);
    let element_name = &open_tag[1..name_end];
    let prefix = element_name
        .rsplit_once(':')
        .map_or(String::new(), |(prefix, _)| format!("{prefix}:"));
    let one_cell = element_name.ends_with("oneCellAnchor");
    let edit_as = ["editAs=\"", "editAs='"]
        .iter()
        .find_map(|marker| {
            let start = open_tag.find(marker)? + marker.len();
            let end = open_tag[start..].find(['"', '\''])?;
            Some(&open_tag[start..start + end])
        })
        .unwrap_or("twoCell");
    if edit_as == "absolute" {
        return Ok(anchor.to_string());
    }
    let move_only = one_cell || edit_as == "oneCell";
    let marker = |anchor: &str, name: &str| -> Option<(u32, u32)> {
        let open = format!("<{prefix}{name}>");
        let close = format!("</{prefix}{name}>");
        let start = anchor.find(&open)? + open.len();
        let body = &anchor[start..start + anchor[start..].find(&close)?];
        let value = |tag: &str| -> Option<u32> {
            let open = format!("<{prefix}{tag}>");
            let start = body.find(&open)? + open.len();
            body[start..start + body[start..].find('<')?]
                .trim()
                .parse()
                .ok()
        };
        Some((value("col")?, value("row")?))
    };
    let mut anchor = anchor.to_string();
    for &shift in shifts {
        let Some(from) = marker(&anchor, "from") else {
            return Err(OmError::parse("drawing anchor has no from marker"));
        };
        let to = if one_cell {
            None
        } else {
            marker(&anchor, "to")
        };
        let axis = |cell: (u32, u32)| match shift.axis {
            StructuralAxis::Rows => cell.1,
            StructuralAxis::Columns => cell.0,
        };
        let moved = |zero_based: u32| match shift.shift_index(zero_based + 1) {
            Some(moved) => (moved - 1, false),
            None if shift.insert => (shift.axis_max() - 1, false),
            None => (shift.first - 1, true),
        };
        let (new_from, from_collapsed) = moved(axis(from));
        let new_to = to.map(|to| {
            if move_only {
                let span = axis(to).saturating_sub(axis(from));
                ((new_from + span).min(shift.axis_max() - 1), false)
            } else {
                let (value, collapsed) = moved(axis(to));
                (value.max(new_from), collapsed)
            }
        });
        let (tag, offset_tag) = match shift.axis {
            StructuralAxis::Rows => ("row", "rowOff"),
            StructuralAxis::Columns => ("col", "colOff"),
        };
        let mut markers = vec![("from", new_from, from_collapsed)];
        if let Some((value, collapsed)) = new_to {
            markers.push(("to", value, collapsed));
        }
        for (name, value, collapsed) in markers {
            let open = format!("<{prefix}{name}>");
            let close = format!("</{prefix}{name}>");
            let start = anchor.find(&open).map(|start| start + open.len());
            let Some(start) = start else { continue };
            let end = start + anchor[start..].find(&close).unwrap_or(0);
            let mut body =
                replace_tag_text(&anchor[start..end], &format!("{prefix}{tag}"), |_| {
                    Ok(value.to_string())
                })?;
            if collapsed {
                body = replace_tag_text(&body, &format!("{prefix}{offset_tag}"), |_| {
                    Ok("0".to_string())
                })?;
            }
            anchor.replace_range(start..end, &body);
        }
    }
    Ok(anchor)
}

/// Replays structural shifts onto a table part's `table`, `autoFilter`, `sortState`, and
/// `sortCondition` references. The model only admits shifts the table can follow.
pub(crate) fn shift_table_part(xml: &[u8], shifts: &[StructuralShift]) -> OmResult<Vec<u8>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    let shifted = |element: &BytesStart<'_>| -> OmResult<BytesStart<'static>> {
        let mut shifted =
            BytesStart::new(String::from_utf8_lossy(element.name().as_ref()).into_owned());
        for attribute in element.attributes() {
            let attribute = attribute.map_err(xml_error)?;
            if attribute.key.as_ref() == b"ref" {
                let reference = attribute.unescape_value().map_err(xml_error)?;
                let moved = replay_shifts_on_ref(shifts, &reference)?.ok_or_else(|| {
                    OmError::invalid_state(format!(
                        "table reference {reference} was deleted by a structural shift"
                    ))
                })?;
                shifted.push_attribute(("ref", moved.as_str()));
            } else {
                shifted.push_attribute(attribute);
            }
        }
        Ok(shifted)
    };
    let is_ranged = |element: &BytesStart<'_>| {
        matches!(
            xml_local_name(element.name().as_ref()),
            b"table" | b"autoFilter" | b"sortState" | b"sortCondition"
        )
    };
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        let event = match event {
            Event::Start(element) if is_ranged(&element) => Event::Start(shifted(&element)?),
            Event::Empty(element) if is_ranged(&element) => Event::Empty(shifted(&element)?),
            Event::Eof => break,
            event => event.into_owned(),
        };
        writer.write_event(event).map_err(xml_error)?;
        buffer.clear();
    }
    Ok(writer.into_inner().into_inner())
}

/// Rewrites the zero-based `x:Row`/`x:Column` cell and `x:Anchor` box of every legacy VML
/// shape's client data. VML is edited textually because Excel's VML is not reliably XML.
pub(crate) fn shift_vml_part(xml: &[u8], shift: StructuralShift) -> OmResult<Vec<u8>> {
    let text = std::str::from_utf8(xml)
        .map_err(|_| OmError::parse("legacy VML drawing part is not UTF-8"))?;
    let mut output = String::with_capacity(text.len());
    let mut rest = text;
    const OPEN: &str = "<x:ClientData";
    const CLOSE: &str = "</x:ClientData>";
    while let Some(start) = rest.find(OPEN) {
        let Some(end) = rest[start..]
            .find(CLOSE)
            .map(|end| start + end + CLOSE.len())
        else {
            break;
        };
        output.push_str(&rest[..start]);
        output.push_str(&shift_client_data(&rest[start..end], shift)?);
        rest = &rest[end..];
    }
    output.push_str(rest);
    Ok(output.into_bytes())
}

/// The one-based position a zero-based VML row or column index moves to on the shifted axis.
fn shift_zero_based(value: u32, shift: StructuralShift) -> Option<u32> {
    shift.shift_index(value + 1).map(|moved| moved - 1)
}

fn shift_client_data(block: &str, shift: StructuralShift) -> OmResult<String> {
    if !shift.is_whole() {
        return shift_banded_client_data(block, shift);
    }
    let (cell_tag, anchor_indices): (&str, [usize; 2]) = match shift.axis {
        StructuralAxis::Rows => ("x:Row", [2, 6]),
        StructuralAxis::Columns => ("x:Column", [0, 4]),
    };
    let mut block = replace_tag_text(block, cell_tag, |value| {
        let index = value.trim().parse::<u32>().map_err(|_| {
            OmError::parse(format!("legacy VML {cell_tag} is not a number: {value}"))
        })?;
        let moved = shift_zero_based(index, shift).ok_or_else(|| {
            OmError::unsupported(
                "structural comment removal is not implemented for a VML note on a deleted cell",
            )
        })?;
        Ok(moved.to_string())
    })?;
    block = replace_tag_text(&block, "x:Anchor", |value| {
        let mut parts = value
            .split(',')
            .map(|part| {
                part.trim().parse::<u32>().map_err(|_| {
                    OmError::parse(format!("legacy VML anchor is not numeric: {value}"))
                })
            })
            .collect::<OmResult<Vec<_>>>()?;
        if parts.len() != 8 {
            return Err(OmError::parse(format!(
                "legacy VML anchor needs 8 values: {value}"
            )));
        }
        // A box edge inside a deleted span lands on the first index after the span.
        let collapse = shift.first - 1;
        for index in anchor_indices {
            parts[index] = shift_zero_based(parts[index], shift).unwrap_or(collapse);
        }
        if parts[anchor_indices[1]] < parts[anchor_indices[0]] {
            parts[anchor_indices[1]] = parts[anchor_indices[0]];
        }
        let leading = &value[..value.len() - value.trim_start().len()];
        let trailing = &value[value.trim_end().len()..];
        Ok(format!(
            "{leading}{}{trailing}",
            parts
                .iter()
                .map(u32::to_string)
                .collect::<Vec<_>>()
                .join(", ")
        ))
    })?;
    Ok(block)
}

/// Moves a note whose cell lies in a banded shift by the cell's displacement, keeping the size of
/// its box; notes outside the band stay put.
fn shift_banded_client_data(block: &str, shift: StructuralShift) -> OmResult<String> {
    let value_of = |tag: &str| -> OmResult<Option<u32>> {
        let open = format!("<{tag}>");
        let Some(start) = block.find(&open).map(|start| start + open.len()) else {
            return Ok(None);
        };
        let end = start + block[start..].find('<').unwrap_or(0);
        block[start..end]
            .trim()
            .parse::<u32>()
            .map(Some)
            .map_err(|_| OmError::parse(format!("legacy VML {tag} is not a number")))
    };
    let (Some(row), Some(col)) = (value_of("x:Row")?, value_of("x:Column")?) else {
        return Ok(block.to_string());
    };
    let moved = shift.shift_cell((row + 1, col + 1)).ok_or_else(|| {
        OmError::unsupported(
            "structural comment removal is not implemented for a VML note on a deleted cell",
        )
    })?;
    let (old, new, tag, anchor_indices) = match shift.axis {
        StructuralAxis::Rows => (row + 1, moved.0, "x:Row", [2, 6]),
        StructuralAxis::Columns => (col + 1, moved.1, "x:Column", [0, 4]),
    };
    if old == new {
        return Ok(block.to_string());
    }
    let displace = |value: u32| -> u32 {
        if new > old {
            value.saturating_add(new - old)
        } else {
            value.saturating_sub(old - new)
        }
    };
    let block = replace_tag_text(block, tag, |_| Ok((new - 1).to_string()))?;
    replace_tag_text(&block, "x:Anchor", |value| {
        let mut parts = value
            .split(',')
            .map(|part| part.trim().parse::<u32>())
            .collect::<Result<Vec<_>, _>>()
            .map_err(|_| OmError::parse(format!("legacy VML anchor is not numeric: {value}")))?;
        if parts.len() != 8 {
            return Err(OmError::parse(format!(
                "legacy VML anchor needs 8 values: {value}"
            )));
        }
        for index in anchor_indices {
            parts[index] = displace(parts[index]);
        }
        let leading = &value[..value.len() - value.trim_start().len()];
        let trailing = &value[value.trim_end().len()..];
        Ok(format!(
            "{leading}{}{trailing}",
            parts
                .iter()
                .map(u32::to_string)
                .collect::<Vec<_>>()
                .join(", ")
        ))
    })
}

/// Replaces the text of every `<tag>…</tag>` in `block`.
fn replace_tag_text(
    block: &str,
    tag: &str,
    mut replace: impl FnMut(&str) -> OmResult<String>,
) -> OmResult<String> {
    let open = format!("<{tag}>");
    let close = format!("</{tag}>");
    let mut output = String::with_capacity(block.len());
    let mut rest = block;
    while let Some(start) = rest.find(&open) {
        let value_start = start + open.len();
        let Some(value_end) = rest[value_start..]
            .find(&close)
            .map(|end| value_start + end)
        else {
            break;
        };
        output.push_str(&rest[..value_start]);
        output.push_str(&replace(&rest[value_start..value_end])?);
        output.push_str(&close);
        rest = &rest[value_end + close.len()..];
    }
    output.push_str(rest);
    Ok(output)
}

impl WorksheetSupportParts {
    /// Moves the worksheet snapshot's hyperlink references through a whole-row or whole-column
    /// shift so save validation expects the rewritten worksheet. Shifts that would remove a
    /// hyperlink, or that would need comment and VML anchors moved, are refused.
    pub fn apply_structural_shift(&mut self, shift: StructuralShift) -> OmResult<()> {
        if self.comment_part_uris.len() != self.comment_part_source_bytes.len()
            || self.vml_drawing_part_uris.len() != self.vml_drawing_part_source_bytes.len()
        {
            return Err(OmError::unsupported(
                "structural comment anchor retarget requires the source comment and VML parts",
            ));
        }
        // Comment and VML parts move as bytes; their expected summaries are re-read from the
        // moved bytes so save validation checks the parts the save writes.
        for (part_uri, bytes) in &mut self.comment_part_source_bytes {
            *bytes = shift_comment_part(bytes, shift)?;
            let summary = parse_comment_part_summary(bytes)?;
            if let Some(anchor_refs) = self.comment_anchor_refs.get_mut(part_uri) {
                *anchor_refs = summary
                    .comments
                    .iter()
                    .map(|comment| comment.reference.clone())
                    .collect();
            }
            if self.comment_summaries.contains_key(part_uri) {
                self.comment_summaries.insert(part_uri.clone(), summary);
            }
        }
        for (part_uri, bytes) in &mut self.vml_drawing_part_source_bytes {
            *bytes = shift_vml_part(bytes, shift)?;
            if self.vml_drawing_summaries.contains_key(part_uri) {
                self.vml_drawing_summaries
                    .insert(part_uri.clone(), parse_vml_drawing_part_summary(bytes)?);
            }
        }
        let shifts = [shift];
        let moved = |reference: &str| -> OmResult<String> {
            replay_shifts_on_ref(&shifts, reference)?.ok_or_else(|| {
                OmError::unsupported(format!(
                    "structural hyperlink removal is not implemented for hyperlink {reference}"
                ))
            })
        };
        for reference in &mut self.hyperlink_refs {
            *reference = moved(reference)?;
        }
        for summary in &mut self.hyperlink_summaries {
            summary.reference = moved(&summary.reference)?;
        }
        for binding in &mut self.hyperlink_bindings {
            binding.reference = moved(&binding.reference)?;
        }
        if let Some(summary) = self.hyperlinks_part_summary.as_mut() {
            for attributes in &mut summary.hyperlink_attr_maps {
                if let Some(reference) = attributes.get_mut("ref") {
                    *reference = moved(reference)?;
                }
            }
        }
        Ok(())
    }
}

/// Replays a worksheet's recorded structural shifts onto its table parts.
pub(crate) fn shift_table_parts(
    package: &mut OpcPackage,
    sheet_data: &WorksheetData,
) -> OmResult<()> {
    if sheet_data.structural_shifts.is_empty() {
        return Ok(());
    }
    for table_owner in &sheet_data.structural_owners.table_owners {
        let Some(source) = package.part(&table_owner.part_uri) else {
            continue;
        };
        let shifted = shift_table_part(&source.bytes, &sheet_data.structural_shifts)?;
        package.replace_part_bytes(&table_owner.part_uri, shifted)?;
    }
    Ok(())
}

/// Moves pictures and shapes in each drawing part through its host worksheet's recorded
/// structural shifts. This runs after every chart-frame rewrite, which moves chart anchors
/// from the model.
pub(crate) fn shift_drawing_parts(package: &mut OpcPackage, state: &WorkbookState) -> OmResult<()> {
    let mut shifted_drawing_parts = BTreeSet::new();
    for drawing in state.drawings.values() {
        let Some(part_uri) = drawing.raw_part_uri.as_deref() else {
            continue;
        };
        if !shifted_drawing_parts.insert(part_uri) {
            continue;
        }
        let Some(shifts) = state
            .worksheet_data()
            .get(&drawing.host_sheet_id)
            .map(|worksheet| worksheet.structural_shifts.as_slice())
            .filter(|shifts| !shifts.is_empty())
        else {
            continue;
        };
        let Some(part) = package.part(part_uri) else {
            continue;
        };
        // Banded shifts leave drawing objects in place; the model refuses any they reach.
        let whole_shifts = shifts
            .iter()
            .copied()
            .filter(|shift| shift.is_whole())
            .collect::<Vec<_>>();
        if whole_shifts.is_empty() {
            continue;
        }
        let shifted = shift_drawing_part(&part.bytes, &whole_shifts)?;
        package.replace_part_bytes(part_uri, shifted)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn rows(first: u32, count: u32, insert: bool) -> StructuralShift {
        StructuralShift {
            axis: StructuralAxis::Rows,
            first,
            count,
            insert,
            band: None,
        }
    }

    #[test]
    fn comment_refs_move_and_deleted_comments_are_refused() {
        let xml = br#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><commentList><comment ref="B4" authorId="0"><text><t>x</t></text></comment><comment ref="A1" authorId="0"/></commentList></comments>"#;
        let shifted =
            String::from_utf8(shift_comment_part(xml, rows(3, 2, true)).unwrap()).unwrap();
        assert!(
            shifted.contains(r#"<comment ref="B6" authorId="0">"#),
            "{shifted}"
        );
        assert!(
            shifted.contains(r#"<comment ref="A1" authorId="0"/>"#),
            "{shifted}"
        );
        let error = shift_comment_part(xml, rows(4, 1, false)).unwrap_err();
        assert!(error.message.contains("comment B4"), "{error:?}");
    }

    #[test]
    fn drawing_anchors_move_except_charts_and_absolute_objects() {
        let drawing = br#"<xdr:wsDr xmlns:xdr="urn:xdr"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>5</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>7</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>9</xdr:rowOff></xdr:to><xdr:pic/><xdr:clientData/></xdr:twoCellAnchor><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp/></xdr:twoCellAnchor><xdr:twoCellAnchor editAs="absolute"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp/></xdr:twoCellAnchor><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shifted =
            String::from_utf8(shift_drawing_part(drawing, &[rows(5, 2, true)]).unwrap()).unwrap();
        assert!(
            shifted.contains("<xdr:row>3</xdr:row><xdr:rowOff>7</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row>"),
            "a picture spanning the insertion grows: {shifted}"
        );
        assert!(
            shifted.contains("<xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row>"),
            "a move-only shape keeps its height: {shifted}"
        );
        assert_eq!(
            shifted.matches("<xdr:row>9</xdr:row>").count(),
            2,
            "{shifted}"
        );
        let deleted =
            String::from_utf8(shift_drawing_part(drawing, &[rows(3, 2, false)]).unwrap()).unwrap();
        assert!(
            deleted.contains("<xdr:from><xdr:col>1</xdr:col><xdr:colOff>5</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"),
            "an edge on a deleted row collapses: {deleted}"
        );
    }

    #[test]
    fn vml_note_rows_and_anchors_move() {
        let vml = b"<v:shape><x:ClientData ObjectType=\"Note\"><x:Anchor>\n    2, 15, 2, 10, 4, 15, 6, 4</x:Anchor><x:Row>3</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>";
        let shifted = String::from_utf8(shift_vml_part(vml, rows(3, 2, true)).unwrap()).unwrap();
        assert!(
            shifted.contains("<x:Row>5</x:Row><x:Column>1</x:Column>"),
            "{shifted}"
        );
        assert!(
            shifted.contains("<x:Anchor>\n    2, 15, 4, 10, 4, 15, 8, 4</x:Anchor>"),
            "{shifted}"
        );
        let columns = StructuralShift {
            axis: StructuralAxis::Columns,
            first: 1,
            count: 1,
            insert: true,
            band: None,
        };
        let shifted = String::from_utf8(shift_vml_part(vml, columns).unwrap()).unwrap();
        assert!(
            shifted.contains("<x:Row>3</x:Row><x:Column>2</x:Column>"),
            "{shifted}"
        );
        assert!(shifted.contains("3, 15, 2, 10, 5, 15, 6, 4"), "{shifted}");
    }
}
