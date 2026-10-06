//! Moving worksheet comment anchors and legacy VML shape anchors through whole-row and
//! whole-column shifts.

use super::{replay_shifts_on_sqref, xml_error, xml_local_name};
use office_common::{OmError, OmResult, StructuralAxis, StructuralShift};
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer};
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
            let moved = replay_shifts_on_sqref(&[shift], &reference)?.ok_or_else(|| {
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
                let moved = replay_shifts_on_sqref(shifts, &reference)?.ok_or_else(|| {
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

#[cfg(test)]
mod tests {
    use super::*;

    fn rows(first: u32, count: u32, insert: bool) -> StructuralShift {
        StructuralShift {
            axis: StructuralAxis::Rows,
            first,
            count,
            insert,
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
        };
        let shifted = String::from_utf8(shift_vml_part(vml, columns).unwrap()).unwrap();
        assert!(
            shifted.contains("<x:Row>3</x:Row><x:Column>2</x:Column>"),
            "{shifted}"
        );
        assert!(shifted.contains("3, 15, 2, 10, 5, 15, 6, 4"), "{shifted}");
    }
}
