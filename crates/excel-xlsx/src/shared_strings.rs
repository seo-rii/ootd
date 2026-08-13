use super::{xml::expanded_name_is, xml_error};
use office_common::{CellValue, OmError, OmResult, RichTextSource, RichTextValue};
use quick_xml::events::{BytesStart, Event};
use quick_xml::{NsReader, encoding::Decoder};
use std::io::Cursor;

struct CurrentStringItem {
    item_start: usize,
    inner_start: usize,
    depth: usize,
    preserve_raw_hint: bool,
    root_attributes: Vec<(String, String)>,
    namespace_declarations: Vec<(String, String)>,
}

pub(super) fn parse_shared_strings(
    shared_strings_xml: &[u8],
    spreadsheet_namespace: &str,
) -> OmResult<Vec<CellValue>> {
    let mut reader = NsReader::from_reader(Cursor::new(shared_strings_xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut values = Vec::new();
    let mut root_namespace_declarations = Vec::<(String, String)>::new();
    let mut current = None::<CurrentStringItem>;

    let collect_attributes =
        |element: &BytesStart<'_>, decoder: Decoder| -> OmResult<Vec<(String, String)>> {
            let mut attributes = Vec::new();
            for attribute in element.attributes() {
                let attribute = attribute.map_err(xml_error)?;
                attributes.push((
                    String::from_utf8_lossy(attribute.key.as_ref()).into_owned(),
                    attribute
                        .decode_and_unescape_value(decoder)
                        .map_err(xml_error)?
                        .into_owned(),
                ));
            }
            Ok(attributes)
        };

    loop {
        let event_start = reader.buffer_position() as usize;
        let decoder = reader.decoder();
        let event = reader
            .read_resolved_event_into(&mut buffer)
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
        let event_end = reader.buffer_position() as usize;

        match event {
            Ok((namespace, Event::Start(element)))
                if expanded_name_is(
                    namespace.as_deref(),
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"sst",
                ) =>
            {
                root_namespace_declarations = collect_attributes(&element, decoder)?
                    .into_iter()
                    .filter(|(name, _)| name == "xmlns" || name.starts_with("xmlns:"))
                    .collect();
            }
            Ok((namespace, Event::Start(element)))
                if expanded_name_is(
                    namespace.as_deref(),
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"si",
                ) =>
            {
                if current.is_some() {
                    return Err(OmError::parse("shared string items cannot be nested"));
                }
                let attributes = collect_attributes(&element, decoder)?;
                let item_has_namespace_declarations = attributes
                    .iter()
                    .any(|(name, _)| name == "xmlns" || name.starts_with("xmlns:"));
                let mut namespace_declarations = root_namespace_declarations.clone();
                for (name, value) in attributes
                    .iter()
                    .filter(|(name, _)| name == "xmlns" || name.starts_with("xmlns:"))
                {
                    if let Some(existing) = namespace_declarations
                        .iter_mut()
                        .find(|(existing_name, _)| existing_name == name)
                    {
                        existing.1.clone_from(value);
                    } else {
                        namespace_declarations.push((name.clone(), value.clone()));
                    }
                }
                let root_attributes = attributes
                    .into_iter()
                    .filter(|(name, _)| name != "xmlns" && !name.starts_with("xmlns:"))
                    .collect::<Vec<_>>();
                current = Some(CurrentStringItem {
                    item_start: event_start,
                    inner_start: event_end,
                    depth: 1,
                    preserve_raw_hint: item_has_namespace_declarations
                        || !root_attributes.is_empty(),
                    root_attributes,
                    namespace_declarations,
                });
            }
            Ok((namespace, Event::Empty(element)))
                if expanded_name_is(
                    namespace.as_deref(),
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"si",
                ) =>
            {
                let attributes = collect_attributes(&element, decoder)?;
                let item_has_namespace_declarations = attributes
                    .iter()
                    .any(|(name, _)| name == "xmlns" || name.starts_with("xmlns:"));
                let mut namespace_declarations = root_namespace_declarations.clone();
                for (name, value) in attributes
                    .iter()
                    .filter(|(name, _)| name == "xmlns" || name.starts_with("xmlns:"))
                {
                    if let Some(existing) = namespace_declarations
                        .iter_mut()
                        .find(|(existing_name, _)| existing_name == name)
                    {
                        existing.1.clone_from(value);
                    } else {
                        namespace_declarations.push((name.clone(), value.clone()));
                    }
                }
                let root_attributes = attributes
                    .iter()
                    .filter(|(name, _)| name != "xmlns" && !name.starts_with("xmlns:"))
                    .cloned()
                    .collect::<Vec<_>>();
                let raw_item_xml = shared_strings_xml[event_start..event_end].to_vec();
                let raw_item_len = raw_item_xml.len();
                values.push(parse_preserved_text_item(
                    raw_item_xml,
                    raw_item_len,
                    raw_item_len,
                    root_attributes,
                    namespace_declarations,
                    item_has_namespace_declarations,
                    RichTextSource::SharedString,
                    spreadsheet_namespace,
                )?);
            }
            Ok((_, Event::Start(_))) if current.is_some() => {
                current
                    .as_mut()
                    .expect("checked current shared string")
                    .depth += 1;
            }
            Ok((namespace, Event::End(element)))
                if current.is_some()
                    && expanded_name_is(
                        namespace.as_deref(),
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"si",
                    ) =>
            {
                let item = current.take().expect("checked current shared string");
                if item.depth != 1 {
                    return Err(OmError::parse(
                        "shared string item has invalid element nesting",
                    ));
                }
                let raw_item_xml = shared_strings_xml[item.item_start..event_end].to_vec();
                values.push(parse_preserved_text_item(
                    raw_item_xml,
                    item.inner_start - item.item_start,
                    event_start - item.item_start,
                    item.root_attributes,
                    item.namespace_declarations,
                    item.preserve_raw_hint,
                    RichTextSource::SharedString,
                    spreadsheet_namespace,
                )?);
            }
            Ok((_, Event::End(_))) if current.is_some() => {
                let item = current.as_mut().expect("checked current shared string");
                item.depth = item
                    .depth
                    .checked_sub(1)
                    .ok_or_else(|| OmError::parse("shared string item has invalid nesting"))?;
            }
            Ok((_, Event::Eof)) => {
                if current.is_some() {
                    return Err(OmError::parse("shared string item is not closed"));
                }
                break;
            }
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }

    Ok(values)
}

#[allow(clippy::too_many_arguments)]
pub(super) fn parse_preserved_text_item(
    raw_item_xml: Vec<u8>,
    inner_start: usize,
    inner_end: usize,
    root_attributes: Vec<(String, String)>,
    namespace_declarations: Vec<(String, String)>,
    preserve_raw_hint: bool,
    source: RichTextSource,
    spreadsheet_namespace: &str,
) -> OmResult<CellValue> {
    let value = RichTextValue::try_from_preserved_ooxml(
        raw_item_xml,
        inner_start,
        inner_end,
        root_attributes,
        namespace_declarations,
        source,
        spreadsheet_namespace.to_string(),
    )?;
    if preserve_raw_hint || value.requires_raw_preservation() {
        Ok(CellValue::RichText(value))
    } else {
        Ok(CellValue::Text(value.into_string()))
    }
}
