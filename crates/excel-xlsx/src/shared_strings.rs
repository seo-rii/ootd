use super::{xml::resolved_element_is, xml_error};
use office_common::OmResult;
use quick_xml::NsReader;
use quick_xml::events::Event;
use std::io::Cursor;

pub(super) fn parse_shared_strings(
    shared_strings_xml: &[u8],
    spreadsheet_namespace: &str,
) -> OmResult<Vec<String>> {
    let mut reader = NsReader::from_reader(Cursor::new(shared_strings_xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut values = Vec::new();
    let mut current = String::new();
    let mut inside_text = false;
    let mut inside_string_item = false;

    loop {
        match reader.read_resolved_event_into(&mut buffer) {
            Ok((namespace, Event::Start(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"si",
                ) =>
            {
                current.clear();
                inside_string_item = true;
            }
            Ok((namespace, Event::Start(element)))
                if inside_string_item
                    && resolved_element_is(
                        &namespace,
                        element.local_name(),
                        spreadsheet_namespace.as_bytes(),
                        b"t",
                    ) =>
            {
                inside_text = true;
            }
            Ok((namespace, Event::End(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"t",
                ) =>
            {
                inside_text = false;
            }
            Ok((namespace, Event::End(element)))
                if resolved_element_is(
                    &namespace,
                    element.local_name(),
                    spreadsheet_namespace.as_bytes(),
                    b"si",
                ) =>
            {
                values.push(current.clone());
                inside_string_item = false;
            }
            Ok((_, Event::Text(text))) if inside_text => {
                current.push_str(&text.xml_content().map_err(xml_error)?);
            }
            Ok((_, Event::CData(text))) if inside_text => {
                current.push_str(&text.xml_content().map_err(xml_error)?);
            }
            Ok((_, Event::Eof)) => break,
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }

    Ok(values)
}
