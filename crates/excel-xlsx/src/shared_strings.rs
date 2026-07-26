use super::xml_error;
use office_common::OmResult;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::io::Cursor;

pub(super) fn parse_shared_strings(shared_strings_xml: &[u8]) -> OmResult<Vec<String>> {
    let mut reader = Reader::from_reader(Cursor::new(shared_strings_xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut values = Vec::new();
    let mut current = String::new();
    let mut inside_text = false;
    let mut inside_string_item = false;

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.name().as_ref() == b"si" => {
                current.clear();
                inside_string_item = true;
            }
            Ok(Event::Start(element)) if inside_string_item && element.name().as_ref() == b"t" => {
                inside_text = true;
            }
            Ok(Event::End(element)) if element.name().as_ref() == b"t" => inside_text = false,
            Ok(Event::End(element)) if element.name().as_ref() == b"si" => {
                values.push(current.clone());
                inside_string_item = false;
            }
            Ok(Event::Text(text)) if inside_text => {
                current.push_str(&text.xml_content().map_err(xml_error)?);
            }
            Ok(Event::CData(text)) if inside_text => {
                current.push_str(&text.xml_content().map_err(xml_error)?);
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(xml_error(error)),
        }
        buffer.clear();
    }

    Ok(values)
}
