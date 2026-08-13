use std::error::Error;
use std::fmt::{Display, Formatter};
use std::io::{Cursor, Write};
use std::sync::Arc;

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{NsReader, Writer};
use serde::{Deserialize, Serialize};

mod name;
mod reference;

pub use name::{
    BuiltinName, DefinedName, DefinedNameId, DefinedNameKind, DefinedNameMetadata, NameKey,
    NameScope, NameValidationMode, canonicalize_excel_name,
};
pub use reference::{
    ExternalReference, RangeArea, RangeSet, ReferenceTarget, formula_contains_a1_reference,
};

pub type OmResult<T> = Result<T, OmError>;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum OmErrorCode {
    InvalidArgument,
    NotFound,
    TypeMismatch,
    Unsupported,
    InvalidState,
    Io,
    Parse,
    ResourceLimit,
    EncryptedWorkbookUnsupported,
    SignedPackageMutationUnsupported,
    ActiveContentConversionUnsupported,
    ActiveContentPolicyRefused,
    ExternalDataPolicyRefused,
    Calculation,
    External,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct OmError {
    pub code: OmErrorCode,
    pub message: String,
}

impl OmError {
    pub fn new(code: OmErrorCode, message: impl Into<String>) -> Self {
        Self {
            code,
            message: message.into(),
        }
    }

    pub fn invalid_argument(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::InvalidArgument, message)
    }

    pub fn invalid_state(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::InvalidState, message)
    }

    pub fn io(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::Io, message)
    }

    pub fn parse(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::Parse, message)
    }

    pub fn resource_limit(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::ResourceLimit, message)
    }

    pub fn encrypted_workbook_unsupported(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::EncryptedWorkbookUnsupported, message)
    }

    pub fn signed_package_mutation_unsupported(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::SignedPackageMutationUnsupported, message)
    }

    pub fn active_content_conversion_unsupported(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::ActiveContentConversionUnsupported, message)
    }

    pub fn active_content_policy_refused(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::ActiveContentPolicyRefused, message)
    }

    pub fn external_data_policy_refused(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::ExternalDataPolicyRefused, message)
    }

    pub fn unsupported(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::Unsupported, message)
    }
}

impl Display for OmError {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        write!(f, "{:?}: {}", self.code, self.message)
    }
}

impl Error for OmError {}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct WorkbookId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct SheetId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct StyleId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct ChartId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct ChartObjectId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct DrawingId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct DrawingObjectId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct Emu(pub i64);

impl Emu {
    pub fn to_points(self) -> Points {
        Points(self.0 as f64 / 12_700.0)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub struct Points(pub f64);

impl Points {
    pub fn to_emu(self) -> Emu {
        Emu((self.0 * 12_700.0).round() as i64)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct ObjectHandle(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct WorkbookHandle(pub ObjectHandle);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct WorksheetHandle(pub ObjectHandle);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct RangeHandle(pub ObjectHandle);

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
pub enum ExcelProfile {
    Excel2016,
    Excel2021,
    #[default]
    Excel365,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum FileFormat {
    Xlsx,
    Xlsm,
    Xltx,
    Xltm,
    StrictXlsx,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum CellError {
    Null,
    Div0,
    Value,
    Ref,
    Name,
    Num,
    NA,
    GettingData,
    Spill,
    Calc,
    Field,
    Blocked,
    Busy,
    Connect,
    Python,
    Timeout,
    Unknown,
    UnknownLexical(String),
}

impl CellError {
    pub fn from_lexical(value: &str) -> Self {
        match value {
            "#NULL!" => Self::Null,
            "#DIV/0!" => Self::Div0,
            "#VALUE!" => Self::Value,
            "#REF!" => Self::Ref,
            "#NAME?" => Self::Name,
            "#NUM!" => Self::Num,
            "#N/A" => Self::NA,
            "#GETTING_DATA" => Self::GettingData,
            "#SPILL!" => Self::Spill,
            "#CALC!" => Self::Calc,
            "#FIELD!" => Self::Field,
            "#BLOCKED!" => Self::Blocked,
            "#BUSY!" => Self::Busy,
            "#CONNECT!" => Self::Connect,
            "#PYTHON!" => Self::Python,
            "#TIMEOUT!" => Self::Timeout,
            "#UNKNOWN!" => Self::Unknown,
            _ => Self::UnknownLexical(value.to_string()),
        }
    }

    pub fn as_lexical_str(&self) -> &str {
        match self {
            Self::Null => "#NULL!",
            Self::Div0 => "#DIV/0!",
            Self::Value => "#VALUE!",
            Self::Ref => "#REF!",
            Self::Name => "#NAME?",
            Self::Num => "#NUM!",
            Self::NA => "#N/A",
            Self::GettingData => "#GETTING_DATA",
            Self::Spill => "#SPILL!",
            Self::Calc => "#CALC!",
            Self::Field => "#FIELD!",
            Self::Blocked => "#BLOCKED!",
            Self::Busy => "#BUSY!",
            Self::Connect => "#CONNECT!",
            Self::Python => "#PYTHON!",
            Self::Timeout => "#TIMEOUT!",
            Self::Unknown => "#UNKNOWN!",
            Self::UnknownLexical(value) => value,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ExcelDateSystem {
    Date1900,
    Date1904,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum IsoDateTimeOffsetPolicy {
    PreserveWallClock,
    NormalizeToUtc,
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
#[serde(try_from = "String", into = "String")]
pub struct IsoDateTime {
    lexical: String,
    year: u16,
    month: u8,
    day: u8,
    hour: u8,
    minute: u8,
    second: u8,
    fractional_second: String,
    offset_minutes: Option<i16>,
}

impl IsoDateTime {
    pub fn parse(value: impl Into<String>) -> OmResult<Self> {
        let value = value.into();
        let invalid = || {
            OmError::invalid_argument(format!(
                "invalid ISO 8601 worksheet date/time lexical: {value}"
            ))
        };
        if value.is_empty() || value.len() > 64 || !value.is_ascii() || value.trim() != value {
            return Err(invalid());
        }

        let parse_two_digits = |bytes: &[u8]| -> Option<u32> {
            if bytes.len() != 2 || !bytes.iter().all(u8::is_ascii_digit) {
                return None;
            }
            Some(u32::from(bytes[0] - b'0') * 10 + u32::from(bytes[1] - b'0'))
        };
        let mut core_end = value.len();
        let mut offset_minutes = None;
        if value.ends_with('Z') {
            core_end -= 1;
            offset_minutes = Some(0);
        } else if let Some(offset_index) = value
            .as_bytes()
            .iter()
            .enumerate()
            .skip(10)
            .find_map(|(index, byte)| matches!(byte, b'+' | b'-').then_some(index))
        {
            let offset = &value.as_bytes()[offset_index..];
            if offset.len() != 6 || offset[3] != b':' {
                return Err(invalid());
            }
            let Some(offset_hour) = parse_two_digits(&offset[1..3]) else {
                return Err(invalid());
            };
            let Some(offset_minute) = parse_two_digits(&offset[4..6]) else {
                return Err(invalid());
            };
            if offset_hour > 14 || offset_minute > 59 || (offset_hour == 14 && offset_minute != 0) {
                return Err(invalid());
            }
            let offset = i16::try_from(offset_hour * 60 + offset_minute).map_err(|_| invalid())?;
            offset_minutes = Some(if value.as_bytes()[offset_index] == b'-' {
                -offset
            } else {
                offset
            });
            core_end = offset_index;
        }

        let core = &value[..core_end];
        let (date, time) = match core.split_once('T') {
            Some((date, time)) if !time.is_empty() && !time.contains('T') => (date, Some(time)),
            Some(_) => return Err(invalid()),
            None => (core, None),
        };
        let date_bytes = date.as_bytes();
        if date_bytes.len() != 10
            || date_bytes[4] != b'-'
            || date_bytes[7] != b'-'
            || !date_bytes[..4].iter().all(u8::is_ascii_digit)
        {
            return Err(invalid());
        }
        let year = date_bytes[..4]
            .iter()
            .fold(0_u32, |year, digit| year * 10 + u32::from(*digit - b'0'));
        let Some(month) = parse_two_digits(&date_bytes[5..7]) else {
            return Err(invalid());
        };
        let Some(day) = parse_two_digits(&date_bytes[8..10]) else {
            return Err(invalid());
        };
        let leap_year =
            year.is_multiple_of(4) && (!year.is_multiple_of(100) || year.is_multiple_of(400));
        let days_in_month = match month {
            1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
            4 | 6 | 9 | 11 => 30,
            2 if leap_year => 29,
            2 => 28,
            _ => return Err(invalid()),
        };
        if year == 0 || day == 0 || day > days_in_month {
            return Err(invalid());
        }

        let mut hour = 0;
        let mut minute = 0;
        let mut second = 0;
        let mut fractional_second = String::new();
        if let Some(time) = time {
            let time_bytes = time.as_bytes();
            if time_bytes.len() < 8 || time_bytes[2] != b':' || time_bytes[5] != b':' {
                return Err(invalid());
            }
            let Some(parsed_hour) = parse_two_digits(&time_bytes[..2]) else {
                return Err(invalid());
            };
            let Some(parsed_minute) = parse_two_digits(&time_bytes[3..5]) else {
                return Err(invalid());
            };
            let Some(parsed_second) = parse_two_digits(&time_bytes[6..8]) else {
                return Err(invalid());
            };
            if parsed_hour > 23 || parsed_minute > 59 || parsed_second > 59 {
                return Err(invalid());
            }
            if time_bytes.len() > 8
                && (time_bytes[8] != b'.'
                    || time_bytes.len() == 9
                    || !time_bytes[9..].iter().all(u8::is_ascii_digit))
            {
                return Err(invalid());
            }
            hour = u8::try_from(parsed_hour).map_err(|_| invalid())?;
            minute = u8::try_from(parsed_minute).map_err(|_| invalid())?;
            second = u8::try_from(parsed_second).map_err(|_| invalid())?;
            if time_bytes.len() > 8 {
                fractional_second = time[9..].to_string();
            }
        }

        let year = u16::try_from(year).map_err(|_| invalid())?;
        let month = u8::try_from(month).map_err(|_| invalid())?;
        let day = u8::try_from(day).map_err(|_| invalid())?;
        Ok(Self {
            lexical: value,
            year,
            month,
            day,
            hour,
            minute,
            second,
            fractional_second,
            offset_minutes,
        })
    }

    pub fn as_str(&self) -> &str {
        &self.lexical
    }

    pub fn into_string(self) -> String {
        self.lexical
    }

    /// Converts the ISO value to an Excel serial only under explicit workbook and offset policies.
    /// ISO lexical parsing is locale-independent; no locale-specific coercion is performed here.
    pub fn to_excel_serial(
        &self,
        date_system: ExcelDateSystem,
        offset_policy: IsoDateTimeOffsetPolicy,
    ) -> f64 {
        let absolute_day = |year: i64, month: i64, day: i64| {
            let prior_year = year - 1;
            let days_before_year =
                365 * prior_year + prior_year / 4 - prior_year / 100 + prior_year / 400;
            let leap_year = year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
            let days_before_month = match month {
                1 => 0,
                2 => 31,
                3 => 59,
                4 => 90,
                5 => 120,
                6 => 151,
                7 => 181,
                8 => 212,
                9 => 243,
                10 => 273,
                11 => 304,
                12 => 334,
                _ => unreachable!("validated ISO month"),
            } + i64::from(leap_year && month > 2);
            days_before_year + days_before_month + day - 1
        };
        let year = i64::from(self.year);
        let month = i64::from(self.month);
        let day = i64::from(self.day);
        let date_day = absolute_day(year, month, day);
        let mut serial_day = match date_system {
            ExcelDateSystem::Date1900 => {
                let mut serial = date_day - absolute_day(1899, 12, 31);
                if (year, month, day) >= (1900, 3, 1) {
                    serial += 1;
                }
                serial
            }
            ExcelDateSystem::Date1904 => date_day - absolute_day(1904, 1, 1),
        } as f64;
        let fractional_second = if self.fractional_second.is_empty() {
            0.0
        } else {
            format!("0.{}", self.fractional_second)
                .parse::<f64>()
                .expect("validated ISO fractional second")
        };
        let seconds = f64::from(self.hour) * 3_600.0
            + f64::from(self.minute) * 60.0
            + f64::from(self.second)
            + fractional_second;
        serial_day += seconds / 86_400.0;
        if offset_policy == IsoDateTimeOffsetPolicy::NormalizeToUtc {
            serial_day -= f64::from(self.offset_minutes.unwrap_or(0)) / 1_440.0;
        }
        serial_day
    }
}

impl TryFrom<String> for IsoDateTime {
    type Error = OmError;

    fn try_from(value: String) -> OmResult<Self> {
        Self::parse(value)
    }
}

impl From<IsoDateTime> for String {
    fn from(value: IsoDateTime) -> Self {
        value.into_string()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum RichTextSource {
    SharedString,
    InlineString,
}

#[derive(Debug, Clone, Copy)]
enum RichTextTarget {
    Display,
    Phonetic,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
struct PreservedRichTextOoxml {
    raw_item_xml: Vec<u8>,
    inner_start: usize,
    inner_end: usize,
    root_attributes: Vec<(String, String)>,
    namespace_declarations: Vec<(String, String)>,
}

#[derive(Deserialize)]
struct PreservedRichTextOoxmlWire {
    raw_item_xml: Vec<u8>,
    inner_start: usize,
    inner_end: usize,
    root_attributes: Vec<(String, String)>,
    namespace_declarations: Vec<(String, String)>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct RichTextValue {
    display_text: String,
    phonetic_text: String,
    source: RichTextSource,
    spreadsheet_namespace: String,
    #[serde(skip)]
    requires_raw_preservation: bool,
    preserved_xml: Arc<PreservedRichTextOoxml>,
}

#[derive(Deserialize)]
struct RichTextValueWire {
    display_text: String,
    phonetic_text: String,
    source: RichTextSource,
    spreadsheet_namespace: String,
    preserved_xml: PreservedRichTextOoxmlWire,
}

impl<'de> Deserialize<'de> for RichTextValue {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        let wire = RichTextValueWire::deserialize(deserializer)?;
        let parsed = Self::try_from_preserved_ooxml(
            wire.preserved_xml.raw_item_xml,
            wire.preserved_xml.inner_start,
            wire.preserved_xml.inner_end,
            wire.preserved_xml.root_attributes,
            wire.preserved_xml.namespace_declarations,
            wire.source,
            wire.spreadsheet_namespace,
        )
        .map_err(<D::Error as serde::de::Error>::custom)?;
        if parsed.display_text != wire.display_text || parsed.phonetic_text != wire.phonetic_text {
            return Err(<D::Error as serde::de::Error>::custom(
                "preserved rich-text projection does not match its OOXML payload",
            ));
        }
        Ok(parsed)
    }
}

impl RichTextValue {
    pub fn try_from_preserved_ooxml(
        raw_item_xml: Vec<u8>,
        inner_start: usize,
        inner_end: usize,
        root_attributes: Vec<(String, String)>,
        namespace_declarations: Vec<(String, String)>,
        source: RichTextSource,
        spreadsheet_namespace: String,
    ) -> OmResult<Self> {
        let invalid = |reason: &str| {
            OmError::invalid_argument(format!(
                "invalid preserved rich-text OOXML metadata: {reason}"
            ))
        };
        let ncname_is_valid = |name: &str| {
            let is_start = |character: char| {
                character == '_'
                    || character.is_ascii_alphabetic()
                    || matches!(
                        character,
                        '\u{00c0}'..='\u{00d6}'
                            | '\u{00d8}'..='\u{00f6}'
                            | '\u{00f8}'..='\u{02ff}'
                            | '\u{0370}'..='\u{037d}'
                            | '\u{037f}'..='\u{1fff}'
                            | '\u{200c}'..='\u{200d}'
                            | '\u{2070}'..='\u{218f}'
                            | '\u{2c00}'..='\u{2fef}'
                            | '\u{3001}'..='\u{d7ff}'
                            | '\u{f900}'..='\u{fdcf}'
                            | '\u{fdf0}'..='\u{fffd}'
                            | '\u{10000}'..='\u{effff}'
                    )
            };
            let mut characters = name.chars();
            characters.next().is_some_and(is_start)
                && characters.all(|character| {
                    is_start(character)
                        || character.is_ascii_digit()
                        || matches!(
                            character,
                            '-' | '.'
                                | '\u{00b7}'
                                | '\u{0300}'..='\u{036f}'
                                | '\u{203f}'..='\u{2040}'
                        )
                })
        };
        let qname_is_valid = |name: &str| {
            if let Some((prefix, local_name)) = name.split_once(':') {
                !local_name.contains(':') && ncname_is_valid(prefix) && ncname_is_valid(local_name)
            } else {
                ncname_is_valid(name)
            }
        };
        let xml_character_is_valid = |character: char| {
            matches!(
                character,
                '\u{0009}' | '\u{000a}' | '\u{000d}' | '\u{0020}'..='\u{d7ff}'
                    | '\u{e000}'..='\u{fffd}'
                    | '\u{10000}'..='\u{10ffff}'
            )
        };
        let invalid_root_attribute =
            root_attributes
                .iter()
                .enumerate()
                .any(|(index, (name, value))| {
                    name.is_empty()
                        || !qname_is_valid(name)
                        || !value.chars().all(xml_character_is_valid)
                        || name == "xmlns"
                        || name.starts_with("xmlns:")
                        || root_attributes[..index]
                            .iter()
                            .any(|(previous, _)| previous == name)
                });
        let invalid_namespace =
            namespace_declarations
                .iter()
                .enumerate()
                .any(|(index, (name, namespace))| {
                    let invalid_binding = if !namespace.chars().all(xml_character_is_valid) {
                        true
                    } else if name == "xmlns" {
                        namespace == "http://www.w3.org/XML/1998/namespace"
                            || namespace == "http://www.w3.org/2000/xmlns/"
                    } else if let Some(prefix) = name.strip_prefix("xmlns:") {
                        !ncname_is_valid(prefix)
                            || prefix == "xmlns"
                            || namespace.is_empty()
                            || (prefix == "xml"
                                && namespace != "http://www.w3.org/XML/1998/namespace")
                            || (prefix != "xml"
                                && namespace == "http://www.w3.org/XML/1998/namespace")
                            || namespace == "http://www.w3.org/2000/xmlns/"
                    } else {
                        true
                    };
                    invalid_binding
                        || namespace_declarations[..index]
                            .iter()
                            .any(|(previous, _)| previous == name)
                });
        if raw_item_xml.is_empty()
            || inner_start > inner_end
            || inner_end > raw_item_xml.len()
            || invalid_root_attribute
            || invalid_namespace
            || spreadsheet_namespace.is_empty()
            || !spreadsheet_namespace.chars().all(xml_character_is_valid)
        {
            return Err(invalid("invalid bounds, names, or namespace"));
        }

        let mut wrapper_writer = Writer::new(Cursor::new(Vec::new()));
        let mut wrapper = BytesStart::new("ootd-rich-text-root");
        for (name, value) in &namespace_declarations {
            wrapper.push_attribute((name.as_str(), value.as_str()));
        }
        wrapper_writer
            .write_event(Event::Start(wrapper))
            .map_err(|error| invalid(&error.to_string()))?;
        let wrapped_item_start = usize::try_from(wrapper_writer.get_mut().position())
            .map_err(|_| invalid("wrapper offset exceeds the platform limit"))?;
        wrapper_writer
            .get_mut()
            .write_all(&raw_item_xml)
            .map_err(|error| invalid(&error.to_string()))?;
        wrapper_writer
            .write_event(Event::End(BytesEnd::new("ootd-rich-text-root")))
            .map_err(|error| invalid(&error.to_string()))?;
        let wrapped_xml = wrapper_writer.into_inner().into_inner();
        let wrapped_item_end = wrapped_item_start
            .checked_add(raw_item_xml.len())
            .ok_or_else(|| invalid("wrapped item offset overflow"))?;

        let item_local_name = match source {
            RichTextSource::SharedString => b"si".as_slice(),
            RichTextSource::InlineString => b"is".as_slice(),
        };
        let expected_namespace = spreadsheet_namespace.as_bytes();
        let mut reader = NsReader::from_reader(Cursor::new(wrapped_xml));
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let mut item_depth = None::<usize>;
        let mut phonetic_depth = None::<usize>;
        let mut text_target = None::<(usize, RichTextTarget)>;
        let mut text_element_count = 0usize;
        let mut display_text = String::new();
        let mut phonetic_text = String::new();
        let mut saw_item = false;
        let mut closed_item = false;
        let mut requires_raw_preservation = false;

        loop {
            let event_start = reader.buffer_position() as usize;
            let decoder = reader.decoder();
            let event = match reader.read_resolved_event_into(&mut buffer) {
                Ok((namespace, event)) => {
                    let namespace = match namespace {
                        quick_xml::name::ResolveResult::Bound(namespace) => {
                            Some(namespace.as_ref().to_vec())
                        }
                        quick_xml::name::ResolveResult::Unbound => None,
                        quick_xml::name::ResolveResult::Unknown(prefix) => {
                            return Err(invalid(&format!(
                                "undeclared element prefix {}",
                                String::from_utf8_lossy(&prefix)
                            )));
                        }
                    };
                    Ok((namespace, event))
                }
                Err(error) => Err(error),
            };
            let event_end = reader.buffer_position() as usize;

            match event {
                Ok((namespace, Event::Start(element)))
                    if item_depth.is_none()
                        && element.local_name().as_ref() == item_local_name
                        && namespace.as_deref() == Some(expected_namespace) =>
                {
                    if saw_item || event_start != wrapped_item_start {
                        return Err(invalid("payload must contain exactly one item element"));
                    }
                    let item_qname = element.name();
                    let item_name = std::str::from_utf8(item_qname.as_ref())
                        .map_err(|error| invalid(&error.to_string()))?;
                    if !qname_is_valid(item_name) {
                        return Err(invalid("source item has an invalid QName"));
                    }
                    let mut actual_root_attributes = Vec::new();
                    let mut actual_namespace_declarations = Vec::new();
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| invalid(&error.to_string()))?;
                        let name = std::str::from_utf8(attribute.key.as_ref())
                            .map_err(|error| invalid(&error.to_string()))?
                            .to_string();
                        if !qname_is_valid(&name) {
                            return Err(invalid("source item attribute has an invalid QName"));
                        }
                        let value = attribute
                            .decode_and_unescape_value(decoder)
                            .map_err(|error| invalid(&error.to_string()))?
                            .into_owned();
                        if !value.chars().all(xml_character_is_valid) {
                            return Err(invalid(
                                "source item attribute has an invalid XML character",
                            ));
                        }
                        if name == "xmlns" || name.starts_with("xmlns:") {
                            actual_namespace_declarations.push((name, value));
                        } else {
                            if let quick_xml::name::ResolveResult::Unknown(prefix) =
                                reader.resolver().resolve_attribute(attribute.key).0
                            {
                                return Err(invalid(&format!(
                                    "undeclared attribute prefix {}",
                                    String::from_utf8_lossy(&prefix)
                                )));
                            }
                            actual_root_attributes.push((name, value));
                        }
                    }
                    if actual_root_attributes != root_attributes
                        || actual_namespace_declarations
                            .iter()
                            .any(|declaration| !namespace_declarations.contains(declaration))
                    {
                        return Err(invalid("root metadata does not match the raw item"));
                    }
                    let actual_inner_start = event_end
                        .checked_sub(wrapped_item_start)
                        .ok_or_else(|| invalid("item start precedes its wrapper offset"))?;
                    if inner_start != actual_inner_start {
                        return Err(invalid("inner start is not an XML token boundary"));
                    }
                    requires_raw_preservation = !actual_root_attributes.is_empty()
                        || !actual_namespace_declarations.is_empty();
                    saw_item = true;
                    item_depth = Some(1);
                }
                Ok((namespace, Event::Empty(element)))
                    if item_depth.is_none()
                        && element.local_name().as_ref() == item_local_name
                        && namespace.as_deref() == Some(expected_namespace) =>
                {
                    if saw_item
                        || event_start != wrapped_item_start
                        || event_end != wrapped_item_end
                        || inner_start != raw_item_xml.len()
                        || inner_end != raw_item_xml.len()
                    {
                        return Err(invalid("empty payload item has invalid boundaries"));
                    }
                    let item_qname = element.name();
                    let item_name = std::str::from_utf8(item_qname.as_ref())
                        .map_err(|error| invalid(&error.to_string()))?;
                    if !qname_is_valid(item_name) {
                        return Err(invalid("source item has an invalid QName"));
                    }
                    let mut actual_root_attributes = Vec::new();
                    let mut actual_namespace_declarations = Vec::new();
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| invalid(&error.to_string()))?;
                        let name = std::str::from_utf8(attribute.key.as_ref())
                            .map_err(|error| invalid(&error.to_string()))?
                            .to_string();
                        if !qname_is_valid(&name) {
                            return Err(invalid("source item attribute has an invalid QName"));
                        }
                        let value = attribute
                            .decode_and_unescape_value(decoder)
                            .map_err(|error| invalid(&error.to_string()))?
                            .into_owned();
                        if !value.chars().all(xml_character_is_valid) {
                            return Err(invalid(
                                "source item attribute has an invalid XML character",
                            ));
                        }
                        if name == "xmlns" || name.starts_with("xmlns:") {
                            actual_namespace_declarations.push((name, value));
                        } else {
                            if let quick_xml::name::ResolveResult::Unknown(prefix) =
                                reader.resolver().resolve_attribute(attribute.key).0
                            {
                                return Err(invalid(&format!(
                                    "undeclared attribute prefix {}",
                                    String::from_utf8_lossy(&prefix)
                                )));
                            }
                            actual_root_attributes.push((name, value));
                        }
                    }
                    if actual_root_attributes != root_attributes
                        || actual_namespace_declarations
                            .iter()
                            .any(|declaration| !namespace_declarations.contains(declaration))
                    {
                        return Err(invalid("root metadata does not match the raw item"));
                    }
                    requires_raw_preservation = !actual_root_attributes.is_empty()
                        || !actual_namespace_declarations.is_empty();
                    saw_item = true;
                    closed_item = true;
                }
                Ok((namespace, Event::Start(element))) if item_depth.is_some() => {
                    let depth = item_depth
                        .expect("checked rich-text item depth")
                        .checked_add(1)
                        .ok_or_else(|| invalid("item nesting depth overflow"))?;
                    item_depth = Some(depth);
                    let element_qname = element.name();
                    let element_name = std::str::from_utf8(element_qname.as_ref())
                        .map_err(|error| invalid(&error.to_string()))?;
                    if !qname_is_valid(element_name) {
                        return Err(invalid("rich-text child has an invalid QName"));
                    }
                    let mut has_attributes = false;
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| invalid(&error.to_string()))?;
                        let attribute_name = std::str::from_utf8(attribute.key.as_ref())
                            .map_err(|error| invalid(&error.to_string()))?;
                        if !qname_is_valid(attribute_name) {
                            return Err(invalid("rich-text attribute has an invalid QName"));
                        }
                        let attribute_value = attribute
                            .decode_and_unescape_value(decoder)
                            .map_err(|error| invalid(&error.to_string()))?;
                        if !attribute_value.chars().all(xml_character_is_valid) {
                            return Err(invalid(
                                "rich-text attribute has an invalid XML character",
                            ));
                        }
                        if attribute.key.as_namespace_binding().is_none()
                            && let quick_xml::name::ResolveResult::Unknown(prefix) =
                                reader.resolver().resolve_attribute(attribute.key).0
                        {
                            return Err(invalid(&format!(
                                "undeclared attribute prefix {}",
                                String::from_utf8_lossy(&prefix)
                            )));
                        }
                        has_attributes = true;
                    }
                    let is_spreadsheet_element = namespace.as_deref() == Some(expected_namespace);
                    let local_name = element.local_name();
                    if is_spreadsheet_element && local_name.as_ref() == b"rPh" {
                        phonetic_depth = Some(depth);
                        requires_raw_preservation = true;
                    } else if is_spreadsheet_element && local_name.as_ref() == b"t" {
                        text_element_count = text_element_count
                            .checked_add(1)
                            .ok_or_else(|| invalid("text element count overflow"))?;
                        if has_attributes {
                            requires_raw_preservation = true;
                        }
                        if text_element_count > 1 {
                            requires_raw_preservation = true;
                        }
                        text_target = Some((
                            depth,
                            if phonetic_depth.is_some() {
                                RichTextTarget::Phonetic
                            } else {
                                RichTextTarget::Display
                            },
                        ));
                    } else {
                        requires_raw_preservation = true;
                    }
                }
                Ok((namespace, Event::Empty(element))) if item_depth.is_some() => {
                    let element_qname = element.name();
                    let element_name = std::str::from_utf8(element_qname.as_ref())
                        .map_err(|error| invalid(&error.to_string()))?;
                    if !qname_is_valid(element_name) {
                        return Err(invalid("rich-text child has an invalid QName"));
                    }
                    let mut has_attributes = false;
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| invalid(&error.to_string()))?;
                        let attribute_name = std::str::from_utf8(attribute.key.as_ref())
                            .map_err(|error| invalid(&error.to_string()))?;
                        if !qname_is_valid(attribute_name) {
                            return Err(invalid("rich-text attribute has an invalid QName"));
                        }
                        let attribute_value = attribute
                            .decode_and_unescape_value(decoder)
                            .map_err(|error| invalid(&error.to_string()))?;
                        if !attribute_value.chars().all(xml_character_is_valid) {
                            return Err(invalid(
                                "rich-text attribute has an invalid XML character",
                            ));
                        }
                        if attribute.key.as_namespace_binding().is_none()
                            && let quick_xml::name::ResolveResult::Unknown(prefix) =
                                reader.resolver().resolve_attribute(attribute.key).0
                        {
                            return Err(invalid(&format!(
                                "undeclared attribute prefix {}",
                                String::from_utf8_lossy(&prefix)
                            )));
                        }
                        has_attributes = true;
                    }
                    let is_plain_empty_text = namespace.as_deref() == Some(expected_namespace)
                        && element.local_name().as_ref() == b"t"
                        && !has_attributes
                        && text_element_count == 0;
                    if is_plain_empty_text {
                        text_element_count = 1;
                    } else {
                        requires_raw_preservation = true;
                    }
                }
                Ok((_, Event::Text(text))) if item_depth.is_some() => {
                    if let Some((target_depth, target)) = text_target
                        && Some(target_depth) == item_depth
                    {
                        let text = text
                            .xml_content()
                            .map_err(|error| invalid(&error.to_string()))?;
                        if !text.chars().all(xml_character_is_valid) {
                            return Err(invalid("rich text has an invalid XML character"));
                        }
                        match target {
                            RichTextTarget::Display => display_text.push_str(&text),
                            RichTextTarget::Phonetic => phonetic_text.push_str(&text),
                        }
                    } else {
                        let text = text
                            .xml_content()
                            .map_err(|error| invalid(&error.to_string()))?;
                        if !text.chars().all(xml_character_is_valid) {
                            return Err(invalid("rich text has an invalid XML character"));
                        }
                        if text.chars().any(|character| !character.is_whitespace()) {
                            requires_raw_preservation = true;
                        }
                    }
                }
                Ok((_, Event::CData(text))) if item_depth.is_some() => {
                    requires_raw_preservation = true;
                    let text = text
                        .xml_content()
                        .map_err(|error| invalid(&error.to_string()))?;
                    if !text.chars().all(xml_character_is_valid) {
                        return Err(invalid("rich text has an invalid XML character"));
                    }
                    if let Some((target_depth, target)) = text_target
                        && Some(target_depth) == item_depth
                    {
                        match target {
                            RichTextTarget::Display => display_text.push_str(&text),
                            RichTextTarget::Phonetic => phonetic_text.push_str(&text),
                        }
                    }
                }
                Ok((_, Event::GeneralRef(reference))) if item_depth.is_some() => {
                    requires_raw_preservation = true;
                    let reference = reference
                        .decode()
                        .map_err(|error| invalid(&error.to_string()))?;
                    let text = if let Some(number) = reference.strip_prefix("#x") {
                        u32::from_str_radix(number, 16)
                            .ok()
                            .and_then(char::from_u32)
                            .filter(|character| xml_character_is_valid(*character))
                            .map(String::from)
                    } else if let Some(number) = reference.strip_prefix("#X") {
                        u32::from_str_radix(number, 16)
                            .ok()
                            .and_then(char::from_u32)
                            .filter(|character| xml_character_is_valid(*character))
                            .map(String::from)
                    } else if let Some(number) = reference.strip_prefix('#') {
                        number
                            .parse::<u32>()
                            .ok()
                            .and_then(char::from_u32)
                            .filter(|character| xml_character_is_valid(*character))
                            .map(String::from)
                    } else {
                        match reference.as_ref() {
                            "amp" => Some("&".to_string()),
                            "lt" => Some("<".to_string()),
                            "gt" => Some(">".to_string()),
                            "quot" => Some("\"".to_string()),
                            "apos" => Some("'".to_string()),
                            _ => None,
                        }
                    }
                    .ok_or_else(|| invalid("unknown or invalid XML entity reference"))?;
                    if let Some((target_depth, target)) = text_target
                        && Some(target_depth) == item_depth
                    {
                        match target {
                            RichTextTarget::Display => display_text.push_str(&text),
                            RichTextTarget::Phonetic => phonetic_text.push_str(&text),
                        }
                    }
                }
                Ok((_, Event::Comment(_) | Event::PI(_))) if item_depth.is_some() => {
                    requires_raw_preservation = true;
                }
                Ok((_, Event::Decl(_) | Event::DocType(_))) if item_depth.is_some() => {
                    return Err(invalid(
                        "declarations are not valid inside a rich-text item",
                    ));
                }
                Ok((namespace, Event::End(element))) if item_depth.is_some() => {
                    let depth = item_depth.expect("checked rich-text item depth");
                    let is_spreadsheet_element = namespace.as_deref() == Some(expected_namespace);
                    let local_name = element.local_name();
                    if depth == 1
                        && is_spreadsheet_element
                        && local_name.as_ref() == item_local_name
                    {
                        let actual_inner_end = event_start
                            .checked_sub(wrapped_item_start)
                            .ok_or_else(|| invalid("item end precedes its wrapper offset"))?;
                        if inner_end != actual_inner_end || event_end != wrapped_item_end {
                            return Err(invalid("inner end is not an XML token boundary"));
                        }
                        item_depth = None;
                        closed_item = true;
                    } else {
                        if text_target.is_some_and(|(target_depth, _)| target_depth == depth)
                            && is_spreadsheet_element
                            && local_name.as_ref() == b"t"
                        {
                            text_target = None;
                        }
                        if phonetic_depth == Some(depth)
                            && is_spreadsheet_element
                            && local_name.as_ref() == b"rPh"
                        {
                            phonetic_depth = None;
                        }
                        item_depth = Some(
                            depth
                                .checked_sub(1)
                                .ok_or_else(|| invalid("item nesting underflow"))?,
                        );
                    }
                }
                Ok((_, Event::Eof)) => break,
                Ok(_) => {}
                Err(error) => return Err(invalid(&error.to_string())),
            }
            buffer.clear();
        }

        if !saw_item || !closed_item || item_depth.is_some() {
            return Err(invalid("payload does not contain one complete source item"));
        }

        Ok(Self {
            display_text,
            phonetic_text,
            source,
            spreadsheet_namespace,
            requires_raw_preservation,
            preserved_xml: Arc::new(PreservedRichTextOoxml {
                raw_item_xml,
                inner_start,
                inner_end,
                root_attributes,
                namespace_declarations,
            }),
        })
    }

    fn validate_preserved_ooxml(&self) -> OmResult<()> {
        if self.spreadsheet_namespace.is_empty()
            || self.preserved_xml.raw_item_xml.is_empty()
            || self.preserved_xml.inner_start > self.preserved_xml.inner_end
            || self.preserved_xml.inner_end > self.preserved_xml.raw_item_xml.len()
        {
            return Err(OmError::invalid_argument(
                "invalid preserved rich-text OOXML metadata",
            ));
        }
        Ok(())
    }

    pub fn as_str(&self) -> &str {
        &self.display_text
    }

    pub fn phonetic_text(&self) -> &str {
        &self.phonetic_text
    }

    pub fn source(&self) -> RichTextSource {
        self.source
    }

    pub fn spreadsheet_namespace(&self) -> &str {
        &self.spreadsheet_namespace
    }

    pub fn requires_raw_preservation(&self) -> bool {
        self.requires_raw_preservation
    }

    pub fn raw_item_xml(&self) -> &[u8] {
        &self.preserved_xml.raw_item_xml
    }

    pub fn raw_inner_xml(&self) -> &[u8] {
        &self.preserved_xml.raw_item_xml
            [self.preserved_xml.inner_start..self.preserved_xml.inner_end]
    }

    pub fn root_attributes(&self) -> &[(String, String)] {
        &self.preserved_xml.root_attributes
    }

    pub fn namespace_declarations(&self) -> &[(String, String)] {
        &self.preserved_xml.namespace_declarations
    }

    pub fn into_string(self) -> String {
        self.display_text
    }
}

#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub enum CellValue {
    #[default]
    Blank,
    Bool(bool),
    Number(f64),
    Text(String),
    Error(CellError),
    IsoDateTime(IsoDateTime),
    RichText(RichTextValue),
}

impl CellValue {
    pub fn as_text(&self) -> Option<&str> {
        match self {
            Self::Text(value) => Some(value),
            Self::IsoDateTime(value) => Some(value.as_str()),
            Self::RichText(value) => Some(value.as_str()),
            Self::Blank | Self::Bool(_) | Self::Number(_) | Self::Error(_) => None,
        }
    }

    pub fn try_number(value: f64) -> OmResult<Self> {
        if !value.is_finite() {
            return Err(OmError::invalid_argument(
                "worksheet cell numeric value must be finite",
            ));
        }
        Ok(Self::Number(value))
    }

    pub fn validate(&self) -> OmResult<()> {
        match self {
            Self::Number(value) => {
                Self::try_number(*value)?;
            }
            Self::RichText(value) => value.validate_preserved_ooxml()?,
            Self::Blank | Self::Bool(_) | Self::Text(_) | Self::Error(_) | Self::IsoDateTime(_) => {
            }
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum OmValue {
    Missing,
    Empty,
    Null,
    Bool(bool),
    Number(f64),
    Text(String),
    Error(CellError),
    Object(ObjectHandle),
    Array(OmArray),
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct OmArray {
    pub rows: usize,
    pub cols: usize,
    pub values: Vec<OmValue>,
}

impl OmArray {
    pub fn new(rows: usize, cols: usize, values: Vec<OmValue>) -> OmResult<Self> {
        if rows.saturating_mul(cols) != values.len() {
            return Err(OmError::invalid_argument(
                "array dimensions do not match the number of values",
            ));
        }

        Ok(Self { rows, cols, values })
    }

    pub fn scalar(value: OmValue) -> Self {
        Self {
            rows: 1,
            cols: 1,
            values: vec![value],
        }
    }

    pub fn get(&self, row: usize, col: usize) -> Option<&OmValue> {
        if row >= self.rows || col >= self.cols {
            return None;
        }

        self.values.get(row * self.cols + col)
    }
}

impl From<CellValue> for OmValue {
    fn from(value: CellValue) -> Self {
        match value {
            CellValue::Blank => OmValue::Empty,
            CellValue::Bool(value) => OmValue::Bool(value),
            CellValue::Number(value) => OmValue::Number(value),
            CellValue::Text(value) => OmValue::Text(value),
            CellValue::Error(value) => OmValue::Error(value),
            CellValue::IsoDateTime(value) => OmValue::Text(value.into_string()),
            CellValue::RichText(value) => OmValue::Text(value.into_string()),
        }
    }
}

impl TryFrom<OmValue> for CellValue {
    type Error = OmError;

    fn try_from(value: OmValue) -> OmResult<Self> {
        match value {
            OmValue::Missing | OmValue::Empty | OmValue::Null => Ok(CellValue::Blank),
            OmValue::Bool(value) => Ok(CellValue::Bool(value)),
            OmValue::Number(value) => CellValue::try_number(value),
            OmValue::Text(value) => Ok(CellValue::Text(value)),
            OmValue::Error(value) => Ok(CellValue::Error(value)),
            OmValue::Object(_) | OmValue::Array(_) => Err(OmError::type_mismatch(
                "cannot coerce object or array values into a worksheet cell",
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct CellRef {
    pub sheet_id: SheetId,
    pub row: u32,
    pub col: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct Rect {
    pub row_first: u32,
    pub row_last: u32,
    pub col_first: u32,
    pub col_last: u32,
}

impl Rect {
    pub fn single_cell(row: u32, col: u32) -> Self {
        Self {
            row_first: row,
            row_last: row,
            col_first: col,
            col_last: col,
        }
    }

    pub fn width(&self) -> u32 {
        self.col_last
            .saturating_sub(self.col_first)
            .saturating_add(1)
    }

    pub fn height(&self) -> u32 {
        self.row_last
            .saturating_sub(self.row_first)
            .saturating_add(1)
    }

    pub fn checked_cell_count(&self) -> OmResult<u64> {
        ExcelLimits::validate_rect(*self)?;
        u64::from(self.height())
            .checked_mul(u64::from(self.width()))
            .ok_or_else(|| OmError::resource_limit("worksheet range cell count overflowed u64"))
    }

    pub fn checked_cell_count_usize(&self) -> OmResult<usize> {
        usize::try_from(self.checked_cell_count()?).map_err(|_| {
            OmError::resource_limit("worksheet range cell count does not fit this platform's usize")
        })
    }
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ExcelLimits;

impl ExcelLimits {
    pub const MAX_ROW_INDEX: u32 = 1_048_576;
    pub const MAX_COLUMN_INDEX: u32 = 16_384;
    pub const MAX_CELL_COUNT: u64 = Self::MAX_ROW_INDEX as u64 * Self::MAX_COLUMN_INDEX as u64;

    pub fn validate_cell(row: u32, col: u32) -> OmResult<()> {
        if row == 0 || col == 0 {
            return Err(OmError::invalid_argument(
                "worksheet coordinates are 1-based and must be greater than zero",
            ));
        }
        if row > Self::MAX_ROW_INDEX || col > Self::MAX_COLUMN_INDEX {
            return Err(OmError::invalid_argument(format!(
                "worksheet coordinate R{row}C{col} exceeds Excel grid XFD1048576"
            )));
        }
        Ok(())
    }

    pub fn validate_rect(rect: Rect) -> OmResult<()> {
        if rect.row_first > rect.row_last || rect.col_first > rect.col_last {
            return Err(OmError::invalid_argument(
                "worksheet range rectangle is inverted",
            ));
        }
        Self::validate_cell(rect.row_first, rect.col_first)?;
        Self::validate_cell(rect.row_last, rect.col_last)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SheetScope {
    Single(SheetId),
    Multi3D { start: SheetId, end: SheetId },
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct RangeRef {
    pub workbook_id: WorkbookId,
    pub scope: SheetScope,
    pub areas: Vec<Rect>,
}

impl RangeRef {
    pub fn single_cell(workbook_id: WorkbookId, sheet_id: SheetId, row: u32, col: u32) -> Self {
        Self {
            workbook_id,
            scope: SheetScope::Single(sheet_id),
            areas: vec![Rect::single_cell(row, col)],
        }
    }

    pub fn single_rect(workbook_id: WorkbookId, sheet_id: SheetId, rect: Rect) -> Self {
        Self {
            workbook_id,
            scope: SheetScope::Single(sheet_id),
            areas: vec![rect],
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct FormulaSource {
    pub text: String,
    pub is_r1c1: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct WorkbookModel {
    pub id: WorkbookId,
    pub display_name: String,
    pub format: FileFormat,
    #[serde(default)]
    pub date1904: bool,
    #[serde(default)]
    pub is_addin: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
pub enum SheetVisibility {
    #[default]
    Visible,
    Hidden,
    VeryHidden,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
pub enum SheetKind {
    #[default]
    Worksheet,
    ChartSheet,
    MacroSheet,
    DialogSheet,
}

impl SheetKind {
    pub fn is_worksheet(&self) -> bool {
        *self == Self::Worksheet
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
pub enum ObjectPlacement {
    MoveAndSize,
    MoveOnly,
    FreeFloating,
    #[default]
    Unknown,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum DrawingAnchor {
    TwoCell(TwoCellAnchor),
    OneCell(OneCellAnchor),
    Absolute(AbsoluteAnchor),
    UnsupportedRaw,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct TwoCellAnchor {
    pub from: CellMarker,
    pub to: CellMarker,
    pub edit_as: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct OneCellAnchor {
    pub from: CellMarker,
    pub extents: SizeEmu,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct AbsoluteAnchor {
    pub position: PointEmu,
    pub extents: SizeEmu,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct CellMarker {
    pub col_zero_based: u32,
    pub col_offset: Emu,
    pub row_zero_based: u32,
    pub row_offset: Emu,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct PointEmu {
    pub x: Emu,
    pub y: Emu,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct SizeEmu {
    pub cx: Emu,
    pub cy: Emu,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct WorksheetModel {
    pub id: SheetId,
    pub workbook_id: WorkbookId,
    pub name: String,
    #[serde(default, skip_serializing_if = "SheetKind::is_worksheet")]
    pub kind: SheetKind,
    #[serde(default)]
    pub visibility: SheetVisibility,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub relationship_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub part_uri: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct OpaquePart {
    pub uri: String,
    pub content_type: String,
    pub bytes: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct OpenWorkbookSpec {
    pub bytes: Vec<u8>,
    pub format_hint: Option<FileFormat>,
    pub profile: ExcelProfile,
    pub read_only: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct SaveWorkbookSpec {
    pub format: FileFormat,
    pub profile: ExcelProfile,
    pub lossless: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct GetRangeValuesSpec {
    pub workbook: WorkbookHandle,
    pub range: RangeRef,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SetRangeValuesSpec {
    pub workbook: WorkbookHandle,
    pub range: RangeRef,
    pub values: OmArray,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct LoadOptions {
    pub profile: ExcelProfile,
    pub preserve_unknown_parts: bool,
    pub read_calc_chain: bool,
}

impl Default for LoadOptions {
    fn default() -> Self {
        Self {
            profile: ExcelProfile::default(),
            preserve_unknown_parts: true,
            read_calc_chain: true,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ActiveContentPolicy {
    #[default]
    Preserve,
    Refuse,
    Strip,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ActiveContentKind {
    VbaProject,
    VbaProjectSignature,
    VbaData,
    XlmMacroSheet,
    DialogSheet,
    ActiveXControl,
    EmbeddedObject,
    CustomUi,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ActiveContentContentTypeEntryKind {
    Default,
    Override,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ActiveContentRemovedPart {
    pub part_uri: String,
    pub content_type: Option<String>,
    pub byte_len: u64,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ActiveContentRemovedRelationship {
    pub relationship_part_uri: String,
    pub owner_part_uri: Option<String>,
    pub id: String,
    pub relationship_type: String,
    pub target: String,
    pub target_mode: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ActiveContentRemovedContentTypeEntry {
    pub entry_kind: ActiveContentContentTypeEntryKind,
    pub selector: String,
    pub content_type: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ActiveContentAuditManifest {
    pub policy: ActiveContentPolicy,
    pub detected_kinds: Vec<ActiveContentKind>,
    pub removed_parts: Vec<ActiveContentRemovedPart>,
    pub removed_relationships: Vec<ActiveContentRemovedRelationship>,
    pub removed_content_type_entries: Vec<ActiveContentRemovedContentTypeEntry>,
    pub rewritten_owner_part_uris: Vec<String>,
    pub retained_shared_part_uris: Vec<String>,
}

impl ActiveContentAuditManifest {
    pub fn observed(
        policy: ActiveContentPolicy,
        mut detected_kinds: Vec<ActiveContentKind>,
    ) -> Self {
        detected_kinds.sort();
        detected_kinds.dedup();
        Self {
            policy,
            detected_kinds,
            removed_parts: Vec::new(),
            removed_relationships: Vec::new(),
            removed_content_type_entries: Vec::new(),
            rewritten_owner_part_uris: Vec::new(),
            retained_shared_part_uris: Vec::new(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ExternalDataPolicy {
    #[default]
    OfflinePreserve,
    Refuse,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ExternalDataKind {
    ExternalLink,
    ExternalWorkbook,
    DdeLink,
    OleLink,
    Connection,
    QueryTable,
    DataModel,
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ExternalDataRelationship {
    pub relationship_part_uri: String,
    pub id: String,
    pub relationship_type: String,
    pub target: String,
    pub target_mode: Option<String>,
    pub kinds: Vec<ExternalDataKind>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ExternalDataInventory {
    kinds: Vec<ExternalDataKind>,
    part_uris: Vec<String>,
    relationship_part_uris: Vec<String>,
    relationships: Vec<ExternalDataRelationship>,
    has_content_type_markers: bool,
}

impl ExternalDataInventory {
    pub fn new(
        mut kinds: Vec<ExternalDataKind>,
        mut part_uris: Vec<String>,
        mut relationships: Vec<ExternalDataRelationship>,
        has_content_type_markers: bool,
    ) -> Self {
        kinds.sort();
        kinds.dedup();
        part_uris.sort_by_key(|part_uri| part_uri.to_ascii_lowercase());
        part_uris.dedup_by(|left, right| left.eq_ignore_ascii_case(right));
        for relationship in &mut relationships {
            relationship.kinds.sort();
            relationship.kinds.dedup();
        }
        relationships.sort();
        relationships.dedup();
        let mut relationship_part_uris = relationships
            .iter()
            .map(|relationship| relationship.relationship_part_uri.clone())
            .collect::<Vec<_>>();
        relationship_part_uris.sort_by_key(|part_uri| part_uri.to_ascii_lowercase());
        relationship_part_uris.dedup_by(|left, right| left.eq_ignore_ascii_case(right));
        Self {
            kinds,
            part_uris,
            relationship_part_uris,
            relationships,
            has_content_type_markers,
        }
    }

    pub fn has_artifacts(&self) -> bool {
        !self.kinds.is_empty()
            || !self.part_uris.is_empty()
            || !self.relationships.is_empty()
            || self.has_content_type_markers
    }

    pub fn kinds(&self) -> &[ExternalDataKind] {
        &self.kinds
    }

    pub fn part_uris(&self) -> &[String] {
        &self.part_uris
    }

    pub fn relationship_part_uris(&self) -> &[String] {
        &self.relationship_part_uris
    }

    pub fn relationships(&self) -> &[ExternalDataRelationship] {
        &self.relationships
    }

    pub fn has_content_type_markers(&self) -> bool {
        self.has_content_type_markers
    }

    pub fn merged_with(&self, other: &Self) -> Self {
        Self::new(
            self.kinds.iter().chain(&other.kinds).copied().collect(),
            self.part_uris
                .iter()
                .chain(&other.part_uris)
                .cloned()
                .collect(),
            self.relationships
                .iter()
                .chain(&other.relationships)
                .cloned()
                .collect(),
            self.has_content_type_markers || other.has_content_type_markers,
        )
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ExternalDataAccessReport {
    pub policy: ExternalDataPolicy,
    pub inventory: ExternalDataInventory,
    pub link_update_attempted: bool,
    pub refresh_attempted: bool,
    pub external_access_attempted: bool,
}

impl ExternalDataAccessReport {
    pub fn offline(policy: ExternalDataPolicy, inventory: ExternalDataInventory) -> Self {
        Self {
            policy,
            inventory,
            link_update_attempted: false,
            refresh_attempted: false,
            external_access_attempted: false,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct SaveOptions {
    pub profile: ExcelProfile,
    pub lossless: bool,
    #[serde(default)]
    pub active_content_policy: ActiveContentPolicy,
}

impl Default for SaveOptions {
    fn default() -> Self {
        Self {
            profile: ExcelProfile::default(),
            lossless: true,
            active_content_policy: ActiveContentPolicy::default(),
        }
    }
}

impl OmError {
    pub fn type_mismatch(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::TypeMismatch, message)
    }
}

#[cfg(test)]
mod tests {
    use super::{
        ActiveContentPolicy, CellValue, Emu, ExcelDateSystem, ExcelProfile,
        ExternalDataAccessReport, ExternalDataInventory, ExternalDataPolicy, IsoDateTime,
        IsoDateTimeOffsetPolicy, LoadOptions, ObjectHandle, OmArray, OmErrorCode, OmValue, Points,
        RangeRef, Rect, RichTextSource, RichTextValue, SaveOptions, SheetId, SheetScope,
        WorkbookId,
    };

    #[test]
    fn om_array_rejects_invalid_dimensions() {
        let error =
            OmArray::new(2, 2, vec![OmValue::Empty]).expect_err("invalid dimensions should fail");
        assert_eq!(error.code, OmErrorCode::InvalidArgument);
    }

    #[test]
    fn rect_single_cell_has_unit_extent() {
        let rect = Rect::single_cell(3, 7);
        assert_eq!(rect.width(), 1);
        assert_eq!(rect.height(), 1);
        assert_eq!(rect.row_first, 3);
        assert_eq!(rect.col_first, 7);
    }

    #[test]
    fn range_ref_single_cell_uses_single_sheet_scope() {
        let range = RangeRef::single_cell(WorkbookId(1), SheetId(2), 4, 5);
        assert_eq!(range.areas, vec![Rect::single_cell(4, 5)]);
        assert_eq!(range.scope, SheetScope::Single(SheetId(2)));
    }

    #[test]
    fn cell_value_round_trips_through_om_value() {
        let value = CellValue::Text("hello".to_string());
        let om_value = OmValue::from(value.clone());
        let restored = CellValue::try_from(om_value).expect("convert back to cell");

        assert_eq!(restored, value);
    }

    #[test]
    fn non_finite_numbers_cannot_become_cell_values() {
        for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            let error = CellValue::try_from(OmValue::Number(value))
                .expect_err("non-finite worksheet number must fail");
            assert_eq!(error.code, OmErrorCode::InvalidArgument);
            assert_eq!(error.message, "worksheet cell numeric value must be finite");
        }
    }

    #[test]
    fn iso_date_time_preserves_valid_date_time_and_offset_lexicals() {
        for lexical in [
            "2024-02-29",
            "2026-08-11Z",
            "2026-08-11+09:00",
            "2026-08-11T12:34:56",
            "2026-08-11T12:34:56.1200Z",
            "2026-08-11T12:34:56.1200-09:30",
        ] {
            let value = IsoDateTime::parse(lexical).expect("valid ISO date/time");
            assert_eq!(value.as_str(), lexical);
        }
    }

    #[test]
    fn iso_date_time_rejects_invalid_calendar_time_and_offset_lexicals() {
        for lexical in [
            "",
            "0000-01-01",
            "2023-02-29",
            "2024-13-01",
            "2024-04-31",
            "2024-01-01T24:00:00",
            "2024-01-01T12:60:00",
            "2024-01-01T12:00:60",
            "2024-01-01T12:00",
            "2024-01-01T12:00:00.",
            "2024-01-01T12:00:00+14:01",
            " 2024-01-01",
        ] {
            let error = IsoDateTime::parse(lexical).expect_err("invalid ISO date/time must fail");
            assert_eq!(error.code, OmErrorCode::InvalidArgument, "{lexical}");
        }
    }

    #[test]
    fn iso_date_time_cell_projects_to_raw_text_without_serial_conversion() {
        let lexical = "2026-08-11T12:34:56+09:00";
        let value = CellValue::IsoDateTime(IsoDateTime::parse(lexical).expect("valid date/time"));

        assert_eq!(OmValue::from(value), OmValue::Text(lexical.to_string()));
    }

    #[test]
    fn iso_date_time_serial_conversion_requires_date_system_and_offset_policy() {
        let serial_1900 = |lexical: &str| {
            IsoDateTime::parse(lexical)
                .expect("valid date/time")
                .to_excel_serial(
                    ExcelDateSystem::Date1900,
                    IsoDateTimeOffsetPolicy::PreserveWallClock,
                )
        };
        assert_eq!(serial_1900("1900-01-01"), 1.0);
        assert_eq!(serial_1900("1900-02-28"), 59.0);
        assert_eq!(serial_1900("1900-03-01"), 61.0);
        assert_eq!(
            serial_1900("2026-08-11T12:00:00"),
            serial_1900("2026-08-11") + 0.5
        );

        let offset =
            IsoDateTime::parse("1904-01-01T00:00:00+09:00").expect("valid offset date/time");
        assert_eq!(
            offset.to_excel_serial(
                ExcelDateSystem::Date1904,
                IsoDateTimeOffsetPolicy::PreserveWallClock,
            ),
            0.0
        );
        assert_eq!(
            offset.to_excel_serial(
                ExcelDateSystem::Date1904,
                IsoDateTimeOffsetPolicy::NormalizeToUtc,
            ),
            -0.375
        );
    }

    #[test]
    fn rich_text_projects_display_text_and_shares_preserved_xml_on_clone() {
        let spreadsheet_namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let raw_xml = b"<si><r><t>display</t></r><rPh><t>phonetic</t></rPh></si>".to_vec();
        let value = RichTextValue::try_from_preserved_ooxml(
            raw_xml.clone(),
            4,
            raw_xml.len() - 5,
            Vec::new(),
            vec![("xmlns".to_string(), spreadsheet_namespace.to_string())],
            RichTextSource::SharedString,
            spreadsheet_namespace.to_string(),
        )
        .expect("valid preserved rich-text XML");
        let cloned = value.clone();

        assert_eq!(value.as_str(), "display");
        assert_eq!(value.phonetic_text(), "phonetic");
        assert_eq!(value.source(), RichTextSource::SharedString);
        assert_eq!(value.spreadsheet_namespace(), spreadsheet_namespace);
        assert!(value.requires_raw_preservation());
        assert_eq!(value.raw_item_xml(), raw_xml);
        assert_eq!(value.raw_inner_xml(), &raw_xml[4..raw_xml.len() - 5]);
        assert_eq!(
            OmValue::from(CellValue::RichText(value.clone())),
            OmValue::Text("display".to_string())
        );
        assert!(std::ptr::eq(
            value.raw_item_xml().as_ptr(),
            cloned.raw_item_xml().as_ptr()
        ));
    }

    #[test]
    fn rich_text_rejects_invalid_preserved_xml_bounds() {
        let error = RichTextValue::try_from_preserved_ooxml(
            b"<is><t>display</t></is>".to_vec(),
            10,
            5,
            Vec::new(),
            Vec::new(),
            RichTextSource::InlineString,
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
        )
        .expect_err("reversed rich-text bounds must fail");

        assert_eq!(error.code, OmErrorCode::InvalidArgument);
    }

    #[test]
    fn rich_text_deserialization_rejects_invalid_preserved_xml_bounds() {
        let spreadsheet_namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let raw_xml = b"<is><t>display</t></is>".to_vec();
        let value = RichTextValue::try_from_preserved_ooxml(
            raw_xml.clone(),
            4,
            raw_xml.len() - 5,
            Vec::new(),
            vec![("xmlns".to_string(), spreadsheet_namespace.to_string())],
            RichTextSource::InlineString,
            spreadsheet_namespace.to_string(),
        )
        .expect("valid preserved rich-text XML");
        let mut serialized = serde_json::to_value(value).expect("serialize rich text");
        serialized["preserved_xml"]["inner_end"] = serde_json::json!(usize::MAX);

        let error = serde_json::from_value::<RichTextValue>(serialized)
            .expect_err("invalid serialized rich-text bounds must fail");
        assert!(
            error
                .to_string()
                .contains("invalid preserved rich-text OOXML metadata")
        );
    }

    #[test]
    fn rich_text_deserialization_rejects_projection_drift() {
        let spreadsheet_namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let raw_xml = b"<si><r><t>display</t></r><rPh><t>phonetic</t></rPh></si>".to_vec();
        let value = RichTextValue::try_from_preserved_ooxml(
            raw_xml.clone(),
            4,
            raw_xml.len() - 5,
            Vec::new(),
            vec![("xmlns".to_string(), spreadsheet_namespace.to_string())],
            RichTextSource::SharedString,
            spreadsheet_namespace.to_string(),
        )
        .expect("valid preserved rich-text XML");
        let mut serialized = serde_json::to_value(value).expect("serialize rich text");
        serialized["display_text"] = serde_json::json!("forged");

        let error = serde_json::from_value::<RichTextValue>(serialized)
            .expect_err("projection drift must fail deserialization");
        assert!(
            error
                .to_string()
                .contains("preserved rich-text projection does not match its OOXML payload")
        );
    }

    #[test]
    fn rich_text_rejects_non_token_inner_bounds_and_root_metadata_drift() {
        let spreadsheet_namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let raw_xml = b"<si data-owner=\"source\"><t>display</t></si>".to_vec();
        let error = RichTextValue::try_from_preserved_ooxml(
            raw_xml.clone(),
            5,
            raw_xml.len() - 5,
            vec![("data-owner".to_string(), "source".to_string())],
            vec![("xmlns".to_string(), spreadsheet_namespace.to_string())],
            RichTextSource::SharedString,
            spreadsheet_namespace.to_string(),
        )
        .expect_err("non-token inner bound must fail");
        assert_eq!(error.code, OmErrorCode::InvalidArgument);

        let error = RichTextValue::try_from_preserved_ooxml(
            raw_xml.clone(),
            raw_xml
                .iter()
                .position(|byte| *byte == b'>')
                .expect("start-tag end")
                + 1,
            raw_xml.len() - 5,
            vec![("data-owner".to_string(), "forged".to_string())],
            vec![("xmlns".to_string(), spreadsheet_namespace.to_string())],
            RichTextSource::SharedString,
            spreadsheet_namespace.to_string(),
        )
        .expect_err("root metadata drift must fail");
        assert_eq!(error.code, OmErrorCode::InvalidArgument);

        let error = RichTextValue::try_from_preserved_ooxml(
            raw_xml.clone(),
            raw_xml
                .iter()
                .position(|byte| *byte == b'>')
                .expect("start-tag end")
                + 1,
            raw_xml.len() - 5,
            vec![("bad name".to_string(), "source".to_string())],
            vec![("xmlns".to_string(), spreadsheet_namespace.to_string())],
            RichTextSource::SharedString,
            spreadsheet_namespace.to_string(),
        )
        .expect_err("invalid attribute QName must fail before raw splice");
        assert_eq!(error.code, OmErrorCode::InvalidArgument);
    }

    #[test]
    fn rich_text_rejects_source_mismatch_and_markup_outside_the_item() {
        let spreadsheet_namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let raw_xml = b"<si><t>display</t></si>".to_vec();
        let error = RichTextValue::try_from_preserved_ooxml(
            raw_xml.clone(),
            4,
            raw_xml.len() - 5,
            Vec::new(),
            vec![("xmlns".to_string(), spreadsheet_namespace.to_string())],
            RichTextSource::InlineString,
            spreadsheet_namespace.to_string(),
        )
        .expect_err("source item mismatch must fail");
        assert_eq!(error.code, OmErrorCode::InvalidArgument);

        let raw_xml = b"<si><t>display</t></si><!--outside-->".to_vec();
        let error = RichTextValue::try_from_preserved_ooxml(
            raw_xml.clone(),
            4,
            b"<si><t>display</t>".len(),
            Vec::new(),
            vec![("xmlns".to_string(), spreadsheet_namespace.to_string())],
            RichTextSource::SharedString,
            spreadsheet_namespace.to_string(),
        )
        .expect_err("markup outside the source item must fail");
        assert_eq!(error.code, OmErrorCode::InvalidArgument);
    }

    #[test]
    fn rich_text_rejects_undeclared_element_and_attribute_prefixes() {
        let spreadsheet_namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let raw_xml = b"<si bad:owner=\"x\"><t>display</t></si>".to_vec();
        let error = RichTextValue::try_from_preserved_ooxml(
            raw_xml.clone(),
            raw_xml
                .iter()
                .position(|byte| *byte == b'>')
                .expect("start-tag end")
                + 1,
            raw_xml.len() - 5,
            vec![("bad:owner".to_string(), "x".to_string())],
            vec![("xmlns".to_string(), spreadsheet_namespace.to_string())],
            RichTextSource::SharedString,
            spreadsheet_namespace.to_string(),
        )
        .expect_err("undeclared root attribute prefix must fail");
        assert_eq!(error.code, OmErrorCode::InvalidArgument);

        let raw_xml = b"<si><t>display</t><bad:meta/></si>".to_vec();
        let error = RichTextValue::try_from_preserved_ooxml(
            raw_xml.clone(),
            4,
            raw_xml.len() - 5,
            Vec::new(),
            vec![("xmlns".to_string(), spreadsheet_namespace.to_string())],
            RichTextSource::SharedString,
            spreadsheet_namespace.to_string(),
        )
        .expect_err("undeclared child element prefix must fail");
        assert_eq!(error.code, OmErrorCode::InvalidArgument);
    }

    #[test]
    fn rich_text_rejects_invalid_child_qnames_and_entity_references() {
        let spreadsheet_namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        for raw_xml in [
            b"<si><1bad/></si>".as_slice(),
            b"<si><meta>&future;</meta></si>".as_slice(),
            b"<si><t>&#0;</t></si>".as_slice(),
        ] {
            let raw_xml = raw_xml.to_vec();
            let error = RichTextValue::try_from_preserved_ooxml(
                raw_xml.clone(),
                4,
                raw_xml.len() - 5,
                Vec::new(),
                vec![("xmlns".to_string(), spreadsheet_namespace.to_string())],
                RichTextSource::SharedString,
                spreadsheet_namespace.to_string(),
            )
            .expect_err("invalid raw rich-text markup must fail");
            assert_eq!(error.code, OmErrorCode::InvalidArgument);
        }
    }

    #[test]
    fn om_value_array_cannot_become_cell_value() {
        let error = CellValue::try_from(OmValue::Array(OmArray::scalar(OmValue::Number(1.0))))
            .expect_err("array coercion should fail");
        assert_eq!(error.code, OmErrorCode::TypeMismatch);
    }

    #[test]
    fn om_value_object_cannot_become_cell_value() {
        let error = CellValue::try_from(OmValue::Object(ObjectHandle(7)))
            .expect_err("object coercion should fail");
        assert_eq!(error.code, OmErrorCode::TypeMismatch);
    }

    #[test]
    fn missing_empty_and_null_om_values_coerce_to_blank_cells() {
        assert_eq!(
            CellValue::try_from(OmValue::Missing).expect("missing"),
            CellValue::Blank
        );
        assert_eq!(
            CellValue::try_from(OmValue::Empty).expect("empty"),
            CellValue::Blank
        );
        assert_eq!(
            CellValue::try_from(OmValue::Null).expect("null"),
            CellValue::Blank
        );
    }

    #[test]
    fn om_array_get_returns_none_for_out_of_bounds_indices() {
        let array = OmArray::new(
            2,
            2,
            vec![
                OmValue::Number(1.0),
                OmValue::Number(2.0),
                OmValue::Number(3.0),
                OmValue::Number(4.0),
            ],
        )
        .expect("array");

        assert_eq!(array.get(0, 0), Some(&OmValue::Number(1.0)));
        assert_eq!(array.get(2, 0), None);
        assert_eq!(array.get(0, 2), None);
    }

    #[test]
    fn default_load_and_save_options_preserve_lossless_profile_defaults() {
        let load = LoadOptions::default();
        let save = SaveOptions::default();

        assert_eq!(load.profile, ExcelProfile::Excel365);
        assert!(load.preserve_unknown_parts);
        assert!(load.read_calc_chain);
        assert_eq!(save.profile, ExcelProfile::Excel365);
        assert!(save.lossless);
        assert_eq!(save.active_content_policy, ActiveContentPolicy::Preserve);
        let external_data_report = ExternalDataAccessReport::offline(
            ExternalDataPolicy::default(),
            ExternalDataInventory::default(),
        );
        assert_eq!(
            external_data_report.policy,
            ExternalDataPolicy::OfflinePreserve
        );
        assert!(!external_data_report.inventory.has_artifacts());
        assert!(!external_data_report.external_access_attempted);
    }

    #[test]
    fn emu_and_points_convert_with_excel_geometry_scale() {
        assert_eq!(Emu(12_700).to_points(), Points(1.0));
        assert_eq!(Points(1.5).to_emu(), Emu(19_050));
    }
}
