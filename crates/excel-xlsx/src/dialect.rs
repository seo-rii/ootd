use std::io::Cursor;

use office_common::{FileFormat, OmError, OmErrorCode, OmResult};
use quick_xml::NsReader;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;

pub(crate) const TRANSITIONAL_SPREADSHEETML_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const STRICT_SPREADSHEETML_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(crate) const TRANSITIONAL_DRAWINGML_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(crate) const STRICT_DRAWINGML_NAMESPACE: &str = "http://purl.oclc.org/ooxml/drawingml/main";
pub(crate) const TRANSITIONAL_OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const STRICT_OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";

pub(crate) const TRANSITIONAL_OFFICE_DOCUMENT_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
pub(crate) const TRANSITIONAL_CALC_CHAIN_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
pub(crate) const TRANSITIONAL_WORKSHEET_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
pub(crate) const TRANSITIONAL_CHARTSHEET_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
pub(crate) const TRANSITIONAL_DIALOGSHEET_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
pub(crate) const TRANSITIONAL_DRAWING_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
pub(crate) const TRANSITIONAL_CHART_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
pub(crate) const TRANSITIONAL_STYLES_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
pub(crate) const TRANSITIONAL_THEME_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
pub(crate) const TRANSITIONAL_SHARED_STRINGS_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
pub(crate) const TRANSITIONAL_HYPERLINK_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
pub(crate) const TRANSITIONAL_COMMENTS_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
pub(crate) const TRANSITIONAL_VML_DRAWING_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
pub(crate) const TRANSITIONAL_TABLE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

const STRICT_OFFICE_DOCUMENT_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
const STRICT_CALC_CHAIN_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/calcChain";
const STRICT_WORKSHEET_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet";
const STRICT_CHARTSHEET_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
const STRICT_DRAWING_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/drawing";
const STRICT_CHART_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chart";
const STRICT_STYLES_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/styles";
const STRICT_THEME_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/theme";
const STRICT_SHARED_STRINGS_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings";
const STRICT_HYPERLINK_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink";
const STRICT_COMMENTS_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/comments";
const STRICT_TABLE_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/table";

const XLSX_MAIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const STRICT_MAIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.main+xml";
const XLSM_MAIN_CONTENT_TYPE: &str = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
const XLTX_MAIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
const XLTM_MAIN_CONTENT_TYPE: &str = "application/vnd.ms-excel.template.macroEnabled.main+xml";

/// OOXML namespace and relationship dialect used by a loaded workbook package.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum OoxmlDialect {
    #[default]
    Transitional,
    Strict,
}

impl OoxmlDialect {
    pub fn spreadsheetml_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_SPREADSHEETML_NAMESPACE,
            Self::Strict => STRICT_SPREADSHEETML_NAMESPACE,
        }
    }

    pub fn drawingml_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_DRAWINGML_NAMESPACE,
            Self::Strict => STRICT_DRAWINGML_NAMESPACE,
        }
    }

    pub fn office_document_relationships_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE,
            Self::Strict => STRICT_OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum OoxmlRelationshipKind {
    OfficeDocument,
    Worksheet,
    ChartSheet,
    DialogSheet,
    Drawing,
    Chart,
    Styles,
    Theme,
    SharedStrings,
    CalcChain,
    Hyperlink,
    Comments,
    VmlDrawing,
    Table,
}

const RELATIONSHIP_TYPES: &[(OoxmlRelationshipKind, &str, Option<&str>)] = &[
    (
        OoxmlRelationshipKind::OfficeDocument,
        TRANSITIONAL_OFFICE_DOCUMENT_RELATIONSHIP_TYPE,
        Some(STRICT_OFFICE_DOCUMENT_RELATIONSHIP_TYPE),
    ),
    (
        OoxmlRelationshipKind::Worksheet,
        TRANSITIONAL_WORKSHEET_RELATIONSHIP_TYPE,
        Some(STRICT_WORKSHEET_RELATIONSHIP_TYPE),
    ),
    (
        OoxmlRelationshipKind::ChartSheet,
        TRANSITIONAL_CHARTSHEET_RELATIONSHIP_TYPE,
        Some(STRICT_CHARTSHEET_RELATIONSHIP_TYPE),
    ),
    (
        OoxmlRelationshipKind::DialogSheet,
        TRANSITIONAL_DIALOGSHEET_RELATIONSHIP_TYPE,
        None,
    ),
    (
        OoxmlRelationshipKind::Drawing,
        TRANSITIONAL_DRAWING_RELATIONSHIP_TYPE,
        Some(STRICT_DRAWING_RELATIONSHIP_TYPE),
    ),
    (
        OoxmlRelationshipKind::Chart,
        TRANSITIONAL_CHART_RELATIONSHIP_TYPE,
        Some(STRICT_CHART_RELATIONSHIP_TYPE),
    ),
    (
        OoxmlRelationshipKind::Styles,
        TRANSITIONAL_STYLES_RELATIONSHIP_TYPE,
        Some(STRICT_STYLES_RELATIONSHIP_TYPE),
    ),
    (
        OoxmlRelationshipKind::Theme,
        TRANSITIONAL_THEME_RELATIONSHIP_TYPE,
        Some(STRICT_THEME_RELATIONSHIP_TYPE),
    ),
    (
        OoxmlRelationshipKind::SharedStrings,
        TRANSITIONAL_SHARED_STRINGS_RELATIONSHIP_TYPE,
        Some(STRICT_SHARED_STRINGS_RELATIONSHIP_TYPE),
    ),
    (
        OoxmlRelationshipKind::CalcChain,
        TRANSITIONAL_CALC_CHAIN_RELATIONSHIP_TYPE,
        Some(STRICT_CALC_CHAIN_RELATIONSHIP_TYPE),
    ),
    (
        OoxmlRelationshipKind::Hyperlink,
        TRANSITIONAL_HYPERLINK_RELATIONSHIP_TYPE,
        Some(STRICT_HYPERLINK_RELATIONSHIP_TYPE),
    ),
    (
        OoxmlRelationshipKind::Comments,
        TRANSITIONAL_COMMENTS_RELATIONSHIP_TYPE,
        Some(STRICT_COMMENTS_RELATIONSHIP_TYPE),
    ),
    (
        OoxmlRelationshipKind::VmlDrawing,
        TRANSITIONAL_VML_DRAWING_RELATIONSHIP_TYPE,
        None,
    ),
    (
        OoxmlRelationshipKind::Table,
        TRANSITIONAL_TABLE_RELATIONSHIP_TYPE,
        Some(STRICT_TABLE_RELATIONSHIP_TYPE),
    ),
];

pub(crate) fn relationship_type(
    dialect: OoxmlDialect,
    kind: OoxmlRelationshipKind,
) -> Option<&'static str> {
    RELATIONSHIP_TYPES
        .iter()
        .find(|(candidate, _, _)| *candidate == kind)
        .and_then(|(_, transitional, strict)| match dialect {
            OoxmlDialect::Transitional => Some(*transitional),
            OoxmlDialect::Strict => *strict,
        })
}

pub(crate) fn classify_relationship_type(
    relationship_type: &str,
) -> Option<(OoxmlDialect, OoxmlRelationshipKind)> {
    RELATIONSHIP_TYPES
        .iter()
        .find_map(|(kind, transitional, strict)| {
            if relationship_type == *transitional {
                Some((OoxmlDialect::Transitional, *kind))
            } else if strict.is_some_and(|strict| relationship_type == strict) {
                Some((OoxmlDialect::Strict, *kind))
            } else {
                None
            }
        })
}

pub(crate) fn relationship_type_is(
    dialect: OoxmlDialect,
    actual: &str,
    kind: OoxmlRelationshipKind,
) -> bool {
    relationship_type(dialect, kind) == Some(actual)
}

pub(crate) fn validate_known_relationship_dialect<'a>(
    dialect: OoxmlDialect,
    relationship_types: impl IntoIterator<Item = &'a str>,
    context: &str,
) -> OmResult<()> {
    for relationship_type in relationship_types {
        if let Some((actual_dialect, kind)) = classify_relationship_type(relationship_type)
            && actual_dialect != dialect
        {
            return Err(OmError::parse(format!(
                "{context} contains a {actual_dialect:?} {kind:?} relationship in a {dialect:?} workbook: {relationship_type}"
            )));
        }
    }
    Ok(())
}

pub(crate) fn workbook_format(
    dialect: OoxmlDialect,
    content_type: Option<&str>,
    workbook_part_uri: &str,
) -> OmResult<FileFormat> {
    let content_type = content_type.ok_or_else(|| {
        OmError::parse(format!(
            "workbook main part {workbook_part_uri} has no resolved content type"
        ))
    })?;
    match (dialect, content_type) {
        (OoxmlDialect::Transitional, XLSX_MAIN_CONTENT_TYPE) => Ok(FileFormat::Xlsx),
        (OoxmlDialect::Transitional, XLSM_MAIN_CONTENT_TYPE) => Ok(FileFormat::Xlsm),
        (OoxmlDialect::Transitional, XLTX_MAIN_CONTENT_TYPE) => Ok(FileFormat::Xltx),
        (OoxmlDialect::Transitional, XLTM_MAIN_CONTENT_TYPE) => Ok(FileFormat::Xltm),
        (OoxmlDialect::Strict, XLSX_MAIN_CONTENT_TYPE | STRICT_MAIN_CONTENT_TYPE) => {
            Ok(FileFormat::StrictXlsx)
        }
        (OoxmlDialect::Strict, XLTX_MAIN_CONTENT_TYPE) => Err(OmError::unsupported(
            "Strict OOXML template packages are not represented by the current FileFormat API",
        )),
        (OoxmlDialect::Strict, XLSM_MAIN_CONTENT_TYPE | XLTM_MAIN_CONTENT_TYPE) => {
            Err(OmError::unsupported(
                "macro-enabled packages cannot use the Strict OOXML relationship dialect",
            ))
        }
        _ => Err(OmError::new(
            OmErrorCode::Parse,
            format!("unsupported workbook main content type for {dialect:?} OOXML: {content_type}"),
        )),
    }
}

pub(crate) fn validate_xml_root_namespace(
    xml: &[u8],
    expected_local_name: &[u8],
    expected_namespace: &str,
    part_uri: &str,
) -> OmResult<()> {
    let mut reader = NsReader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader.read_resolved_event_into(&mut buffer) {
            Ok((namespace, Event::Start(element) | Event::Empty(element))) => {
                let actual_local_name = element.local_name();
                let actual_namespace = match namespace {
                    ResolveResult::Bound(namespace) => {
                        String::from_utf8_lossy(namespace.as_ref()).into_owned()
                    }
                    ResolveResult::Unbound => "<unbound>".to_string(),
                    ResolveResult::Unknown(prefix) => format!(
                        "<unknown prefix {}>",
                        String::from_utf8_lossy(prefix.as_ref())
                    ),
                };
                if actual_local_name.as_ref() != expected_local_name
                    || actual_namespace != expected_namespace
                {
                    return Err(OmError::parse(format!(
                        "{part_uri} root must be {{{expected_namespace}}}{} but found {{{actual_namespace}}}{}",
                        String::from_utf8_lossy(expected_local_name),
                        String::from_utf8_lossy(actual_local_name.as_ref())
                    )));
                }
                return Ok(());
            }
            Ok((_, Event::Eof)) => {
                return Err(OmError::parse(format!(
                    "{part_uri} does not contain a root element"
                )));
            }
            Ok(_) => {}
            Err(error) => return Err(OmError::parse(error.to_string())),
        }
        buffer.clear();
    }
}
