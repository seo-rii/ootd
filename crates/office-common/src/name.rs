use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};

use crate::{FormulaSource, SheetId};

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub enum NameScope {
    Workbook,
    Worksheet(SheetId),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct DefinedNameId(pub u32);

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct NameKey {
    pub scope: NameScope,
    pub canonical_name: String,
}

impl NameKey {
    pub fn new(scope: NameScope, name: &str) -> Self {
        Self {
            scope,
            canonical_name: canonicalize_excel_name(name),
        }
    }
}

pub fn canonicalize_excel_name(name: &str) -> String {
    name.to_ascii_uppercase()
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum NameValidationMode {
    StrictExcel,
    PreserveLoadedInvalid,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum BuiltinName {
    PrintArea,
    PrintTitles,
    FilterDatabase,
    AutoOpen,
    AutoClose,
    Extract,
    Criteria,
    Database,
    Other(String),
}

impl BuiltinName {
    pub fn from_excel_name(name: &str) -> Option<Self> {
        match canonicalize_excel_name(name).as_str() {
            "_XLNM.PRINT_AREA" => Some(Self::PrintArea),
            "_XLNM.PRINT_TITLES" => Some(Self::PrintTitles),
            "_XLNM._FILTERDATABASE" => Some(Self::FilterDatabase),
            "_XLNM.AUTO_OPEN" => Some(Self::AutoOpen),
            "_XLNM.AUTO_CLOSE" => Some(Self::AutoClose),
            "_XLNM.EXTRACT" => Some(Self::Extract),
            "_XLNM.CRITERIA" => Some(Self::Criteria),
            "_XLNM.DATABASE" => Some(Self::Database),
            value if value.starts_with("_XLNM.") => Some(Self::Other(name.to_string())),
            _ => None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum DefinedNameKind {
    Unknown,
    Range,
    Formula,
    Constant,
    External,
    Builtin,
    Unsupported,
}

#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct DefinedNameMetadata {
    pub hidden: bool,
    pub builtin: Option<BuiltinName>,
    pub function: bool,
    pub vb_procedure: bool,
    pub xlm: bool,
    #[serde(default)]
    pub function_group_id: Option<u32>,
    #[serde(default)]
    pub shortcut_key: Option<String>,
    pub workbook_parameter: bool,
    pub description: Option<String>,
    pub comment: Option<String>,
    pub custom_xml_attrs: BTreeMap<String, String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct DefinedName {
    pub id: DefinedNameId,
    pub display_name: String,
    pub canonical_name: String,
    pub scope: NameScope,
    pub refers_to: FormulaSource,
    pub target_kind: DefinedNameKind,
    pub metadata: DefinedNameMetadata,
}

impl DefinedName {
    pub fn new(
        id: DefinedNameId,
        scope: NameScope,
        display_name: impl Into<String>,
        refers_to: FormulaSource,
    ) -> Self {
        let display_name = display_name.into();
        let builtin = BuiltinName::from_excel_name(&display_name);
        let target_kind = if builtin.is_some() {
            DefinedNameKind::Builtin
        } else {
            DefinedNameKind::Unknown
        };
        Self {
            id,
            canonical_name: canonicalize_excel_name(&display_name),
            display_name,
            scope,
            refers_to,
            target_kind,
            metadata: DefinedNameMetadata {
                builtin,
                ..DefinedNameMetadata::default()
            },
        }
    }

    pub fn key(&self) -> NameKey {
        NameKey {
            scope: self.scope,
            canonical_name: self.canonical_name.clone(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{
        BuiltinName, DefinedName, DefinedNameId, NameKey, NameScope, canonicalize_excel_name,
    };
    use crate::{FormulaSource, SheetId};

    #[test]
    fn name_key_canonicalizes_case() {
        let key = NameKey::new(NameScope::Workbook, "Revenue_2026");

        assert_eq!(key.canonical_name, "REVENUE_2026");
    }

    #[test]
    fn name_scope_distinguishes_workbook_and_sheet() {
        let workbook_key = NameKey::new(NameScope::Workbook, "Total");
        let worksheet_key = NameKey::new(NameScope::Worksheet(SheetId(7)), "Total");

        assert_ne!(workbook_key, worksheet_key);
        assert_eq!(workbook_key.canonical_name, worksheet_key.canonical_name);
    }

    #[test]
    fn builtin_name_classifies_print_area() {
        assert_eq!(
            BuiltinName::from_excel_name("_xlnm.Print_Area"),
            Some(BuiltinName::PrintArea)
        );
    }

    #[test]
    fn builtin_name_preserves_unknown_xlnm_name() {
        assert_eq!(
            BuiltinName::from_excel_name("_xlnm.Custom"),
            Some(BuiltinName::Other("_xlnm.Custom".to_string()))
        );
    }

    #[test]
    fn defined_name_preserves_display_case_and_refers_to() {
        let defined_name = DefinedName::new(
            DefinedNameId(3),
            NameScope::Worksheet(SheetId(9)),
            "Revenue_2026",
            FormulaSource {
                text: "Sheet1!$A$1:$A$4".to_string(),
                is_r1c1: false,
            },
        );

        assert_eq!(defined_name.display_name, "Revenue_2026");
        assert_eq!(defined_name.canonical_name, "REVENUE_2026");
        assert_eq!(defined_name.refers_to.text, "Sheet1!$A$1:$A$4");
        assert_eq!(
            defined_name.key(),
            NameKey::new(NameScope::Worksheet(SheetId(9)), "revenue_2026")
        );
    }

    #[test]
    fn excel_name_canonicalization_is_ascii_case_fold() {
        assert_eq!(canonicalize_excel_name("print_area"), "PRINT_AREA");
    }
}
