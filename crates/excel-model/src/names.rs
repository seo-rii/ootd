use std::collections::BTreeMap;

use office_common::{
    DefinedName, DefinedNameId, DefinedNameMetadata, FormulaSource, NameKey, NameScope,
    NameValidationMode, OmError, OmErrorCode, OmResult, SheetId, canonicalize_excel_name,
};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DefinedNameTable {
    next_id: u32,
    names_by_id: BTreeMap<DefinedNameId, DefinedName>,
    ids_by_key: BTreeMap<NameKey, DefinedNameId>,
    dirty: bool,
}

impl Default for DefinedNameTable {
    fn default() -> Self {
        Self {
            next_id: 1,
            names_by_id: BTreeMap::new(),
            ids_by_key: BTreeMap::new(),
            dirty: false,
        }
    }
}

impl DefinedNameTable {
    pub fn len(&self) -> usize {
        self.names_by_id.len()
    }

    pub fn is_empty(&self) -> bool {
        self.names_by_id.is_empty()
    }

    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    pub fn mark_clean(&mut self) {
        self.dirty = false;
    }

    pub fn iter(&self) -> impl Iterator<Item = &DefinedName> {
        self.names_by_id.values()
    }

    pub fn get(&self, id: DefinedNameId) -> Option<&DefinedName> {
        self.names_by_id.get(&id)
    }

    pub fn set_hidden_by_id(&mut self, id: DefinedNameId, hidden: bool) -> OmResult<bool> {
        let defined_name = self.names_by_id.get_mut(&id).ok_or_else(|| {
            OmError::new(
                OmErrorCode::NotFound,
                format!("defined name id {} was not found", id.0),
            )
        })?;
        if defined_name.metadata.hidden != hidden {
            defined_name.metadata.hidden = hidden;
            self.dirty = true;
            return Ok(true);
        }
        Ok(false)
    }

    pub fn set_refers_to_by_id(
        &mut self,
        id: DefinedNameId,
        refers_to: FormulaSource,
    ) -> OmResult<bool> {
        let defined_name = self.names_by_id.get_mut(&id).ok_or_else(|| {
            OmError::new(
                OmErrorCode::NotFound,
                format!("defined name id {} was not found", id.0),
            )
        })?;
        if defined_name.refers_to != refers_to {
            defined_name.refers_to = refers_to;
            self.dirty = true;
            return Ok(true);
        }
        Ok(false)
    }

    pub fn lookup_in_scope(&self, scope: NameScope, name: &str) -> Option<&DefinedName> {
        let key = NameKey::new(scope, name);
        self.ids_by_key
            .get(&key)
            .and_then(|id| self.names_by_id.get(id))
    }

    pub fn lookup(&self, current_sheet: Option<SheetId>, name: &str) -> Option<&DefinedName> {
        if let Some(sheet_id) = current_sheet {
            if let Some(defined_name) = self.lookup_in_scope(NameScope::Worksheet(sheet_id), name) {
                return Some(defined_name);
            }
        }

        self.lookup_in_scope(NameScope::Workbook, name)
    }

    pub fn add(
        &mut self,
        scope: NameScope,
        display_name: impl Into<String>,
        refers_to: FormulaSource,
        validation_mode: NameValidationMode,
    ) -> OmResult<DefinedNameId> {
        self.add_with_metadata(
            scope,
            display_name,
            refers_to,
            DefinedNameMetadata::default(),
            validation_mode,
        )
    }

    pub fn add_with_metadata(
        &mut self,
        scope: NameScope,
        display_name: impl Into<String>,
        refers_to: FormulaSource,
        mut metadata: DefinedNameMetadata,
        validation_mode: NameValidationMode,
    ) -> OmResult<DefinedNameId> {
        let display_name = display_name.into();
        validate_defined_name(&display_name, validation_mode)?;

        let key = NameKey::new(scope, &display_name);
        if self.ids_by_key.contains_key(&key) {
            return Err(OmError::invalid_argument(format!(
                "defined name '{}' already exists in this scope",
                display_name,
            )));
        }

        let id = DefinedNameId(self.next_id);
        self.next_id = self
            .next_id
            .checked_add(1)
            .ok_or_else(|| OmError::new(OmErrorCode::InvalidState, "defined name id overflow"))?;
        let mut defined_name = DefinedName::new(id, scope, display_name, refers_to);
        if metadata.builtin.is_none() {
            metadata.builtin = defined_name.metadata.builtin.clone();
        }
        defined_name.metadata = metadata;

        self.ids_by_key.insert(key, id);
        self.names_by_id.insert(id, defined_name);
        self.dirty = true;
        Ok(id)
    }

    pub fn insert(
        &mut self,
        defined_name: DefinedName,
        validation_mode: NameValidationMode,
    ) -> OmResult<()> {
        self.insert_with_dirty(defined_name, validation_mode, true)
    }

    pub fn insert_loaded(
        &mut self,
        defined_name: DefinedName,
        validation_mode: NameValidationMode,
    ) -> OmResult<()> {
        self.insert_with_dirty(defined_name, validation_mode, false)
    }

    fn insert_with_dirty(
        &mut self,
        defined_name: DefinedName,
        validation_mode: NameValidationMode,
        dirty: bool,
    ) -> OmResult<()> {
        validate_defined_name(&defined_name.display_name, validation_mode)?;

        let key = defined_name.key();
        if self.ids_by_key.contains_key(&key) {
            return Err(OmError::invalid_argument(format!(
                "defined name '{}' already exists in this scope",
                defined_name.display_name,
            )));
        }
        if self.names_by_id.contains_key(&defined_name.id) {
            return Err(OmError::invalid_argument(format!(
                "defined name id {} already exists",
                defined_name.id.0,
            )));
        }

        self.next_id =
            self.next_id
                .max(defined_name.id.0.checked_add(1).ok_or_else(|| {
                    OmError::new(OmErrorCode::InvalidState, "defined name id overflow")
                })?);
        self.ids_by_key.insert(key, defined_name.id);
        self.names_by_id.insert(defined_name.id, defined_name);
        self.dirty |= dirty;
        Ok(())
    }

    pub fn remove(&mut self, scope: NameScope, name: &str) -> OmResult<DefinedName> {
        let key = NameKey::new(scope, name);
        let id = self.ids_by_key.remove(&key).ok_or_else(|| {
            OmError::new(
                OmErrorCode::NotFound,
                format!("defined name '{}' was not found", name),
            )
        })?;
        let defined_name = self.names_by_id.remove(&id).ok_or_else(|| {
            OmError::new(
                OmErrorCode::InvalidState,
                format!("defined name id {} is missing", id.0),
            )
        })?;
        self.dirty = true;
        Ok(defined_name)
    }

    pub fn remove_by_id(&mut self, id: DefinedNameId) -> OmResult<DefinedName> {
        let defined_name = self.names_by_id.remove(&id).ok_or_else(|| {
            OmError::new(
                OmErrorCode::NotFound,
                format!("defined name id {} was not found", id.0),
            )
        })?;
        self.ids_by_key.remove(&defined_name.key());
        self.dirty = true;
        Ok(defined_name)
    }
}

fn validate_defined_name(name: &str, mode: NameValidationMode) -> OmResult<()> {
    if matches!(mode, NameValidationMode::PreserveLoadedInvalid) {
        return Ok(());
    }

    if name.is_empty() {
        return Err(OmError::invalid_argument("defined name must not be empty"));
    }

    let Some(first_char) = name.chars().next() else {
        return Err(OmError::invalid_argument("defined name must not be empty"));
    };
    if first_char.is_ascii_digit() {
        return Err(OmError::invalid_argument(
            "defined name must not start with a number",
        ));
    }
    if !(first_char.is_ascii_alphabetic() || first_char == '_' || first_char == '\\') {
        return Err(OmError::invalid_argument(
            "defined name must start with a letter, underscore, or backslash",
        ));
    }

    for ch in name.chars() {
        if ch.is_ascii_whitespace() || matches!(ch, ':' | ',' | '!') {
            return Err(OmError::invalid_argument(format!(
                "defined name '{}' contains characters that conflict with reference syntax",
                name,
            )));
        }
        if !(ch.is_ascii_alphanumeric() || matches!(ch, '_' | '.' | '\\')) {
            return Err(OmError::invalid_argument(format!(
                "defined name '{}' contains unsupported characters",
                name,
            )));
        }
    }

    let upper = canonicalize_excel_name(name);
    let bytes = upper.as_bytes();
    let mut letter_count = 0usize;
    while letter_count < bytes.len() && bytes[letter_count].is_ascii_uppercase() {
        letter_count += 1;
    }
    if letter_count > 0 && letter_count < bytes.len() {
        let row_part = &upper[letter_count..];
        if row_part.bytes().all(|byte| byte.is_ascii_digit()) {
            let mut column_index = 0u32;
            for byte in &bytes[..letter_count] {
                column_index = column_index
                    .saturating_mul(26)
                    .saturating_add(u32::from(byte - b'A' + 1));
            }
            if column_index <= 16_384 {
                if let Ok(row_index) = row_part.parse::<u32>() {
                    if (1..=1_048_576).contains(&row_index) {
                        return Err(OmError::invalid_argument(format!(
                            "defined name '{}' must not look like an A1 reference",
                            name,
                        )));
                    }
                }
            }
        }
    }

    if let Some(after_r) = upper.strip_prefix('R') {
        if let Some((row_part, col_part)) = after_r.split_once('C') {
            if !row_part.is_empty()
                && !col_part.is_empty()
                && row_part.bytes().all(|byte| byte.is_ascii_digit())
                && col_part.bytes().all(|byte| byte.is_ascii_digit())
            {
                return Err(OmError::invalid_argument(format!(
                    "defined name '{}' must not look like an R1C1 reference",
                    name,
                )));
            }
        }
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::DefinedNameTable;
    use office_common::{
        DefinedName, DefinedNameId, FormulaSource, NameScope, NameValidationMode, OmErrorCode,
        SheetId,
    };

    fn source(text: &str) -> FormulaSource {
        FormulaSource {
            text: text.to_string(),
            is_r1c1: false,
        }
    }

    #[test]
    fn defined_name_key_is_case_insensitive() {
        let mut table = DefinedNameTable::default();
        table
            .add(
                NameScope::Workbook,
                "Revenue",
                source("Sheet1!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("add name");

        let error = table
            .add(
                NameScope::Workbook,
                "revenue",
                source("Sheet1!$B$1"),
                NameValidationMode::StrictExcel,
            )
            .expect_err("duplicate name should fail");

        assert_eq!(error.code, OmErrorCode::InvalidArgument);
    }

    #[test]
    fn workbook_lookup_falls_back_after_sheet_scope() {
        let mut table = DefinedNameTable::default();
        table
            .add(
                NameScope::Workbook,
                "Total",
                source("Sheet1!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("workbook name");

        let found = table
            .lookup(Some(SheetId(3)), "total")
            .expect("workbook fallback");
        assert_eq!(found.refers_to.text, "Sheet1!$A$1");
    }

    #[test]
    fn set_refers_to_by_id_updates_source_and_dirty_state() {
        let mut table = DefinedNameTable::default();
        let id = table
            .add(
                NameScope::Workbook,
                "Total",
                source("Sheet1!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("add name");
        table.mark_clean();

        assert!(
            table
                .set_refers_to_by_id(
                    id,
                    FormulaSource {
                        text: "Sheet1!$B$2".to_string(),
                        is_r1c1: false,
                    },
                )
                .expect("set RefersTo")
        );
        assert!(table.is_dirty());
        assert_eq!(
            table.get(id).expect("updated name").refers_to.text,
            "Sheet1!$B$2"
        );

        table.mark_clean();
        assert!(
            !table
                .set_refers_to_by_id(
                    id,
                    FormulaSource {
                        text: "Sheet1!$B$2".to_string(),
                        is_r1c1: false,
                    },
                )
                .expect("set same RefersTo")
        );
        assert!(!table.is_dirty());
    }

    #[test]
    fn sheet_scoped_name_shadows_workbook_name() {
        let mut table = DefinedNameTable::default();
        table
            .add(
                NameScope::Workbook,
                "Total",
                source("Sheet1!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("workbook name");
        table
            .add(
                NameScope::Worksheet(SheetId(7)),
                "Total",
                source("Sheet1!$B$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("sheet name");

        let found = table
            .lookup(Some(SheetId(7)), "TOTAL")
            .expect("sheet scoped name");
        assert_eq!(found.scope, NameScope::Worksheet(SheetId(7)));
        assert_eq!(found.refers_to.text, "Sheet1!$B$1");
    }

    #[test]
    fn same_name_is_allowed_on_different_sheets() {
        let mut table = DefinedNameTable::default();
        table
            .add(
                NameScope::Worksheet(SheetId(1)),
                "Total",
                source("Sheet1!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("sheet 1 name");
        table
            .add(
                NameScope::Worksheet(SheetId(2)),
                "total",
                source("Sheet2!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("sheet 2 name");

        assert_eq!(table.len(), 2);
    }

    #[test]
    fn insert_preserves_display_case_and_raw_refers_to() {
        let mut table = DefinedNameTable::default();
        let defined_name = DefinedName::new(
            DefinedNameId(42),
            NameScope::Workbook,
            "Revenue_2026",
            source("Sheet1!$A$1:$A$4"),
        );
        table
            .insert(defined_name, NameValidationMode::StrictExcel)
            .expect("insert loaded name");

        let found = table
            .lookup_in_scope(NameScope::Workbook, "revenue_2026")
            .expect("loaded name");
        assert_eq!(found.id, DefinedNameId(42));
        assert_eq!(found.display_name, "Revenue_2026");
        assert_eq!(found.refers_to.text, "Sheet1!$A$1:$A$4");
    }

    #[test]
    fn insert_loaded_preserves_clean_state() {
        let mut table = DefinedNameTable::default();
        let defined_name = DefinedName::new(
            DefinedNameId(1),
            NameScope::Workbook,
            "Total",
            source("Sheet1!$A$1"),
        );
        table
            .insert_loaded(defined_name, NameValidationMode::StrictExcel)
            .expect("insert loaded name");

        assert!(!table.is_dirty());
    }

    #[test]
    fn add_remove_and_mark_clean_update_dirty_state() {
        let mut table = DefinedNameTable::default();
        table
            .add(
                NameScope::Workbook,
                "Total",
                source("Sheet1!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("add name");
        assert!(table.is_dirty());

        table.mark_clean();
        assert!(!table.is_dirty());

        table
            .remove(NameScope::Workbook, "Total")
            .expect("remove name");
        assert!(table.is_dirty());
    }

    #[test]
    fn add_with_metadata_preserves_hidden_flag() {
        let mut table = DefinedNameTable::default();
        let mut metadata = office_common::DefinedNameMetadata::default();
        metadata.hidden = true;

        let id = table
            .add_with_metadata(
                NameScope::Workbook,
                "HiddenTotal",
                source("Sheet1!$A$1"),
                metadata,
                NameValidationMode::StrictExcel,
            )
            .expect("add hidden name");

        let defined_name = table.get(id).expect("hidden name by id");
        assert!(defined_name.metadata.hidden);
        assert!(table.is_dirty());
    }

    #[test]
    fn set_hidden_by_id_updates_metadata_and_dirty_state() {
        let mut table = DefinedNameTable::default();
        let id = table
            .add(
                NameScope::Workbook,
                "Total",
                source("Sheet1!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("add name");
        table.mark_clean();

        assert!(table.set_hidden_by_id(id, true).expect("set hidden true"));
        assert!(table.get(id).expect("defined name by id").metadata.hidden);
        assert!(table.is_dirty());

        table.mark_clean();
        assert!(
            !table
                .set_hidden_by_id(id, true)
                .expect("set hidden true again")
        );
        assert!(!table.is_dirty());
    }

    #[test]
    fn strict_validation_rejects_reference_like_names() {
        for name in ["A1", "R1C1", "1Name", "Name With Space", "A:B"] {
            let mut table = DefinedNameTable::default();
            let error = table
                .add(
                    NameScope::Workbook,
                    name,
                    source("Sheet1!$A$1"),
                    NameValidationMode::StrictExcel,
                )
                .expect_err("invalid name should fail");

            assert_eq!(error.code, OmErrorCode::InvalidArgument);
        }
    }

    #[test]
    fn preserve_loaded_invalid_mode_keeps_invalid_names() {
        let mut table = DefinedNameTable::default();
        table
            .add(
                NameScope::Workbook,
                "A1",
                source("Sheet1!$A$1"),
                NameValidationMode::PreserveLoadedInvalid,
            )
            .expect("preserve invalid loaded name");

        assert!(table.lookup_in_scope(NameScope::Workbook, "A1").is_some());
    }

    #[test]
    fn remove_by_scope_clears_lookup_key() {
        let mut table = DefinedNameTable::default();
        table
            .add(
                NameScope::Workbook,
                "Total",
                source("Sheet1!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("add name");

        let removed = table
            .remove(NameScope::Workbook, "total")
            .expect("remove name");

        assert_eq!(removed.display_name, "Total");
        assert!(
            table
                .lookup_in_scope(NameScope::Workbook, "Total")
                .is_none()
        );
    }
}
