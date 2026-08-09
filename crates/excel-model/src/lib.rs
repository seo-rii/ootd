use std::collections::{BTreeMap, BTreeSet};

use office_common::{
    CellValue, ChartId, DefinedName, DefinedNameId, DrawingId, ExcelLimits, FormulaSource,
    NameScope, NameValidationMode, OmArray, OmError, OmErrorCode, OmResult, OmValue, OpaquePart,
    RangeRef, RangeSet, Rect, ReferenceTarget, SheetId, SheetKind, SheetScope, SheetVisibility,
    StyleId, WorkbookId, WorkbookModel, WorksheetModel, formula_contains_a1_reference,
};

mod charts;
mod names;

pub use charts::{
    AxisModel, ChartAxisCrosses, ChartAxisDisplayUnit, ChartAxisGroup, ChartAxisKind,
    ChartAxisScaleType, ChartAxisTimeUnit, ChartBarShape, ChartBuiltInDisplayUnit, ChartCacheKind,
    ChartCacheSnapshot, ChartCellMarkerXmlAttrs, ChartDataLabelPosition, ChartDataLabelsModel,
    ChartDataTableModel, ChartDisplayBlanksAs, ChartGroupModel, ChartLayoutMode, ChartLayoutTarget,
    ChartLegendPosition, ChartManualLayout, ChartMarkerStyle, ChartMarkerXmlAttrs, ChartModel,
    ChartObjectModel, ChartPointModel, ChartProtectionModel, ChartSheetBinding,
    ChartSizeRepresents, ChartSourceExpr, ChartSourceReference, ChartSplitType, ChartText,
    ChartTickLabelPosition, ChartTickMark, ChartType, ChartView3DModel, DrawingModel,
    DrawingObjectModel, LegendModel, SeriesModel, resolve_chart_source_reference,
    resolve_chart_source_reference_with_names,
};
pub use names::DefinedNameTable;

#[derive(Debug, Clone, PartialEq)]
pub struct CellData {
    pub value: CellValue,
    pub formula: Option<FormulaSource>,
    pub style_id: Option<StyleId>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CellBatchMutationKind {
    Replace,
    Rearrange,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellShiftDirection {
    Up,
    Left,
    Down,
    Right,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct WorksheetData {
    pub cells: BTreeMap<(u32, u32), CellData>,
    pub source_xml: Vec<u8>,
    pub dirty: bool,
    pub dirty_cells: BTreeSet<(u32, u32)>,
    pub dynamic_array_formulas: BTreeSet<(u32, u32)>,
    pub spill_ranges: BTreeMap<(u32, u32), Rect>,
    pub spill_owners: BTreeMap<(u32, u32), (u32, u32)>,
    pub structural_owners: WorksheetStructuralOwners,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct WorksheetStructuralOwners {
    pub merged_ranges: Vec<Rect>,
    pub data_validation_ranges: Vec<Rect>,
    pub data_validation_formulas: Vec<String>,
    pub row_metadata_ranges: Vec<Rect>,
    pub column_metadata_ranges: Vec<Rect>,
    pub table_relationship_ids: Vec<String>,
    pub table_owners: Vec<TableStructuralOwner>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableStructuralOwner {
    pub relationship_id: String,
    pub part_uri: String,
    pub range: Rect,
    pub formulas: Vec<String>,
}

impl WorksheetData {
    fn spill_owner_for_key(&self, key: (u32, u32)) -> Option<(u32, u32)> {
        self.spill_owners.get(&key).copied().or_else(|| {
            self.spill_ranges.iter().find_map(|(anchor, spill_range)| {
                (key != *anchor
                    && key.0 >= spill_range.row_first
                    && key.0 <= spill_range.row_last
                    && key.1 >= spill_range.col_first
                    && key.1 <= spill_range.col_last)
                    .then_some(*anchor)
            })
        })
    }

    fn ensure_spill_children_are_not_edited(
        &self,
        keys: impl IntoIterator<Item = (u32, u32)>,
    ) -> OmResult<()> {
        for key in keys {
            if let Some(anchor) = self.spill_owner_for_key(key) {
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    format!(
                        "cannot edit spill child R{}C{}; spill anchor is R{}C{}",
                        key.0, key.1, anchor.0, anchor.1,
                    ),
                ));
            }
        }
        Ok(())
    }

    fn ensure_spill_topology_is_not_modified(
        &self,
        keys: impl IntoIterator<Item = (u32, u32)>,
        operation: &str,
    ) -> OmResult<()> {
        for key in keys {
            if let Some(anchor) = self.spill_owner_for_key(key) {
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    format!(
                        "cannot {operation} spill child R{}C{}; spill anchor is R{}C{}",
                        key.0, key.1, anchor.0, anchor.1
                    ),
                ));
            }
            if self.spill_ranges.contains_key(&key) || self.dynamic_array_formulas.contains(&key) {
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    format!("cannot {operation} spill anchor R{}C{}", key.0, key.1),
                ));
            }
        }
        Ok(())
    }

    fn prepare_cell_for_edit(&mut self, key: (u32, u32)) {
        self.prepare_cell_for_edit_with_change(key);
    }

    fn prepare_cell_for_edit_with_change(&mut self, key: (u32, u32)) -> bool {
        let removed_spill_owner = self.spill_owners.remove(&key).is_some();
        let removed_spill_range = self.spill_ranges.contains_key(&key);
        if removed_spill_range {
            self.clear_owned_spill_unchecked(key);
        }
        let removed_dynamic_formula = self.dynamic_array_formulas.remove(&key);
        removed_spill_owner || removed_spill_range || removed_dynamic_formula
    }

    pub fn clear_owned_spill(&mut self, anchor: (u32, u32)) -> OmResult<()> {
        ExcelLimits::validate_cell(anchor.0, anchor.1)?;
        self.clear_owned_spill_unchecked(anchor);
        Ok(())
    }

    fn clear_owned_spill_unchecked(&mut self, anchor: (u32, u32)) {
        let owned_cells = self
            .spill_owners
            .iter()
            .filter_map(|(&key, &owner)| (owner == anchor).then_some(key))
            .collect::<Vec<_>>();
        for key in owned_cells {
            self.spill_owners.remove(&key);
            let remove = if let Some(cell) = self.cells.get_mut(&key) {
                cell.value = CellValue::Blank;
                cell.formula = None;
                cell.style_id.is_none()
            } else {
                false
            };
            if remove {
                self.cells.remove(&key);
            }
            self.dirty = true;
            self.dirty_cells.insert(key);
        }
        self.spill_ranges.remove(&anchor);
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct WorkbookState {
    model: WorkbookModel,
    worksheets: Vec<WorksheetModel>,
    worksheet_data: BTreeMap<SheetId, WorksheetData>,
    pub defined_names: DefinedNameTable,
    pub charts: BTreeMap<ChartId, ChartModel>,
    pub drawings: BTreeMap<DrawingId, DrawingModel>,
    pub chart_sheets: BTreeMap<SheetId, ChartSheetBinding>,
    pub opaque_parts: Vec<OpaquePart>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct WorkbookStateParts {
    pub model: WorkbookModel,
    pub worksheets: Vec<WorksheetModel>,
    pub worksheet_data: BTreeMap<SheetId, WorksheetData>,
    pub defined_names: DefinedNameTable,
    pub charts: BTreeMap<ChartId, ChartModel>,
    pub drawings: BTreeMap<DrawingId, DrawingModel>,
    pub chart_sheets: BTreeMap<SheetId, ChartSheetBinding>,
    pub opaque_parts: Vec<OpaquePart>,
}

impl WorkbookState {
    pub fn try_new(parts: WorkbookStateParts) -> OmResult<Self> {
        let WorkbookStateParts {
            model,
            worksheets,
            worksheet_data,
            defined_names,
            charts,
            drawings,
            chart_sheets,
            opaque_parts,
        } = parts;
        let state = Self {
            model,
            worksheets,
            worksheet_data,
            defined_names,
            charts,
            drawings,
            chart_sheets,
            opaque_parts,
        };
        state.validate_for_save()?;
        Ok(state)
    }

    pub fn into_parts(self) -> WorkbookStateParts {
        let Self {
            model,
            worksheets,
            worksheet_data,
            defined_names,
            charts,
            drawings,
            chart_sheets,
            opaque_parts,
        } = self;
        WorkbookStateParts {
            model,
            worksheets,
            worksheet_data,
            defined_names,
            charts,
            drawings,
            chart_sheets,
            opaque_parts,
        }
    }

    pub fn model(&self) -> &WorkbookModel {
        &self.model
    }

    pub fn worksheets(&self) -> &[WorksheetModel] {
        &self.worksheets
    }

    pub fn set_display_name(&mut self, display_name: impl Into<String>) -> bool {
        let display_name = display_name.into();
        if self.model.display_name == display_name {
            return false;
        }
        self.model.display_name = display_name;
        true
    }

    pub fn set_date1904(&mut self, date1904: bool) -> bool {
        if self.model.date1904 == date1904 {
            return false;
        }
        self.model.date1904 = date1904;
        true
    }

    pub fn set_is_addin(&mut self, is_addin: bool) -> bool {
        if self.model.is_addin == is_addin {
            return false;
        }
        self.model.is_addin = is_addin;
        true
    }

    pub fn set_format(&mut self, format: office_common::FileFormat) -> bool {
        if self.model.format == format {
            return false;
        }
        self.model.format = format;
        true
    }

    pub fn rename_worksheet(
        &mut self,
        sheet_id: SheetId,
        name: impl Into<String>,
    ) -> OmResult<bool> {
        let worksheet_index = self.worksheet_index(sheet_id)?;
        let mut renamed = self.worksheets[worksheet_index].clone();
        renamed.name = name.into();
        self.validate_worksheet_metadata(&renamed)?;
        if self.worksheets.iter().any(|worksheet| {
            worksheet.id != sheet_id && worksheet.name.eq_ignore_ascii_case(&renamed.name)
        }) {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!("duplicate worksheet name: {}", renamed.name),
            ));
        }
        if self.worksheets[worksheet_index].name == renamed.name {
            return Ok(false);
        }
        self.worksheets[worksheet_index] = renamed;
        Ok(true)
    }

    pub fn set_worksheet_visibility(
        &mut self,
        sheet_id: SheetId,
        visibility: SheetVisibility,
    ) -> OmResult<bool> {
        let worksheet_index = self.worksheet_index(sheet_id)?;
        let current_visibility = self.worksheets[worksheet_index].visibility;
        if current_visibility == visibility {
            return Ok(false);
        }
        self.worksheets[worksheet_index].visibility = visibility;
        Ok(true)
    }

    pub fn bind_chart_sheet_package(
        &mut self,
        sheet_id: SheetId,
        relationship_id: impl Into<String>,
        part_uri: impl Into<String>,
    ) -> OmResult<bool> {
        let worksheet_index = self.worksheet_index(sheet_id)?;
        let relationship_id = relationship_id.into();
        let part_uri = part_uri.into();
        let worksheet = &self.worksheets[worksheet_index];
        if worksheet.kind != SheetKind::ChartSheet {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!("worksheet {} is not a chart sheet", sheet_id.0),
            ));
        }
        let chart_sheet = self.chart_sheets.get(&sheet_id).ok_or_else(|| {
            OmError::new(
                OmErrorCode::InvalidState,
                format!("chart sheet {} has no chart binding", sheet_id.0),
            )
        })?;
        if chart_sheet.sheet_id != sheet_id {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!(
                    "chart sheet binding key {} does not match model sheet {}",
                    sheet_id.0, chart_sheet.sheet_id.0
                ),
            ));
        }
        match (
            worksheet.relationship_id.as_deref(),
            worksheet.part_uri.as_deref(),
            chart_sheet.raw_part_uri.as_deref(),
        ) {
            (None, None, None) => {}
            (Some(current_relationship_id), Some(current_part_uri), Some(current_raw_part_uri))
                if current_relationship_id == relationship_id
                    && current_part_uri == part_uri
                    && current_raw_part_uri == part_uri => {}
            (Some(_), Some(_), Some(_)) => {
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    format!(
                        "chart sheet {} is already bound to a different package part",
                        sheet_id.0
                    ),
                ));
            }
            _ => {
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    format!(
                        "chart sheet {} has an incomplete package binding",
                        sheet_id.0
                    ),
                ));
            }
        }

        let mut bound = worksheet.clone();
        bound.relationship_id = Some(relationship_id.clone());
        bound.part_uri = Some(part_uri.clone());
        self.validate_worksheet_metadata(&bound)?;
        if self.worksheets.iter().any(|existing| {
            existing.id != sheet_id
                && existing.relationship_id.as_deref() == Some(relationship_id.as_str())
        }) {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!("duplicate worksheet relationship id {relationship_id}"),
            ));
        }
        if self.worksheets.iter().any(|existing| {
            existing.id != sheet_id
                && existing
                    .part_uri
                    .as_deref()
                    .is_some_and(|current| current.eq_ignore_ascii_case(&part_uri))
        }) {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!("duplicate worksheet part URI {part_uri}"),
            ));
        }

        let changed = self.worksheets[worksheet_index] != bound
            || chart_sheet.raw_part_uri.as_deref() != Some(part_uri.as_str());
        if !changed {
            return Ok(false);
        }
        self.worksheets[worksheet_index] = bound;
        self.chart_sheets
            .get_mut(&sheet_id)
            .expect("chart sheet binding presence was checked before mutation")
            .raw_part_uri = Some(part_uri);
        Ok(true)
    }

    pub fn validate_worksheet_reorder(&self, ordered_sheet_ids: &[SheetId]) -> OmResult<bool> {
        if ordered_sheet_ids.len() != self.worksheets.len() {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!(
                    "worksheet order contains {} ids for {} worksheets",
                    ordered_sheet_ids.len(),
                    self.worksheets.len()
                ),
            ));
        }

        let current_ids = self
            .worksheets
            .iter()
            .map(|worksheet| worksheet.id)
            .collect::<Vec<_>>();
        let current_id_set = current_ids.iter().copied().collect::<BTreeSet<_>>();
        if current_id_set.len() != current_ids.len() {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                "worksheet collection contains duplicate ids",
            ));
        }
        let requested_ids = ordered_sheet_ids.iter().copied().collect::<BTreeSet<_>>();
        if requested_ids.len() != ordered_sheet_ids.len() || requested_ids != current_id_set {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                "worksheet order must be an exact permutation of the current worksheet ids",
            ));
        }

        Ok(current_ids != ordered_sheet_ids)
    }

    pub fn reorder_worksheets(&mut self, ordered_sheet_ids: &[SheetId]) -> OmResult<bool> {
        if !self.validate_worksheet_reorder(ordered_sheet_ids)? {
            return Ok(false);
        }
        let worksheets_by_id = self
            .worksheets
            .iter()
            .map(|worksheet| (worksheet.id, worksheet.clone()))
            .collect::<BTreeMap<_, _>>();
        let reordered = ordered_sheet_ids
            .iter()
            .map(|sheet_id| {
                worksheets_by_id
                    .get(sheet_id)
                    .expect("worksheet order identity set was validated above")
                    .clone()
            })
            .collect::<Vec<_>>();
        self.worksheets = reordered;
        Ok(true)
    }

    pub fn validate_for_save(&self) -> OmResult<()> {
        if self.worksheets.is_empty() {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                "workbook must contain at least one worksheet",
            ));
        }

        let mut worksheet_ids = BTreeSet::new();
        let mut worksheet_names = BTreeSet::new();
        let mut worksheet_relationship_ids = BTreeSet::new();
        let mut worksheet_part_uris = BTreeSet::new();

        for worksheet in &self.worksheets {
            self.validate_worksheet_metadata(worksheet)?;
            if !worksheet_ids.insert(worksheet.id) {
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    format!("duplicate worksheet id {}", worksheet.id.0),
                ));
            }
            if !worksheet_names.insert(worksheet.name.to_ascii_lowercase()) {
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    format!("duplicate worksheet name: {}", worksheet.name),
                ));
            }
            if let (Some(relationship_id), Some(part_uri)) =
                (&worksheet.relationship_id, &worksheet.part_uri)
            {
                if !worksheet_relationship_ids.insert(relationship_id) {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        format!("duplicate worksheet relationship id {relationship_id}"),
                    ));
                }
                if !worksheet_part_uris.insert(part_uri.to_ascii_lowercase()) {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        format!("duplicate worksheet part URI {part_uri}"),
                    ));
                }
            }
            if !self.worksheet_data.contains_key(&worksheet.id) {
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    format!(
                        "worksheet {} ({}) has no worksheet data",
                        worksheet.name, worksheet.id.0
                    ),
                ));
            }
        }

        for defined_name in self.defined_names.iter() {
            if let NameScope::Worksheet(sheet_id) = defined_name.scope
                && !worksheet_ids.contains(&sheet_id)
            {
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    format!(
                        "defined name {} is scoped to unknown worksheet {}",
                        defined_name.display_name, sheet_id.0
                    ),
                ));
            }
        }

        for (&sheet_id, worksheet) in &self.worksheet_data {
            if !worksheet_ids.contains(&sheet_id) {
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    format!("worksheet data references unknown worksheet {}", sheet_id.0),
                ));
            }
            let table_relationship_ids = worksheet
                .structural_owners
                .table_relationship_ids
                .iter()
                .map(String::as_str)
                .collect::<BTreeSet<_>>();
            if table_relationship_ids.len()
                != worksheet.structural_owners.table_relationship_ids.len()
            {
                return Err(OmError::invalid_state(format!(
                    "worksheet {} has duplicate structural table relationship IDs",
                    sheet_id.0,
                )));
            }
            let table_owner_relationship_ids = worksheet
                .structural_owners
                .table_owners
                .iter()
                .map(|owner| owner.relationship_id.as_str())
                .collect::<BTreeSet<_>>();
            if table_owner_relationship_ids.len() != worksheet.structural_owners.table_owners.len()
                || table_owner_relationship_ids != table_relationship_ids
            {
                return Err(OmError::invalid_state(format!(
                    "worksheet {} structural table owner bindings do not match table relationships",
                    sheet_id.0,
                )));
            }
            for table_owner in &worksheet.structural_owners.table_owners {
                if table_owner.relationship_id.is_empty() || table_owner.part_uri.is_empty() {
                    return Err(OmError::invalid_state(format!(
                        "worksheet {} has an incomplete structural table owner binding",
                        sheet_id.0,
                    )));
                }
                ExcelLimits::validate_rect(table_owner.range).map_err(|_| {
                    OmError::invalid_state(format!(
                        "worksheet {} table relationship {} has an out-of-grid range",
                        sheet_id.0, table_owner.relationship_id,
                    ))
                })?;
            }
            let mut row_metadata_rows = BTreeSet::new();
            for range in &worksheet.structural_owners.row_metadata_ranges {
                if ExcelLimits::validate_rect(*range).is_err()
                    || range.row_first != range.row_last
                    || range.col_first != 1
                    || range.col_last != ExcelLimits::MAX_COLUMN_INDEX
                    || !row_metadata_rows.insert(range.row_first)
                {
                    return Err(OmError::invalid_state(format!(
                        "worksheet {} has invalid structural row metadata ownership",
                        sheet_id.0,
                    )));
                }
            }
            for range in &worksheet.structural_owners.column_metadata_ranges {
                if ExcelLimits::validate_rect(*range).is_err()
                    || range.row_first != 1
                    || range.row_last != ExcelLimits::MAX_ROW_INDEX
                {
                    return Err(OmError::invalid_state(format!(
                        "worksheet {} has invalid structural column metadata ownership",
                        sheet_id.0,
                    )));
                }
            }
            for (&(row, col), cell) in &worksheet.cells {
                if cell.value.validate().is_err() {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        format!(
                            "worksheet {} cell R{}C{} numeric value must be finite",
                            sheet_id.0, row, col
                        ),
                    ));
                }
            }
            for &anchor in &worksheet.dynamic_array_formulas {
                if worksheet
                    .cells
                    .get(&anchor)
                    .is_none_or(|cell| cell.formula.is_none())
                {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        format!(
                            "worksheet {} dynamic-array anchor R{}C{} has no formula cell",
                            sheet_id.0, anchor.0, anchor.1
                        ),
                    ));
                }
            }

            let mut previous_spill_ranges = Vec::new();
            for (&anchor, spill_range) in &worksheet.spill_ranges {
                if !worksheet.dynamic_array_formulas.contains(&anchor) {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        format!(
                            "worksheet {} spill range at R{}C{} has no dynamic-array formula owner",
                            sheet_id.0, anchor.0, anchor.1
                        ),
                    ));
                }
                if anchor.0 < spill_range.row_first
                    || anchor.0 > spill_range.row_last
                    || anchor.1 < spill_range.col_first
                    || anchor.1 > spill_range.col_last
                {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        format!(
                            "worksheet {} spill anchor R{}C{} is outside its spill range",
                            sheet_id.0, anchor.0, anchor.1
                        ),
                    ));
                }
                if previous_spill_ranges.iter().any(|existing: &&Rect| {
                    existing.row_first <= spill_range.row_last
                        && spill_range.row_first <= existing.row_last
                        && existing.col_first <= spill_range.col_last
                        && spill_range.col_first <= existing.col_last
                }) {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        format!(
                            "worksheet {} has overlapping spill range at R{}C{}",
                            sheet_id.0, anchor.0, anchor.1
                        ),
                    ));
                }
                previous_spill_ranges.push(spill_range);
            }

            for (&child, &anchor) in &worksheet.spill_owners {
                let Some(spill_range) = worksheet.spill_ranges.get(&anchor) else {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        format!(
                            "worksheet {} spill child R{}C{} references unknown anchor R{}C{}",
                            sheet_id.0, child.0, child.1, anchor.0, anchor.1
                        ),
                    ));
                };
                if child == anchor
                    || child.0 < spill_range.row_first
                    || child.0 > spill_range.row_last
                    || child.1 < spill_range.col_first
                    || child.1 > spill_range.col_last
                {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        format!(
                            "worksheet {} spill child R{}C{} is outside owner R{}C{} spill range",
                            sheet_id.0, child.0, child.1, anchor.0, anchor.1
                        ),
                    ));
                }
                let Some(child_cell) = worksheet.cells.get(&child) else {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        format!(
                            "worksheet {} spill child R{}C{} has no materialized cell",
                            sheet_id.0, child.0, child.1
                        ),
                    ));
                };
                if child_cell.formula.is_some() {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        format!(
                            "worksheet {} spill child R{}C{} cannot contain a formula",
                            sheet_id.0, child.0, child.1
                        ),
                    ));
                }
            }
        }

        Ok(())
    }

    fn validate_worksheet_metadata(&self, worksheet: &WorksheetModel) -> OmResult<()> {
        if worksheet.id.0 == 0 || worksheet.id.0 > u64::from(u32::MAX) {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!(
                    "worksheet {} has id {} outside the supported unsigned 32-bit range",
                    worksheet.name, worksheet.id.0
                ),
            ));
        }
        if worksheet.name.trim().is_empty() {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!("worksheet {} has an empty name", worksheet.id.0),
            ));
        }
        if worksheet.name.chars().count() > 31 {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!(
                    "worksheet {} name exceeds Excel's 31-character limit",
                    worksheet.name
                ),
            ));
        }
        if worksheet
            .name
            .chars()
            .any(|ch| matches!(ch, ':' | '\\' | '/' | '?' | '*' | '[' | ']') || ch.is_control())
        {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!(
                    "worksheet {} name contains invalid characters",
                    worksheet.name
                ),
            ));
        }
        if worksheet.workbook_id != self.model.id {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!(
                    "worksheet {} ({}) has workbook id {}, expected {}",
                    worksheet.name, worksheet.id.0, worksheet.workbook_id.0, self.model.id.0
                ),
            ));
        }
        match (&worksheet.relationship_id, &worksheet.part_uri) {
            (None, None) if worksheet.kind == office_common::SheetKind::ChartSheet => Ok(()),
            (None, None) => Err(OmError::new(
                OmErrorCode::InvalidState,
                format!(
                    "worksheet {} ({}) has no package binding; only an unbound chart-sheet record awaiting graph preflight may omit both relationship id and part URI",
                    worksheet.name, worksheet.id.0
                ),
            )),
            (Some(relationship_id), Some(part_uri))
                if !relationship_id.is_empty() && !part_uri.is_empty() =>
            {
                Ok(())
            }
            _ => Err(OmError::new(
                OmErrorCode::InvalidState,
                format!(
                    "worksheet {} ({}) must have both a non-empty relationship id and part URI, or neither",
                    worksheet.name, worksheet.id.0
                ),
            )),
        }
    }

    pub fn assign_workbook_id(&mut self, workbook_id: WorkbookId) -> OmResult<()> {
        let mut rebound_charts = self.charts.clone();
        for (&chart_id, chart) in &mut rebound_charts {
            chart.workbook_id = workbook_id;
            for (series_index, series) in chart.series.iter_mut().enumerate() {
                for (source_name, source) in [
                    ("name", series.name.as_mut()),
                    ("x-values", series.x_values.as_mut()),
                    ("values", series.values.as_mut()),
                    ("bubble-size", series.bubble_size.as_mut()),
                ] {
                    let Some(source) = source else {
                        continue;
                    };
                    if let Some(ReferenceTarget::Range(range)) = source.resolved.as_mut() {
                        let updated_range = RangeSet::new(workbook_id, range.areas().to_vec())
                            .map_err(|error| {
                                OmError::new(
                                    OmErrorCode::InvalidState,
                                    format!(
                                        "cannot reassign workbook id for chart {} series {} {} range: {}",
                                        chart_id.0,
                                        series_index + 1,
                                        source_name,
                                        error.message
                                    ),
                                )
                            })?;
                        *range = updated_range;
                    }
                    if let Some(ReferenceTarget::Range(range)) = source
                        .full_reference
                        .as_mut()
                        .and_then(|reference| reference.resolved.as_mut())
                    {
                        let updated_range = RangeSet::new(workbook_id, range.areas().to_vec())
                            .map_err(|error| {
                                OmError::new(
                                    OmErrorCode::InvalidState,
                                    format!(
                                        "cannot reassign workbook id for chart {} series {} {} full range: {}",
                                        chart_id.0,
                                        series_index + 1,
                                        source_name,
                                        error.message
                                    ),
                                )
                            })?;
                        *range = updated_range;
                    }
                }
            }
        }

        self.model.id = workbook_id;
        for worksheet in &mut self.worksheets {
            worksheet.workbook_id = workbook_id;
        }
        self.charts = rebound_charts;
        for drawing in self.drawings.values_mut() {
            drawing.workbook_id = workbook_id;
            for object in &mut drawing.objects {
                if let DrawingObjectModel::ChartFrame(chart_object) = object {
                    chart_object.workbook_id = workbook_id;
                }
            }
        }
        Ok(())
    }

    pub fn set_worksheet_source_xml(
        &mut self,
        sheet_id: SheetId,
        source_xml: Vec<u8>,
    ) -> OmResult<()> {
        self.worksheet_data_for_sheet_mut(sheet_id)?.source_xml = source_xml;
        Ok(())
    }

    pub fn insert_cell(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
        cell: CellData,
    ) -> OmResult<()> {
        ExcelLimits::validate_cell(row, col)?;
        cell.value.validate()?;
        let worksheet = self.worksheet_data_for_sheet_mut(sheet_id)?;
        worksheet.cells.insert((row, col), cell);
        worksheet.dirty = true;
        worksheet.dirty_cells.insert((row, col));
        Ok(())
    }

    pub fn cell(&self, sheet_id: SheetId, row: u32, col: u32) -> Option<&CellData> {
        self.worksheet_data
            .get(&sheet_id)
            .and_then(|worksheet| worksheet.cells.get(&(row, col)))
    }

    pub fn defined_names(&self) -> &DefinedNameTable {
        &self.defined_names
    }

    pub fn defined_names_mut(&mut self) -> &mut DefinedNameTable {
        &mut self.defined_names
    }

    pub fn charts(&self) -> &BTreeMap<ChartId, ChartModel> {
        &self.charts
    }

    pub fn charts_mut(&mut self) -> &mut BTreeMap<ChartId, ChartModel> {
        &mut self.charts
    }

    pub fn drawings(&self) -> &BTreeMap<DrawingId, DrawingModel> {
        &self.drawings
    }

    pub fn drawings_mut(&mut self) -> &mut BTreeMap<DrawingId, DrawingModel> {
        &mut self.drawings
    }

    pub fn chart_sheets(&self) -> &BTreeMap<SheetId, ChartSheetBinding> {
        &self.chart_sheets
    }

    pub fn chart_sheets_mut(&mut self) -> &mut BTreeMap<SheetId, ChartSheetBinding> {
        &mut self.chart_sheets
    }

    pub fn chart_sheet(&self, sheet_id: SheetId) -> Option<&ChartSheetBinding> {
        self.chart_sheets.get(&sheet_id)
    }

    pub fn add_defined_name(
        &mut self,
        scope: NameScope,
        display_name: impl Into<String>,
        refers_to: FormulaSource,
        validation_mode: NameValidationMode,
    ) -> OmResult<DefinedNameId> {
        self.defined_names
            .add(scope, display_name, refers_to, validation_mode)
    }

    pub fn remove_defined_name(&mut self, scope: NameScope, name: &str) -> OmResult<DefinedName> {
        self.defined_names.remove(scope, name)
    }

    pub fn lookup_name(&self, current_sheet: Option<SheetId>, name: &str) -> Option<&DefinedName> {
        self.defined_names.lookup(current_sheet, name)
    }

    pub fn lookup_name_in_scope(&self, scope: NameScope, name: &str) -> Option<&DefinedName> {
        self.defined_names.lookup_in_scope(scope, name)
    }

    pub fn worksheet_data_for_sheet(&self, sheet_id: SheetId) -> OmResult<&WorksheetData> {
        self.ensure_worksheet_exists(sheet_id)?;
        self.worksheet_data.get(&sheet_id).ok_or_else(|| {
            OmError::new(
                OmErrorCode::NotFound,
                format!("unknown worksheet data for sheet {}", sheet_id.0),
            )
        })
    }

    pub fn worksheet_data_for_sheet_mut(
        &mut self,
        sheet_id: SheetId,
    ) -> OmResult<&mut WorksheetData> {
        self.ensure_worksheet_exists(sheet_id)?;
        self.worksheet_data.get_mut(&sheet_id).ok_or_else(|| {
            OmError::new(
                OmErrorCode::NotFound,
                format!("unknown worksheet data for sheet {}", sheet_id.0),
            )
        })
    }

    pub fn worksheet_data(&self) -> &BTreeMap<SheetId, WorksheetData> {
        &self.worksheet_data
    }

    pub fn replace_worksheet_data_for_sheet(
        &mut self,
        sheet_id: SheetId,
        worksheet_data: WorksheetData,
    ) -> OmResult<WorksheetData> {
        let current = self.worksheet_data_for_sheet_mut(sheet_id)?;
        Ok(std::mem::replace(current, worksheet_data))
    }

    pub fn insert_worksheet_with_data(
        &mut self,
        index: usize,
        worksheet: WorksheetModel,
        worksheet_data: WorksheetData,
    ) -> OmResult<()> {
        if index > self.worksheets.len() {
            return Err(OmError::invalid_argument(format!(
                "worksheet insertion index {index} exceeds worksheet count {}",
                self.worksheets.len()
            )));
        }
        self.validate_worksheet_metadata(&worksheet)?;
        if self
            .worksheets
            .iter()
            .any(|existing| existing.id == worksheet.id)
            || self.worksheet_data.contains_key(&worksheet.id)
        {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!("duplicate worksheet id {}", worksheet.id.0),
            ));
        }
        if self
            .worksheets
            .iter()
            .any(|existing| existing.name.eq_ignore_ascii_case(&worksheet.name))
        {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!("duplicate worksheet name: {}", worksheet.name),
            ));
        }
        if let Some(relationship_id) = worksheet.relationship_id.as_deref()
            && self
                .worksheets
                .iter()
                .any(|existing| existing.relationship_id.as_deref() == Some(relationship_id))
        {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!("duplicate worksheet relationship id {relationship_id}"),
            ));
        }
        if let Some(part_uri) = worksheet.part_uri.as_deref()
            && self.worksheets.iter().any(|existing| {
                existing
                    .part_uri
                    .as_deref()
                    .is_some_and(|existing| existing.eq_ignore_ascii_case(part_uri))
            })
        {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!("duplicate worksheet part URI {part_uri}"),
            ));
        }

        let sheet_id = worksheet.id;
        self.worksheets.insert(index, worksheet);
        self.worksheet_data.insert(sheet_id, worksheet_data);
        Ok(())
    }

    pub fn remove_worksheet_with_data(
        &mut self,
        sheet_id: SheetId,
    ) -> OmResult<(usize, WorksheetModel, WorksheetData)> {
        if self.worksheets.len() == 1 {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                "workbook must contain at least one worksheet",
            ));
        }
        let worksheet_index = self
            .worksheets
            .iter()
            .position(|worksheet| worksheet.id == sheet_id)
            .ok_or_else(|| {
                OmError::new(
                    OmErrorCode::NotFound,
                    format!("unknown worksheet {}", sheet_id.0),
                )
            })?;
        if !self.worksheet_data.contains_key(&sheet_id) {
            return Err(OmError::new(
                OmErrorCode::NotFound,
                format!("unknown worksheet data for sheet {}", sheet_id.0),
            ));
        }

        let worksheet = self.worksheets.remove(worksheet_index);
        let worksheet_data = self
            .worksheet_data
            .remove(&sheet_id)
            .expect("worksheet data presence was checked before mutation");
        Ok((worksheet_index, worksheet, worksheet_data))
    }

    pub fn mark_worksheet_data_clean(&mut self) {
        for worksheet in self.worksheet_data.values_mut() {
            worksheet.dirty = false;
            worksheet.dirty_cells.clear();
        }
    }

    fn ensure_worksheet_exists(&self, sheet_id: SheetId) -> OmResult<()> {
        self.worksheet_index(sheet_id).map(|_| ())
    }

    fn worksheet_index(&self, sheet_id: SheetId) -> OmResult<usize> {
        self.worksheets
            .iter()
            .position(|worksheet| worksheet.id == sheet_id)
            .ok_or_else(|| {
                OmError::new(
                    OmErrorCode::NotFound,
                    format!("unknown worksheet {}", sheet_id.0),
                )
            })
    }

    pub fn get_range_values(&self, range: &RangeRef) -> OmResult<OmArray> {
        let (sheet_id, rect) = self.single_sheet_rect(range)?;
        let worksheet = self.worksheet_data_for_sheet(sheet_id)?;
        let mut values = Vec::with_capacity(rect.checked_cell_count_usize()?);

        for row in rect.row_first..=rect.row_last {
            for col in rect.col_first..=rect.col_last {
                let value = worksheet
                    .cells
                    .get(&(row, col))
                    .map(|cell| OmValue::from(cell.value.clone()))
                    .unwrap_or(OmValue::Empty);
                values.push(value);
            }
        }

        OmArray::new(rect.height() as usize, rect.width() as usize, values)
    }

    pub fn get_range_formulas(&self, range: &RangeRef) -> OmResult<OmArray> {
        let (sheet_id, rect) = self.single_sheet_rect(range)?;
        let worksheet = self.worksheet_data_for_sheet(sheet_id)?;
        let mut values = Vec::with_capacity(rect.checked_cell_count_usize()?);

        for row in rect.row_first..=rect.row_last {
            for col in rect.col_first..=rect.col_last {
                let value = match worksheet.cells.get(&(row, col)) {
                    Some(cell) => match &cell.formula {
                        Some(formula) => OmValue::Text(format!("={}", formula.text)),
                        None => OmValue::from(cell.value.clone()),
                    },
                    None => OmValue::Empty,
                };
                values.push(value);
            }
        }

        OmArray::new(rect.height() as usize, rect.width() as usize, values)
    }

    pub fn set_range_values(&mut self, range: &RangeRef, values: &OmArray) -> OmResult<()> {
        self.set_range_values_with_change(range, values).map(|_| ())
    }

    pub fn set_range_values_with_change(
        &mut self,
        range: &RangeRef,
        values: &OmArray,
    ) -> OmResult<bool> {
        let (sheet_id, rects) = self.same_sheet_rects(range)?;
        let mut updates = Vec::new();

        if rects.len() == 1 {
            let rect = rects[0];
            if values.rows != rect.height() as usize || values.cols != rect.width() as usize {
                return Err(OmError::invalid_argument(format!(
                    "range dimensions {}x{} do not match value matrix {}x{}",
                    rect.height(),
                    rect.width(),
                    values.rows,
                    values.cols,
                )));
            }

            updates.reserve(values.values.len());
            for row_offset in 0..values.rows {
                for col_offset in 0..values.cols {
                    let row = rect.row_first + row_offset as u32;
                    let col = rect.col_first + col_offset as u32;
                    let value = CellValue::try_from(
                        values.values[row_offset * values.cols + col_offset].clone(),
                    )?;
                    updates.push(((row, col), value));
                }
            }
        } else {
            if values.rows != 1 || values.cols != 1 {
                return Err(OmError::unsupported(
                    "multi-area range value assignment currently supports scalar values only",
                ));
            }

            let value = CellValue::try_from(values.values[0].clone())?;
            updates.reserve(checked_rect_cell_count_sum(&rects)?);
            for rect in &rects {
                for row in rect.row_first..=rect.row_last {
                    for col in rect.col_first..=rect.col_last {
                        updates.push(((row, col), value.clone()));
                    }
                }
            }
        }

        let worksheet = self.worksheet_data_for_sheet_mut(sheet_id)?;
        worksheet.ensure_spill_children_are_not_edited(updates.iter().map(|(key, _)| *key))?;
        let mut changed = false;
        for (key, value) in updates {
            let unchanged = worksheet.cells.get(&key).is_some_and(|existing| {
                existing.value == value
                    && existing.formula.is_none()
                    && !worksheet.spill_owners.contains_key(&key)
                    && !worksheet.spill_ranges.contains_key(&key)
            });
            if unchanged {
                continue;
            }
            let metadata_changed = worksheet.prepare_cell_for_edit_with_change(key);
            if let Some(existing) = worksheet.cells.get_mut(&key) {
                existing.value = value;
                existing.formula = None;
                if matches!(existing.value, CellValue::Blank) && existing.style_id.is_none() {
                    worksheet.cells.remove(&key);
                }
                worksheet.dirty = true;
                worksheet.dirty_cells.insert(key);
                changed = true;
                continue;
            }

            if !matches!(value, CellValue::Blank) {
                worksheet.cells.insert(
                    key,
                    CellData {
                        value,
                        formula: None,
                        style_id: None,
                    },
                );
                worksheet.dirty = true;
                worksheet.dirty_cells.insert(key);
                changed = true;
            } else if metadata_changed {
                worksheet.dirty = true;
                worksheet.dirty_cells.insert(key);
                changed = true;
            }
        }

        Ok(changed)
    }

    pub fn set_range_formulas(&mut self, range: &RangeRef, values: &OmArray) -> OmResult<()> {
        self.set_range_formulas_with_change(range, values)
            .map(|_| ())
    }

    pub fn set_range_formulas_with_change(
        &mut self,
        range: &RangeRef,
        values: &OmArray,
    ) -> OmResult<bool> {
        self.set_range_formulas_impl(range, values, false, false)
    }

    pub fn set_range_dynamic_array_formulas(
        &mut self,
        range: &RangeRef,
        values: &OmArray,
    ) -> OmResult<()> {
        self.set_range_dynamic_array_formulas_with_change(range, values)
            .map(|_| ())
    }

    pub fn set_range_dynamic_array_formulas_with_change(
        &mut self,
        range: &RangeRef,
        values: &OmArray,
    ) -> OmResult<bool> {
        self.set_range_formulas_impl(range, values, true, false)
    }

    pub fn set_range_r1c1_formulas(&mut self, range: &RangeRef, values: &OmArray) -> OmResult<()> {
        self.set_range_r1c1_formulas_with_change(range, values)
            .map(|_| ())
    }

    pub fn set_range_r1c1_formulas_with_change(
        &mut self,
        range: &RangeRef,
        values: &OmArray,
    ) -> OmResult<bool> {
        self.set_range_formulas_impl(range, values, false, true)
    }

    pub fn set_range_dynamic_array_r1c1_formulas(
        &mut self,
        range: &RangeRef,
        values: &OmArray,
    ) -> OmResult<()> {
        self.set_range_dynamic_array_r1c1_formulas_with_change(range, values)
            .map(|_| ())
    }

    pub fn set_range_dynamic_array_r1c1_formulas_with_change(
        &mut self,
        range: &RangeRef,
        values: &OmArray,
    ) -> OmResult<bool> {
        self.set_range_formulas_impl(range, values, true, true)
    }

    fn set_range_formulas_impl(
        &mut self,
        range: &RangeRef,
        values: &OmArray,
        dynamic_array: bool,
        is_r1c1: bool,
    ) -> OmResult<bool> {
        let (sheet_id, rects) = self.same_sheet_rects(range)?;
        let mut updates = Vec::new();

        if rects.len() == 1 {
            let rect = rects[0];
            if values.rows != rect.height() as usize || values.cols != rect.width() as usize {
                return Err(OmError::invalid_argument(format!(
                    "range dimensions {}x{} do not match formula matrix {}x{}",
                    rect.height(),
                    rect.width(),
                    values.rows,
                    values.cols,
                )));
            }

            updates.reserve(values.values.len());
            for row_offset in 0..values.rows {
                for col_offset in 0..values.cols {
                    let row = rect.row_first + row_offset as u32;
                    let col = rect.col_first + col_offset as u32;
                    let value = values.values[row_offset * values.cols + col_offset].clone();
                    let (cell_value, formula) = match value {
                        OmValue::Text(text) => {
                            if let Some(formula_text) = text.strip_prefix('=') {
                                (
                                    CellValue::Blank,
                                    Some(FormulaSource {
                                        text: formula_text.to_string(),
                                        is_r1c1,
                                    }),
                                )
                            } else {
                                (CellValue::Text(text), None)
                            }
                        }
                        other => (CellValue::try_from(other)?, None),
                    };
                    updates.push(((row, col), cell_value, formula));
                }
            }
        } else {
            if values.rows != 1 || values.cols != 1 {
                return Err(OmError::unsupported(
                    "multi-area range formula assignment currently supports scalar values only",
                ));
            }

            let (cell_value, formula) = match values.values[0].clone() {
                OmValue::Text(text) => {
                    if let Some(formula_text) = text.strip_prefix('=') {
                        (
                            CellValue::Blank,
                            Some(FormulaSource {
                                text: formula_text.to_string(),
                                is_r1c1,
                            }),
                        )
                    } else {
                        (CellValue::Text(text), None)
                    }
                }
                other => (CellValue::try_from(other)?, None),
            };
            updates.reserve(checked_rect_cell_count_sum(&rects)?);
            for rect in &rects {
                for row in rect.row_first..=rect.row_last {
                    for col in rect.col_first..=rect.col_last {
                        updates.push(((row, col), cell_value.clone(), formula.clone()));
                    }
                }
            }
        }

        let worksheet = self.worksheet_data_for_sheet_mut(sheet_id)?;
        worksheet.ensure_spill_children_are_not_edited(updates.iter().map(|(key, _, _)| *key))?;
        let mut changed = false;
        for (key, cell_value, formula) in updates {
            let is_dynamic_formula = dynamic_array && formula.is_some();
            let unchanged = worksheet.cells.get(&key).is_some_and(|existing| {
                existing.value == cell_value
                    && existing.formula == formula
                    && worksheet.dynamic_array_formulas.contains(&key) == is_dynamic_formula
            });
            if unchanged {
                continue;
            }
            let metadata_changed = worksheet.prepare_cell_for_edit_with_change(key);
            if let Some(existing) = worksheet.cells.get_mut(&key) {
                existing.value = cell_value;
                existing.formula = formula;
                if matches!(existing.value, CellValue::Blank)
                    && existing.style_id.is_none()
                    && existing.formula.is_none()
                {
                    worksheet.cells.remove(&key);
                }
                worksheet.dirty = true;
                worksheet.dirty_cells.insert(key);
                changed = true;
            } else if !matches!(cell_value, CellValue::Blank) || formula.is_some() {
                worksheet.cells.insert(
                    key,
                    CellData {
                        value: cell_value,
                        formula,
                        style_id: None,
                    },
                );
                worksheet.dirty = true;
                worksheet.dirty_cells.insert(key);
                changed = true;
            } else if metadata_changed {
                worksheet.dirty = true;
                worksheet.dirty_cells.insert(key);
                changed = true;
            }
            if is_dynamic_formula {
                worksheet.dynamic_array_formulas.insert(key);
            }
        }

        Ok(changed)
    }

    pub fn replace_cells_with_change(
        &mut self,
        sheet_id: SheetId,
        replacements: BTreeMap<(u32, u32), Option<CellData>>,
    ) -> OmResult<bool> {
        self.apply_cell_batch_with_change(sheet_id, replacements, CellBatchMutationKind::Replace)
    }

    pub fn rearrange_cells_with_change(
        &mut self,
        sheet_id: SheetId,
        replacements: BTreeMap<(u32, u32), Option<CellData>>,
    ) -> OmResult<bool> {
        self.apply_cell_batch_with_change(sheet_id, replacements, CellBatchMutationKind::Rearrange)
    }

    pub fn shift_cells_with_change(
        &mut self,
        sheet_id: SheetId,
        rect: Rect,
        direction: CellShiftDirection,
    ) -> OmResult<bool> {
        ExcelLimits::validate_rect(rect)?;
        let (operation, member) = match direction {
            CellShiftDirection::Up | CellShiftDirection::Left => ("delete", "Delete"),
            CellShiftDirection::Down | CellShiftDirection::Right => ("insert", "Insert"),
        };
        for (&owner_sheet_id, owner) in &self.worksheet_data {
            for (&(row, col), cell) in &owner.cells {
                if let Some(formula) = &cell.formula
                    && (formula.is_r1c1 || formula_contains_a1_reference(&formula.text))
                {
                    return Err(OmError::unsupported(format!(
                        "Range.{member} structural formula retarget is not implemented for worksheet {} cell R{}C{}",
                        owner_sheet_id.0, row, col
                    )));
                }
            }
            for (formula_index, formula) in owner
                .structural_owners
                .data_validation_formulas
                .iter()
                .enumerate()
            {
                if formula_contains_a1_reference(formula) {
                    return Err(OmError::unsupported(format!(
                        "Range.{member} structural data-validation formula retarget is not implemented for worksheet {} formula {}",
                        owner_sheet_id.0,
                        formula_index + 1,
                    )));
                }
            }
            let table_relationship_ids = owner
                .structural_owners
                .table_relationship_ids
                .iter()
                .map(String::as_str)
                .collect::<BTreeSet<_>>();
            let table_owner_relationship_ids = owner
                .structural_owners
                .table_owners
                .iter()
                .map(|table_owner| table_owner.relationship_id.as_str())
                .collect::<BTreeSet<_>>();
            if table_relationship_ids.len() != owner.structural_owners.table_relationship_ids.len()
                || table_owner_relationship_ids.len() != owner.structural_owners.table_owners.len()
                || table_relationship_ids != table_owner_relationship_ids
            {
                let relationship_id = owner
                    .structural_owners
                    .table_relationship_ids
                    .first()
                    .map(String::as_str)
                    .or_else(|| {
                        owner
                            .structural_owners
                            .table_owners
                            .first()
                            .map(|table_owner| table_owner.relationship_id.as_str())
                    })
                    .unwrap_or("<unknown>");
                return Err(OmError::unsupported(format!(
                    "Range.{member} structural table owner retarget is not implemented for worksheet {} relationship {relationship_id}",
                    owner_sheet_id.0,
                )));
            }
            for table_owner in &owner.structural_owners.table_owners {
                for (formula_index, formula) in table_owner.formulas.iter().enumerate() {
                    if formula_contains_a1_reference(formula) {
                        return Err(OmError::unsupported(format!(
                            "Range.{member} structural table formula retarget is not implemented for worksheet {} relationship {} part {} formula {}",
                            owner_sheet_id.0,
                            table_owner.relationship_id,
                            table_owner.part_uri,
                            formula_index + 1,
                        )));
                    }
                }
            }
        }
        for defined_name in self.defined_names.iter() {
            if defined_name.refers_to.is_r1c1
                || formula_contains_a1_reference(&defined_name.refers_to.text)
            {
                let scope = match defined_name.scope {
                    NameScope::Workbook => "workbook".to_string(),
                    NameScope::Worksheet(sheet_id) => format!("worksheet {}", sheet_id.0),
                };
                return Err(OmError::unsupported(format!(
                    "Range.{member} structural defined-name retarget is not implemented for {scope} name '{}'",
                    defined_name.display_name
                )));
            }
        }
        let affected_rect = match direction {
            CellShiftDirection::Up | CellShiftDirection::Down => Rect {
                row_first: rect.row_first,
                row_last: ExcelLimits::MAX_ROW_INDEX,
                col_first: rect.col_first,
                col_last: rect.col_last,
            },
            CellShiftDirection::Left | CellShiftDirection::Right => Rect {
                row_first: rect.row_first,
                row_last: rect.row_last,
                col_first: rect.col_first,
                col_last: ExcelLimits::MAX_COLUMN_INDEX,
            },
        };
        for (&chart_id, chart) in &self.charts {
            for (series_index, series) in chart.series.iter().enumerate() {
                for (source_name, source) in [
                    ("name", series.name.as_ref()),
                    ("x-values", series.x_values.as_ref()),
                    ("values", series.values.as_ref()),
                    ("bubble-size", series.bubble_size.as_ref()),
                ] {
                    let Some(source) = source else {
                        continue;
                    };
                    for (is_full_reference, reference) in [
                        (false, Some((&source.raw, source.resolved.as_ref()))),
                        (
                            true,
                            source
                                .full_reference
                                .as_ref()
                                .map(|reference| (&reference.raw, reference.resolved.as_ref())),
                        ),
                    ] {
                        let Some((raw, resolved)) = reference else {
                            continue;
                        };
                        let reference_suffix = if is_full_reference {
                            " full-reference"
                        } else {
                            ""
                        };
                        let Some(resolved) = resolved else {
                            if raw.is_r1c1 || formula_contains_a1_reference(&raw.text) {
                                return Err(OmError::unsupported(format!(
                                    "Range.{member} structural chart source retarget is not implemented for chart {} series {} {source_name}{reference_suffix} unresolved reference",
                                    chart_id.0,
                                    series_index + 1,
                                )));
                            }
                            continue;
                        };
                        let ReferenceTarget::Range(range) = resolved else {
                            continue;
                        };
                        if range.workbook_id() != self.model.id {
                            return Err(OmError::invalid_state(format!(
                                "chart {} series {} {source_name}{reference_suffix} range belongs to workbook {}, expected {}",
                                chart_id.0,
                                series_index + 1,
                                range.workbook_id().0,
                                self.model.id.0,
                            )));
                        }
                        if range.areas().is_empty() {
                            return Err(OmError::invalid_state(format!(
                                "chart {} series {} {source_name}{reference_suffix} range has no areas",
                                chart_id.0,
                                series_index + 1,
                            )));
                        }
                        for area in range.areas() {
                            ExcelLimits::validate_rect(area.rect).map_err(|error| {
                                OmError::invalid_state(format!(
                                    "chart {} series {} {source_name}{reference_suffix} range is invalid: {}",
                                    chart_id.0,
                                    series_index + 1,
                                    error.message,
                                ))
                            })?;
                            let owner_sheet_id = match area.scope {
                                SheetScope::Single(owner_sheet_id) => owner_sheet_id,
                                SheetScope::Multi3D { start, end } => {
                                    return Err(OmError::unsupported(format!(
                                        "Range.{member} structural chart source retarget is not implemented for chart {} series {} {source_name}{reference_suffix} 3D range worksheets {}:{}",
                                        chart_id.0,
                                        series_index + 1,
                                        start.0,
                                        end.0,
                                    )));
                                }
                            };
                            if owner_sheet_id == sheet_id
                                && affected_rect.row_first <= area.rect.row_last
                                && area.rect.row_first <= affected_rect.row_last
                                && affected_rect.col_first <= area.rect.col_last
                                && area.rect.col_first <= affected_rect.col_last
                            {
                                return Err(OmError::unsupported(format!(
                                    "Range.{member} structural chart source retarget is not implemented for chart {} series {} {source_name}{reference_suffix} worksheet {} range R{}C{}:R{}C{}",
                                    chart_id.0,
                                    series_index + 1,
                                    owner_sheet_id.0,
                                    area.rect.row_first,
                                    area.rect.col_first,
                                    area.rect.row_last,
                                    area.rect.col_last,
                                )));
                            }
                        }
                    }
                }
            }
        }
        let worksheet = self.worksheet_data_for_sheet(sheet_id)?;
        if let Some(merged_range) = worksheet
            .structural_owners
            .merged_ranges
            .iter()
            .find(|range| {
                affected_rect.row_first <= range.row_last
                    && range.row_first <= affected_rect.row_last
                    && affected_rect.col_first <= range.col_last
                    && range.col_first <= affected_rect.col_last
            })
        {
            return Err(OmError::unsupported(format!(
                "Range.{member} structural merged-cell retarget is not implemented for worksheet {} range R{}C{}:R{}C{}",
                sheet_id.0,
                merged_range.row_first,
                merged_range.col_first,
                merged_range.row_last,
                merged_range.col_last,
            )));
        }
        if let Some(validation_range) = worksheet
            .structural_owners
            .data_validation_ranges
            .iter()
            .find(|range| {
                affected_rect.row_first <= range.row_last
                    && range.row_first <= affected_rect.row_last
                    && affected_rect.col_first <= range.col_last
                    && range.col_first <= affected_rect.col_last
            })
        {
            return Err(OmError::unsupported(format!(
                "Range.{member} structural data-validation retarget is not implemented for worksheet {} range R{}C{}:R{}C{}",
                sheet_id.0,
                validation_range.row_first,
                validation_range.col_first,
                validation_range.row_last,
                validation_range.col_last,
            )));
        }
        if let Some(table_owner) =
            worksheet
                .structural_owners
                .table_owners
                .iter()
                .find(|table_owner| {
                    affected_rect.row_first <= table_owner.range.row_last
                        && table_owner.range.row_first <= affected_rect.row_last
                        && affected_rect.col_first <= table_owner.range.col_last
                        && table_owner.range.col_first <= affected_rect.col_last
                })
        {
            return Err(OmError::unsupported(format!(
                "Range.{member} structural table range retarget is not implemented for worksheet {} relationship {} part {} range R{}C{}:R{}C{}",
                sheet_id.0,
                table_owner.relationship_id,
                table_owner.part_uri,
                table_owner.range.row_first,
                table_owner.range.col_first,
                table_owner.range.row_last,
                table_owner.range.col_last,
            )));
        }
        if let Some(row_metadata_range) = worksheet
            .structural_owners
            .row_metadata_ranges
            .iter()
            .find(|range| {
                affected_rect.row_first <= range.row_last
                    && range.row_first <= affected_rect.row_last
                    && affected_rect.col_first <= range.col_last
                    && range.col_first <= affected_rect.col_last
            })
        {
            return Err(OmError::unsupported(format!(
                "Range.{member} structural row metadata retarget is not implemented for worksheet {} row {}",
                sheet_id.0, row_metadata_range.row_first,
            )));
        }
        if let Some(column_metadata_range) = worksheet
            .structural_owners
            .column_metadata_ranges
            .iter()
            .find(|range| {
                affected_rect.row_first <= range.row_last
                    && range.row_first <= affected_rect.row_last
                    && affected_rect.col_first <= range.col_last
                    && range.col_first <= affected_rect.col_last
            })
        {
            return Err(OmError::unsupported(format!(
                "Range.{member} structural column metadata retarget is not implemented for worksheet {} columns C{}:C{}",
                sheet_id.0, column_metadata_range.col_first, column_metadata_range.col_last,
            )));
        }
        let mut protected_keys = Vec::new();
        for spill_range in worksheet.spill_ranges.values() {
            if affected_rect.row_first <= spill_range.row_last
                && spill_range.row_first <= affected_rect.row_last
                && affected_rect.col_first <= spill_range.col_last
                && spill_range.col_first <= affected_rect.col_last
            {
                protected_keys.push((
                    affected_rect.row_first.max(spill_range.row_first),
                    affected_rect.col_first.max(spill_range.col_first),
                ));
            }
        }
        for &anchor in &worksheet.dynamic_array_formulas {
            if (affected_rect.row_first..=affected_rect.row_last).contains(&anchor.0)
                && (affected_rect.col_first..=affected_rect.col_last).contains(&anchor.1)
            {
                protected_keys.push(anchor);
            }
        }
        for &child in worksheet.spill_owners.keys() {
            if (affected_rect.row_first..=affected_rect.row_last).contains(&child.0)
                && (affected_rect.col_first..=affected_rect.col_last).contains(&child.1)
            {
                protected_keys.push(child);
            }
        }
        worksheet.ensure_spill_topology_is_not_modified(protected_keys, operation)?;

        let source_cells = worksheet
            .cells
            .iter()
            .filter_map(|(&key, cell)| {
                ((affected_rect.row_first..=affected_rect.row_last).contains(&key.0)
                    && (affected_rect.col_first..=affected_rect.col_last).contains(&key.1))
                .then_some((key, cell.clone()))
            })
            .collect::<Vec<_>>();
        let mut replacements = BTreeMap::new();
        for (key, _) in &source_cells {
            replacements.insert(*key, None);
        }
        match direction {
            CellShiftDirection::Up => {
                let shift_height = rect.height();
                for ((row, col), cell) in source_cells {
                    if row > rect.row_last {
                        replacements.insert((row - shift_height, col), Some(cell));
                    }
                }
            }
            CellShiftDirection::Left => {
                let shift_width = rect.width();
                for ((row, col), cell) in source_cells {
                    if col > rect.col_last {
                        replacements.insert((row, col - shift_width), Some(cell));
                    }
                }
            }
            CellShiftDirection::Down => {
                let shift_height = rect.height();
                for ((row, col), cell) in source_cells {
                    let target_row = row.checked_add(shift_height).ok_or_else(|| {
                        OmError::invalid_argument(
                            "Range.Insert would shift cells beyond worksheet rows",
                        )
                    })?;
                    if target_row > ExcelLimits::MAX_ROW_INDEX {
                        return Err(OmError::invalid_argument(
                            "Range.Insert would shift cells beyond worksheet rows",
                        ));
                    }
                    replacements.insert((target_row, col), Some(cell));
                }
            }
            CellShiftDirection::Right => {
                let shift_width = rect.width();
                for ((row, col), cell) in source_cells {
                    let target_col = col.checked_add(shift_width).ok_or_else(|| {
                        OmError::invalid_argument(
                            "Range.Insert would shift cells beyond worksheet columns",
                        )
                    })?;
                    if target_col > ExcelLimits::MAX_COLUMN_INDEX {
                        return Err(OmError::invalid_argument(
                            "Range.Insert would shift cells beyond worksheet columns",
                        ));
                    }
                    replacements.insert((row, target_col), Some(cell));
                }
            }
        }

        self.apply_cell_batch_with_change(sheet_id, replacements, CellBatchMutationKind::Rearrange)
    }

    pub fn fill_cells_with_change(
        &mut self,
        sheet_id: SheetId,
        source_keys: BTreeSet<(u32, u32)>,
        replacements: BTreeMap<(u32, u32), Option<CellData>>,
    ) -> OmResult<bool> {
        let worksheet = self.worksheet_data_for_sheet(sheet_id)?;
        for &(row, col) in &source_keys {
            ExcelLimits::validate_cell(row, col)?;
        }
        for &(row, col) in replacements.keys() {
            ExcelLimits::validate_cell(row, col)?;
        }
        if replacements.is_empty() {
            return Ok(false);
        }
        if source_keys.is_empty() {
            return Err(OmError::invalid_argument(
                "fill cell batch requires at least one source cell",
            ));
        }
        worksheet.ensure_spill_topology_is_not_modified(
            source_keys
                .iter()
                .copied()
                .chain(replacements.keys().copied()),
            "fill",
        )?;
        self.apply_cell_batch_with_change(sheet_id, replacements, CellBatchMutationKind::Rearrange)
    }

    pub fn validate_copy_source_cells(
        &self,
        sheet_id: SheetId,
        source_keys: &BTreeSet<(u32, u32)>,
    ) -> OmResult<()> {
        self.validate_transfer_source_cells(sheet_id, source_keys, "copy")
    }

    pub fn validate_cut_source_cells(
        &self,
        sheet_id: SheetId,
        source_keys: &BTreeSet<(u32, u32)>,
    ) -> OmResult<()> {
        self.validate_transfer_source_cells(sheet_id, source_keys, "cut")
    }

    fn validate_transfer_source_cells(
        &self,
        sheet_id: SheetId,
        source_keys: &BTreeSet<(u32, u32)>,
        operation: &str,
    ) -> OmResult<()> {
        if source_keys.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "{operation} source requires at least one cell"
            )));
        }
        for &(row, col) in source_keys {
            ExcelLimits::validate_cell(row, col)?;
        }
        self.worksheet_data_for_sheet(sheet_id)?
            .ensure_spill_topology_is_not_modified(source_keys.iter().copied(), operation)
    }

    pub fn copy_cells_with_change(
        &mut self,
        sheet_id: SheetId,
        replacements: BTreeMap<(u32, u32), Option<CellData>>,
    ) -> OmResult<bool> {
        self.transfer_cells_with_change(sheet_id, replacements, "copy into")
    }

    pub fn cut_cells_with_change(
        &mut self,
        sheet_id: SheetId,
        replacements: BTreeMap<(u32, u32), Option<CellData>>,
    ) -> OmResult<bool> {
        self.transfer_cells_with_change(sheet_id, replacements, "cut")
    }

    fn transfer_cells_with_change(
        &mut self,
        sheet_id: SheetId,
        replacements: BTreeMap<(u32, u32), Option<CellData>>,
        operation: &str,
    ) -> OmResult<bool> {
        let worksheet = self.worksheet_data_for_sheet(sheet_id)?;
        for &(row, col) in replacements.keys() {
            ExcelLimits::validate_cell(row, col)?;
        }
        if replacements.is_empty() {
            return Ok(false);
        }
        worksheet.ensure_spill_topology_is_not_modified(replacements.keys().copied(), operation)?;
        self.apply_cell_batch_with_change(sheet_id, replacements, CellBatchMutationKind::Rearrange)
    }

    fn apply_cell_batch_with_change(
        &mut self,
        sheet_id: SheetId,
        replacements: BTreeMap<(u32, u32), Option<CellData>>,
        mutation_kind: CellBatchMutationKind,
    ) -> OmResult<bool> {
        let worksheet = self.worksheet_data_for_sheet(sheet_id)?;
        let mut updates = Vec::with_capacity(replacements.len());
        for (key, replacement) in replacements {
            ExcelLimits::validate_cell(key.0, key.1)?;
            if let Some(cell) = &replacement {
                cell.value.validate()?;
            }
            if worksheet.cells.get(&key) != replacement.as_ref() {
                updates.push((key, replacement));
            }
        }
        match mutation_kind {
            CellBatchMutationKind::Replace => {
                worksheet
                    .ensure_spill_children_are_not_edited(updates.iter().map(|(key, _)| *key))?;
            }
            CellBatchMutationKind::Rearrange => {
                worksheet.ensure_spill_topology_is_not_modified(
                    updates.iter().map(|(key, _)| *key),
                    "rearrange",
                )?;
            }
        }

        let worksheet = self.worksheet_data_for_sheet_mut(sheet_id)?;
        for (key, replacement) in &updates {
            let preserve_dynamic_formula = mutation_kind == CellBatchMutationKind::Replace
                && worksheet.dynamic_array_formulas.contains(key)
                && replacement
                    .as_ref()
                    .is_some_and(|cell| cell.formula.is_some());
            worksheet.prepare_cell_for_edit_with_change(*key);
            match replacement {
                Some(cell) => {
                    worksheet.cells.insert(*key, cell.clone());
                }
                None => {
                    worksheet.cells.remove(key);
                }
            }
            if preserve_dynamic_formula {
                worksheet.dynamic_array_formulas.insert(*key);
            }
            worksheet.dirty = true;
            worksheet.dirty_cells.insert(*key);
        }

        Ok(!updates.is_empty())
    }

    pub fn clear_range_contents(&mut self, range: &RangeRef) -> OmResult<()> {
        self.clear_range_contents_with_change(range).map(|_| ())
    }

    pub fn clear_range_contents_with_change(&mut self, range: &RangeRef) -> OmResult<bool> {
        let (sheet_id, rects) = self.same_sheet_rects(range)?;
        let worksheet = self.worksheet_data_for_sheet_mut(sheet_id)?;
        worksheet.ensure_spill_children_are_not_edited(rects.iter().flat_map(|rect| {
            (rect.row_first..=rect.row_last)
                .flat_map(move |row| (rect.col_first..=rect.col_last).map(move |col| (row, col)))
        }))?;
        let mut changed_any = false;
        for rect in rects {
            for row in rect.row_first..=rect.row_last {
                for col in rect.col_first..=rect.col_last {
                    let key = (row, col);
                    let mut changed = false;
                    let mut remove = false;

                    if worksheet.spill_owners.contains_key(&key)
                        || worksheet.spill_ranges.contains_key(&key)
                        || worksheet.dynamic_array_formulas.contains(&key)
                    {
                        worksheet.prepare_cell_for_edit(key);
                        changed = true;
                    }

                    if let Some(existing) = worksheet.cells.get_mut(&key) {
                        if matches!(existing.value, CellValue::Blank) && existing.formula.is_none()
                        {
                            if !changed {
                                continue;
                            }
                        } else {
                            existing.value = CellValue::Blank;
                            existing.formula = None;
                            remove = existing.style_id.is_none();
                            changed = true;
                        }
                    }

                    if changed {
                        changed_any = true;
                        if remove {
                            worksheet.cells.remove(&key);
                        }
                        worksheet.dirty = true;
                        worksheet.dirty_cells.insert(key);
                    }
                }
            }
        }

        Ok(changed_any)
    }

    pub fn clear_range_with_change(&mut self, range: &RangeRef) -> OmResult<bool> {
        let (sheet_id, rects) = self.same_sheet_rects(range)?;
        let worksheet = self.worksheet_data_for_sheet_mut(sheet_id)?;
        worksheet.ensure_spill_children_are_not_edited(rects.iter().flat_map(|rect| {
            (rect.row_first..=rect.row_last)
                .flat_map(move |row| (rect.col_first..=rect.col_last).map(move |col| (row, col)))
        }))?;
        let mut changed_any = false;
        for rect in rects {
            for row in rect.row_first..=rect.row_last {
                for col in rect.col_first..=rect.col_last {
                    let key = (row, col);
                    let metadata_changed = worksheet.prepare_cell_for_edit_with_change(key);
                    let cell_changed = worksheet.cells.remove(&key).is_some();
                    if metadata_changed || cell_changed {
                        worksheet.dirty = true;
                        worksheet.dirty_cells.insert(key);
                        changed_any = true;
                    }
                }
            }
        }

        Ok(changed_any)
    }

    pub fn clear_range_formats_with_change(&mut self, range: &RangeRef) -> OmResult<bool> {
        let (sheet_id, rects) = self.same_sheet_rects(range)?;
        let worksheet = self.worksheet_data_for_sheet_mut(sheet_id)?;
        let mut changed_any = false;
        for rect in rects {
            for row in rect.row_first..=rect.row_last {
                for col in rect.col_first..=rect.col_last {
                    let key = (row, col);
                    let preserve_topology_cell = worksheet.spill_owner_for_key(key).is_some()
                        || worksheet.spill_ranges.contains_key(&key)
                        || worksheet.dynamic_array_formulas.contains(&key);
                    let mut remove_cell = false;
                    if let Some(existing) = worksheet.cells.get_mut(&key)
                        && existing.style_id.is_some()
                    {
                        existing.style_id = None;
                        remove_cell = !preserve_topology_cell
                            && matches!(existing.value, CellValue::Blank)
                            && existing.formula.is_none();
                        worksheet.dirty = true;
                        worksheet.dirty_cells.insert(key);
                        changed_any = true;
                    }
                    if remove_cell {
                        worksheet.cells.remove(&key);
                    }
                }
            }
        }

        Ok(changed_any)
    }

    fn single_sheet_rect(&self, range: &RangeRef) -> OmResult<(SheetId, Rect)> {
        let (sheet_id, rects) = self.same_sheet_rects(range)?;
        if rects.len() != 1 {
            return Err(OmError::unsupported(
                "multi-area ranges are not supported for this worksheet operation",
            ));
        }

        Ok((sheet_id, rects[0]))
    }

    fn same_sheet_rects(&self, range: &RangeRef) -> OmResult<(SheetId, Vec<Rect>)> {
        if range.workbook_id != self.model.id {
            return Err(OmError::invalid_argument(format!(
                "range workbook {} does not match loaded workbook {}",
                range.workbook_id.0, self.model.id.0,
            )));
        }

        if range.areas.is_empty() {
            return Err(OmError::invalid_argument(
                "worksheet range must contain at least one area",
            ));
        }

        for rect in &range.areas {
            ExcelLimits::validate_rect(*rect)?;
        }

        let sheet_id = match range.scope {
            SheetScope::Single(sheet_id) => sheet_id,
            SheetScope::Multi3D { .. } => {
                return Err(OmError::unsupported(
                    "3D ranges are not supported in the core worksheet model yet",
                ));
            }
        };

        self.ensure_worksheet_exists(sheet_id)?;

        Ok((sheet_id, range.areas.clone()))
    }
}

fn checked_rect_cell_count_sum(rects: &[Rect]) -> OmResult<usize> {
    rects.iter().try_fold(0usize, |total, rect| {
        total
            .checked_add(rect.checked_cell_count_usize()?)
            .ok_or_else(|| OmError::resource_limit("worksheet range cell count exceeds usize"))
    })
}

#[cfg(test)]
mod tests {
    use super::{
        CellData, CellShiftDirection, ChartModel, ChartObjectModel, ChartSheetBinding,
        ChartSourceExpr, ChartType, DefinedNameTable, DrawingModel, DrawingObjectModel,
        SeriesModel, TableStructuralOwner, WorkbookState, WorkbookStateParts, WorksheetData,
        WorksheetStructuralOwners,
    };
    use std::collections::{BTreeMap, BTreeSet};

    use office_common::{
        CellValue, ChartId, ChartObjectId, DrawingId, ExcelLimits, FileFormat, FormulaSource,
        NameScope, NameValidationMode, ObjectHandle, ObjectPlacement, OmArray, OmErrorCode,
        OmValue, RangeRef, RangeSet, Rect, ReferenceTarget, SheetId, SheetKind, SheetScope,
        SheetVisibility, StyleId, WorkbookId, WorkbookModel, WorksheetModel,
    };

    fn sample_state() -> WorkbookState {
        let workbook_id = WorkbookId(7);
        let sheet_id = SheetId(3);
        WorkbookState {
            model: WorkbookModel {
                id: workbook_id,
                display_name: "Workbook".to_string(),
                format: FileFormat::Xlsx,
                date1904: false,
                is_addin: false,
            },
            worksheets: vec![WorksheetModel {
                id: sheet_id,
                workbook_id,
                name: "Sheet1".to_string(),
                kind: SheetKind::Worksheet,
                visibility: SheetVisibility::Visible,
                relationship_id: Some("rId1".to_string()),
                part_uri: Some("xl/worksheets/sheet1.xml".to_string()),
            }],
            worksheet_data: BTreeMap::from([(
                sheet_id,
                WorksheetData {
                    cells: BTreeMap::from([
                        (
                            (1, 1),
                            CellData {
                                value: CellValue::Number(42.0),
                                formula: None,
                                style_id: None,
                            },
                        ),
                        (
                            (2, 2),
                            CellData {
                                value: CellValue::Text("hello".to_string()),
                                formula: None,
                                style_id: Some(StyleId(9)),
                            },
                        ),
                    ]),
                    source_xml: b"<worksheet/>".to_vec(),
                    dirty: false,
                    dirty_cells: BTreeSet::new(),
                    ..WorksheetData::default()
                },
            )]),
            defined_names: DefinedNameTable::default(),
            charts: BTreeMap::new(),
            drawings: BTreeMap::new(),
            chart_sheets: BTreeMap::new(),
            opaque_parts: Vec::new(),
        }
    }

    fn seed_two_by_two_spill(state: &mut WorkbookState) {
        let anchor = (3, 3);
        let worksheet = state
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data");
        worksheet.cells.insert(
            anchor,
            CellData {
                value: CellValue::Number(1.0),
                formula: Some(FormulaSource {
                    text: "SEQUENCE(2,2)".to_string(),
                    is_r1c1: false,
                }),
                style_id: None,
            },
        );
        for (key, value, style_id) in [
            ((3, 4), 2.0, None),
            ((4, 3), 3.0, None),
            ((4, 4), 4.0, Some(StyleId(17))),
        ] {
            worksheet.cells.insert(
                key,
                CellData {
                    value: CellValue::Number(value),
                    formula: None,
                    style_id,
                },
            );
            worksheet.spill_owners.insert(key, anchor);
        }
        worksheet.dynamic_array_formulas.insert(anchor);
        worksheet.spill_ranges.insert(
            anchor,
            Rect {
                row_first: 3,
                row_last: 4,
                col_first: 3,
                col_last: 4,
            },
        );
    }

    fn add_second_worksheet(state: &mut WorkbookState) {
        let mut second = state.worksheets[0].clone();
        second.id = SheetId(4);
        second.name = "Sheet2".to_string();
        second.relationship_id = Some("rId2".to_string());
        second.part_uri = Some("xl/worksheets/sheet2.xml".to_string());
        state.worksheets.push(second);
        state
            .worksheet_data
            .insert(SheetId(4), WorksheetData::default());
    }

    fn formula_source(text: &str) -> FormulaSource {
        FormulaSource {
            text: text.to_string(),
            is_r1c1: false,
        }
    }

    fn chart_model(
        chart_id: ChartId,
        workbook_id: WorkbookId,
        series: Vec<SeriesModel>,
    ) -> ChartModel {
        ChartModel {
            id: chart_id,
            workbook_id,
            chart_type: ChartType::Bar,
            style: None,
            series,
            title: None,
            legend: None,
            axes: Vec::new(),
            groups: Vec::new(),
            vary_by_categories: None,
            gap_width: None,
            gap_depth: None,
            overlap: None,
            bar_shape: None,
            has_series_lines: None,
            has_drop_lines: None,
            has_hi_lo_lines: None,
            has_up_down_bars: None,
            first_slice_angle: None,
            explosion: None,
            bubble_scale: None,
            show_negative_bubbles: None,
            has_3d_shading: None,
            doughnut_hole_size: None,
            second_plot_size: None,
            size_represents: None,
            split_type: None,
            split_value: None,
            data_labels: None,
            data_table: None,
            data_table_dirty: false,
            plot_area_layout: None,
            plot_area_layout_dirty: false,
            show_data_labels_over_maximum: None,
            display_blanks_as: None,
            plot_visible_only: None,
            view_3d: None,
            view_3d_dirty: false,
            rounded_corners: None,
            protection: None,
            protection_dirty: false,
            raw_part_uri: Some("xl/charts/chart1.xml".to_string()),
            series_topology_dirty: false,
            content_dirty: false,
            dirty: false,
        }
    }

    fn series_model_with_values(values: ChartSourceExpr) -> SeriesModel {
        SeriesModel {
            name: None,
            x_values: None,
            values: Some(values),
            bubble_size: None,
            bar_shape: None,
            smooth: None,
            marker_style: None,
            marker_size: None,
            invert_if_negative: None,
            points: BTreeMap::new(),
            data_labels: None,
            point_data_labels: BTreeMap::new(),
            raw_index: None,
            order: Some(0),
            axis_group: super::ChartAxisGroup::Primary,
            is_filtered: false,
            filter_dirty: false,
        }
    }

    fn malformed_empty_range_set() -> RangeSet {
        serde_json::from_str(r#"{"workbook_id":7,"areas":[]}"#)
            .expect("serde can currently materialize an invalid private-field range set")
    }

    #[test]
    fn workbook_state_validate_for_save_accepts_consistent_model() {
        sample_state()
            .validate_for_save()
            .expect("consistent workbook state");
    }

    #[test]
    fn non_finite_cell_values_fail_model_mutation_and_save_boundaries() {
        for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            let invalid_cell = CellData {
                value: CellValue::Number(value),
                formula: None,
                style_id: None,
            };

            let mut insert_state = sample_state();
            let insert_before = insert_state.clone();
            let insert_error = insert_state
                .insert_cell(SheetId(3), 5, 5, invalid_cell.clone())
                .expect_err("insert_cell must reject a non-finite number");
            assert_eq!(insert_error.code, OmErrorCode::InvalidArgument);
            assert_eq!(insert_state, insert_before);

            let mut batch_state = sample_state();
            let batch_before = batch_state.clone();
            let batch_error = batch_state
                .copy_cells_with_change(
                    SheetId(3),
                    BTreeMap::from([((5, 5), Some(invalid_cell.clone()))]),
                )
                .expect_err("cell batch must reject a non-finite number");
            assert_eq!(batch_error.code, OmErrorCode::InvalidArgument);
            assert_eq!(batch_state, batch_before);

            let mut invalid_state = sample_state();
            invalid_state
                .worksheet_data
                .get_mut(&SheetId(3))
                .expect("worksheet data")
                .cells
                .insert((5, 5), invalid_cell);
            let save_error = invalid_state
                .validate_for_save()
                .expect_err("save preflight must reject a non-finite number");
            assert_eq!(save_error.code, OmErrorCode::InvalidState);
            assert_eq!(
                save_error.message,
                "worksheet 3 cell R5C5 numeric value must be finite"
            );
        }
    }

    #[test]
    fn structural_cell_shifts_fail_closed_for_reference_formulas() {
        for (formula_text, is_r1c1) in [("A1", false), ("R[-1]C", true)] {
            let mut state = sample_state();
            state
                .worksheet_data
                .get_mut(&SheetId(3))
                .expect("worksheet data")
                .cells
                .insert(
                    (10, 10),
                    CellData {
                        value: CellValue::Blank,
                        formula: Some(FormulaSource {
                            text: formula_text.to_string(),
                            is_r1c1,
                        }),
                        style_id: None,
                    },
                );
            let before = state.clone();

            let error = state
                .shift_cells_with_change(
                    SheetId(3),
                    Rect::single_cell(1, 1),
                    CellShiftDirection::Down,
                )
                .expect_err("reference-bearing structural shift must fail closed");

            assert_eq!(error.code, OmErrorCode::Unsupported, "{formula_text}");
            assert_eq!(
                error.message,
                "Range.Insert structural formula retarget is not implemented for worksheet 3 cell R10C10",
                "{formula_text}",
            );
            assert_eq!(state, before, "{formula_text}");
        }
    }

    #[test]
    fn structural_cell_shifts_fail_closed_for_reference_defined_names() {
        for (scope, refers_to, is_r1c1, expected_scope) in [
            (NameScope::Workbook, "Sheet1!$M$50", false, "workbook"),
            (
                NameScope::Worksheet(SheetId(3)),
                "Sheet1!R1C1",
                true,
                "worksheet 3",
            ),
        ] {
            let mut state = sample_state();
            state
                .add_defined_name(
                    scope,
                    "ShiftOwner",
                    FormulaSource {
                        text: refers_to.to_string(),
                        is_r1c1,
                    },
                    NameValidationMode::StrictExcel,
                )
                .expect("seed defined name");
            let before = state.clone();

            let error = state
                .shift_cells_with_change(
                    SheetId(3),
                    Rect::single_cell(1, 1),
                    CellShiftDirection::Down,
                )
                .expect_err("reference-bearing defined name must fail closed");

            assert_eq!(error.code, OmErrorCode::Unsupported, "{refers_to}");
            assert_eq!(
                error.message,
                format!(
                    "Range.Insert structural defined-name retarget is not implemented for {expected_scope} name 'ShiftOwner'"
                ),
                "{refers_to}",
            );
            assert_eq!(state, before, "{refers_to}");
        }

        let mut state = sample_state();
        state
            .add_defined_name(
                NameScope::Workbook,
                "ConstantOwner",
                FormulaSource {
                    text: r#""A1""#.to_string(),
                    is_r1c1: false,
                },
                NameValidationMode::StrictExcel,
            )
            .expect("seed reference-free defined name");

        assert!(
            state
                .shift_cells_with_change(
                    SheetId(3),
                    Rect::single_cell(1, 1),
                    CellShiftDirection::Down,
                )
                .expect("reference-free defined name must remain eligible"),
        );
        assert_eq!(
            state
                .lookup_name_in_scope(NameScope::Workbook, "ConstantOwner")
                .expect("reference-free defined name after shift")
                .refers_to
                .text,
            r#""A1""#,
        );
    }

    #[test]
    fn structural_cell_shifts_inventory_chart_source_owners_atomically() {
        let chart_id = ChartId(11);
        let chart_range = Rect {
            row_first: 4,
            row_last: 5,
            col_first: 4,
            col_last: 5,
        };
        let chart_source = |workbook_id| ChartSourceExpr {
            raw: formula_source("Sheet1!$D$4:$E$5"),
            resolved: Some(ReferenceTarget::Range(
                RangeSet::single_rect(workbook_id, SheetId(3), chart_range)
                    .expect("chart source range"),
            )),
            full_reference: None,
            cache: None,
            dirty: false,
        };

        let mut blocked = sample_state();
        blocked.charts.insert(
            chart_id,
            chart_model(
                chart_id,
                blocked.model.id,
                vec![series_model_with_values(chart_source(blocked.model.id))],
            ),
        );
        let before = blocked.clone();
        let error = blocked
            .shift_cells_with_change(
                SheetId(3),
                Rect {
                    row_first: 1,
                    row_last: 1,
                    col_first: 4,
                    col_last: 5,
                },
                CellShiftDirection::Down,
            )
            .expect_err("intersecting chart source must fail closed");
        assert_eq!(error.code, OmErrorCode::Unsupported);
        assert_eq!(
            error.message,
            "Range.Insert structural chart source retarget is not implemented for chart 11 series 1 values worksheet 3 range R4C4:R5C5",
        );
        assert_eq!(blocked, before);

        let mut full_reference = sample_state();
        full_reference.charts.insert(
            chart_id,
            chart_model(
                chart_id,
                full_reference.model.id,
                vec![series_model_with_values(ChartSourceExpr {
                    raw: formula_source("42"),
                    resolved: Some(ReferenceTarget::Value(CellValue::Number(42.0))),
                    full_reference: Some(super::ChartSourceReference {
                        raw: formula_source("Sheet1!$D$4:$E$5"),
                        resolved: Some(ReferenceTarget::Range(
                            RangeSet::single_rect(full_reference.model.id, SheetId(3), chart_range)
                                .expect("full chart source range"),
                        )),
                    }),
                    cache: None,
                    dirty: false,
                })],
            ),
        );
        let before = full_reference.clone();
        let error = full_reference
            .shift_cells_with_change(
                SheetId(3),
                Rect {
                    row_first: 1,
                    row_last: 1,
                    col_first: 4,
                    col_last: 5,
                },
                CellShiftDirection::Down,
            )
            .expect_err("intersecting full chart source must fail closed");
        assert_eq!(error.code, OmErrorCode::Unsupported);
        assert!(
            error
                .message
                .contains("chart 11 series 1 values full-reference worksheet 3 range R4C4:R5C5"),
            "{error:?}",
        );
        assert_eq!(full_reference, before);

        let mut unresolved = sample_state();
        unresolved.charts.insert(
            chart_id,
            chart_model(
                chart_id,
                unresolved.model.id,
                vec![series_model_with_values(ChartSourceExpr {
                    raw: FormulaSource {
                        text: "R1C1:R2C2".to_string(),
                        is_r1c1: true,
                    },
                    resolved: None,
                    full_reference: None,
                    cache: None,
                    dirty: false,
                })],
            ),
        );
        let before = unresolved.clone();
        let error = unresolved
            .shift_cells_with_change(
                SheetId(3),
                Rect::single_cell(1, 1),
                CellShiftDirection::Down,
            )
            .expect_err("unresolved chart reference must fail closed");
        assert_eq!(error.code, OmErrorCode::Unsupported);
        assert!(
            error
                .message
                .contains("chart 11 series 1 values unresolved reference"),
            "{error:?}",
        );
        assert_eq!(unresolved, before);

        let mut allowed = sample_state();
        allowed.charts.insert(
            chart_id,
            chart_model(
                chart_id,
                allowed.model.id,
                vec![series_model_with_values(chart_source(allowed.model.id))],
            ),
        );
        let chart_before = allowed.charts.get(&chart_id).expect("chart").clone();
        assert!(
            allowed
                .shift_cells_with_change(
                    SheetId(3),
                    Rect::single_cell(1, 1),
                    CellShiftDirection::Down,
                )
                .expect("non-intersecting chart source must remain eligible"),
        );
        assert_eq!(allowed.charts.get(&chart_id), Some(&chart_before));
    }

    #[test]
    fn structural_cell_shifts_reject_only_intersecting_merged_ranges() {
        let merged_range = Rect {
            row_first: 4,
            row_last: 5,
            col_first: 4,
            col_last: 5,
        };
        let mut blocked = sample_state();
        blocked
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .structural_owners = WorksheetStructuralOwners {
            merged_ranges: vec![merged_range],
            ..WorksheetStructuralOwners::default()
        };
        let before = blocked.clone();

        let error = blocked
            .shift_cells_with_change(
                SheetId(3),
                Rect {
                    row_first: 1,
                    row_last: 1,
                    col_first: 4,
                    col_last: 5,
                },
                CellShiftDirection::Down,
            )
            .expect_err("intersecting merged range must fail closed");

        assert_eq!(error.code, OmErrorCode::Unsupported);
        assert_eq!(
            error.message,
            "Range.Insert structural merged-cell retarget is not implemented for worksheet 3 range R4C4:R5C5",
        );
        assert_eq!(blocked, before);

        let mut allowed = sample_state();
        allowed
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .structural_owners = WorksheetStructuralOwners {
            merged_ranges: vec![merged_range],
            ..WorksheetStructuralOwners::default()
        };
        assert!(
            allowed
                .shift_cells_with_change(
                    SheetId(3),
                    Rect::single_cell(1, 1),
                    CellShiftDirection::Down,
                )
                .expect("non-intersecting merged range must remain eligible"),
        );
        assert_eq!(
            allowed
                .worksheet_data_for_sheet(SheetId(3))
                .expect("worksheet data after shift")
                .structural_owners
                .merged_ranges,
            vec![merged_range],
        );
    }

    #[test]
    fn structural_cell_shifts_reject_only_intersecting_data_validation_ranges() {
        let validation_range = Rect {
            row_first: 4,
            row_last: 5,
            col_first: 4,
            col_last: 5,
        };
        let mut blocked = sample_state();
        blocked
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .structural_owners = WorksheetStructuralOwners {
            data_validation_ranges: vec![validation_range],
            ..WorksheetStructuralOwners::default()
        };
        let before = blocked.clone();

        let error = blocked
            .shift_cells_with_change(
                SheetId(3),
                Rect {
                    row_first: 1,
                    row_last: 1,
                    col_first: 4,
                    col_last: 5,
                },
                CellShiftDirection::Down,
            )
            .expect_err("intersecting data validation must fail closed");

        assert_eq!(error.code, OmErrorCode::Unsupported);
        assert_eq!(
            error.message,
            "Range.Insert structural data-validation retarget is not implemented for worksheet 3 range R4C4:R5C5",
        );
        assert_eq!(blocked, before);

        let mut allowed = sample_state();
        allowed
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .structural_owners = WorksheetStructuralOwners {
            data_validation_ranges: vec![validation_range],
            ..WorksheetStructuralOwners::default()
        };
        assert!(
            allowed
                .shift_cells_with_change(
                    SheetId(3),
                    Rect::single_cell(1, 1),
                    CellShiftDirection::Down,
                )
                .expect("non-intersecting data validation must remain eligible"),
        );
        assert_eq!(
            allowed
                .worksheet_data_for_sheet(SheetId(3))
                .expect("worksheet data after shift")
                .structural_owners
                .data_validation_ranges,
            vec![validation_range],
        );
    }

    #[test]
    fn structural_cell_shifts_reject_data_validation_reference_formulas() {
        let mut blocked = sample_state();
        blocked
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .structural_owners = WorksheetStructuralOwners {
            data_validation_formulas: vec!["=$D$4>0".to_string()],
            ..WorksheetStructuralOwners::default()
        };
        let before = blocked.clone();

        let error = blocked
            .shift_cells_with_change(
                SheetId(3),
                Rect::single_cell(1, 1),
                CellShiftDirection::Down,
            )
            .expect_err("data-validation reference formula must fail closed");

        assert_eq!(error.code, OmErrorCode::Unsupported);
        assert_eq!(
            error.message,
            "Range.Insert structural data-validation formula retarget is not implemented for worksheet 3 formula 1",
        );
        assert_eq!(blocked, before);

        let mut allowed = sample_state();
        allowed
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .structural_owners = WorksheetStructuralOwners {
            data_validation_formulas: vec!["1".to_string(), r#""A1""#.to_string()],
            ..WorksheetStructuralOwners::default()
        };
        assert!(
            allowed
                .shift_cells_with_change(
                    SheetId(3),
                    Rect::single_cell(1, 1),
                    CellShiftDirection::Down,
                )
                .expect("reference-free data-validation formulas must remain eligible"),
        );
    }

    #[test]
    fn structural_cell_shifts_reject_table_relationship_owners() {
        let mut state = sample_state();
        state
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .structural_owners = WorksheetStructuralOwners {
            table_relationship_ids: vec!["rIdTable1".to_string()],
            ..WorksheetStructuralOwners::default()
        };
        let before = state.clone();

        let error = state
            .shift_cells_with_change(
                SheetId(3),
                Rect::single_cell(1, 1),
                CellShiftDirection::Down,
            )
            .expect_err("table relationship owner must fail closed");

        assert_eq!(error.code, OmErrorCode::Unsupported);
        assert_eq!(
            error.message,
            "Range.Insert structural table owner retarget is not implemented for worksheet 3 relationship rIdTable1",
        );
        assert_eq!(state, before);
        assert_eq!(
            state
                .validate_for_save()
                .expect_err("unresolved table relationship must fail save validation")
                .code,
            OmErrorCode::InvalidState,
        );
    }

    #[test]
    fn structural_cell_shifts_use_resolved_table_ranges_and_formulas() {
        let table_range = Rect {
            row_first: 4,
            row_last: 5,
            col_first: 4,
            col_last: 5,
        };
        let owners = |formulas| WorksheetStructuralOwners {
            table_relationship_ids: vec!["rIdTable1".to_string()],
            table_owners: vec![TableStructuralOwner {
                relationship_id: "rIdTable1".to_string(),
                part_uri: "xl/tables/table1.xml".to_string(),
                range: table_range,
                formulas,
            }],
            ..WorksheetStructuralOwners::default()
        };

        let mut blocked = sample_state();
        blocked
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .structural_owners = owners(Vec::new());
        let before = blocked.clone();
        let error = blocked
            .shift_cells_with_change(
                SheetId(3),
                Rect {
                    row_first: 4,
                    row_last: 4,
                    col_first: 4,
                    col_last: 5,
                },
                CellShiftDirection::Up,
            )
            .expect_err("intersecting resolved table range must fail closed");
        assert_eq!(error.code, OmErrorCode::Unsupported);
        assert_eq!(
            error.message,
            "Range.Delete structural table range retarget is not implemented for worksheet 3 relationship rIdTable1 part xl/tables/table1.xml range R4C4:R5C5",
        );
        assert_eq!(blocked, before);

        let mut allowed = sample_state();
        allowed
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .structural_owners = owners(Vec::new());
        assert!(
            allowed
                .shift_cells_with_change(
                    SheetId(3),
                    Rect::single_cell(1, 1),
                    CellShiftDirection::Down,
                )
                .expect("non-intersecting resolved table range must remain eligible"),
        );
        allowed
            .validate_for_save()
            .expect("resolved table owner must pass save validation");

        let mut formula_blocked = sample_state();
        formula_blocked
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .structural_owners = owners(vec!["=$A$1+1".to_string()]);
        let formula_before = formula_blocked.clone();
        let formula_error = formula_blocked
            .shift_cells_with_change(
                SheetId(3),
                Rect::single_cell(1, 1),
                CellShiftDirection::Down,
            )
            .expect_err("table reference formula must fail closed");
        assert_eq!(formula_error.code, OmErrorCode::Unsupported);
        assert_eq!(
            formula_error.message,
            "Range.Insert structural table formula retarget is not implemented for worksheet 3 relationship rIdTable1 part xl/tables/table1.xml formula 1",
        );
        assert_eq!(formula_blocked, formula_before);
    }

    #[test]
    fn structural_cell_shifts_fail_closed_on_row_and_column_metadata() {
        let row_metadata = Rect {
            row_first: 4,
            row_last: 4,
            col_first: 1,
            col_last: ExcelLimits::MAX_COLUMN_INDEX,
        };
        let column_metadata = Rect {
            row_first: 1,
            row_last: ExcelLimits::MAX_ROW_INDEX,
            col_first: 4,
            col_last: 5,
        };

        let mut row_blocked = sample_state();
        row_blocked
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .structural_owners
            .row_metadata_ranges = vec![row_metadata];
        let row_before = row_blocked.clone();
        let row_error = row_blocked
            .shift_cells_with_change(
                SheetId(3),
                Rect::single_cell(1, 1),
                CellShiftDirection::Down,
            )
            .expect_err("vertical shift through row metadata must fail closed");
        assert_eq!(row_error.code, OmErrorCode::Unsupported);
        assert_eq!(
            row_error.message,
            "Range.Insert structural row metadata retarget is not implemented for worksheet 3 row 4",
        );
        assert_eq!(row_blocked, row_before);

        let mut row_allowed = sample_state();
        row_allowed
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .structural_owners
            .row_metadata_ranges = vec![row_metadata];
        assert!(
            row_allowed
                .shift_cells_with_change(
                    SheetId(3),
                    Rect::single_cell(1, 1),
                    CellShiftDirection::Right,
                )
                .expect("shift outside row metadata must remain eligible"),
        );
        row_allowed
            .validate_for_save()
            .expect("row metadata owner must remain valid");

        let mut column_blocked = sample_state();
        column_blocked
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .structural_owners
            .column_metadata_ranges = vec![column_metadata];
        let column_before = column_blocked.clone();
        let column_error = column_blocked
            .shift_cells_with_change(
                SheetId(3),
                Rect::single_cell(1, 1),
                CellShiftDirection::Right,
            )
            .expect_err("horizontal shift through column metadata must fail closed");
        assert_eq!(column_error.code, OmErrorCode::Unsupported);
        assert_eq!(
            column_error.message,
            "Range.Insert structural column metadata retarget is not implemented for worksheet 3 columns C4:C5",
        );
        assert_eq!(column_blocked, column_before);

        let mut column_allowed = sample_state();
        column_allowed
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .structural_owners
            .column_metadata_ranges = vec![column_metadata];
        assert!(
            column_allowed
                .shift_cells_with_change(
                    SheetId(3),
                    Rect::single_cell(1, 1),
                    CellShiftDirection::Down,
                )
                .expect("shift outside column metadata must remain eligible"),
        );
        column_allowed
            .validate_for_save()
            .expect("column metadata owner must remain valid");

        let mut invalid = sample_state();
        invalid
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .structural_owners
            .row_metadata_ranges = vec![Rect {
            col_last: ExcelLimits::MAX_COLUMN_INDEX - 1,
            ..row_metadata
        }];
        assert_eq!(
            invalid
                .validate_for_save()
                .expect_err("partial row metadata owner must fail validation")
                .code,
            OmErrorCode::InvalidState,
        );
    }

    #[test]
    fn workbook_state_validate_for_save_accepts_blocked_and_materialized_spills() {
        let mut blocked = sample_state();
        let blocked_worksheet = blocked
            .worksheet_data
            .get_mut(&SheetId(3))
            .expect("worksheet data");
        blocked_worksheet.cells.insert(
            (3, 3),
            CellData {
                value: CellValue::Error(office_common::CellError::Spill),
                formula: Some(formula_source("SEQUENCE(2,2)")),
                style_id: None,
            },
        );
        blocked_worksheet.dynamic_array_formulas.insert((3, 3));
        blocked
            .validate_for_save()
            .expect("blocked spill has no materialized range");

        let mut materialized = sample_state();
        seed_two_by_two_spill(&mut materialized);
        materialized
            .validate_for_save()
            .expect("materialized spill topology");
    }

    #[test]
    fn workbook_state_validate_for_save_rejects_empty_worksheet_collection() {
        let mut state = sample_state();
        state.worksheets.clear();

        let error = state
            .validate_for_save()
            .expect_err("empty workbook must fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("at least one worksheet"));
    }

    #[test]
    fn workbook_state_validate_for_save_rejects_worksheet_workbook_id_drift() {
        let mut state = sample_state();
        state.worksheets[0].workbook_id = WorkbookId(99);

        let error = state
            .validate_for_save()
            .expect_err("worksheet ownership drift must fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("worksheet Sheet1"));
        assert!(error.message.contains("workbook id 99"));
        assert!(error.message.contains("expected 7"));
    }

    #[test]
    fn workbook_state_validate_for_save_rejects_duplicate_worksheet_identity() {
        let mut state = sample_state();
        let mut duplicate = state.worksheets[0].clone();
        duplicate.name = "Sheet2".to_string();
        state.worksheets.push(duplicate);

        let error = state
            .validate_for_save()
            .expect_err("duplicate worksheet id must fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("duplicate worksheet id 3"));
    }

    #[test]
    fn workbook_state_validate_for_save_rejects_duplicate_worksheet_name() {
        let mut state = sample_state();
        add_second_worksheet(&mut state);
        state.worksheets[1].name = "sHeEt1".to_string();

        let error = state
            .validate_for_save()
            .expect_err("duplicate worksheet name must fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("duplicate worksheet name"));
        assert!(error.message.contains("sHeEt1"));
    }

    #[test]
    fn workbook_state_validate_for_save_rejects_duplicate_worksheet_relationship_id() {
        let mut state = sample_state();
        add_second_worksheet(&mut state);
        state.worksheets[1].relationship_id = Some("rId1".to_string());

        let error = state
            .validate_for_save()
            .expect_err("duplicate worksheet relationship id must fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(
            error
                .message
                .contains("duplicate worksheet relationship id rId1")
        );
    }

    #[test]
    fn workbook_state_validate_for_save_rejects_duplicate_worksheet_part_uri() {
        let mut state = sample_state();
        add_second_worksheet(&mut state);
        state.worksheets[1].part_uri = Some("xl/worksheets/sheet1.xml".to_string());

        let error = state
            .validate_for_save()
            .expect_err("duplicate worksheet part URI must fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(
            error
                .message
                .contains("duplicate worksheet part URI xl/worksheets/sheet1.xml")
        );
    }

    #[test]
    fn workbook_state_validate_for_save_allows_unbound_chart_sheet_record_for_graph_preflight() {
        let mut state = sample_state();
        add_second_worksheet(&mut state);
        state.worksheets[1].kind = SheetKind::ChartSheet;
        state.worksheets[1].relationship_id = None;
        state.worksheets[1].part_uri = None;

        state
            .validate_for_save()
            .expect("XLSX graph layer validates the unbound chart-sheet record");
    }

    #[test]
    fn workbook_state_validate_for_save_rejects_unbound_ordinary_worksheet() {
        let mut state = sample_state();
        state.worksheets[0].relationship_id = None;
        state.worksheets[0].part_uri = None;

        let error = state
            .validate_for_save()
            .expect_err("ordinary worksheet requires a package binding");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("only an unbound chart-sheet record"));
    }

    #[test]
    fn workbook_state_validate_for_save_rejects_incomplete_worksheet_binding() {
        let mut state = sample_state();
        state.worksheets[0].part_uri = None;

        let error = state
            .validate_for_save()
            .expect_err("partial relationship binding must fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("relationship id and part URI"));
    }

    #[test]
    fn workbook_state_validate_for_save_rejects_missing_worksheet_data() {
        let mut state = sample_state();
        state.worksheet_data.remove(&SheetId(3));

        let error = state
            .validate_for_save()
            .expect_err("missing worksheet data must fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(
            error
                .message
                .contains("worksheet Sheet1 (3) has no worksheet data")
        );
    }

    #[test]
    fn workbook_state_validate_for_save_rejects_orphan_worksheet_data() {
        let mut state = sample_state();
        state
            .worksheet_data
            .insert(SheetId(404), WorksheetData::default());

        let error = state
            .validate_for_save()
            .expect_err("orphan worksheet data should fail validation");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("unknown worksheet 404"));
    }

    #[test]
    fn workbook_state_validate_for_save_rejects_dangling_defined_name_scope() {
        let mut state = sample_state();
        state
            .add_defined_name(
                NameScope::Worksheet(SheetId(99)),
                "LocalName",
                formula_source("Sheet1!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("public API currently permits a dangling scope");

        let error = state
            .validate_for_save()
            .expect_err("dangling defined-name scope must fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("defined name LocalName"));
        assert!(error.message.contains("unknown worksheet 99"));
    }

    #[test]
    fn workbook_state_validate_for_save_rejects_phantom_spill_owner() {
        let mut state = sample_state();
        state
            .worksheet_data
            .get_mut(&SheetId(3))
            .expect("worksheet data")
            .spill_owners
            .insert((2, 2), (9, 9));

        let error = state
            .validate_for_save()
            .expect_err("spill child with no owner range must fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("spill child R2C2"));
        assert!(error.message.contains("unknown anchor R9C9"));
    }

    #[test]
    fn workbook_state_validate_for_save_rejects_formula_spill_child() {
        let mut state = sample_state();
        seed_two_by_two_spill(&mut state);
        state
            .worksheet_data
            .get_mut(&SheetId(3))
            .expect("worksheet data")
            .cells
            .get_mut(&(3, 4))
            .expect("spill child")
            .formula = Some(formula_source("1+1"));

        let error = state
            .validate_for_save()
            .expect_err("spill child formula must fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("spill child R3C4"));
        assert!(error.message.contains("formula"));
    }

    #[test]
    fn workbook_state_tracks_chart_models_drawings_and_chart_sheet_bindings() {
        let mut state = sample_state();
        let workbook_id = state.model.id;
        let sheet_id = state.worksheets[0].id;
        let chart_id = ChartId(11);
        let drawing_id = DrawingId(12);

        state
            .charts_mut()
            .insert(chart_id, chart_model(chart_id, workbook_id, Vec::new()));
        state.drawings_mut().insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id,
                host_sheet_id: sheet_id,
                objects: vec![DrawingObjectModel::ChartFrame(ChartObjectModel {
                    id: ChartObjectId(21),
                    anchor_attrs: BTreeMap::new(),
                    position_attrs: BTreeMap::new(),
                    extents_attrs: BTreeMap::new(),
                    marker_attrs: Default::default(),
                    graphic_frame_attrs: BTreeMap::new(),
                    graphic_frame_transform_xml: None,
                    graphic_data_attrs: BTreeMap::new(),
                    graphic_data_child_xmls: Vec::new(),
                    chart_reference_attrs: BTreeMap::new(),
                    non_visual_frame_attrs: BTreeMap::new(),
                    graphic_attrs: BTreeMap::new(),
                    non_visual_id: Some(2),
                    non_visual_attrs: BTreeMap::new(),
                    non_visual_child_xml: None,
                    non_visual_frame_properties_xml: None,
                    client_data_attrs: BTreeMap::new(),
                    client_data_xml: None,
                    anchor_extension_xmls: Vec::new(),
                    workbook_id,
                    host_sheet_id: sheet_id,
                    chart_id,
                    name: "Chart 1".to_string(),
                    anchor: None,
                    placement: ObjectPlacement::Unknown,
                    z_order: Some(0),
                    raw_binding: Some("xl/drawings/drawing1.xml#rIdChart1".to_string()),
                    dirty: false,
                })],
                raw_part_uri: Some("xl/drawings/drawing1.xml".to_string()),
                dirty: false,
            },
        );
        state.chart_sheets_mut().insert(
            sheet_id,
            ChartSheetBinding {
                sheet_id,
                chart_id,
                drawing_id: Some(drawing_id),
                raw_part_uri: Some("xl/chartsheets/sheet1.xml".to_string()),
            },
        );

        assert_eq!(
            state.charts().get(&chart_id).expect("chart").chart_type,
            ChartType::Bar
        );
        assert_eq!(
            state
                .drawings()
                .get(&drawing_id)
                .expect("drawing")
                .objects
                .len(),
            1
        );
        assert_eq!(
            state.chart_sheet(sheet_id).expect("chart sheet").chart_id,
            chart_id
        );
    }

    #[test]
    fn workbook_state_defined_name_workbook_scope_lookup_is_case_insensitive() {
        let mut state = sample_state();
        state
            .add_defined_name(
                NameScope::Workbook,
                "Total",
                formula_source("Sheet1!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("add workbook name");

        let found = state
            .lookup_name(Some(SheetId(3)), "total")
            .expect("lookup workbook name");

        assert_eq!(found.display_name, "Total");
        assert_eq!(found.refers_to.text, "Sheet1!$A$1");
    }

    #[test]
    fn workbook_state_defined_name_sheet_scope_shadows_workbook_scope() {
        let mut state = sample_state();
        state
            .add_defined_name(
                NameScope::Workbook,
                "Total",
                formula_source("Sheet1!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("add workbook name");
        state
            .add_defined_name(
                NameScope::Worksheet(SheetId(3)),
                "total",
                formula_source("Sheet1!$B$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("add sheet name");

        let found = state
            .lookup_name(Some(SheetId(3)), "TOTAL")
            .expect("lookup shadowing name");

        assert_eq!(found.scope, NameScope::Worksheet(SheetId(3)));
        assert_eq!(found.refers_to.text, "Sheet1!$B$1");
    }

    #[test]
    fn workbook_state_defined_name_duplicate_same_scope_is_rejected() {
        let mut state = sample_state();
        state
            .add_defined_name(
                NameScope::Workbook,
                "Total",
                formula_source("Sheet1!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("add workbook name");

        let error = state
            .add_defined_name(
                NameScope::Workbook,
                "total",
                formula_source("Sheet1!$B$1"),
                NameValidationMode::StrictExcel,
            )
            .expect_err("duplicate should fail");

        assert_eq!(error.code, OmErrorCode::InvalidArgument);
    }

    #[test]
    fn workbook_state_defined_name_allows_same_name_on_different_sheets() {
        let mut state = sample_state();
        state
            .add_defined_name(
                NameScope::Worksheet(SheetId(3)),
                "Total",
                formula_source("Sheet1!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("add sheet name");
        state
            .add_defined_name(
                NameScope::Worksheet(SheetId(4)),
                "total",
                formula_source("Sheet2!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("add same name on another sheet");

        assert_eq!(state.defined_names().len(), 2);
    }

    #[test]
    fn workbook_state_defined_name_remove_clears_lookup() {
        let mut state = sample_state();
        state
            .add_defined_name(
                NameScope::Workbook,
                "Total",
                formula_source("Sheet1!$A$1"),
                NameValidationMode::StrictExcel,
            )
            .expect("add workbook name");

        let removed = state
            .remove_defined_name(NameScope::Workbook, "total")
            .expect("remove name");

        assert_eq!(removed.display_name, "Total");
        assert!(state.lookup_name(None, "Total").is_none());
    }

    #[test]
    fn get_range_values_reads_sparse_cells_as_empty() {
        let state = sample_state();
        let values = state
            .get_range_values(&RangeRef::single_rect(
                WorkbookId(7),
                SheetId(3),
                Rect {
                    row_first: 1,
                    row_last: 2,
                    col_first: 1,
                    col_last: 2,
                },
            ))
            .expect("range values");

        assert_eq!(values.rows, 2);
        assert_eq!(values.cols, 2);
        assert_eq!(values.get(0, 0), Some(&OmValue::Number(42.0)));
        assert_eq!(values.get(0, 1), Some(&OmValue::Empty));
        assert_eq!(values.get(1, 0), Some(&OmValue::Empty));
        assert_eq!(values.get(1, 1), Some(&OmValue::Text("hello".to_string())));
    }

    #[test]
    fn set_range_values_updates_cells_in_row_major_order() {
        let mut state = sample_state();
        state
            .set_range_values(
                &RangeRef::single_rect(
                    WorkbookId(7),
                    SheetId(3),
                    Rect {
                        row_first: 1,
                        row_last: 2,
                        col_first: 1,
                        col_last: 2,
                    },
                ),
                &OmArray::new(
                    2,
                    2,
                    vec![
                        OmValue::Number(1.0),
                        OmValue::Text("two".to_string()),
                        OmValue::Bool(true),
                        OmValue::Empty,
                    ],
                )
                .expect("array"),
            )
            .expect("set range values");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(worksheet.dirty);
        assert_eq!(
            worksheet.dirty_cells,
            BTreeSet::from([(1, 1), (1, 2), (2, 1), (2, 2)])
        );
        assert_eq!(
            worksheet.cells.get(&(1, 1)).expect("A1").value,
            CellValue::Number(1.0)
        );
        assert_eq!(
            worksheet.cells.get(&(1, 2)).expect("B1").value,
            CellValue::Text("two".to_string())
        );
        assert_eq!(
            worksheet.cells.get(&(2, 1)).expect("A2").value,
            CellValue::Bool(true)
        );
        assert_eq!(
            worksheet.cells.get(&(2, 2)).expect("B2").style_id,
            Some(StyleId(9))
        );
        assert_eq!(
            worksheet.cells.get(&(2, 2)).expect("B2").value,
            CellValue::Blank
        );
    }

    #[test]
    fn set_range_values_rejects_mismatched_dimensions() {
        let mut state = sample_state();
        let result = state.set_range_values(
            &RangeRef::single_rect(
                WorkbookId(7),
                SheetId(3),
                Rect {
                    row_first: 1,
                    row_last: 2,
                    col_first: 1,
                    col_last: 2,
                },
            ),
            &OmArray::scalar(OmValue::Number(1.0)),
        );

        assert!(result.is_err());
    }

    #[test]
    fn set_range_values_scalar_broadcasts_to_multi_area_ranges() {
        let mut state = sample_state();
        state
            .set_range_values(
                &RangeRef {
                    workbook_id: WorkbookId(7),
                    scope: SheetScope::Single(SheetId(3)),
                    areas: vec![Rect::single_cell(1, 1), Rect::single_cell(2, 2)],
                },
                &OmArray::scalar(OmValue::Number(9.0)),
            )
            .expect("broadcast scalar");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(worksheet.dirty);
        assert_eq!(worksheet.dirty_cells, BTreeSet::from([(1, 1), (2, 2)]));
        assert_eq!(
            worksheet.cells.get(&(1, 1)).expect("A1").value,
            CellValue::Number(9.0)
        );
        assert_eq!(
            worksheet.cells.get(&(2, 2)).expect("B2").value,
            CellValue::Number(9.0)
        );
        assert_eq!(
            worksheet.cells.get(&(2, 2)).expect("B2").style_id,
            Some(StyleId(9))
        );
    }

    #[test]
    fn set_range_values_rejects_array_assignment_to_multi_area_ranges() {
        let mut state = sample_state();
        let error = state
            .set_range_values(
                &RangeRef {
                    workbook_id: WorkbookId(7),
                    scope: SheetScope::Single(SheetId(3)),
                    areas: vec![Rect::single_cell(1, 1), Rect::single_cell(2, 2)],
                },
                &OmArray::new(1, 2, vec![OmValue::Number(1.0), OmValue::Number(2.0)])
                    .expect("array"),
            )
            .expect_err("multi-area array assignment should fail");

        assert_eq!(error.code, OmErrorCode::Unsupported);
        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(!worksheet.dirty);
        assert!(worksheet.dirty_cells.is_empty());
    }

    #[test]
    fn blank_write_removes_unstyled_cells() {
        let mut state = sample_state();
        state
            .set_range_values(
                &RangeRef::single_cell(WorkbookId(7), SheetId(3), 1, 1),
                &OmArray::scalar(OmValue::Empty),
            )
            .expect("clear cell");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(!worksheet.cells.contains_key(&(1, 1)));
        assert!(worksheet.dirty_cells.contains(&(1, 1)));
    }

    #[test]
    fn get_range_values_rejects_foreign_workbook_id() {
        let state = sample_state();
        let result =
            state.get_range_values(&RangeRef::single_cell(WorkbookId(999), SheetId(3), 1, 1));

        assert!(result.is_err());
    }

    #[test]
    fn get_range_formulas_returns_formula_text_and_constant_values() {
        let mut state = sample_state();
        let worksheet = state
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data");
        worksheet.cells.get_mut(&(1, 1)).expect("A1").formula = Some(FormulaSource {
            text: "SUM(B1:B2)".to_string(),
            is_r1c1: false,
        });

        let formulas = state
            .get_range_formulas(&RangeRef::single_rect(
                WorkbookId(7),
                SheetId(3),
                Rect {
                    row_first: 1,
                    row_last: 2,
                    col_first: 1,
                    col_last: 2,
                },
            ))
            .expect("range formulas");

        assert_eq!(
            formulas,
            OmArray::new(
                2,
                2,
                vec![
                    OmValue::Text("=SUM(B1:B2)".to_string()),
                    OmValue::Empty,
                    OmValue::Empty,
                    OmValue::Text("hello".to_string()),
                ],
            )
            .expect("array"),
        );
    }

    #[test]
    fn set_range_values_rejects_foreign_workbook_id() {
        let mut state = sample_state();
        let result = state.set_range_values(
            &RangeRef::single_cell(WorkbookId(999), SheetId(3), 1, 1),
            &OmArray::scalar(OmValue::Number(1.0)),
        );

        assert!(result.is_err());
    }

    #[test]
    fn get_range_values_rejects_zero_based_coordinates() {
        let state = sample_state();
        let result = state.get_range_values(&RangeRef::single_rect(
            WorkbookId(7),
            SheetId(3),
            Rect {
                row_first: 0,
                row_last: 1,
                col_first: 1,
                col_last: 1,
            },
        ));

        assert!(result.is_err());
    }

    #[test]
    fn set_range_values_rejects_unordered_rect_bounds() {
        let mut state = sample_state();
        let result = state.set_range_values(
            &RangeRef::single_rect(
                WorkbookId(7),
                SheetId(3),
                Rect {
                    row_first: 3,
                    row_last: 2,
                    col_first: 1,
                    col_last: 1,
                },
            ),
            &OmArray::scalar(OmValue::Number(1.0)),
        );

        assert!(result.is_err());
    }

    #[test]
    fn set_range_values_clears_formula_even_when_value_is_unchanged() {
        let mut state = sample_state();
        let worksheet = state
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data");
        worksheet.cells.get_mut(&(1, 1)).expect("A1").formula = Some(FormulaSource {
            text: "SUM(B1:B2)".to_string(),
            is_r1c1: false,
        });

        state
            .set_range_values(
                &RangeRef::single_cell(WorkbookId(7), SheetId(3), 1, 1),
                &OmArray::scalar(OmValue::Number(42.0)),
            )
            .expect("overwrite formula");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        let cell = worksheet.cells.get(&(1, 1)).expect("A1");
        assert_eq!(cell.value, CellValue::Number(42.0));
        assert!(cell.formula.is_none());
        assert!(worksheet.dirty);
        assert!(worksheet.dirty_cells.contains(&(1, 1)));
    }

    #[test]
    fn set_range_values_keeps_clean_when_existing_cell_state_is_unchanged() {
        let mut state = sample_state();
        state
            .set_range_values(
                &RangeRef::single_cell(WorkbookId(7), SheetId(3), 2, 2),
                &OmArray::scalar(OmValue::Text("hello".to_string())),
            )
            .expect("rewrite same value");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(!worksheet.dirty);
        assert!(worksheet.dirty_cells.is_empty());
    }

    #[test]
    fn insert_cell_marks_worksheet_dirty_and_tracks_coordinate() {
        let mut state = sample_state();
        state
            .insert_cell(
                SheetId(3),
                3,
                1,
                CellData {
                    value: CellValue::Number(9.0),
                    formula: None,
                    style_id: None,
                },
            )
            .expect("valid test cell coordinate");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(worksheet.dirty);
        assert!(worksheet.dirty_cells.contains(&(3, 1)));
        assert_eq!(
            worksheet.cells.get(&(3, 1)).expect("A3").value,
            CellValue::Number(9.0)
        );
    }

    #[test]
    fn insert_cell_rejects_unknown_sheet_without_creating_worksheet_data() {
        let mut state = sample_state();
        let original = state.clone();

        let error = state
            .insert_cell(
                SheetId(404),
                1,
                1,
                CellData {
                    value: CellValue::Number(9.0),
                    formula: None,
                    style_id: None,
                },
            )
            .expect_err("unknown sheet insertion should fail");

        assert_eq!(error.code, OmErrorCode::NotFound);
        assert_eq!(state, original);
    }

    #[test]
    fn set_worksheet_source_xml_rejects_unknown_sheet_without_creating_worksheet_data() {
        let mut state = sample_state();
        let original = state.clone();

        let error = state
            .set_worksheet_source_xml(SheetId(404), b"<worksheet/>".to_vec())
            .expect_err("unknown sheet source update should fail");

        assert_eq!(error.code, OmErrorCode::NotFound);
        assert_eq!(state, original);
    }

    #[test]
    fn set_worksheet_source_xml_updates_existing_worksheet_data() {
        let mut state = sample_state();
        let source_xml = b"<worksheet><sheetData/></worksheet>".to_vec();

        state
            .set_worksheet_source_xml(SheetId(3), source_xml.clone())
            .expect("existing worksheet source update should succeed");

        assert_eq!(
            state
                .worksheet_data_for_sheet(SheetId(3))
                .expect("worksheet data")
                .source_xml,
            source_xml
        );
    }

    #[test]
    fn range_operations_reject_coordinates_outside_the_excel_grid() {
        let state = sample_state();

        for range in [
            RangeRef::single_cell(WorkbookId(7), SheetId(3), 1_048_577, 1),
            RangeRef::single_cell(WorkbookId(7), SheetId(3), 1, 16_385),
        ] {
            let error = state
                .get_range_values(&range)
                .expect_err("out-of-grid range should be rejected before iteration");

            assert_eq!(error.code, OmErrorCode::InvalidArgument);
        }
    }

    #[test]
    fn insert_cell_rejects_coordinates_outside_the_excel_grid_without_mutation() {
        let mut state = sample_state();

        let error = state
            .insert_cell(
                SheetId(3),
                1_048_577,
                1,
                CellData {
                    value: CellValue::Number(9.0),
                    formula: None,
                    style_id: None,
                },
            )
            .expect_err("out-of-grid cell insertion should fail");

        assert_eq!(error.code, OmErrorCode::InvalidArgument);
        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(!worksheet.cells.contains_key(&(1_048_577, 1)));
        assert!(!worksheet.dirty_cells.contains(&(1_048_577, 1)));
    }

    #[test]
    fn clear_owned_spill_rejects_an_out_of_grid_anchor_without_mutation() {
        let mut state = sample_state();
        let worksheet = state
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data");

        let error = worksheet
            .clear_owned_spill((1, 16_385))
            .expect_err("out-of-grid spill anchor should fail");

        assert_eq!(error.code, OmErrorCode::InvalidArgument);
        assert!(!worksheet.dirty);
        assert!(worksheet.dirty_cells.is_empty());
    }

    #[test]
    fn assign_workbook_id_updates_model_and_worksheet_ownership() {
        let mut state = sample_state();
        let sheet_id = state.worksheets[0].id;
        let chart_id = ChartId(11);
        let drawing_id = DrawingId(12);
        state.charts.insert(
            chart_id,
            ChartModel {
                id: chart_id,
                workbook_id: state.model.id,
                chart_type: ChartType::Bar,
                style: None,
                series: vec![super::SeriesModel {
                    name: Some(ChartSourceExpr {
                        raw: formula_source("Sheet1!$A$1"),
                        resolved: Some(ReferenceTarget::Range(
                            RangeSet::single_rect(
                                state.model.id,
                                sheet_id,
                                Rect::single_cell(1, 1),
                            )
                            .expect("name range set"),
                        )),
                        full_reference: None,
                        cache: None,
                        dirty: false,
                    }),
                    x_values: Some(ChartSourceExpr {
                        raw: formula_source("Sheet1!$A$2:$A$3"),
                        resolved: Some(ReferenceTarget::Range(
                            RangeSet::single_rect(
                                state.model.id,
                                sheet_id,
                                Rect {
                                    row_first: 2,
                                    row_last: 3,
                                    col_first: 1,
                                    col_last: 1,
                                },
                            )
                            .expect("x values range set"),
                        )),
                        full_reference: Some(super::ChartSourceReference {
                            raw: formula_source("Sheet1!$A$1:$A$3"),
                            resolved: Some(ReferenceTarget::Range(
                                RangeSet::single_rect(
                                    state.model.id,
                                    sheet_id,
                                    Rect {
                                        row_first: 1,
                                        row_last: 3,
                                        col_first: 1,
                                        col_last: 1,
                                    },
                                )
                                .expect("full x values range set"),
                            )),
                        }),
                        cache: None,
                        dirty: false,
                    }),
                    values: Some(ChartSourceExpr {
                        raw: formula_source("Sheet1!$A$1"),
                        resolved: Some(ReferenceTarget::Range(
                            RangeSet::single_rect(
                                state.model.id,
                                sheet_id,
                                Rect::single_cell(1, 1),
                            )
                            .expect("range set"),
                        )),
                        full_reference: None,
                        cache: None,
                        dirty: false,
                    }),
                    bubble_size: Some(ChartSourceExpr {
                        raw: formula_source("Sheet1!$B$2:$B$3"),
                        resolved: Some(ReferenceTarget::Range(
                            RangeSet::single_rect(
                                state.model.id,
                                sheet_id,
                                Rect {
                                    row_first: 2,
                                    row_last: 3,
                                    col_first: 2,
                                    col_last: 2,
                                },
                            )
                            .expect("bubble size range set"),
                        )),
                        full_reference: None,
                        cache: None,
                        dirty: false,
                    }),
                    bar_shape: None,
                    smooth: None,
                    marker_style: None,
                    marker_size: None,
                    invert_if_negative: None,
                    points: BTreeMap::new(),
                    data_labels: None,
                    point_data_labels: BTreeMap::new(),
                    raw_index: None,
                    order: Some(0),
                    axis_group: super::ChartAxisGroup::Primary,
                    is_filtered: false,
                    filter_dirty: false,
                }],
                title: None,
                legend: None,
                axes: Vec::new(),
                groups: Vec::new(),
                vary_by_categories: None,
                gap_width: None,
                gap_depth: None,
                overlap: None,
                bar_shape: None,
                has_series_lines: None,
                has_drop_lines: None,
                has_hi_lo_lines: None,
                has_up_down_bars: None,
                first_slice_angle: None,
                explosion: None,
                bubble_scale: None,
                show_negative_bubbles: None,
                has_3d_shading: None,
                doughnut_hole_size: None,
                second_plot_size: None,
                size_represents: None,
                split_type: None,
                split_value: None,
                data_labels: None,
                data_table: None,
                data_table_dirty: false,
                plot_area_layout: None,
                plot_area_layout_dirty: false,
                show_data_labels_over_maximum: None,
                display_blanks_as: None,
                plot_visible_only: None,
                view_3d: None,
                view_3d_dirty: false,
                rounded_corners: None,
                protection: None,
                protection_dirty: false,
                raw_part_uri: Some("xl/charts/chart1.xml".to_string()),
                series_topology_dirty: false,
                content_dirty: false,
                dirty: false,
            },
        );
        state.drawings.insert(
            drawing_id,
            DrawingModel {
                id: drawing_id,
                workbook_id: state.model.id,
                host_sheet_id: sheet_id,
                objects: vec![DrawingObjectModel::ChartFrame(ChartObjectModel {
                    id: ChartObjectId(21),
                    anchor_attrs: BTreeMap::new(),
                    position_attrs: BTreeMap::new(),
                    extents_attrs: BTreeMap::new(),
                    marker_attrs: Default::default(),
                    graphic_frame_attrs: BTreeMap::new(),
                    graphic_frame_transform_xml: None,
                    graphic_data_attrs: BTreeMap::new(),
                    graphic_data_child_xmls: Vec::new(),
                    chart_reference_attrs: BTreeMap::new(),
                    non_visual_frame_attrs: BTreeMap::new(),
                    graphic_attrs: BTreeMap::new(),
                    non_visual_id: Some(2),
                    non_visual_attrs: BTreeMap::new(),
                    non_visual_child_xml: None,
                    non_visual_frame_properties_xml: None,
                    client_data_attrs: BTreeMap::new(),
                    client_data_xml: None,
                    anchor_extension_xmls: Vec::new(),
                    workbook_id: state.model.id,
                    host_sheet_id: sheet_id,
                    chart_id,
                    name: "Chart 1".to_string(),
                    anchor: None,
                    placement: ObjectPlacement::Unknown,
                    z_order: Some(0),
                    raw_binding: None,
                    dirty: false,
                })],
                raw_part_uri: Some("xl/drawings/drawing1.xml".to_string()),
                dirty: false,
            },
        );

        state
            .assign_workbook_id(WorkbookId(99))
            .expect("valid workbook references should rebind");

        assert_eq!(state.model.id, WorkbookId(99));
        assert_eq!(state.worksheets[0].workbook_id, WorkbookId(99));
        let chart = state.charts.get(&chart_id).expect("chart");
        assert_eq!(chart.workbook_id, WorkbookId(99));
        for source in [
            chart.series[0].name.as_ref().expect("name"),
            chart.series[0].x_values.as_ref().expect("x values"),
            chart.series[0].values.as_ref().expect("values"),
            chart.series[0].bubble_size.as_ref().expect("bubble size"),
        ] {
            let Some(ReferenceTarget::Range(range)) = source.resolved.as_ref() else {
                panic!("expected range source");
            };
            assert_eq!(range.workbook_id(), WorkbookId(99));
        }
        let Some(ReferenceTarget::Range(full_range)) = chart.series[0]
            .x_values
            .as_ref()
            .and_then(|source| source.full_reference.as_ref())
            .and_then(|reference| reference.resolved.as_ref())
        else {
            panic!("expected full x values range source");
        };
        assert_eq!(full_range.workbook_id(), WorkbookId(99));
        let drawing = state.drawings.get(&drawing_id).expect("drawing");
        assert_eq!(drawing.workbook_id, WorkbookId(99));
        let DrawingObjectModel::ChartFrame(chart_object) = &drawing.objects[0] else {
            panic!("expected chart object");
        };
        assert_eq!(chart_object.workbook_id, WorkbookId(99));
    }

    #[test]
    fn assign_workbook_id_does_not_partially_update_malformed_chart_references() {
        let mut state = sample_state();
        let chart_id = ChartId(11);
        let series = series_model_with_values(ChartSourceExpr {
            raw: formula_source("Sheet1!$A$1"),
            resolved: Some(ReferenceTarget::Range(malformed_empty_range_set())),
            full_reference: None,
            cache: None,
            dirty: false,
        });
        state.charts.insert(
            chart_id,
            chart_model(chart_id, state.model.id, vec![series]),
        );
        let original = state.clone();

        let error = state
            .assign_workbook_id(WorkbookId(99))
            .expect_err("malformed chart source should reject workbook-id reassignment");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("chart 11 series 1 values range"));
        assert_eq!(state, original);
    }

    #[test]
    fn assign_workbook_id_rejects_malformed_full_reference_atomically() {
        let mut state = sample_state();
        let chart_id = ChartId(11);
        let series = series_model_with_values(ChartSourceExpr {
            raw: formula_source("Sheet1!$A$1"),
            resolved: None,
            full_reference: Some(super::ChartSourceReference {
                raw: formula_source("Sheet1!$A$1"),
                resolved: Some(ReferenceTarget::Range(malformed_empty_range_set())),
            }),
            cache: None,
            dirty: false,
        });
        state.charts.insert(
            chart_id,
            chart_model(chart_id, state.model.id, vec![series]),
        );
        let original = state.clone();

        let error = state
            .assign_workbook_id(WorkbookId(99))
            .expect_err("malformed full reference should reject workbook-id reassignment");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(
            error
                .message
                .contains("chart 11 series 1 values full range")
        );
        assert_eq!(state, original);
    }

    #[test]
    fn worksheet_data_lookup_returns_not_found_for_unknown_sheet() {
        let state = sample_state();

        let error = state
            .worksheet_data_for_sheet(SheetId(999))
            .expect_err("unknown worksheet data should fail");

        assert_eq!(error.code, OmErrorCode::NotFound);
        assert!(error.message.contains("999"));
    }

    #[test]
    fn worksheet_data_mutation_rejects_preexisting_orphan_entry() {
        let mut state = sample_state();
        state
            .worksheet_data
            .insert(SheetId(404), WorksheetData::default());

        let error = state
            .worksheet_data_for_sheet_mut(SheetId(404))
            .expect_err("orphan worksheet data must not be mutable through the command accessor");

        assert_eq!(error.code, OmErrorCode::NotFound);
        assert!(error.message.contains("unknown worksheet 404"));
    }

    #[test]
    fn worksheet_data_replacement_rejects_unknown_owner_atomically() {
        let mut state = sample_state();
        let original = state.clone();

        let error = state
            .replace_worksheet_data_for_sheet(SheetId(404), WorksheetData::default())
            .expect_err("unknown worksheet data owner should be rejected");

        assert_eq!(error.code, OmErrorCode::NotFound);
        assert!(error.message.contains("unknown worksheet 404"));
        assert_eq!(state, original);
    }

    #[test]
    fn worksheet_and_data_insert_remove_commands_keep_owner_keys_together() {
        let mut state = sample_state();
        let original = state.clone();
        let worksheet = WorksheetModel {
            id: SheetId(4),
            workbook_id: state.model.id,
            name: "Sheet2".to_string(),
            kind: office_common::SheetKind::Worksheet,
            visibility: office_common::SheetVisibility::Visible,
            relationship_id: Some("rId2".to_string()),
            part_uri: Some("xl/worksheets/sheet2.xml".to_string()),
        };
        let worksheet_data = WorksheetData {
            source_xml: b"<worksheet/>".to_vec(),
            ..WorksheetData::default()
        };

        state
            .insert_worksheet_with_data(1, worksheet.clone(), worksheet_data.clone())
            .expect("valid worksheet and data should be inserted together");
        assert_eq!(state.worksheets[1], worksheet);
        assert_eq!(
            state
                .worksheet_data_for_sheet(SheetId(4))
                .expect("inserted worksheet data"),
            &worksheet_data
        );

        let (index, removed_worksheet, removed_data) = state
            .remove_worksheet_with_data(SheetId(4))
            .expect("worksheet and data should be removed together");
        assert_eq!(index, 1);
        assert_eq!(removed_worksheet, worksheet);
        assert_eq!(removed_data, worksheet_data);
        assert_eq!(state, original);
    }

    #[test]
    fn worksheet_and_data_insert_rolls_back_invalid_owner_metadata() {
        let mut state = sample_state();
        let original = state.clone();
        let worksheet = WorksheetModel {
            id: SheetId(4),
            workbook_id: state.model.id,
            name: state.worksheets[0].name.clone(),
            kind: office_common::SheetKind::Worksheet,
            visibility: office_common::SheetVisibility::Visible,
            relationship_id: Some("rId2".to_string()),
            part_uri: Some("xl/worksheets/sheet2.xml".to_string()),
        };

        let error = state
            .insert_worksheet_with_data(1, worksheet, WorksheetData::default())
            .expect_err("duplicate worksheet name should reject the paired insertion");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("duplicate worksheet name"));
        assert_eq!(state, original);
    }

    #[test]
    fn worksheet_metadata_commands_validate_before_mutation() {
        let mut state = sample_state();
        add_second_worksheet(&mut state);

        assert!(
            state
                .rename_worksheet(SheetId(4), "Renamed")
                .expect("rename second worksheet")
        );
        assert!(
            !state
                .rename_worksheet(SheetId(4), "Renamed")
                .expect("same worksheet name should be a no-op")
        );
        assert_eq!(state.worksheets[1].name, "Renamed");

        let renamed = state.clone();
        let error = state
            .rename_worksheet(SheetId(4), "sHeEt1")
            .expect_err("case-insensitive duplicate name should be rejected");
        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert_eq!(state, renamed);

        let error = state
            .rename_worksheet(SheetId(404), "Missing")
            .expect_err("unknown worksheet rename should be rejected");
        assert_eq!(error.code, OmErrorCode::NotFound);
        assert_eq!(state, renamed);

        assert!(
            state
                .set_worksheet_visibility(SheetId(4), SheetVisibility::Hidden)
                .expect("hide second worksheet")
        );
        assert!(
            !state
                .set_worksheet_visibility(SheetId(4), SheetVisibility::Hidden)
                .expect("same visibility should be a no-op")
        );

        let hidden = state.clone();
        let error = state
            .set_worksheet_visibility(SheetId(404), SheetVisibility::VeryHidden)
            .expect_err("unknown worksheet visibility should be rejected");
        assert_eq!(error.code, OmErrorCode::NotFound);
        assert_eq!(state, hidden);
    }

    #[test]
    fn chart_sheet_package_binding_command_updates_owner_atomically() {
        let mut state = sample_state();
        state
            .insert_worksheet_with_data(
                1,
                WorksheetModel {
                    id: SheetId(4),
                    workbook_id: state.model.id,
                    name: "Chart1".to_string(),
                    kind: SheetKind::ChartSheet,
                    visibility: SheetVisibility::Visible,
                    relationship_id: None,
                    part_uri: None,
                },
                WorksheetData::default(),
            )
            .expect("insert unbound chart sheet");
        state.chart_sheets.insert(
            SheetId(4),
            ChartSheetBinding {
                sheet_id: SheetId(4),
                chart_id: ChartId(1),
                drawing_id: Some(DrawingId(1)),
                raw_part_uri: None,
            },
        );

        let original = state.clone();
        let error = state
            .bind_chart_sheet_package(SheetId(4), "rId1", "xl/chartsheets/sheet1.xml")
            .expect_err("duplicate relationship id should be rejected");
        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert_eq!(state, original);

        let mut partial = original.clone();
        partial.worksheets[1].relationship_id = Some("rId2".to_string());
        let partial_original = partial.clone();
        let error = partial
            .bind_chart_sheet_package(SheetId(4), "rId2", "xl/chartsheets/sheet1.xml")
            .expect_err("partial chart sheet package identity should be rejected");
        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert_eq!(partial, partial_original);

        assert!(
            state
                .bind_chart_sheet_package(SheetId(4), "rId2", "xl/chartsheets/sheet1.xml",)
                .expect("bind chart sheet package identity")
        );
        assert!(
            !state
                .bind_chart_sheet_package(SheetId(4), "rId2", "xl/chartsheets/sheet1.xml",)
                .expect("same chart sheet package identity should be a no-op")
        );
        assert_eq!(state.worksheets[1].relationship_id.as_deref(), Some("rId2"));
        assert_eq!(
            state.worksheets[1].part_uri.as_deref(),
            Some("xl/chartsheets/sheet1.xml")
        );
        assert_eq!(
            state.chart_sheets[&SheetId(4)].raw_part_uri.as_deref(),
            Some("xl/chartsheets/sheet1.xml")
        );

        let bound = state.clone();
        let error = state
            .bind_chart_sheet_package(SheetId(4), "rId3", "xl/chartsheets/sheet2.xml")
            .expect_err("a bound chart sheet should not be retargeted");
        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert_eq!(state, bound);
    }

    #[test]
    fn worksheet_reorder_command_requires_exact_identity_permutation() {
        let mut state = sample_state();
        add_second_worksheet(&mut state);
        let original_data = state.worksheet_data.clone();

        assert!(
            state
                .reorder_worksheets(&[SheetId(4), SheetId(3)])
                .expect("reverse worksheet order")
        );
        assert!(
            !state
                .reorder_worksheets(&[SheetId(4), SheetId(3)])
                .expect("same worksheet order should be a no-op")
        );
        assert_eq!(
            state
                .worksheets
                .iter()
                .map(|worksheet| worksheet.id)
                .collect::<Vec<_>>(),
            vec![SheetId(4), SheetId(3)]
        );
        assert_eq!(state.worksheet_data, original_data);

        let reordered = state.clone();
        for invalid_order in [
            vec![SheetId(4)],
            vec![SheetId(4), SheetId(4)],
            vec![SheetId(4), SheetId(404)],
        ] {
            let error = state
                .reorder_worksheets(invalid_order.as_slice())
                .expect_err("invalid worksheet order should be rejected");
            assert_eq!(error.code, OmErrorCode::InvalidState);
            assert_eq!(state, reordered);
        }
    }

    #[test]
    fn workbook_state_constructor_rejects_orphan_worksheet_data() {
        let mut state = sample_state();
        state
            .worksheet_data
            .insert(SheetId(404), WorksheetData::default());

        let error = WorkbookState::try_new(WorkbookStateParts {
            model: state.model,
            worksheets: state.worksheets,
            worksheet_data: state.worksheet_data,
            defined_names: state.defined_names,
            charts: state.charts,
            drawings: state.drawings,
            chart_sheets: state.chart_sheets,
            opaque_parts: state.opaque_parts,
        })
        .expect_err("validated construction should reject orphan worksheet data");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(
            error
                .message
                .contains("worksheet data references unknown worksheet 404")
        );
    }

    #[test]
    fn workbook_model_metadata_commands_only_change_requested_fields() {
        let mut state = sample_state();
        let workbook_id = state.model.id;

        assert!(state.set_display_name("Renamed"));
        assert!(!state.set_display_name("Renamed"));
        assert!(state.set_date1904(true));
        assert!(!state.set_date1904(true));
        assert!(state.set_is_addin(true));
        assert!(!state.set_is_addin(true));
        assert!(state.set_format(FileFormat::Xltx));
        assert!(!state.set_format(FileFormat::Xltx));

        assert_eq!(state.model().id, workbook_id);
        assert_eq!(state.model().display_name, "Renamed");
        assert!(state.model().date1904);
        assert!(state.model().is_addin);
        assert_eq!(state.model().format, FileFormat::Xltx);
    }

    #[test]
    fn workbook_state_parts_round_trip_preserves_private_collections() {
        let state = sample_state();

        assert_eq!(state.worksheets().len(), 1);
        assert_eq!(state.worksheets()[0].id, SheetId(3));
        let rebuilt = WorkbookState::try_new(state.clone().into_parts())
            .expect("validated state parts should reconstruct the same workbook");

        assert_eq!(rebuilt, state);
    }

    #[test]
    fn get_range_values_rejects_unknown_sheet_id_with_not_found() {
        let state = sample_state();
        let error = state
            .get_range_values(&RangeRef::single_cell(WorkbookId(7), SheetId(999), 1, 1))
            .expect_err("unknown sheet should fail");

        assert_eq!(error.code, OmErrorCode::NotFound);
        assert!(error.message.contains("999"));
    }

    #[test]
    fn set_range_values_rejects_3d_ranges_with_unsupported_code() {
        let mut state = sample_state();
        let error = state
            .set_range_values(
                &RangeRef {
                    workbook_id: WorkbookId(7),
                    scope: SheetScope::Multi3D {
                        start: SheetId(3),
                        end: SheetId(4),
                    },
                    areas: vec![Rect::single_cell(1, 1)],
                },
                &OmArray::scalar(OmValue::Number(1.0)),
            )
            .expect_err("3D ranges should fail");

        assert_eq!(error.code, OmErrorCode::Unsupported);
    }

    #[test]
    fn set_range_values_rejects_object_values_with_type_mismatch_and_keeps_clean_state() {
        let mut state = sample_state();
        let error = state
            .set_range_values(
                &RangeRef::single_cell(WorkbookId(7), SheetId(3), 1, 1),
                &OmArray::scalar(OmValue::Object(ObjectHandle(5))),
            )
            .expect_err("object values should fail");

        assert_eq!(error.code, OmErrorCode::TypeMismatch);
        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(!worksheet.dirty);
        assert!(worksheet.dirty_cells.is_empty());
        assert_eq!(
            worksheet.cells.get(&(1, 1)).expect("A1").value,
            CellValue::Number(42.0)
        );
    }

    #[test]
    fn set_range_formulas_updates_formula_cells_and_plain_text_constants() {
        let mut state = sample_state();
        state
            .set_range_formulas(
                &RangeRef::single_rect(
                    WorkbookId(7),
                    SheetId(3),
                    Rect {
                        row_first: 1,
                        row_last: 1,
                        col_first: 1,
                        col_last: 2,
                    },
                ),
                &OmArray::new(
                    1,
                    2,
                    vec![
                        OmValue::Text("=SUM(B1:B2)".to_string()),
                        OmValue::Text("literal".to_string()),
                    ],
                )
                .expect("array"),
            )
            .expect("set formulas");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(worksheet.dirty);
        assert_eq!(worksheet.dirty_cells, BTreeSet::from([(1, 1), (1, 2)]));
        let a1 = worksheet.cells.get(&(1, 1)).expect("A1");
        assert_eq!(a1.value, CellValue::Blank);
        assert_eq!(
            a1.formula,
            Some(FormulaSource {
                text: "SUM(B1:B2)".to_string(),
                is_r1c1: false,
            })
        );
        let b1 = worksheet.cells.get(&(1, 2)).expect("B1");
        assert_eq!(b1.value, CellValue::Text("literal".to_string()));
        assert!(b1.formula.is_none());
    }

    #[test]
    fn set_range_formulas_scalar_broadcasts_to_multi_area_ranges() {
        let mut state = sample_state();
        state
            .set_range_formulas(
                &RangeRef {
                    workbook_id: WorkbookId(7),
                    scope: SheetScope::Single(SheetId(3)),
                    areas: vec![Rect::single_cell(1, 1), Rect::single_cell(3, 3)],
                },
                &OmArray::scalar(OmValue::Text("=SUM(B1:B2)".to_string())),
            )
            .expect("broadcast formula");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(worksheet.dirty);
        assert_eq!(worksheet.dirty_cells, BTreeSet::from([(1, 1), (3, 3)]));
        for key in [(1, 1), (3, 3)] {
            let cell = worksheet.cells.get(&key).expect("formula cell");
            assert_eq!(cell.value, CellValue::Blank);
            assert_eq!(
                cell.formula,
                Some(FormulaSource {
                    text: "SUM(B1:B2)".to_string(),
                    is_r1c1: false,
                })
            );
        }
    }

    #[test]
    fn set_range_values_rejects_spill_child_edits_atomically() {
        let mut state = sample_state();
        seed_two_by_two_spill(&mut state);
        let before = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data")
            .clone();

        let error = state
            .set_range_values(
                &RangeRef {
                    workbook_id: WorkbookId(7),
                    scope: SheetScope::Single(SheetId(3)),
                    areas: vec![Rect::single_cell(5, 1), Rect::single_cell(4, 4)],
                },
                &OmArray::scalar(OmValue::Number(99.0)),
            )
            .expect_err("spill child value edit should fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("R4C4"));
        assert!(error.message.contains("R3C3"));
        assert_eq!(
            state
                .worksheet_data_for_sheet(SheetId(3))
                .expect("worksheet data"),
            &before
        );
    }

    #[test]
    fn set_range_formulas_rejects_unchanged_spill_child_edits_atomically() {
        let mut state = sample_state();
        seed_two_by_two_spill(&mut state);
        let before = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data")
            .clone();

        let error = state
            .set_range_formulas(
                &RangeRef {
                    workbook_id: WorkbookId(7),
                    scope: SheetScope::Single(SheetId(3)),
                    areas: vec![Rect::single_cell(5, 2), Rect::single_cell(4, 4)],
                },
                &OmArray::scalar(OmValue::Number(4.0)),
            )
            .expect_err("spill child formula edit should fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert_eq!(
            state
                .worksheet_data_for_sheet(SheetId(3))
                .expect("worksheet data"),
            &before
        );
    }

    #[test]
    fn clear_range_contents_rejects_spill_child_edits_atomically() {
        let mut state = sample_state();
        seed_two_by_two_spill(&mut state);
        let before = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data")
            .clone();

        let error = state
            .clear_range_contents(&RangeRef {
                workbook_id: WorkbookId(7),
                scope: SheetScope::Single(SheetId(3)),
                areas: vec![Rect::single_cell(1, 1), Rect::single_cell(4, 4)],
            })
            .expect_err("spill child clear should fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert_eq!(
            state
                .worksheet_data_for_sheet(SheetId(3))
                .expect("worksheet data"),
            &before
        );
    }

    #[test]
    fn clear_range_rejects_spill_child_edits_atomically() {
        let mut state = sample_state();
        seed_two_by_two_spill(&mut state);
        let before = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data")
            .clone();

        let error = state
            .clear_range_with_change(&RangeRef {
                workbook_id: WorkbookId(7),
                scope: SheetScope::Single(SheetId(3)),
                areas: vec![Rect::single_cell(1, 1), Rect::single_cell(4, 4)],
            })
            .expect_err("spill child clear should fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("R4C4"));
        assert!(error.message.contains("R3C3"));
        assert_eq!(
            state
                .worksheet_data_for_sheet(SheetId(3))
                .expect("worksheet data"),
            &before
        );
    }

    #[test]
    fn replace_cells_rejects_spill_child_updates_atomically() {
        let mut state = sample_state();
        seed_two_by_two_spill(&mut state);
        let before = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data")
            .clone();

        let error = state
            .replace_cells_with_change(
                SheetId(3),
                BTreeMap::from([
                    (
                        (1, 1),
                        Some(CellData {
                            value: CellValue::Text("changed".to_string()),
                            formula: None,
                            style_id: None,
                        }),
                    ),
                    ((4, 4), None),
                ]),
            )
            .expect_err("spill child replacement should fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("R4C4"));
        assert!(error.message.contains("R3C3"));
        assert_eq!(
            state
                .worksheet_data_for_sheet(SheetId(3))
                .expect("worksheet data"),
            &before
        );
    }

    #[test]
    fn replace_cells_clears_owned_spill_and_preserves_dynamic_formula_kind() {
        let mut state = sample_state();
        seed_two_by_two_spill(&mut state);
        let replacement = CellData {
            value: CellValue::Blank,
            formula: Some(FormulaSource {
                text: "SEQUENCE(1,3)".to_string(),
                is_r1c1: false,
            }),
            style_id: None,
        };

        assert!(
            state
                .replace_cells_with_change(
                    SheetId(3),
                    BTreeMap::from([((3, 3), Some(replacement.clone()))]),
                )
                .expect("replace dynamic spill anchor")
        );

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data after replacement");
        assert_eq!(worksheet.cells.get(&(3, 3)), Some(&replacement));
        assert!(worksheet.dynamic_array_formulas.contains(&(3, 3)));
        assert!(!worksheet.spill_ranges.contains_key(&(3, 3)));
        assert!(worksheet.spill_owners.is_empty());
        assert!(!worksheet.cells.contains_key(&(3, 4)));
        assert!(!worksheet.cells.contains_key(&(4, 3)));
        let styled_child = worksheet
            .cells
            .get(&(4, 4))
            .expect("styled child should remain as a blank shell");
        assert_eq!(styled_child.value, CellValue::Blank);
        assert!(styled_child.formula.is_none());
        assert_eq!(styled_child.style_id, Some(StyleId(17)));
        for key in [(3, 3), (3, 4), (4, 3), (4, 4)] {
            assert!(worksheet.dirty_cells.contains(&key), "{key:?}");
        }
        state.validate_for_save().expect("valid replacement state");
    }

    #[test]
    fn rearrange_cells_rejects_spill_topology_updates_atomically() {
        for protected_key in [(3, 3), (4, 4)] {
            let mut state = sample_state();
            seed_two_by_two_spill(&mut state);
            let before = state
                .worksheet_data_for_sheet(SheetId(3))
                .expect("worksheet data")
                .clone();

            let error = state
                .rearrange_cells_with_change(
                    SheetId(3),
                    BTreeMap::from([
                        (
                            (1, 1),
                            Some(CellData {
                                value: CellValue::Text("changed".to_string()),
                                formula: None,
                                style_id: None,
                            }),
                        ),
                        (protected_key, None),
                    ]),
                )
                .expect_err("spill topology rearrangement should fail");

            assert_eq!(error.code, OmErrorCode::InvalidState, "{protected_key:?}");
            assert!(
                error
                    .message
                    .contains(&format!("R{}C{}", protected_key.0, protected_key.1)),
                "{protected_key:?}: {error:?}",
            );
            assert_eq!(
                state
                    .worksheet_data_for_sheet(SheetId(3))
                    .expect("worksheet data"),
                &before,
                "{protected_key:?}",
            );
        }
    }

    #[test]
    fn rearrange_cells_applies_plain_batch_and_tracks_dirty_cells() {
        let mut state = sample_state();
        let moved = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data")
            .cells
            .get(&(1, 1))
            .expect("source cell")
            .clone();

        assert!(
            state
                .rearrange_cells_with_change(
                    SheetId(3),
                    BTreeMap::from([((1, 1), None), ((1, 2), Some(moved.clone()))]),
                )
                .expect("rearrange plain cells")
        );

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data after rearrangement");
        assert!(!worksheet.cells.contains_key(&(1, 1)));
        assert_eq!(worksheet.cells.get(&(1, 2)), Some(&moved));
        assert!(worksheet.dirty);
        assert!(worksheet.dirty_cells.contains(&(1, 1)));
        assert!(worksheet.dirty_cells.contains(&(1, 2)));
        state.validate_for_save().expect("valid rearranged state");
    }

    #[test]
    fn shift_cells_rejects_late_overflow_atomically() {
        let mut state = sample_state();
        state
            .insert_cell(
                SheetId(3),
                2,
                4,
                CellData {
                    value: CellValue::Number(11.0),
                    formula: None,
                    style_id: None,
                },
            )
            .expect("seed safe shift lane");
        state
            .insert_cell(
                SheetId(3),
                ExcelLimits::MAX_ROW_INDEX,
                5,
                CellData {
                    value: CellValue::Number(22.0),
                    formula: None,
                    style_id: None,
                },
            )
            .expect("seed overflowing shift lane");
        let before = state.clone();

        let error = state
            .shift_cells_with_change(
                SheetId(3),
                Rect {
                    row_first: 1,
                    row_last: 1,
                    col_first: 4,
                    col_last: 5,
                },
                CellShiftDirection::Down,
            )
            .expect_err("late row overflow should reject the whole shift");

        assert_eq!(error.code, OmErrorCode::InvalidArgument);
        assert!(error.message.contains("rows"), "{error:?}");
        assert_eq!(state, before);
    }

    #[test]
    fn shift_cells_rejects_unmaterialized_spill_intersection() {
        let mut state = sample_state();
        seed_two_by_two_spill(&mut state);
        {
            let worksheet = state
                .worksheet_data_for_sheet_mut(SheetId(3))
                .expect("worksheet data");
            for key in [(3, 4), (4, 4)] {
                worksheet.cells.remove(&key);
                worksheet.spill_owners.remove(&key);
            }
            worksheet.cells.insert(
                (10, 4),
                CellData {
                    value: CellValue::Number(17.0),
                    formula: None,
                    style_id: None,
                },
            );
        }
        let before = state.clone();

        let error = state
            .shift_cells_with_change(
                SheetId(3),
                Rect::single_cell(3, 4),
                CellShiftDirection::Down,
            )
            .expect_err("geometric spill intersection should reject the whole shift");

        assert_eq!(error.code, OmErrorCode::InvalidState);
        assert!(error.message.contains("R3C4"), "{error:?}");
        assert!(error.message.contains("R3C3"), "{error:?}");
        assert_eq!(state, before);
    }

    #[test]
    fn fill_cells_rejects_spill_sources_atomically() {
        for source_key in [(3, 3), (4, 4)] {
            let mut state = sample_state();
            seed_two_by_two_spill(&mut state);
            let before = state
                .worksheet_data_for_sheet(SheetId(3))
                .expect("worksheet data")
                .clone();

            let error = state
                .fill_cells_with_change(
                    SheetId(3),
                    BTreeSet::from([source_key]),
                    BTreeMap::from([(
                        (1, 1),
                        Some(CellData {
                            value: CellValue::Text("changed".to_string()),
                            formula: None,
                            style_id: None,
                        }),
                    )]),
                )
                .expect_err("spill source should reject the whole fill");

            assert_eq!(error.code, OmErrorCode::InvalidState, "{source_key:?}");
            assert!(
                error
                    .message
                    .contains(&format!("R{}C{}", source_key.0, source_key.1)),
                "{source_key:?}: {error:?}",
            );
            assert_eq!(
                state
                    .worksheet_data_for_sheet(SheetId(3))
                    .expect("worksheet data"),
                &before,
                "{source_key:?}",
            );
        }
    }

    #[test]
    fn fill_cells_preflights_unchanged_spill_destinations_before_updates() {
        for destination_key in [(3, 3), (4, 4)] {
            let mut state = sample_state();
            seed_two_by_two_spill(&mut state);
            let before = state
                .worksheet_data_for_sheet(SheetId(3))
                .expect("worksheet data")
                .clone();
            let unchanged_spill_cell = before
                .cells
                .get(&destination_key)
                .expect("materialized spill cell")
                .clone();

            let error = state
                .fill_cells_with_change(
                    SheetId(3),
                    BTreeSet::from([(1, 1)]),
                    BTreeMap::from([
                        (
                            (1, 2),
                            Some(CellData {
                                value: CellValue::Text("changed".to_string()),
                                formula: None,
                                style_id: None,
                            }),
                        ),
                        (destination_key, Some(unchanged_spill_cell)),
                    ]),
                )
                .expect_err("spill destination should reject the whole fill");

            assert_eq!(error.code, OmErrorCode::InvalidState, "{destination_key:?}");
            assert!(
                error
                    .message
                    .contains(&format!("R{}C{}", destination_key.0, destination_key.1)),
                "{destination_key:?}: {error:?}",
            );
            assert_eq!(
                state
                    .worksheet_data_for_sheet(SheetId(3))
                    .expect("worksheet data"),
                &before,
                "{destination_key:?}",
            );
        }
    }

    #[test]
    fn fill_cells_applies_plain_batch_and_tracks_dirty_cells() {
        let mut state = sample_state();
        let source = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data")
            .cells
            .get(&(1, 1))
            .expect("source cell")
            .clone();

        assert!(
            state
                .fill_cells_with_change(
                    SheetId(3),
                    BTreeSet::from([(1, 1)]),
                    BTreeMap::from([((1, 2), Some(source.clone()))]),
                )
                .expect("fill plain cell")
        );

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data after fill");
        assert_eq!(worksheet.cells.get(&(1, 1)), Some(&source));
        assert_eq!(worksheet.cells.get(&(1, 2)), Some(&source));
        assert!(worksheet.dirty);
        assert!(worksheet.dirty_cells.contains(&(1, 2)));
        state.validate_for_save().expect("valid filled state");
    }

    #[test]
    fn copy_source_validation_rejects_spill_anchor_and_child() {
        for source_key in [(3, 3), (4, 4)] {
            let mut state = sample_state();
            seed_two_by_two_spill(&mut state);
            let before = state
                .worksheet_data_for_sheet(SheetId(3))
                .expect("worksheet data")
                .clone();

            let error = state
                .validate_copy_source_cells(SheetId(3), &BTreeSet::from([source_key]))
                .expect_err("spill source should be rejected");

            assert_eq!(error.code, OmErrorCode::InvalidState, "{source_key:?}");
            assert!(
                error
                    .message
                    .contains(&format!("R{}C{}", source_key.0, source_key.1)),
                "{source_key:?}: {error:?}",
            );
            assert_eq!(
                state
                    .worksheet_data_for_sheet(SheetId(3))
                    .expect("worksheet data"),
                &before,
                "{source_key:?}",
            );
        }
    }

    #[test]
    fn copy_cells_preflights_unchanged_spill_destinations_before_updates() {
        for destination_key in [(3, 3), (4, 4)] {
            let mut state = sample_state();
            seed_two_by_two_spill(&mut state);
            let before = state
                .worksheet_data_for_sheet(SheetId(3))
                .expect("worksheet data")
                .clone();
            let unchanged_spill_cell = before
                .cells
                .get(&destination_key)
                .expect("materialized spill cell")
                .clone();

            let error = state
                .copy_cells_with_change(
                    SheetId(3),
                    BTreeMap::from([
                        (
                            (1, 2),
                            Some(CellData {
                                value: CellValue::Text("changed".to_string()),
                                formula: None,
                                style_id: None,
                            }),
                        ),
                        (destination_key, Some(unchanged_spill_cell)),
                    ]),
                )
                .expect_err("spill destination should reject the whole copy");

            assert_eq!(error.code, OmErrorCode::InvalidState, "{destination_key:?}");
            assert!(
                error
                    .message
                    .contains(&format!("R{}C{}", destination_key.0, destination_key.1)),
                "{destination_key:?}: {error:?}",
            );
            assert_eq!(
                state
                    .worksheet_data_for_sheet(SheetId(3))
                    .expect("worksheet data"),
                &before,
                "{destination_key:?}",
            );
        }
    }

    #[test]
    fn copy_cells_applies_plain_batch_and_tracks_dirty_cells() {
        let mut state = sample_state();
        let source = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data")
            .cells
            .get(&(1, 1))
            .expect("source cell")
            .clone();
        state
            .validate_copy_source_cells(SheetId(3), &BTreeSet::from([(1, 1)]))
            .expect("plain source should be copyable");

        assert!(
            state
                .copy_cells_with_change(
                    SheetId(3),
                    BTreeMap::from([((1, 2), Some(source.clone()))]),
                )
                .expect("copy plain cell")
        );

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data after copy");
        assert_eq!(worksheet.cells.get(&(1, 1)), Some(&source));
        assert_eq!(worksheet.cells.get(&(1, 2)), Some(&source));
        assert!(worksheet.dirty);
        assert!(worksheet.dirty_cells.contains(&(1, 2)));
        state.validate_for_save().expect("valid copied state");
    }

    #[test]
    fn cut_source_validation_rejects_spill_anchor_and_child() {
        for source_key in [(3, 3), (4, 4)] {
            let mut state = sample_state();
            seed_two_by_two_spill(&mut state);
            let before = state
                .worksheet_data_for_sheet(SheetId(3))
                .expect("worksheet data")
                .clone();

            let error = state
                .validate_cut_source_cells(SheetId(3), &BTreeSet::from([source_key]))
                .expect_err("spill source should be rejected");

            assert_eq!(error.code, OmErrorCode::InvalidState, "{source_key:?}");
            assert!(
                error
                    .message
                    .contains(&format!("R{}C{}", source_key.0, source_key.1)),
                "{source_key:?}: {error:?}",
            );
            assert_eq!(
                state
                    .worksheet_data_for_sheet(SheetId(3))
                    .expect("worksheet data"),
                &before,
                "{source_key:?}",
            );
        }
    }

    #[test]
    fn cut_cells_preflights_unchanged_spill_cells_before_updates() {
        for protected_key in [(3, 3), (4, 4)] {
            let mut state = sample_state();
            seed_two_by_two_spill(&mut state);
            let before = state
                .worksheet_data_for_sheet(SheetId(3))
                .expect("worksheet data")
                .clone();
            let unchanged_spill_cell = before
                .cells
                .get(&protected_key)
                .expect("materialized spill cell")
                .clone();

            let error = state
                .cut_cells_with_change(
                    SheetId(3),
                    BTreeMap::from([
                        (
                            (1, 2),
                            Some(CellData {
                                value: CellValue::Text("changed".to_string()),
                                formula: None,
                                style_id: None,
                            }),
                        ),
                        (protected_key, Some(unchanged_spill_cell)),
                    ]),
                )
                .expect_err("spill cell should reject the whole cut");

            assert_eq!(error.code, OmErrorCode::InvalidState, "{protected_key:?}");
            assert!(
                error
                    .message
                    .contains(&format!("R{}C{}", protected_key.0, protected_key.1)),
                "{protected_key:?}: {error:?}",
            );
            assert_eq!(
                state
                    .worksheet_data_for_sheet(SheetId(3))
                    .expect("worksheet data"),
                &before,
                "{protected_key:?}",
            );
        }
    }

    #[test]
    fn cut_cells_applies_plain_batch_and_tracks_dirty_cells() {
        let mut state = sample_state();
        let source = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data")
            .cells
            .get(&(1, 1))
            .expect("source cell")
            .clone();
        state
            .validate_cut_source_cells(SheetId(3), &BTreeSet::from([(1, 1)]))
            .expect("plain source should be movable");

        assert!(
            state
                .cut_cells_with_change(
                    SheetId(3),
                    BTreeMap::from([((1, 1), None), ((1, 2), Some(source.clone()))]),
                )
                .expect("cut plain cell")
        );

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data after cut");
        assert_eq!(worksheet.cells.get(&(1, 1)), None);
        assert_eq!(worksheet.cells.get(&(1, 2)), Some(&source));
        assert!(worksheet.dirty);
        assert!(worksheet.dirty_cells.contains(&(1, 1)));
        assert!(worksheet.dirty_cells.contains(&(1, 2)));
        state.validate_for_save().expect("valid cut state");
    }

    #[test]
    fn clear_range_formats_preserves_blank_spill_child_cell() {
        let mut state = sample_state();
        seed_two_by_two_spill(&mut state);
        state
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data")
            .cells
            .get_mut(&(4, 4))
            .expect("styled spill child")
            .value = CellValue::Blank;

        assert!(
            state
                .clear_range_formats_with_change(&RangeRef::single_cell(
                    WorkbookId(7),
                    SheetId(3),
                    4,
                    4,
                ))
                .expect("clear spill-child format")
        );

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data after clear");
        let child = worksheet
            .cells
            .get(&(4, 4))
            .expect("blank spill child cell remains materialized");
        assert_eq!(child.value, CellValue::Blank);
        assert!(child.formula.is_none());
        assert!(child.style_id.is_none());
        assert_eq!(worksheet.spill_owners.get(&(4, 4)), Some(&(3, 3)));
        assert!(worksheet.dirty);
        assert!(worksheet.dirty_cells.contains(&(4, 4)));
        state.validate_for_save().expect("valid spill topology");
    }

    #[test]
    fn setting_spill_anchor_value_clears_owned_extent() {
        let mut state = sample_state();
        seed_two_by_two_spill(&mut state);

        state
            .set_range_values(
                &RangeRef::single_cell(WorkbookId(7), SheetId(3), 3, 3),
                &OmArray::scalar(OmValue::Number(99.0)),
            )
            .expect("overwrite spill anchor");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        let anchor = worksheet.cells.get(&(3, 3)).expect("anchor");
        assert_eq!(anchor.value, CellValue::Number(99.0));
        assert!(anchor.formula.is_none());
        assert!(!worksheet.cells.contains_key(&(3, 4)));
        assert!(!worksheet.cells.contains_key(&(4, 3)));
        let styled_child = worksheet.cells.get(&(4, 4)).expect("styled child shell");
        assert_eq!(styled_child.value, CellValue::Blank);
        assert_eq!(styled_child.style_id, Some(StyleId(17)));
        assert!(worksheet.dynamic_array_formulas.is_empty());
        assert!(worksheet.spill_ranges.is_empty());
        assert!(worksheet.spill_owners.is_empty());
    }

    #[test]
    fn clearing_spill_anchor_clears_owned_extent() {
        let mut state = sample_state();
        seed_two_by_two_spill(&mut state);

        state
            .clear_range_contents(&RangeRef::single_cell(WorkbookId(7), SheetId(3), 3, 3))
            .expect("clear spill anchor");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(!worksheet.cells.contains_key(&(3, 3)));
        assert!(!worksheet.cells.contains_key(&(3, 4)));
        assert!(!worksheet.cells.contains_key(&(4, 3)));
        let styled_child = worksheet.cells.get(&(4, 4)).expect("styled child shell");
        assert_eq!(styled_child.value, CellValue::Blank);
        assert_eq!(styled_child.style_id, Some(StyleId(17)));
        assert!(worksheet.dynamic_array_formulas.is_empty());
        assert!(worksheet.spill_ranges.is_empty());
        assert!(worksheet.spill_owners.is_empty());
    }

    #[test]
    fn blank_write_preserves_styled_cell_shell() {
        let mut state = sample_state();
        state
            .set_range_values(
                &RangeRef::single_cell(WorkbookId(7), SheetId(3), 2, 2),
                &OmArray::scalar(OmValue::Empty),
            )
            .expect("clear styled cell");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        let cell = worksheet.cells.get(&(2, 2)).expect("B2");
        assert_eq!(cell.value, CellValue::Blank);
        assert_eq!(cell.style_id, Some(StyleId(9)));
        assert!(worksheet.dirty);
        assert!(worksheet.dirty_cells.contains(&(2, 2)));
    }

    #[test]
    fn clear_range_contents_removes_unstyled_cells_and_preserves_styled_shells() {
        let mut state = sample_state();
        let worksheet = state
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data");
        worksheet.cells.insert(
            (1, 2),
            CellData {
                value: CellValue::Blank,
                formula: Some(FormulaSource {
                    text: "SUM(A1:A2)".to_string(),
                    is_r1c1: false,
                }),
                style_id: None,
            },
        );

        state
            .clear_range_contents(&RangeRef::single_rect(
                WorkbookId(7),
                SheetId(3),
                Rect {
                    row_first: 1,
                    row_last: 2,
                    col_first: 1,
                    col_last: 2,
                },
            ))
            .expect("clear contents");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(worksheet.dirty);
        assert_eq!(
            worksheet.dirty_cells,
            BTreeSet::from([(1, 1), (1, 2), (2, 2)])
        );
        assert!(!worksheet.cells.contains_key(&(1, 1)));
        assert!(!worksheet.cells.contains_key(&(1, 2)));
        let styled_cell = worksheet.cells.get(&(2, 2)).expect("B2");
        assert_eq!(styled_cell.value, CellValue::Blank);
        assert!(styled_cell.formula.is_none());
        assert_eq!(styled_cell.style_id, Some(StyleId(9)));
    }

    #[test]
    fn clear_range_contents_clears_each_multi_area_rect() {
        let mut state = sample_state();
        let worksheet = state
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data");
        worksheet.cells.insert(
            (4, 4),
            CellData {
                value: CellValue::Text("keep".to_string()),
                formula: None,
                style_id: None,
            },
        );
        worksheet.cells.insert(
            (3, 3),
            CellData {
                value: CellValue::Blank,
                formula: Some(FormulaSource {
                    text: "A1+B2".to_string(),
                    is_r1c1: false,
                }),
                style_id: None,
            },
        );

        state
            .clear_range_contents(&RangeRef {
                workbook_id: WorkbookId(7),
                scope: SheetScope::Single(SheetId(3)),
                areas: vec![Rect::single_cell(1, 1), Rect::single_cell(3, 3)],
            })
            .expect("clear multi-area contents");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(worksheet.dirty);
        assert_eq!(worksheet.dirty_cells, BTreeSet::from([(1, 1), (3, 3)]));
        assert!(!worksheet.cells.contains_key(&(1, 1)));
        assert!(!worksheet.cells.contains_key(&(3, 3)));
        assert_eq!(
            worksheet.cells.get(&(2, 2)).expect("B2").value,
            CellValue::Text("hello".to_string())
        );
        assert_eq!(
            worksheet.cells.get(&(4, 4)).expect("D4").value,
            CellValue::Text("keep".to_string())
        );
    }
}
