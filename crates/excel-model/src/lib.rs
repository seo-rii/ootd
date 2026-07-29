use std::collections::{BTreeMap, BTreeSet};

use office_common::{
    CellValue, ChartId, DefinedName, DefinedNameId, DrawingId, ExcelLimits, FormulaSource,
    NameScope, NameValidationMode, OmArray, OmError, OmErrorCode, OmResult, OmValue, OpaquePart,
    RangeRef, RangeSet, Rect, ReferenceTarget, SheetId, SheetScope, StyleId, WorkbookId,
    WorkbookModel, WorksheetModel,
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

#[derive(Debug, Clone, PartialEq, Default)]
pub struct WorksheetData {
    pub cells: BTreeMap<(u32, u32), CellData>,
    pub source_xml: Vec<u8>,
    pub dirty: bool,
    pub dirty_cells: BTreeSet<(u32, u32)>,
    pub dynamic_array_formulas: BTreeSet<(u32, u32)>,
    pub spill_ranges: BTreeMap<(u32, u32), Rect>,
    pub spill_owners: BTreeMap<(u32, u32), (u32, u32)>,
}

impl WorksheetData {
    fn ensure_spill_children_are_not_edited(
        &self,
        keys: impl IntoIterator<Item = (u32, u32)>,
    ) -> OmResult<()> {
        for key in keys {
            let anchor = self.spill_owners.get(&key).copied().or_else(|| {
                self.spill_ranges.iter().find_map(|(anchor, spill_range)| {
                    (key != *anchor
                        && key.0 >= spill_range.row_first
                        && key.0 <= spill_range.row_last
                        && key.1 >= spill_range.col_first
                        && key.1 <= spill_range.col_last)
                        .then_some(*anchor)
                })
            });
            if let Some(anchor) = anchor {
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
            if worksheet.id.0 == 0 || worksheet.id.0 > u64::from(u32::MAX) {
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    format!(
                        "worksheet {} has id {} outside the supported unsigned 32-bit range",
                        worksheet.name, worksheet.id.0
                    ),
                ));
            }
            if !worksheet_ids.insert(worksheet.id) {
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    format!("duplicate worksheet id {}", worksheet.id.0),
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
            if !worksheet_names.insert(worksheet.name.to_ascii_lowercase()) {
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    format!("duplicate worksheet name: {}", worksheet.name),
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
                (None, None) if worksheet.kind == office_common::SheetKind::ChartSheet => {}
                (None, None) => {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        format!(
                            "worksheet {} ({}) has no package binding; only an unbound chart-sheet record awaiting graph preflight may omit both relationship id and part URI",
                            worksheet.name, worksheet.id.0
                        ),
                    ));
                }
                (Some(relationship_id), Some(part_uri))
                    if !relationship_id.is_empty() && !part_uri.is_empty() =>
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
                _ => {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        format!(
                            "worksheet {} ({}) must have both a non-empty relationship id and part URI, or neither",
                            worksheet.name, worksheet.id.0
                        ),
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

    fn ensure_worksheet_exists(&self, sheet_id: SheetId) -> OmResult<()> {
        if self
            .worksheets
            .iter()
            .any(|worksheet| worksheet.id == sheet_id)
        {
            return Ok(());
        }
        Err(OmError::new(
            OmErrorCode::NotFound,
            format!("unknown worksheet {}", sheet_id.0),
        ))
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
        CellData, ChartModel, ChartObjectModel, ChartSheetBinding, ChartSourceExpr, ChartType,
        DefinedNameTable, DrawingModel, DrawingObjectModel, SeriesModel, WorkbookState,
        WorksheetData,
    };
    use std::collections::{BTreeMap, BTreeSet};

    use office_common::{
        CellValue, ChartId, ChartObjectId, DrawingId, FileFormat, FormulaSource, NameScope,
        NameValidationMode, ObjectHandle, ObjectPlacement, OmArray, OmErrorCode, OmValue, RangeRef,
        RangeSet, Rect, ReferenceTarget, SheetId, SheetKind, SheetScope, SheetVisibility, StyleId,
        WorkbookId, WorkbookModel, WorksheetModel,
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
