use std::collections::{BTreeMap, BTreeSet};

use office_common::{
    CellValue, ChartId, DefinedName, DefinedNameId, DrawingId, FormulaSource, NameScope,
    NameValidationMode, OmArray, OmError, OmErrorCode, OmResult, OmValue, OpaquePart, RangeRef,
    RangeSet, Rect, ReferenceTarget, SheetId, SheetScope, StyleId, WorkbookId, WorkbookModel,
    WorksheetModel,
};

mod charts;
mod names;

pub use charts::{
    AxisModel, ChartAxisKind, ChartCacheKind, ChartCacheSnapshot, ChartCellMarkerXmlAttrs,
    ChartDisplayBlanksAs, ChartLegendPosition, ChartMarkerXmlAttrs, ChartModel, ChartObjectModel,
    ChartSheetBinding, ChartSourceExpr, ChartText, ChartType, DrawingModel, DrawingObjectModel,
    LegendModel, SeriesModel, resolve_chart_source_reference,
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
    pub fn assign_workbook_id(&mut self, workbook_id: WorkbookId) {
        self.model.id = workbook_id;
        for worksheet in &mut self.worksheets {
            worksheet.workbook_id = workbook_id;
        }
        for chart in self.charts.values_mut() {
            chart.workbook_id = workbook_id;
            for series in &mut chart.series {
                for source in [&mut series.name, &mut series.x_values, &mut series.values]
                    .into_iter()
                    .flatten()
                {
                    if let Some(ReferenceTarget::Range(range)) = source.resolved.as_mut()
                        && let Ok(updated_range) =
                            RangeSet::new(workbook_id, range.areas().to_vec())
                    {
                        *range = updated_range;
                    }
                }
            }
        }
        for drawing in self.drawings.values_mut() {
            drawing.workbook_id = workbook_id;
            for object in &mut drawing.objects {
                if let DrawingObjectModel::ChartFrame(chart_object) = object {
                    chart_object.workbook_id = workbook_id;
                }
            }
        }
    }

    pub fn set_worksheet_source_xml(&mut self, sheet_id: SheetId, source_xml: Vec<u8>) {
        self.worksheet_data(sheet_id).source_xml = source_xml;
    }

    pub fn insert_cell(&mut self, sheet_id: SheetId, row: u32, col: u32, cell: CellData) {
        let worksheet = self.worksheet_data(sheet_id);
        worksheet.cells.insert((row, col), cell);
        worksheet.dirty = true;
        worksheet.dirty_cells.insert((row, col));
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
        self.worksheet_data.get_mut(&sheet_id).ok_or_else(|| {
            OmError::new(
                OmErrorCode::NotFound,
                format!("unknown worksheet data for sheet {}", sheet_id.0),
            )
        })
    }

    pub fn get_range_values(&self, range: &RangeRef) -> OmResult<OmArray> {
        let (sheet_id, rect) = self.single_sheet_rect(range)?;
        let worksheet = self.worksheet_data_for_sheet(sheet_id)?;
        let mut values = Vec::with_capacity((rect.height() * rect.width()) as usize);

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
        let mut values = Vec::with_capacity((rect.height() * rect.width()) as usize);

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
            updates.reserve(
                rects
                    .iter()
                    .map(|rect| (rect.height() * rect.width()) as usize)
                    .sum(),
            );
            for rect in &rects {
                for row in rect.row_first..=rect.row_last {
                    for col in rect.col_first..=rect.col_last {
                        updates.push(((row, col), value.clone()));
                    }
                }
            }
        }

        let worksheet = self.worksheet_data_for_sheet_mut(sheet_id)?;
        for (key, value) in updates {
            if let Some(existing) = worksheet.cells.get_mut(&key) {
                if existing.value == value && existing.formula.is_none() {
                    continue;
                }

                existing.value = value;
                existing.formula = None;
                if matches!(existing.value, CellValue::Blank) && existing.style_id.is_none() {
                    worksheet.cells.remove(&key);
                }
                worksheet.dirty = true;
                worksheet.dirty_cells.insert(key);
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
            }
        }

        Ok(())
    }

    pub fn set_range_formulas(&mut self, range: &RangeRef, values: &OmArray) -> OmResult<()> {
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
                                        is_r1c1: false,
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
                                is_r1c1: false,
                            }),
                        )
                    } else {
                        (CellValue::Text(text), None)
                    }
                }
                other => (CellValue::try_from(other)?, None),
            };
            updates.reserve(
                rects
                    .iter()
                    .map(|rect| (rect.height() * rect.width()) as usize)
                    .sum(),
            );
            for rect in &rects {
                for row in rect.row_first..=rect.row_last {
                    for col in rect.col_first..=rect.col_last {
                        updates.push(((row, col), cell_value.clone(), formula.clone()));
                    }
                }
            }
        }

        let worksheet = self.worksheet_data_for_sheet_mut(sheet_id)?;
        for (key, cell_value, formula) in updates {
            if let Some(existing) = worksheet.cells.get_mut(&key) {
                if existing.value == cell_value && existing.formula == formula {
                    continue;
                }

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
                continue;
            }

            if !matches!(cell_value, CellValue::Blank) || formula.is_some() {
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
            }
        }

        Ok(())
    }

    pub fn clear_range_contents(&mut self, range: &RangeRef) -> OmResult<()> {
        let (sheet_id, rects) = self.same_sheet_rects(range)?;
        let worksheet = self.worksheet_data_for_sheet_mut(sheet_id)?;
        for rect in rects {
            for row in rect.row_first..=rect.row_last {
                for col in rect.col_first..=rect.col_last {
                    let key = (row, col);
                    let mut changed = false;
                    let mut remove = false;

                    if let Some(existing) = worksheet.cells.get_mut(&key) {
                        if matches!(existing.value, CellValue::Blank) && existing.formula.is_none()
                        {
                            continue;
                        }

                        existing.value = CellValue::Blank;
                        existing.formula = None;
                        remove = existing.style_id.is_none();
                        changed = true;
                    }

                    if changed {
                        if remove {
                            worksheet.cells.remove(&key);
                        }
                        worksheet.dirty = true;
                        worksheet.dirty_cells.insert(key);
                    }
                }
            }
        }

        Ok(())
    }

    fn worksheet_data(&mut self, sheet_id: SheetId) -> &mut WorksheetData {
        self.worksheet_data.entry(sheet_id).or_default()
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
            if rect.row_first == 0
                || rect.row_last == 0
                || rect.col_first == 0
                || rect.col_last == 0
            {
                return Err(OmError::invalid_argument(
                    "worksheet coordinates are 1-based and must be greater than zero",
                ));
            }
            if rect.row_first > rect.row_last || rect.col_first > rect.col_last {
                return Err(OmError::invalid_argument(
                    "worksheet range bounds must be ordered",
                ));
            }
        }

        let sheet_id = match range.scope {
            SheetScope::Single(sheet_id) => sheet_id,
            SheetScope::Multi3D { .. } => {
                return Err(OmError::unsupported(
                    "3D ranges are not supported in the core worksheet model yet",
                ));
            }
        };

        if !self
            .worksheets
            .iter()
            .any(|worksheet| worksheet.id == sheet_id)
        {
            return Err(OmError::new(
                OmErrorCode::NotFound,
                format!("unknown worksheet {}", sheet_id.0),
            ));
        }

        Ok((sheet_id, range.areas.clone()))
    }
}

#[cfg(test)]
mod tests {
    use super::{
        CellData, ChartModel, ChartObjectModel, ChartSheetBinding, ChartSourceExpr, ChartType,
        DefinedNameTable, DrawingModel, DrawingObjectModel, WorkbookState, WorksheetData,
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
                },
            )]),
            defined_names: DefinedNameTable::default(),
            charts: BTreeMap::new(),
            drawings: BTreeMap::new(),
            chart_sheets: BTreeMap::new(),
            opaque_parts: Vec::new(),
        }
    }

    fn formula_source(text: &str) -> FormulaSource {
        FormulaSource {
            text: text.to_string(),
            is_r1c1: false,
        }
    }

    #[test]
    fn workbook_state_tracks_chart_models_drawings_and_chart_sheet_bindings() {
        let mut state = sample_state();
        let workbook_id = state.model.id;
        let sheet_id = state.worksheets[0].id;
        let chart_id = ChartId(11);
        let drawing_id = DrawingId(12);

        state.charts_mut().insert(
            chart_id,
            ChartModel {
                id: chart_id,
                workbook_id,
                chart_type: ChartType::Bar,
                series: Vec::new(),
                title: None,
                legend: None,
                axes: Vec::new(),
                display_blanks_as: None,
                plot_visible_only: None,
                raw_part_uri: Some("xl/charts/chart1.xml".to_string()),
                dirty: false,
            },
        );
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
        state.insert_cell(
            SheetId(3),
            3,
            1,
            CellData {
                value: CellValue::Number(9.0),
                formula: None,
                style_id: None,
            },
        );

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
                series: vec![super::SeriesModel {
                    name: None,
                    x_values: None,
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
                        cache: None,
                        dirty: false,
                    }),
                    order: Some(0),
                }],
                title: None,
                legend: None,
                axes: Vec::new(),
                display_blanks_as: None,
                plot_visible_only: None,
                raw_part_uri: Some("xl/charts/chart1.xml".to_string()),
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

        state.assign_workbook_id(WorkbookId(99));

        assert_eq!(state.model.id, WorkbookId(99));
        assert_eq!(state.worksheets[0].workbook_id, WorkbookId(99));
        let chart = state.charts.get(&chart_id).expect("chart");
        assert_eq!(chart.workbook_id, WorkbookId(99));
        let Some(ReferenceTarget::Range(range)) = chart.series[0]
            .values
            .as_ref()
            .expect("values")
            .resolved
            .as_ref()
        else {
            panic!("expected range source");
        };
        assert_eq!(range.workbook_id(), WorkbookId(99));
        let drawing = state.drawings.get(&drawing_id).expect("drawing");
        assert_eq!(drawing.workbook_id, WorkbookId(99));
        let DrawingObjectModel::ChartFrame(chart_object) = &drawing.objects[0] else {
            panic!("expected chart object");
        };
        assert_eq!(chart_object.workbook_id, WorkbookId(99));
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
        assert!(worksheet.cells.get(&(1, 1)).is_none());
        assert!(worksheet.cells.get(&(1, 2)).is_none());
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
