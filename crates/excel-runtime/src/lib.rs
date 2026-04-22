use excel_model::WorkbookState;
use excel_xlsx::{LoadedXlsxWorkbook, XlsxCodec};
use office_codegen::{OmFocusSurfaceRegistry, build_focus_surface_registry_from_json};
use office_common::{
    ExcelProfile, FileFormat, GetRangeValuesSpec, LoadOptions, ObjectHandle, OmArray, OmError,
    OmErrorCode, OmResult, OmValue, OpaquePart, OpenWorkbookSpec, RangeHandle, RangeRef, Rect,
    SaveOptions, SaveWorkbookSpec, SetRangeValuesSpec, SheetId, WorkbookHandle, WorkbookId,
    WorkbookModel, WorksheetHandle, WorksheetModel,
};
use office_idl::{AccessMode, SupportState};
use office_opc::{CompressionMethod, OpcPackage, OpcPart};
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::{Path, PathBuf};

const ROOT_APPLICATION_HANDLE_VALUE: u64 = 0;
const FIRST_DYNAMIC_OBJECT_HANDLE_VALUE: u64 = 1_000_000;
const EXCEL_MAX_ROW_INDEX: u32 = 1_048_576;
const EXCEL_MAX_COLUMN_INDEX: u32 = 16_384;
const PINNED_OM_TEMPLATE_JSON: &str = include_str!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../specs/pinned/office_idl_excel_om.template.json"
));

#[derive(Debug)]
struct RuntimeWorkbook {
    loaded: LoadedXlsxWorkbook,
    read_only: bool,
    source_path: Option<PathBuf>,
    dirty: bool,
}

#[derive(Debug, Clone, Copy)]
struct RuntimeSelection {
    workbook: WorkbookHandle,
    sheet_id: SheetId,
    rect: Rect,
}

#[derive(Debug, Clone, Copy)]
enum RangeProjection {
    Cells,
    Rows,
    Columns,
}

#[derive(Debug, Clone, Copy)]
enum RuntimeObjectKind {
    Application,
    WorkbooksCollection,
    Workbook {
        workbook: WorkbookHandle,
    },
    WorksheetsCollection {
        workbook: WorkbookHandle,
    },
    Worksheet {
        workbook: WorkbookHandle,
        sheet_id: SheetId,
    },
    Range {
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        rect: Rect,
        projection: RangeProjection,
    },
}

#[derive(Debug)]
pub struct ExcelRuntime {
    codec: XlsxCodec,
    dispatch_registry: OmFocusSurfaceRegistry,
    root_application: ObjectHandle,
    next_handle: u64,
    next_object_handle: u64,
    next_created_workbook_index: u64,
    active_workbook: Option<WorkbookHandle>,
    selection: Option<RuntimeSelection>,
    workbooks: BTreeMap<u64, RuntimeWorkbook>,
    objects: BTreeMap<u64, RuntimeObjectKind>,
    stale_objects: BTreeSet<u64>,
}

impl ExcelRuntime {
    pub fn new() -> Self {
        let mut objects = BTreeMap::new();
        objects.insert(
            ROOT_APPLICATION_HANDLE_VALUE,
            RuntimeObjectKind::Application,
        );

        Self {
            codec: XlsxCodec,
            dispatch_registry: build_focus_surface_registry_from_json(PINNED_OM_TEMPLATE_JSON)
                .expect("pinned OM focus registry"),
            root_application: ObjectHandle(ROOT_APPLICATION_HANDLE_VALUE),
            next_handle: 1,
            next_object_handle: FIRST_DYNAMIC_OBJECT_HANDLE_VALUE,
            next_created_workbook_index: 1,
            active_workbook: None,
            selection: None,
            workbooks: BTreeMap::new(),
            objects,
            stale_objects: BTreeSet::new(),
        }
    }

    pub fn root_application(&self) -> ObjectHandle {
        self.root_application
    }

    pub fn dispatch_registry(&self) -> &OmFocusSurfaceRegistry {
        &self.dispatch_registry
    }

    pub fn open_workbook(&mut self, spec: OpenWorkbookSpec) -> OmResult<WorkbookHandle> {
        self.open_workbook_with_display_name(spec, None, None)
    }

    pub fn create_workbook(&mut self) -> OmResult<WorkbookHandle> {
        let workbook_name = format!("Book{}", self.next_created_workbook_index);
        self.next_created_workbook_index += 1;
        self.open_workbook_with_display_name(
            OpenWorkbookSpec {
                bytes: blank_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            },
            Some(workbook_name),
            None,
        )
    }

    fn open_workbook_with_display_name(
        &mut self,
        spec: OpenWorkbookSpec,
        display_name: Option<String>,
        source_path: Option<PathBuf>,
    ) -> OmResult<WorkbookHandle> {
        let mut loaded = self.codec.load(
            &spec.bytes,
            LoadOptions {
                profile: spec.profile,
                preserve_unknown_parts: true,
                read_calc_chain: true,
            },
        )?;

        if let Some(format_hint) = spec.format_hint
            && format_hint != loaded.detected_format
        {
            return Err(OmError::new(
                OmErrorCode::InvalidArgument,
                format!(
                    "format hint {format_hint:?} does not match detected workbook format {:?}",
                    loaded.detected_format
                ),
            ));
        }

        if let Some(display_name) = display_name.filter(|value| !value.is_empty()) {
            loaded.state.model.display_name = display_name;
        }

        let handle_value = self.next_handle;
        self.next_handle += 1;
        let workbook_id = WorkbookId(handle_value);
        loaded.state.assign_workbook_id(workbook_id);
        let workbook_handle = WorkbookHandle(ObjectHandle(handle_value));
        let default_selection = loaded
            .state
            .worksheets
            .first()
            .map(|worksheet| RuntimeSelection {
                workbook: workbook_handle,
                sheet_id: worksheet.id,
                rect: Rect::single_cell(1, 1),
            });

        self.workbooks.insert(
            handle_value,
            RuntimeWorkbook {
                loaded,
                read_only: spec.read_only,
                source_path,
                dirty: false,
            },
        );
        self.objects.insert(
            handle_value,
            RuntimeObjectKind::Workbook {
                workbook: workbook_handle,
            },
        );
        self.active_workbook = Some(workbook_handle);
        self.selection = default_selection;

        Ok(workbook_handle)
    }

    pub fn save_workbook(
        &self,
        workbook: WorkbookHandle,
        spec: SaveWorkbookSpec,
    ) -> OmResult<Vec<u8>> {
        let runtime = self.runtime_workbook(workbook)?;
        if spec.format != runtime.loaded.detected_format {
            return Err(OmError::new(
                OmErrorCode::Unsupported,
                format!(
                    "save conversion from {:?} to {:?} is not implemented yet",
                    runtime.loaded.detected_format, spec.format
                ),
            ));
        }
        self.codec.save(
            &runtime.loaded,
            SaveOptions {
                profile: spec.profile,
                lossless: spec.lossless,
            },
        )
    }

    pub fn close_workbook(&mut self, workbook: WorkbookHandle) -> OmResult<()> {
        let WorkbookHandle(ObjectHandle(handle_value)) = workbook;
        if self.workbooks.remove(&handle_value).is_none() {
            return if self.stale_objects.contains(&handle_value) {
                Err(OmError::invalid_state("stale workbook handle"))
            } else {
                Err(OmError::new(
                    OmErrorCode::NotFound,
                    "unknown workbook handle",
                ))
            };
        }

        self.objects.remove(&handle_value);
        self.stale_objects.insert(handle_value);
        let owned_object_ids = self
            .objects
            .iter()
            .filter_map(|(&object_id, object)| {
                runtime_object_owner(*object)
                    .filter(|owner| *owner == workbook)
                    .map(|_| object_id)
            })
            .collect::<Vec<_>>();
        for object_id in owned_object_ids {
            self.objects.remove(&object_id);
            self.stale_objects.insert(object_id);
        }

        if self.active_workbook == Some(workbook) {
            self.active_workbook = self
                .workbooks
                .keys()
                .next_back()
                .copied()
                .map(|id| WorkbookHandle(ObjectHandle(id)));
        }
        if self
            .selection
            .is_some_and(|selection| selection.workbook == workbook)
        {
            self.selection = self
                .active_workbook
                .and_then(|active_workbook| self.default_selection(active_workbook).ok());
        }
        Ok(())
    }

    pub fn workbook_model(&self, workbook: WorkbookHandle) -> OmResult<&WorkbookModel> {
        Ok(&self.runtime_workbook(workbook)?.loaded.state.model)
    }

    pub fn workbook_state(&self, workbook: WorkbookHandle) -> OmResult<&WorkbookState> {
        Ok(&self.runtime_workbook(workbook)?.loaded.state)
    }

    pub fn worksheets(&self, workbook: WorkbookHandle) -> OmResult<&[WorksheetModel]> {
        Ok(&self.runtime_workbook(workbook)?.loaded.state.worksheets)
    }

    pub fn opaque_parts(&self, workbook: WorkbookHandle) -> OmResult<&[OpaquePart]> {
        Ok(&self.runtime_workbook(workbook)?.loaded.state.opaque_parts)
    }

    pub fn is_read_only(&self, workbook: WorkbookHandle) -> OmResult<bool> {
        Ok(self.runtime_workbook(workbook)?.read_only)
    }

    pub fn get_range_values(&self, spec: GetRangeValuesSpec) -> OmResult<office_common::OmArray> {
        self.runtime_workbook(spec.workbook)?
            .loaded
            .state
            .get_range_values(&spec.range)
    }

    pub fn get_range_formulas(&self, spec: GetRangeValuesSpec) -> OmResult<office_common::OmArray> {
        self.runtime_workbook(spec.workbook)?
            .loaded
            .state
            .get_range_formulas(&spec.range)
    }

    pub fn set_range_values(&mut self, spec: SetRangeValuesSpec) -> OmResult<()> {
        let runtime = self.runtime_workbook_mut(spec.workbook)?;
        if runtime.read_only {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                "cannot modify a read-only workbook",
            ));
        }

        runtime
            .loaded
            .state
            .set_range_values(&spec.range, &spec.values)
    }

    pub fn set_range_formulas(&mut self, spec: SetRangeValuesSpec) -> OmResult<()> {
        let runtime = self.runtime_workbook_mut(spec.workbook)?;
        if runtime.read_only {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                "cannot modify a read-only workbook",
            ));
        }

        runtime
            .loaded
            .state
            .set_range_formulas(&spec.range, &spec.values)
    }

    pub fn dispatch_get(
        &mut self,
        handle: ObjectHandle,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        match self.runtime_object(handle)? {
            RuntimeObjectKind::Application => self.dispatch_get_application(member, args),
            RuntimeObjectKind::WorkbooksCollection => self.dispatch_get_workbooks(member, args),
            RuntimeObjectKind::Workbook { workbook } => {
                self.dispatch_get_workbook(workbook, member, args)
            }
            RuntimeObjectKind::WorksheetsCollection { workbook } => {
                self.dispatch_get_worksheets(workbook, member, args)
            }
            RuntimeObjectKind::Worksheet { workbook, sheet_id } => {
                self.dispatch_get_worksheet(workbook, sheet_id, member, args)
            }
            RuntimeObjectKind::Range {
                workbook,
                sheet_id,
                rect,
                projection,
            } => self.dispatch_get_range(workbook, sheet_id, rect, projection, member, args),
        }
    }

    pub fn dispatch_set(
        &mut self,
        handle: ObjectHandle,
        member: &str,
        value: OmValue,
        args: &[OmValue],
    ) -> OmResult<()> {
        match self.runtime_object(handle)? {
            RuntimeObjectKind::Workbook { workbook } => {
                self.focus_member_supported("Workbook", member, true)?;
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "Workbook.{member} does not accept index arguments"
                    )));
                }
                match member {
                    "Saved" => {
                        let OmValue::Bool(saved) = value else {
                            return Err(OmError::type_mismatch(
                                "Workbook.Saved expects a boolean value",
                            ));
                        };
                        if saved {
                            self.clear_workbook_dirty_state(workbook)?;
                        } else {
                            self.runtime_workbook_mut(workbook)?.dirty = true;
                        }
                        Ok(())
                    }
                    _ => Err(OmError::unsupported(format!(
                        "Workbook.{member} is not writable"
                    ))),
                }
            }
            RuntimeObjectKind::Worksheet { workbook, sheet_id } => {
                self.focus_member_supported("Worksheet", member, true)?;
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "Worksheet.{member} does not accept index arguments"
                    )));
                }
                match member {
                    "Name" => {
                        let OmValue::Text(new_name) = value else {
                            return Err(OmError::type_mismatch(
                                "Worksheet.Name expects a string value",
                            ));
                        };
                        let runtime = self.runtime_workbook_mut(workbook)?;
                        if runtime.read_only {
                            return Err(OmError::new(
                                OmErrorCode::InvalidState,
                                "cannot modify a read-only workbook",
                            ));
                        }
                        if new_name.trim().is_empty() {
                            return Err(OmError::invalid_argument(
                                "Worksheet.Name cannot be empty",
                            ));
                        }
                        if new_name.chars().count() > 31 {
                            return Err(OmError::invalid_argument(
                                "Worksheet.Name cannot exceed 31 characters",
                            ));
                        }
                        if new_name.chars().any(|ch| {
                            matches!(ch, ':' | '\\' | '/' | '?' | '*' | '[' | ']')
                                || ch.is_control()
                        }) {
                            return Err(OmError::invalid_argument(
                                "Worksheet.Name contains invalid characters",
                            ));
                        }
                        if runtime.loaded.state.worksheets.iter().any(|worksheet| {
                            worksheet.id != sheet_id
                                && worksheet.name.eq_ignore_ascii_case(new_name.as_str())
                        }) {
                            return Err(OmError::invalid_argument(format!(
                                "worksheet name {new_name:?} is already in use"
                            )));
                        }
                        let worksheet = runtime
                            .loaded
                            .state
                            .worksheets
                            .iter_mut()
                            .find(|worksheet| worksheet.id == sheet_id)
                            .ok_or_else(|| {
                                OmError::new(OmErrorCode::NotFound, "unknown worksheet")
                            })?;
                        if worksheet.name != new_name {
                            worksheet.name = new_name;
                            runtime.dirty = true;
                        }
                        Ok(())
                    }
                    _ => Err(OmError::unsupported(format!(
                        "Worksheet.{member} is not writable"
                    ))),
                }
            }
            RuntimeObjectKind::Range {
                workbook,
                sheet_id,
                rect,
                projection,
            } => self.dispatch_set_range(workbook, sheet_id, rect, projection, member, value, args),
            RuntimeObjectKind::Application
            | RuntimeObjectKind::WorkbooksCollection
            | RuntimeObjectKind::WorksheetsCollection { .. } => Err(OmError::unsupported(format!(
                "member {member} is not writable for this object handle"
            ))),
        }
    }

    pub fn dispatch_invoke(
        &mut self,
        handle: ObjectHandle,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        match self.runtime_object(handle)? {
            RuntimeObjectKind::Application => self.dispatch_invoke_application(member, args),
            RuntimeObjectKind::WorkbooksCollection => self.dispatch_invoke_workbooks(member, args),
            RuntimeObjectKind::Workbook { workbook } => {
                self.dispatch_invoke_workbook(workbook, member, args)
            }
            RuntimeObjectKind::WorksheetsCollection { workbook } => {
                self.dispatch_invoke_worksheets(workbook, member, args)
            }
            RuntimeObjectKind::Worksheet { workbook, sheet_id } => {
                self.dispatch_invoke_worksheet(workbook, sheet_id, member, args)
            }
            RuntimeObjectKind::Range {
                workbook,
                sheet_id,
                rect,
                projection,
            } => {
                self.focus_member_supported("Range", member, false)?;
                match member {
                    "Item" => {
                        let item_rect = match args {
                            [index] => {
                                let index = coerce_u32_arg(index, "Range.Item index")?;
                                match projection {
                                    RangeProjection::Rows => {
                                        if index > rect.height() {
                                            return Err(OmError::invalid_argument(
                                                "Range.Item row index is out of bounds",
                                            ));
                                        }
                                        Rect {
                                            row_first: rect.row_first + index - 1,
                                            row_last: rect.row_first + index - 1,
                                            col_first: rect.col_first,
                                            col_last: rect.col_last,
                                        }
                                    }
                                    RangeProjection::Columns => {
                                        if index > rect.width() {
                                            return Err(OmError::invalid_argument(
                                                "Range.Item column index is out of bounds",
                                            ));
                                        }
                                        Rect {
                                            row_first: rect.row_first,
                                            row_last: rect.row_last,
                                            col_first: rect.col_first + index - 1,
                                            col_last: rect.col_first + index - 1,
                                        }
                                    }
                                    RangeProjection::Cells => {
                                        let cell_count =
                                            u64::from(rect.width()) * u64::from(rect.height());
                                        if u64::from(index) > cell_count {
                                            return Err(OmError::invalid_argument(
                                                "Range.Item index is out of bounds",
                                            ));
                                        }
                                        let zero_based = u64::from(index - 1);
                                        let row_offset =
                                            (zero_based / u64::from(rect.width())) as u32;
                                        let col_offset =
                                            (zero_based % u64::from(rect.width())) as u32;
                                        Rect::single_cell(
                                            rect.row_first + row_offset,
                                            rect.col_first + col_offset,
                                        )
                                    }
                                }
                            }
                            [row_index, column_index] => {
                                let row_index = coerce_u32_arg(row_index, "Range.Item row index")?;
                                let column_index =
                                    coerce_u32_arg(column_index, "Range.Item column index")?;
                                if row_index > rect.height() {
                                    return Err(OmError::invalid_argument(
                                        "Range.Item row index is out of bounds",
                                    ));
                                }
                                if column_index > rect.width() {
                                    return Err(OmError::invalid_argument(
                                        "Range.Item column index is out of bounds",
                                    ));
                                }
                                Rect::single_cell(
                                    rect.row_first + row_index - 1,
                                    rect.col_first + column_index - 1,
                                )
                            }
                            _ => {
                                return Err(OmError::invalid_argument(
                                    "Range.Item expects an index and optional column index",
                                ));
                            }
                        };
                        Ok(OmValue::Object(
                            self.register_range_handle(workbook, sheet_id, item_rect).0,
                        ))
                    }
                    "Offset" => {
                        let coerce_offset = |value: &OmValue, label: &str| -> OmResult<i32> {
                            match value {
                                OmValue::Missing | OmValue::Empty | OmValue::Null => Ok(0),
                                OmValue::Number(number) => {
                                    if !number.is_finite()
                                        || number.fract() != 0.0
                                        || *number < i32::MIN as f64
                                        || *number > i32::MAX as f64
                                    {
                                        return Err(OmError::invalid_argument(format!(
                                            "{label} must be a whole number"
                                        )));
                                    }
                                    Ok(*number as i32)
                                }
                                _ => Err(OmError::type_mismatch(format!(
                                    "{label} must be numeric when provided"
                                ))),
                            }
                        };
                        let translate_axis =
                            |value: u32, offset: i32, label: &str| -> OmResult<u32> {
                                let translated = i64::from(value) + i64::from(offset);
                                if !(1..=i64::from(u32::MAX)).contains(&translated) {
                                    return Err(OmError::invalid_argument(format!(
                                        "{label} moves the range outside worksheet bounds"
                                    )));
                                }
                                Ok(translated as u32)
                            };

                        let (row_offset, column_offset) = match args {
                            [] => (0, 0),
                            [row_offset] => {
                                (coerce_offset(row_offset, "Range.Offset row offset")?, 0)
                            }
                            [row_offset, column_offset] => (
                                coerce_offset(row_offset, "Range.Offset row offset")?,
                                coerce_offset(column_offset, "Range.Offset column offset")?,
                            ),
                            _ => {
                                return Err(OmError::invalid_argument(
                                    "Range.Offset expects optional row and column offsets",
                                ));
                            }
                        };
                        let offset_rect = Rect {
                            row_first: translate_axis(
                                rect.row_first,
                                row_offset,
                                "Range.Offset row offset",
                            )?,
                            row_last: translate_axis(
                                rect.row_last,
                                row_offset,
                                "Range.Offset row offset",
                            )?,
                            col_first: translate_axis(
                                rect.col_first,
                                column_offset,
                                "Range.Offset column offset",
                            )?,
                            col_last: translate_axis(
                                rect.col_last,
                                column_offset,
                                "Range.Offset column offset",
                            )?,
                        };
                        Ok(OmValue::Object(
                            self.register_projected_range_handle(
                                workbook,
                                sheet_id,
                                offset_rect,
                                projection,
                            )
                            .0,
                        ))
                    }
                    "Resize" => {
                        let coerce_size = |value: &OmValue,
                                           default: u32,
                                           label: &str|
                         -> OmResult<u32> {
                            match value {
                                OmValue::Missing | OmValue::Empty | OmValue::Null => Ok(default),
                                OmValue::Number(number) => coerce_positive_index(*number, label),
                                _ => Err(OmError::type_mismatch(format!(
                                    "{label} must be numeric when provided"
                                ))),
                            }
                        };

                        let (row_size, column_size) = match args {
                            [] => (rect.height(), rect.width()),
                            [row_size] => (
                                coerce_size(row_size, rect.height(), "Range.Resize row size")?,
                                rect.width(),
                            ),
                            [row_size, column_size] => (
                                coerce_size(row_size, rect.height(), "Range.Resize row size")?,
                                coerce_size(column_size, rect.width(), "Range.Resize column size")?,
                            ),
                            _ => {
                                return Err(OmError::invalid_argument(
                                    "Range.Resize expects optional row and column sizes",
                                ));
                            }
                        };
                        let resized_rect = Rect {
                            row_first: rect.row_first,
                            row_last: rect.row_first.checked_add(row_size - 1).ok_or_else(
                                || {
                                    OmError::invalid_argument(
                                        "Range.Resize row size overflows worksheet bounds",
                                    )
                                },
                            )?,
                            col_first: rect.col_first,
                            col_last: rect.col_first.checked_add(column_size - 1).ok_or_else(
                                || {
                                    OmError::invalid_argument(
                                        "Range.Resize column size overflows worksheet bounds",
                                    )
                                },
                            )?,
                        };
                        Ok(OmValue::Object(
                            self.register_projected_range_handle(
                                workbook,
                                sheet_id,
                                resized_rect,
                                projection,
                            )
                            .0,
                        ))
                    }
                    "Select" => {
                        if !args.is_empty() {
                            return Err(OmError::invalid_argument(
                                "Range.Select does not accept arguments",
                            ));
                        }
                        self.set_selection(workbook, sheet_id, rect);
                        Ok(OmValue::Empty)
                    }
                    "ClearContents" => {
                        if !args.is_empty() {
                            return Err(OmError::invalid_argument(
                                "Range.ClearContents does not accept arguments",
                            ));
                        }
                        let range = self.range_ref(workbook, sheet_id, rect)?;
                        let runtime = self.runtime_workbook_mut(workbook)?;
                        if runtime.read_only {
                            return Err(OmError::new(
                                OmErrorCode::InvalidState,
                                "cannot modify a read-only workbook",
                            ));
                        }
                        runtime.loaded.state.clear_range_contents(&range)?;
                        Ok(OmValue::Empty)
                    }
                    _ => Err(OmError::unsupported(format!(
                        "Range.{member} is not implemented as a method"
                    ))),
                }
            }
        }
    }

    fn runtime_workbook(&self, workbook: WorkbookHandle) -> OmResult<&RuntimeWorkbook> {
        let WorkbookHandle(ObjectHandle(handle_value)) = workbook;
        if self.stale_objects.contains(&handle_value) {
            return Err(OmError::invalid_state("stale workbook handle"));
        }
        self.workbooks
            .get(&handle_value)
            .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "unknown workbook handle"))
    }

    fn runtime_workbook_mut(&mut self, workbook: WorkbookHandle) -> OmResult<&mut RuntimeWorkbook> {
        let WorkbookHandle(ObjectHandle(handle_value)) = workbook;
        if self.stale_objects.contains(&handle_value) {
            return Err(OmError::invalid_state("stale workbook handle"));
        }
        self.workbooks
            .get_mut(&handle_value)
            .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "unknown workbook handle"))
    }

    fn runtime_object(&self, handle: ObjectHandle) -> OmResult<RuntimeObjectKind> {
        if let Some(object) = self.objects.get(&handle.0).copied() {
            return Ok(object);
        }
        if self.stale_objects.contains(&handle.0) {
            return Err(OmError::invalid_state("stale object handle"));
        }
        Err(OmError::new(OmErrorCode::NotFound, "unknown object handle"))
    }

    fn register_object(&mut self, object: RuntimeObjectKind) -> ObjectHandle {
        let handle = ObjectHandle(self.next_object_handle);
        self.next_object_handle += 1;
        self.objects.insert(handle.0, object);
        handle
    }

    fn focus_member_supported(&self, surface: &str, member: &str, write: bool) -> OmResult<()> {
        let Some(surface_entry) = self
            .dispatch_registry
            .focus_surfaces
            .iter()
            .find(|entry| entry.name == surface)
        else {
            return Err(OmError::unsupported(format!(
                "surface {surface} is not available in the pinned OM registry"
            )));
        };
        let Some(member_entry) = surface_entry
            .members
            .iter()
            .find(|entry| entry.name == member)
        else {
            return Err(OmError::new(
                OmErrorCode::NotFound,
                format!("member {surface}.{member} is not available in the pinned OM registry"),
            ));
        };
        if matches!(member_entry.support, SupportState::Unsupported) {
            return Err(OmError::unsupported(format!(
                "member {surface}.{member} is marked unsupported in the pinned OM registry"
            )));
        };
        if write
            && !matches!(
                member_entry.access,
                AccessMode::Write | AccessMode::Readwrite
            )
        {
            return Err(OmError::unsupported(format!(
                "member {surface}.{member} is not writable in the pinned OM registry"
            )));
        }
        Ok(())
    }

    fn dispatch_get_application(&mut self, member: &str, args: &[OmValue]) -> OmResult<OmValue> {
        self.focus_member_supported("Application", member, false)?;
        if member != "Cells" && !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "Application.{member} does not accept index arguments"
            )));
        }

        match member {
            "Workbooks" => Ok(OmValue::Object(
                self.register_object(RuntimeObjectKind::WorkbooksCollection),
            )),
            "ActiveWorkbook" => Ok(self
                .active_workbook
                .map(|workbook| OmValue::Object(workbook.0))
                .unwrap_or(OmValue::Empty)),
            "ActiveSheet" => {
                let Some(active_workbook) = self.active_workbook else {
                    return Ok(OmValue::Empty);
                };
                let sheet_id = self.active_sheet_id(active_workbook)?;
                Ok(OmValue::Object(
                    self.register_worksheet_handle(active_workbook, sheet_id).0,
                ))
            }
            "ActiveCell" => {
                let Some(active_workbook) = self.active_workbook else {
                    return Ok(OmValue::Empty);
                };
                let selection = self
                    .selection
                    .filter(|selection| selection.workbook == active_workbook)
                    .unwrap_or(self.default_selection(active_workbook)?);
                Ok(OmValue::Object(
                    self.register_range_handle(
                        active_workbook,
                        selection.sheet_id,
                        Rect::single_cell(selection.rect.row_first, selection.rect.col_first),
                    )
                    .0,
                ))
            }
            "Selection" => {
                let Some(active_workbook) = self.active_workbook else {
                    return Ok(OmValue::Empty);
                };
                let selection = self
                    .selection
                    .filter(|selection| selection.workbook == active_workbook)
                    .unwrap_or(self.default_selection(active_workbook)?);
                Ok(OmValue::Object(
                    self.register_range_handle(active_workbook, selection.sheet_id, selection.rect)
                        .0,
                ))
            }
            "Cells" => {
                let Some(active_workbook) = self.active_workbook else {
                    return Ok(OmValue::Empty);
                };
                let sheet_id = self.active_sheet_id(active_workbook)?;
                if args.is_empty() {
                    self.dispatch_get_worksheet(active_workbook, sheet_id, "Cells", &[])
                } else {
                    self.dispatch_invoke_worksheet(active_workbook, sheet_id, "Cells", args)
                }
            }
            _ => Err(OmError::unsupported(format!(
                "Application.{member} is not implemented"
            ))),
        }
    }

    fn dispatch_get_workbook(
        &mut self,
        workbook: WorkbookHandle,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Workbook", member, false)?;
        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "Workbook.{member} does not accept index arguments"
            )));
        }

        match member {
            "Name" => Ok(OmValue::Text(
                self.workbook_model(workbook)?.display_name.clone(),
            )),
            "Parent" => Ok(OmValue::Object(self.root_application())),
            "Path" => Ok(OmValue::Text(
                self.runtime_workbook(workbook)?
                    .source_path
                    .as_ref()
                    .and_then(|path| path.parent())
                    .map(|path| path.to_string_lossy().into_owned())
                    .unwrap_or_default(),
            )),
            "FullName" => {
                let runtime = self.runtime_workbook(workbook)?;
                Ok(OmValue::Text(
                    runtime
                        .source_path
                        .as_ref()
                        .map(|path| path.to_string_lossy().into_owned())
                        .unwrap_or_else(|| runtime.loaded.state.model.display_name.clone()),
                ))
            }
            "ReadOnly" => Ok(OmValue::Bool(self.runtime_workbook(workbook)?.read_only)),
            "Saved" => Ok(OmValue::Bool({
                let runtime = self.runtime_workbook(workbook)?;
                !runtime.dirty
                    && runtime
                        .loaded
                        .state
                        .worksheet_data
                        .values()
                        .all(|worksheet| !worksheet.dirty)
            })),
            "Worksheets" => Ok(OmValue::Object(
                self.register_object(RuntimeObjectKind::WorksheetsCollection { workbook }),
            )),
            _ => Err(OmError::unsupported(format!(
                "Workbook.{member} is not implemented as a property"
            ))),
        }
    }

    fn dispatch_get_workbooks(&mut self, member: &str, args: &[OmValue]) -> OmResult<OmValue> {
        match member {
            "Count" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Workbooks.Count does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(self.workbooks.len() as f64))
            }
            "Parent" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Workbooks.Parent does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(self.root_application()))
            }
            "Item" => self.resolve_workbook_item(args),
            _ => Err(OmError::unsupported(format!(
                "Workbooks.{member} is not implemented"
            ))),
        }
    }

    fn dispatch_get_worksheets(
        &mut self,
        workbook: WorkbookHandle,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        match member {
            "Count" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheets.Count does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(
                    self.runtime_workbook(workbook)?
                        .loaded
                        .state
                        .worksheets
                        .len() as f64,
                ))
            }
            "Parent" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheets.Parent does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(workbook.0))
            }
            "Item" => self.resolve_worksheet_item(workbook, args),
            _ => Err(OmError::unsupported(format!(
                "Worksheets.{member} is not implemented"
            ))),
        }
    }

    fn dispatch_get_worksheet(
        &mut self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Worksheet", member, false)?;

        match member {
            "Name" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheet.Name does not accept arguments",
                    ));
                }
                let worksheet = self.worksheet_model(workbook, sheet_id)?;
                Ok(OmValue::Text(worksheet.name.clone()))
            }
            "Parent" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheet.Parent does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(workbook.0))
            }
            "Index" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheet.Index does not accept arguments",
                    ));
                }
                let index = self
                    .runtime_workbook(workbook)?
                    .loaded
                    .state
                    .worksheets
                    .iter()
                    .position(|worksheet| worksheet.id == sheet_id)
                    .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "unknown worksheet"))?;
                Ok(OmValue::Number((index + 1) as f64))
            }
            "UsedRange" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheet.UsedRange does not accept arguments",
                    ));
                }
                let rect = self.used_range_rect(workbook, sheet_id)?;
                Ok(OmValue::Object(
                    self.register_range_handle(workbook, sheet_id, rect).0,
                ))
            }
            "Cells" => {
                let rect = match args {
                    [] | [OmValue::Missing] | [OmValue::Empty] | [OmValue::Null] => Rect {
                        row_first: 1,
                        row_last: EXCEL_MAX_ROW_INDEX,
                        col_first: 1,
                        col_last: EXCEL_MAX_COLUMN_INDEX,
                    },
                    _ => {
                        let (row, col) = parse_cells_args(args)?;
                        Rect::single_cell(row, col)
                    }
                };
                Ok(OmValue::Object(
                    self.register_range_handle(workbook, sheet_id, rect).0,
                ))
            }
            "Rows" => {
                let rect = match args {
                    [] | [OmValue::Missing] | [OmValue::Empty] | [OmValue::Null] => Rect {
                        row_first: 1,
                        row_last: EXCEL_MAX_ROW_INDEX,
                        col_first: 1,
                        col_last: EXCEL_MAX_COLUMN_INDEX,
                    },
                    [OmValue::Text(reference)] => {
                        let reference = reference.trim().replace('$', "");
                        let parts: Vec<_> = reference.split(':').collect();
                        let parse_row = |part: &str| -> OmResult<u32> {
                            if part.is_empty() || !part.chars().all(|ch| ch.is_ascii_digit()) {
                                return Err(OmError::invalid_argument(
                                    "Worksheet.Rows text selector must be a row number or range like \"2:3\"",
                                ));
                            }
                            let index = part.parse::<u32>().map_err(|_| {
                                OmError::invalid_argument(
                                    "Worksheet.Rows text selector is not a valid row number",
                                )
                            })?;
                            if index == 0 || index > EXCEL_MAX_ROW_INDEX {
                                return Err(OmError::invalid_argument(
                                    "Worksheet.Rows text selector is out of bounds",
                                ));
                            }
                            Ok(index)
                        };
                        let (row_first, row_last) = match parts.as_slice() {
                            [single] => {
                                let index = parse_row(single)?;
                                (index, index)
                            }
                            [start, end] => {
                                let start = parse_row(start)?;
                                let end = parse_row(end)?;
                                (start.min(end), start.max(end))
                            }
                            _ => {
                                return Err(OmError::invalid_argument(
                                    "Worksheet.Rows text selector must be a row number or range like \"2:3\"",
                                ));
                            }
                        };
                        Rect {
                            row_first,
                            row_last,
                            col_first: 1,
                            col_last: EXCEL_MAX_COLUMN_INDEX,
                        }
                    }
                    [index] => {
                        let index = coerce_u32_arg(index, "Worksheet.Rows index")?;
                        if index > EXCEL_MAX_ROW_INDEX {
                            return Err(OmError::invalid_argument(
                                "Worksheet.Rows index is out of bounds",
                            ));
                        }
                        Rect {
                            row_first: index,
                            row_last: index,
                            col_first: 1,
                            col_last: EXCEL_MAX_COLUMN_INDEX,
                        }
                    }
                    _ => {
                        return Err(OmError::invalid_argument(
                            "Worksheet.Rows expects an optional row index or text range",
                        ));
                    }
                };
                Ok(OmValue::Object(
                    self.register_projected_range_handle(
                        workbook,
                        sheet_id,
                        rect,
                        RangeProjection::Rows,
                    )
                    .0,
                ))
            }
            "Columns" => {
                let rect = match args {
                    [] | [OmValue::Missing] | [OmValue::Empty] | [OmValue::Null] => Rect {
                        row_first: 1,
                        row_last: EXCEL_MAX_ROW_INDEX,
                        col_first: 1,
                        col_last: EXCEL_MAX_COLUMN_INDEX,
                    },
                    [OmValue::Number(number)] => {
                        let index = coerce_positive_index(*number, "Worksheet.Columns index")?;
                        if index > EXCEL_MAX_COLUMN_INDEX {
                            return Err(OmError::invalid_argument(
                                "Worksheet.Columns index is out of bounds",
                            ));
                        }
                        Rect {
                            row_first: 1,
                            row_last: EXCEL_MAX_ROW_INDEX,
                            col_first: index,
                            col_last: index,
                        }
                    }
                    [OmValue::Text(reference)] => {
                        let reference = reference.trim().replace('$', "").to_ascii_uppercase();
                        let parts: Vec<_> = reference.split(':').collect();
                        let parse_column = |part: &str| -> OmResult<u32> {
                            if part.is_empty() || !part.chars().all(|ch| ch.is_ascii_alphabetic()) {
                                return Err(OmError::invalid_argument(
                                    "Worksheet.Columns text selector must be a column label or range like \"B:C\"",
                                ));
                            }

                            let mut index = 0u32;
                            for ch in part.bytes() {
                                index = index
                                    .checked_mul(26)
                                    .and_then(|value| value.checked_add((ch - b'A' + 1) as u32))
                                    .ok_or_else(|| {
                                        OmError::invalid_argument(
                                            "Worksheet.Columns text selector overflows column bounds",
                                        )
                                    })?;
                            }
                            if index > EXCEL_MAX_COLUMN_INDEX {
                                return Err(OmError::invalid_argument(
                                    "Worksheet.Columns text selector is out of bounds",
                                ));
                            }
                            Ok(index)
                        };
                        let (col_first, col_last) = match parts.as_slice() {
                            [single] => {
                                let index = parse_column(single)?;
                                (index, index)
                            }
                            [start, end] => {
                                let start = parse_column(start)?;
                                let end = parse_column(end)?;
                                (start.min(end), start.max(end))
                            }
                            _ => {
                                return Err(OmError::invalid_argument(
                                    "Worksheet.Columns text selector must be a column label or range like \"B:C\"",
                                ));
                            }
                        };
                        Rect {
                            row_first: 1,
                            row_last: EXCEL_MAX_ROW_INDEX,
                            col_first,
                            col_last,
                        }
                    }
                    [_] => {
                        return Err(OmError::type_mismatch(
                            "Worksheet.Columns expects a numeric index or column label string",
                        ));
                    }
                    _ => {
                        return Err(OmError::invalid_argument(
                            "Worksheet.Columns expects an optional column index or text range",
                        ));
                    }
                };
                Ok(OmValue::Object(
                    self.register_projected_range_handle(
                        workbook,
                        sheet_id,
                        rect,
                        RangeProjection::Columns,
                    )
                    .0,
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "Worksheet.{member} is not implemented as a property"
            ))),
        }
    }

    fn dispatch_get_range(
        &mut self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        rect: Rect,
        projection: RangeProjection,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Range", member, false)?;
        if member != "Address" && !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "Range.{member} does not accept arguments"
            )));
        }

        match member {
            "Value" | "Value2" | "Formula" => {
                let array = if member == "Formula" {
                    self.get_range_formulas(GetRangeValuesSpec {
                        workbook,
                        range: self.range_ref(workbook, sheet_id, rect)?,
                    })?
                } else {
                    self.get_range_values(GetRangeValuesSpec {
                        workbook,
                        range: self.range_ref(workbook, sheet_id, rect)?,
                    })?
                };
                if array.rows == 1 && array.cols == 1 {
                    Ok(array.values.into_iter().next().unwrap_or(OmValue::Empty))
                } else {
                    Ok(OmValue::Array(array))
                }
            }
            "Text" => {
                let array = self.get_range_values(GetRangeValuesSpec {
                    workbook,
                    range: self.range_ref(workbook, sheet_id, rect)?,
                })?;
                let render_text = |value: &OmValue| match value {
                    OmValue::Missing | OmValue::Empty | OmValue::Null => String::new(),
                    OmValue::Bool(true) => "TRUE".to_string(),
                    OmValue::Bool(false) => "FALSE".to_string(),
                    OmValue::Number(number) => number.to_string(),
                    OmValue::Text(text) => text.clone(),
                    OmValue::Error(error) => match error {
                        office_common::CellError::Null => "#NULL!".to_string(),
                        office_common::CellError::Div0 => "#DIV/0!".to_string(),
                        office_common::CellError::Value => "#VALUE!".to_string(),
                        office_common::CellError::Ref => "#REF!".to_string(),
                        office_common::CellError::Name => "#NAME?".to_string(),
                        office_common::CellError::Num => "#NUM!".to_string(),
                        office_common::CellError::NA => "#N/A".to_string(),
                        office_common::CellError::GettingData => "#GETTING_DATA".to_string(),
                        office_common::CellError::Spill => "#SPILL!".to_string(),
                        office_common::CellError::Calc => "#CALC!".to_string(),
                        office_common::CellError::Field => "#FIELD!".to_string(),
                        office_common::CellError::Blocked => "#BLOCKED!".to_string(),
                        office_common::CellError::Unknown => "#UNKNOWN!".to_string(),
                    },
                    OmValue::Object(_) | OmValue::Array(_) => String::new(),
                };
                let Some(first) = array.values.first() else {
                    return Ok(OmValue::Text(String::new()));
                };
                let first_text = render_text(first);
                if array.values.len() == 1 {
                    Ok(OmValue::Text(first_text))
                } else if array
                    .values
                    .iter()
                    .all(|value| render_text(value) == first_text)
                {
                    Ok(OmValue::Text(first_text))
                } else {
                    Ok(OmValue::Null)
                }
            }
            "HasFormula" => {
                let worksheet_data = self
                    .runtime_workbook(workbook)?
                    .loaded
                    .state
                    .worksheet_data_for_sheet(sheet_id)?;
                let mut has_formula = false;
                let mut has_non_formula = false;

                for row in rect.row_first..=rect.row_last {
                    for col in rect.col_first..=rect.col_last {
                        if worksheet_data
                            .cells
                            .get(&(row, col))
                            .and_then(|cell| cell.formula.as_ref())
                            .is_some()
                        {
                            has_formula = true;
                        } else {
                            has_non_formula = true;
                        }

                        if has_formula && has_non_formula {
                            return Ok(OmValue::Null);
                        }
                    }
                }

                Ok(OmValue::Bool(has_formula))
            }
            "Address" => {
                let row_absolute = match args {
                    [] => true,
                    [value, ..] => {
                        coerce_optional_bool_arg(value, true, "Range.Address row absolute")?
                    }
                };
                let column_absolute = match args {
                    [] | [_] => true,
                    [_, value] => {
                        coerce_optional_bool_arg(value, true, "Range.Address column absolute")?
                    }
                    _ => {
                        return Err(OmError::invalid_argument(
                            "Range.Address accepts optional row and column absolute flags",
                        ));
                    }
                };
                Ok(OmValue::Text(format_rect_address_with_flags(
                    rect,
                    row_absolute,
                    column_absolute,
                )))
            }
            "Parent" => Ok(OmValue::Object(
                self.register_worksheet_handle(workbook, sheet_id).0,
            )),
            "Row" => Ok(OmValue::Number(rect.row_first as f64)),
            "Column" => Ok(OmValue::Number(rect.col_first as f64)),
            "Rows" => Ok(OmValue::Object(
                self.register_projected_range_handle(
                    workbook,
                    sheet_id,
                    rect,
                    RangeProjection::Rows,
                )
                .0,
            )),
            "Columns" => Ok(OmValue::Object(
                self.register_projected_range_handle(
                    workbook,
                    sheet_id,
                    rect,
                    RangeProjection::Columns,
                )
                .0,
            )),
            "Cells" => Ok(OmValue::Object(
                self.register_projected_range_handle(
                    workbook,
                    sheet_id,
                    rect,
                    RangeProjection::Cells,
                )
                .0,
            )),
            "CurrentRegion" => Ok(OmValue::Object(
                self.register_range_handle(
                    workbook,
                    sheet_id,
                    self.current_region_rect(workbook, sheet_id, rect)?,
                )
                .0,
            )),
            "EntireRow" => Ok(OmValue::Object(
                self.register_range_handle(
                    workbook,
                    sheet_id,
                    Rect {
                        row_first: rect.row_first,
                        row_last: rect.row_last,
                        col_first: 1,
                        col_last: EXCEL_MAX_COLUMN_INDEX,
                    },
                )
                .0,
            )),
            "EntireColumn" => Ok(OmValue::Object(
                self.register_range_handle(
                    workbook,
                    sheet_id,
                    Rect {
                        row_first: 1,
                        row_last: EXCEL_MAX_ROW_INDEX,
                        col_first: rect.col_first,
                        col_last: rect.col_last,
                    },
                )
                .0,
            )),
            "Count" => Ok(OmValue::Number(match projection {
                RangeProjection::Cells => u64::from(rect.width()) * u64::from(rect.height()),
                RangeProjection::Rows => u64::from(rect.height()),
                RangeProjection::Columns => u64::from(rect.width()),
            } as f64)),
            _ => Err(OmError::unsupported(format!(
                "Range.{member} is not implemented"
            ))),
        }
    }

    fn dispatch_set_range(
        &mut self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        rect: Rect,
        _projection: RangeProjection,
        member: &str,
        value: OmValue,
        args: &[OmValue],
    ) -> OmResult<()> {
        self.focus_member_supported("Range", member, true)?;
        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "Range.{member} does not accept index arguments"
            )));
        }

        match member {
            "Value" | "Value2" | "Formula" => {
                let values = match value {
                    OmValue::Array(array) => array,
                    scalar => OmArray::new(
                        rect.height() as usize,
                        rect.width() as usize,
                        vec![scalar; rect.height() as usize * rect.width() as usize],
                    )?,
                };
                if member == "Formula" {
                    self.set_range_formulas(SetRangeValuesSpec {
                        workbook,
                        range: self.range_ref(workbook, sheet_id, rect)?,
                        values,
                    })?;
                } else {
                    self.set_range_values(SetRangeValuesSpec {
                        workbook,
                        range: self.range_ref(workbook, sheet_id, rect)?,
                        values,
                    })?;
                }
                Ok(())
            }
            _ => Err(OmError::unsupported(format!(
                "Range.{member} is not writable"
            ))),
        }
    }

    fn dispatch_invoke_workbook(
        &mut self,
        workbook: WorkbookHandle,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Workbook", member, false)?;

        match member {
            "Save" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Workbook.Save does not accept arguments",
                    ));
                }
                let format = self.workbook_model(workbook)?.format;
                let bytes = self.save_workbook(
                    workbook,
                    SaveWorkbookSpec {
                        format,
                        profile: ExcelProfile::Excel365,
                        lossless: true,
                    },
                )?;
                if let Some(path) = self.runtime_workbook(workbook)?.source_path.as_ref() {
                    fs::write(path, &bytes).map_err(|error| {
                        OmError::new(
                            OmErrorCode::Io,
                            format!("failed to write workbook {}: {error}", path.display()),
                        )
                    })?;
                    self.clear_workbook_dirty_state(workbook)?;
                }
                Ok(OmValue::Empty)
            }
            "SaveAs" => {
                if args.len() != 1 {
                    return Err(OmError::invalid_argument(
                        "Workbook.SaveAs expects a single filename argument",
                    ));
                }
                let path = match &args[0] {
                    OmValue::Text(path) => PathBuf::from(path),
                    _ => {
                        return Err(OmError::type_mismatch(
                            "Workbook.SaveAs expects a string filename",
                        ));
                    }
                };
                let format = match path
                    .extension()
                    .and_then(|value| value.to_str())
                    .map(|value| value.to_ascii_lowercase())
                    .as_deref()
                {
                    Some("xlsx") => FileFormat::Xlsx,
                    Some("xlsm") => FileFormat::Xlsm,
                    Some("xltx") => FileFormat::Xltx,
                    Some("xltm") => FileFormat::Xltm,
                    _ => self.workbook_model(workbook)?.format,
                };
                let bytes = self.save_workbook(
                    workbook,
                    SaveWorkbookSpec {
                        format,
                        profile: ExcelProfile::Excel365,
                        lossless: true,
                    },
                )?;
                fs::write(&path, &bytes).map_err(|error| {
                    OmError::new(
                        OmErrorCode::Io,
                        format!("failed to write workbook {}: {error}", path.display()),
                    )
                })?;
                let display_name = path
                    .file_name()
                    .and_then(|value| value.to_str())
                    .filter(|value| !value.is_empty())
                    .map(str::to_owned);
                {
                    let runtime = self.runtime_workbook_mut(workbook)?;
                    runtime.source_path = Some(path);
                    runtime.loaded.state.model.format = format;
                    if let Some(display_name) = display_name {
                        runtime.loaded.state.model.display_name = display_name;
                    }
                }
                self.clear_workbook_dirty_state(workbook)?;
                Ok(OmValue::Empty)
            }
            "Close" => {
                if args.len() > 1 {
                    return Err(OmError::invalid_argument(
                        "Workbook.Close accepts at most one save_changes argument",
                    ));
                }
                let save_changes = match args {
                    [] | [OmValue::Missing] | [OmValue::Empty] | [OmValue::Null] => false,
                    [OmValue::Bool(save_changes)] => *save_changes,
                    [_] => {
                        return Err(OmError::type_mismatch(
                            "Workbook.Close save_changes expects a boolean when provided",
                        ));
                    }
                    _ => unreachable!("Workbook.Close argument count already validated"),
                };
                if save_changes {
                    let format = self.workbook_model(workbook)?.format;
                    let bytes = self.save_workbook(
                        workbook,
                        SaveWorkbookSpec {
                            format,
                            profile: ExcelProfile::Excel365,
                            lossless: true,
                        },
                    )?;
                    if let Some(path) = self.runtime_workbook(workbook)?.source_path.as_ref() {
                        fs::write(path, &bytes).map_err(|error| {
                            OmError::new(
                                OmErrorCode::Io,
                                format!("failed to write workbook {}: {error}", path.display()),
                            )
                        })?;
                    }
                }
                self.close_workbook(workbook)?;
                Ok(OmValue::Empty)
            }
            _ => Err(OmError::unsupported(format!(
                "Workbook.{member} is not implemented as a method"
            ))),
        }
    }

    fn dispatch_invoke_application(&mut self, member: &str, args: &[OmValue]) -> OmResult<OmValue> {
        self.focus_member_supported("Application", member, false)?;

        match member {
            "CalculateFullRebuild" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.CalculateFullRebuild does not accept arguments",
                    ));
                }
                Ok(OmValue::Empty)
            }
            "Range" => {
                let Some(active_workbook) = self.active_workbook else {
                    return Err(OmError::invalid_state("application has no active workbook"));
                };
                let sheet_id = self.active_sheet_id(active_workbook)?;
                self.dispatch_invoke_worksheet(active_workbook, sheet_id, "Range", args)
            }
            _ => Err(OmError::unsupported(format!(
                "Application.{member} is not implemented as a method"
            ))),
        }
    }

    fn dispatch_invoke_workbooks(&mut self, member: &str, args: &[OmValue]) -> OmResult<OmValue> {
        match member {
            "Add" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Workbooks.Add does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(self.create_workbook()?.0))
            }
            "Open" => {
                if args.len() != 1 {
                    return Err(OmError::invalid_argument(
                        "Workbooks.Open expects a single filename argument",
                    ));
                }
                let path = match &args[0] {
                    OmValue::Text(path) => path,
                    _ => {
                        return Err(OmError::type_mismatch(
                            "Workbooks.Open expects a string filename",
                        ));
                    }
                };
                let bytes = fs::read(path).map_err(|error| {
                    OmError::new(
                        OmErrorCode::Io,
                        format!("failed to read workbook {path}: {error}"),
                    )
                })?;
                let display_name = Path::new(path)
                    .file_name()
                    .and_then(|value| value.to_str())
                    .map(str::to_owned);
                Ok(OmValue::Object(
                    self.open_workbook_with_display_name(
                        OpenWorkbookSpec {
                            bytes,
                            format_hint: None,
                            profile: ExcelProfile::Excel365,
                            read_only: false,
                        },
                        display_name,
                        Some(PathBuf::from(path)),
                    )?
                    .0,
                ))
            }
            "Item" => self.resolve_workbook_item(args),
            _ => Err(OmError::unsupported(format!(
                "Workbooks.{member} is not implemented as a method"
            ))),
        }
    }

    fn dispatch_invoke_worksheets(
        &mut self,
        workbook: WorkbookHandle,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        match member {
            "Item" => self.resolve_worksheet_item(workbook, args),
            _ => Err(OmError::unsupported(format!(
                "Worksheets.{member} is not implemented as a method"
            ))),
        }
    }

    fn dispatch_invoke_worksheet(
        &mut self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Worksheet", member, false)?;
        match member {
            "Range" => {
                let rect = match args {
                    [OmValue::Text(a1)] => parse_rect_a1(a1)?,
                    [OmValue::Object(handle)] => match self.runtime_object(*handle)? {
                        RuntimeObjectKind::Range {
                            workbook: range_workbook,
                            sheet_id: range_sheet_id,
                            rect,
                            ..
                        } => {
                            if range_workbook != workbook || range_sheet_id != sheet_id {
                                return Err(OmError::invalid_argument(
                                    "Worksheet.Range object argument must belong to the same worksheet",
                                ));
                            }
                            rect
                        }
                        _ => {
                            return Err(OmError::type_mismatch(
                                "Worksheet.Range expects A1 references or Range objects",
                            ));
                        }
                    },
                    [start, end] => {
                        let start = match start {
                            OmValue::Text(a1) => parse_rect_a1(a1)?,
                            OmValue::Object(handle) => match self.runtime_object(*handle)? {
                                RuntimeObjectKind::Range {
                                    workbook: range_workbook,
                                    sheet_id: range_sheet_id,
                                    rect,
                                    ..
                                } => {
                                    if range_workbook != workbook || range_sheet_id != sheet_id {
                                        return Err(OmError::invalid_argument(
                                            "Worksheet.Range object arguments must belong to the same worksheet",
                                        ));
                                    }
                                    rect
                                }
                                _ => {
                                    return Err(OmError::type_mismatch(
                                        "Worksheet.Range expects A1 references or Range objects",
                                    ));
                                }
                            },
                            _ => {
                                return Err(OmError::type_mismatch(
                                    "Worksheet.Range expects A1 references or Range objects",
                                ));
                            }
                        };
                        let end = match end {
                            OmValue::Text(a1) => parse_rect_a1(a1)?,
                            OmValue::Object(handle) => match self.runtime_object(*handle)? {
                                RuntimeObjectKind::Range {
                                    workbook: range_workbook,
                                    sheet_id: range_sheet_id,
                                    rect,
                                    ..
                                } => {
                                    if range_workbook != workbook || range_sheet_id != sheet_id {
                                        return Err(OmError::invalid_argument(
                                            "Worksheet.Range object arguments must belong to the same worksheet",
                                        ));
                                    }
                                    rect
                                }
                                _ => {
                                    return Err(OmError::type_mismatch(
                                        "Worksheet.Range expects A1 references or Range objects",
                                    ));
                                }
                            },
                            _ => {
                                return Err(OmError::type_mismatch(
                                    "Worksheet.Range expects A1 references or Range objects",
                                ));
                            }
                        };
                        Rect {
                            row_first: start.row_first.min(end.row_first),
                            row_last: start.row_last.max(end.row_last),
                            col_first: start.col_first.min(end.col_first),
                            col_last: start.col_last.max(end.col_last),
                        }
                    }
                    _ => {
                        return Err(OmError::invalid_argument(
                            "Worksheet.Range expects one A1 reference or Range object, or two A1/range endpoints",
                        ));
                    }
                };
                self.remember_selection(workbook, sheet_id, rect);
                Ok(OmValue::Object(
                    self.register_range_handle(workbook, sheet_id, rect).0,
                ))
            }
            "Activate" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheet.Activate does not accept arguments",
                    ));
                }
                let rect = self
                    .selection
                    .filter(|selection| {
                        selection.workbook == workbook && selection.sheet_id == sheet_id
                    })
                    .map(|selection| selection.rect)
                    .unwrap_or(Rect::single_cell(1, 1));
                self.set_selection(workbook, sheet_id, rect);
                Ok(OmValue::Empty)
            }
            "Cells" => {
                let (row, col) = parse_cells_args(args)?;
                let rect = Rect::single_cell(row, col);
                self.remember_selection(workbook, sheet_id, rect);
                Ok(OmValue::Object(
                    self.register_range_handle(workbook, sheet_id, rect).0,
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "Worksheet.{member} is not implemented as a method"
            ))),
        }
    }

    fn resolve_workbook_item(&self, args: &[OmValue]) -> OmResult<OmValue> {
        if args.len() != 1 {
            return Err(OmError::invalid_argument(
                "Workbooks.Item expects a single workbook index or name",
            ));
        }

        let workbook = match &args[0] {
            OmValue::Number(index) => {
                let index = coerce_positive_index(*index, "Workbooks.Item index")?;
                self.workbooks
                    .keys()
                    .nth(index as usize - 1)
                    .copied()
                    .map(|handle| WorkbookHandle(ObjectHandle(handle)))
            }
            OmValue::Text(name) => self
                .workbooks
                .iter()
                .find(|(_, runtime)| runtime.loaded.state.model.display_name == *name)
                .map(|(&handle, _)| WorkbookHandle(ObjectHandle(handle))),
            _ => {
                return Err(OmError::type_mismatch(
                    "Workbooks.Item expects a numeric index or workbook name",
                ));
            }
        }
        .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "workbook item not found"))?;

        Ok(OmValue::Object(workbook.0))
    }

    fn resolve_worksheet_item(
        &mut self,
        workbook: WorkbookHandle,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if args.len() != 1 {
            return Err(OmError::invalid_argument(
                "Worksheets.Item expects a single worksheet index or name",
            ));
        }

        let worksheet = match &args[0] {
            OmValue::Number(index) => {
                let index = coerce_positive_index(*index, "Worksheets.Item index")?;
                self.runtime_workbook(workbook)?
                    .loaded
                    .state
                    .worksheets
                    .get(index as usize - 1)
            }
            OmValue::Text(name) => self
                .runtime_workbook(workbook)?
                .loaded
                .state
                .worksheets
                .iter()
                .find(|worksheet| worksheet.name == *name),
            _ => {
                return Err(OmError::type_mismatch(
                    "Worksheets.Item expects a numeric index or worksheet name",
                ));
            }
        }
        .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "worksheet item not found"))?;

        Ok(OmValue::Object(
            self.register_worksheet_handle(workbook, worksheet.id).0,
        ))
    }

    fn register_worksheet_handle(
        &mut self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
    ) -> WorksheetHandle {
        WorksheetHandle(self.register_object(RuntimeObjectKind::Worksheet { workbook, sheet_id }))
    }

    fn register_range_handle(
        &mut self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        rect: Rect,
    ) -> RangeHandle {
        self.register_projected_range_handle(workbook, sheet_id, rect, RangeProjection::Cells)
    }

    fn register_projected_range_handle(
        &mut self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        rect: Rect,
        projection: RangeProjection,
    ) -> RangeHandle {
        RangeHandle(self.register_object(RuntimeObjectKind::Range {
            workbook,
            sheet_id,
            rect,
            projection,
        }))
    }

    fn worksheet_model(
        &self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
    ) -> OmResult<&WorksheetModel> {
        self.runtime_workbook(workbook)?
            .loaded
            .state
            .worksheets
            .iter()
            .find(|worksheet| worksheet.id == sheet_id)
            .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "unknown worksheet"))
    }

    fn used_range_rect(&self, workbook: WorkbookHandle, sheet_id: SheetId) -> OmResult<Rect> {
        let worksheet_data = self
            .runtime_workbook(workbook)?
            .loaded
            .state
            .worksheet_data_for_sheet(sheet_id)?;
        let mut rows = worksheet_data.cells.keys().map(|(row, _)| *row);
        let mut cols = worksheet_data.cells.keys().map(|(_, col)| *col);
        let Some(row_first) = rows.next() else {
            return Ok(Rect::single_cell(1, 1));
        };
        let Some(col_first) = cols.next() else {
            return Ok(Rect::single_cell(1, 1));
        };
        let row_last = worksheet_data
            .cells
            .keys()
            .map(|(row, _)| *row)
            .max()
            .unwrap_or(row_first);
        let col_last = worksheet_data
            .cells
            .keys()
            .map(|(_, col)| *col)
            .max()
            .unwrap_or(col_first);

        Ok(Rect {
            row_first,
            row_last,
            col_first,
            col_last,
        })
    }

    fn default_selection(&self, workbook: WorkbookHandle) -> OmResult<RuntimeSelection> {
        let sheet_id = self
            .runtime_workbook(workbook)?
            .loaded
            .state
            .worksheets
            .first()
            .map(|worksheet| worksheet.id)
            .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "workbook has no worksheets"))?;
        Ok(RuntimeSelection {
            workbook,
            sheet_id,
            rect: Rect::single_cell(1, 1),
        })
    }

    fn active_sheet_id(&self, workbook: WorkbookHandle) -> OmResult<SheetId> {
        Ok(self
            .selection
            .filter(|selection| selection.workbook == workbook)
            .map(|selection| selection.sheet_id)
            .unwrap_or(self.default_selection(workbook)?.sheet_id))
    }

    fn set_selection(&mut self, workbook: WorkbookHandle, sheet_id: SheetId, rect: Rect) {
        self.active_workbook = Some(workbook);
        self.selection = Some(RuntimeSelection {
            workbook,
            sheet_id,
            rect,
        });
    }

    fn remember_selection(&mut self, workbook: WorkbookHandle, sheet_id: SheetId, rect: Rect) {
        if self.active_workbook == Some(workbook) {
            self.set_selection(workbook, sheet_id, rect);
        }
    }

    fn current_region_rect(
        &self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        seed_rect: Rect,
    ) -> OmResult<Rect> {
        let worksheet_data = self
            .runtime_workbook(workbook)?
            .loaded
            .state
            .worksheet_data_for_sheet(sheet_id)?;
        let cell_is_occupied = |row: u32, col: u32| {
            worksheet_data.cells.get(&(row, col)).is_some_and(|cell| {
                cell.formula.is_some() || !matches!(cell.value, office_common::CellValue::Blank)
            })
        };
        let row_has_occupied = |row: u32, col_first: u32, col_last: u32| {
            (col_first..=col_last).any(|col| cell_is_occupied(row, col))
        };
        let col_has_occupied = |col: u32, row_first: u32, row_last: u32| {
            (row_first..=row_last).any(|row| cell_is_occupied(row, col))
        };

        let occupied_cells = worksheet_data
            .cells
            .iter()
            .filter(|((row, col), cell)| {
                *row >= seed_rect.row_first
                    && *row <= seed_rect.row_last
                    && *col >= seed_rect.col_first
                    && *col <= seed_rect.col_last
                    && (cell.formula.is_some()
                        || !matches!(cell.value, office_common::CellValue::Blank))
            })
            .map(|((row, col), _)| (*row, *col))
            .collect::<Vec<_>>();
        let Some((first_row, first_col)) = occupied_cells.first().copied() else {
            return Ok(seed_rect);
        };

        let mut region = Rect {
            row_first: occupied_cells
                .iter()
                .map(|(row, _)| *row)
                .min()
                .unwrap_or(first_row),
            row_last: occupied_cells
                .iter()
                .map(|(row, _)| *row)
                .max()
                .unwrap_or(first_row),
            col_first: occupied_cells
                .iter()
                .map(|(_, col)| *col)
                .min()
                .unwrap_or(first_col),
            col_last: occupied_cells
                .iter()
                .map(|(_, col)| *col)
                .max()
                .unwrap_or(first_col),
        };

        loop {
            let mut changed = false;

            while region.row_first > 1
                && row_has_occupied(region.row_first - 1, region.col_first, region.col_last)
            {
                region.row_first -= 1;
                changed = true;
            }
            while region.row_last < u32::MAX
                && row_has_occupied(region.row_last + 1, region.col_first, region.col_last)
            {
                region.row_last += 1;
                changed = true;
            }
            while region.col_first > 1
                && col_has_occupied(region.col_first - 1, region.row_first, region.row_last)
            {
                region.col_first -= 1;
                changed = true;
            }
            while region.col_last < u32::MAX
                && col_has_occupied(region.col_last + 1, region.row_first, region.row_last)
            {
                region.col_last += 1;
                changed = true;
            }

            if !changed {
                return Ok(region);
            }
        }
    }

    fn range_ref(
        &self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        rect: Rect,
    ) -> OmResult<RangeRef> {
        Ok(RangeRef::single_rect(
            self.workbook_model(workbook)?.id,
            sheet_id,
            rect,
        ))
    }

    fn clear_workbook_dirty_state(&mut self, workbook: WorkbookHandle) -> OmResult<()> {
        let runtime = self.runtime_workbook_mut(workbook)?;
        runtime.dirty = false;
        for worksheet in runtime.loaded.state.worksheet_data.values_mut() {
            worksheet.dirty = false;
            worksheet.dirty_cells.clear();
        }
        Ok(())
    }
}

pub fn supports_format(format: FileFormat) -> bool {
    matches!(
        format,
        FileFormat::Xlsx | FileFormat::Xlsm | FileFormat::Xltx | FileFormat::Xltm
    )
}

fn blank_workbook_bytes() -> Vec<u8> {
    let package = OpcPackage::new(vec![
        OpcPart {
            name: "[Content_Types].xml".to_string(),
            content_type: None,
            compression: CompressionMethod::Stored,
            bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#
                .to_vec(),
        },
        OpcPart {
            name: "_rels/.rels".to_string(),
            content_type: None,
            compression: CompressionMethod::Stored,
            bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#
                .to_vec(),
        },
        OpcPart {
            name: "xl/workbook.xml".to_string(),
            content_type: None,
            compression: CompressionMethod::Stored,
            bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#
                .to_vec(),
        },
        OpcPart {
            name: "xl/_rels/workbook.xml.rels".to_string(),
            content_type: None,
            compression: CompressionMethod::Stored,
            bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#
                .to_vec(),
        },
        OpcPart {
            name: "xl/worksheets/sheet1.xml".to_string(),
            content_type: None,
            compression: CompressionMethod::Stored,
            bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#
                .to_vec(),
        },
    ]);

    package.to_bytes().expect("blank workbook package bytes")
}

fn runtime_object_owner(object: RuntimeObjectKind) -> Option<WorkbookHandle> {
    match object {
        RuntimeObjectKind::Application | RuntimeObjectKind::WorkbooksCollection => None,
        RuntimeObjectKind::Workbook { workbook }
        | RuntimeObjectKind::WorksheetsCollection { workbook }
        | RuntimeObjectKind::Worksheet { workbook, .. }
        | RuntimeObjectKind::Range { workbook, .. } => Some(workbook),
    }
}

fn coerce_u32_arg(value: &OmValue, label: &str) -> OmResult<u32> {
    match value {
        OmValue::Number(number) => coerce_positive_index(*number, label),
        _ => Err(OmError::type_mismatch(format!("{label} must be numeric"))),
    }
}

fn coerce_positive_index(value: f64, label: &str) -> OmResult<u32> {
    if !value.is_finite() || value.fract() != 0.0 || value < 1.0 || value > u32::MAX as f64 {
        return Err(OmError::invalid_argument(format!(
            "{label} must be a positive 1-based integer"
        )));
    }
    Ok(value as u32)
}

fn parse_cells_args(args: &[OmValue]) -> OmResult<(u32, u32)> {
    if args.len() != 2 {
        return Err(OmError::invalid_argument(
            "Worksheet.Cells expects row and column arguments",
        ));
    }
    let column = match &args[1] {
        OmValue::Number(number) => coerce_positive_index(*number, "Worksheet.Cells column")?,
        OmValue::Text(reference) => {
            let reference = reference.trim().replace('$', "").to_ascii_uppercase();
            if reference.is_empty() || !reference.chars().all(|ch| ch.is_ascii_alphabetic()) {
                return Err(OmError::invalid_argument(
                    "Worksheet.Cells column text selector must be a column label like \"B\"",
                ));
            }

            let mut index = 0u32;
            for ch in reference.bytes() {
                index = index
                    .checked_mul(26)
                    .and_then(|value| value.checked_add((ch - b'A' + 1) as u32))
                    .ok_or_else(|| {
                        OmError::invalid_argument(
                            "Worksheet.Cells column text selector overflows column bounds",
                        )
                    })?;
            }
            if index > EXCEL_MAX_COLUMN_INDEX {
                return Err(OmError::invalid_argument(
                    "Worksheet.Cells column text selector is out of bounds",
                ));
            }
            index
        }
        _ => {
            return Err(OmError::type_mismatch(
                "Worksheet.Cells column must be numeric or a column label string",
            ));
        }
    };
    Ok((coerce_u32_arg(&args[0], "Worksheet.Cells row")?, column))
}

fn coerce_optional_bool_arg(value: &OmValue, default: bool, label: &str) -> OmResult<bool> {
    match value {
        OmValue::Missing | OmValue::Empty | OmValue::Null => Ok(default),
        OmValue::Bool(value) => Ok(*value),
        _ => Err(OmError::type_mismatch(format!(
            "{label} must be boolean when provided"
        ))),
    }
}

fn parse_rect_a1(input: &str) -> OmResult<Rect> {
    let input = input.trim();
    let mut parts = input.split(':');
    let first = parts
        .next()
        .ok_or_else(|| OmError::parse("empty A1 reference"))?;
    let second = parts.next();
    if parts.next().is_some() {
        return Err(OmError::parse("A1 range contains too many ':' separators"));
    }
    let first = parse_cell_a1(first)?;
    let second = second.map(parse_cell_a1).transpose()?.unwrap_or(first);
    Ok(Rect {
        row_first: first.0.min(second.0),
        row_last: first.0.max(second.0),
        col_first: first.1.min(second.1),
        col_last: first.1.max(second.1),
    })
}

fn parse_cell_a1(input: &str) -> OmResult<(u32, u32)> {
    let trimmed = input.trim().trim_matches('$');
    let mut letters = String::new();
    let mut digits = String::new();
    for ch in trimmed.chars() {
        if ch == '$' {
            continue;
        }
        if ch.is_ascii_alphabetic() && digits.is_empty() {
            letters.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        } else {
            return Err(OmError::parse(format!("invalid A1 reference {input:?}")));
        }
    }
    if letters.is_empty() || digits.is_empty() {
        return Err(OmError::parse(format!("invalid A1 reference {input:?}")));
    }

    let mut col = 0u32;
    for ch in letters.bytes() {
        col = col
            .checked_mul(26)
            .and_then(|value| value.checked_add((ch - b'A' + 1) as u32))
            .ok_or_else(|| OmError::parse("column index overflow"))?;
    }
    let row = digits
        .parse::<u32>()
        .map_err(|_| OmError::parse(format!("invalid row index in {input:?}")))?;
    if row == 0 || col == 0 {
        return Err(OmError::parse(format!("invalid A1 reference {input:?}")));
    }
    Ok((row, col))
}

fn format_rect_address_with_flags(rect: Rect, row_absolute: bool, column_absolute: bool) -> String {
    let start = format_cell_address(
        rect.row_first,
        rect.col_first,
        row_absolute,
        column_absolute,
    );
    if rect.row_first == rect.row_last && rect.col_first == rect.col_last {
        start
    } else {
        let end = format_cell_address(rect.row_last, rect.col_last, row_absolute, column_absolute);
        format!("{start}:{end}")
    }
}

fn format_cell_address(row: u32, col: u32, row_absolute: bool, column_absolute: bool) -> String {
    let mut address = String::new();
    if column_absolute {
        address.push('$');
    }
    address.push_str(&column_to_letters(col));
    if row_absolute {
        address.push('$');
    }
    address.push_str(&row.to_string());
    address
}

fn column_to_letters(mut col: u32) -> String {
    let mut letters = Vec::new();
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        letters.push((b'A' + rem) as char);
        col = (col - 1) / 26;
    }
    letters.iter().rev().collect()
}

#[cfg(test)]
mod tests {
    use super::{ExcelRuntime, blank_workbook_bytes, supports_format};
    use std::fs;
    use std::time::{SystemTime, UNIX_EPOCH};

    use office_common::{
        CellValue, ExcelProfile, FileFormat, GetRangeValuesSpec, LoadOptions, ObjectHandle,
        OmArray, OmErrorCode, OmValue, OpenWorkbookSpec, RangeRef, Rect, SaveWorkbookSpec,
        SetRangeValuesSpec, WorkbookId,
    };
    use office_opc::{CompressionMethod, OpcPackage, OpcPart};

    fn expect_object_handle(value: OmValue) -> ObjectHandle {
        match value {
            OmValue::Object(handle) => handle,
            other => panic!("expected object handle, got {other:?}"),
        }
    }

    fn expect_text(value: OmValue) -> String {
        match value {
            OmValue::Text(text) => text,
            other => panic!("expected text value, got {other:?}"),
        }
    }

    fn expect_number(value: OmValue) -> f64 {
        match value {
            OmValue::Number(number) => number,
            other => panic!("expected numeric value, got {other:?}"),
        }
    }

    fn expect_bool(value: OmValue) -> bool {
        match value {
            OmValue::Bool(value) => value,
            other => panic!("expected bool value, got {other:?}"),
        }
    }

    #[test]
    fn opens_and_saves_detected_format() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: true,
            })
            .expect("open workbook");

        assert_eq!(
            runtime
                .workbook_model(workbook)
                .expect("model")
                .display_name,
            "Workbook"
        );
        assert_eq!(
            runtime.worksheets(workbook).expect("worksheets")[0].name,
            "Sheet1"
        );
        assert!(runtime.is_read_only(workbook).expect("read_only"));
        assert!(supports_format(FileFormat::Xlsx));

        let bytes = runtime
            .save_workbook(
                workbook,
                SaveWorkbookSpec {
                    format: FileFormat::Xlsx,
                    profile: ExcelProfile::Excel365,
                    lossless: true,
                },
            )
            .expect("save workbook");
        let reparsed = OpcPackage::from_bytes(&bytes).expect("reparse saved workbook");
        assert!(reparsed.contains("customXml/item1.xml"));

        runtime.close_workbook(workbook).expect("close workbook");
    }

    #[test]
    fn gets_and_sets_range_values_for_mutable_workbook() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let sheet_id = runtime.worksheets(workbook).expect("worksheets")[0].id;

        let initial = runtime
            .get_range_values(GetRangeValuesSpec {
                workbook,
                range: RangeRef::single_rect(
                    WorkbookId(1),
                    sheet_id,
                    Rect {
                        row_first: 1,
                        row_last: 1,
                        col_first: 1,
                        col_last: 3,
                    },
                ),
            })
            .expect("initial range");
        assert_eq!(initial.values[0], OmValue::Number(42.0));
        assert_eq!(initial.values[1], OmValue::Text("SHARED".to_string()));
        assert_eq!(initial.values[2], OmValue::Text("shared".to_string()));

        runtime
            .set_range_values(SetRangeValuesSpec {
                workbook,
                range: RangeRef::single_rect(
                    WorkbookId(1),
                    sheet_id,
                    Rect {
                        row_first: 2,
                        row_last: 2,
                        col_first: 1,
                        col_last: 2,
                    },
                ),
                values: OmArray::new(
                    1,
                    2,
                    vec![OmValue::Text("changed".to_string()), OmValue::Bool(true)],
                )
                .expect("values"),
            })
            .expect("set range");

        let updated = runtime
            .get_range_values(GetRangeValuesSpec {
                workbook,
                range: RangeRef::single_rect(
                    WorkbookId(1),
                    sheet_id,
                    Rect {
                        row_first: 2,
                        row_last: 2,
                        col_first: 1,
                        col_last: 2,
                    },
                ),
            })
            .expect("updated range");
        assert_eq!(updated.values[0], OmValue::Text("changed".to_string()));
        assert_eq!(updated.values[1], OmValue::Bool(true));

        let saved = runtime
            .save_workbook(
                workbook,
                SaveWorkbookSpec {
                    format: FileFormat::Xlsx,
                    profile: ExcelProfile::Excel365,
                    lossless: true,
                },
            )
            .expect("save workbook");
        let reopened = ExcelRuntime::new()
            .codec
            .load(&saved, office_common::LoadOptions::default())
            .expect("reopen saved workbook");
        let reopened_sheet_id = reopened.state.worksheets[0].id;
        assert_eq!(
            reopened
                .state
                .cell(reopened_sheet_id, 2, 1)
                .expect("A2")
                .value,
            office_common::CellValue::Text("changed".to_string())
        );
        assert_eq!(
            reopened
                .state
                .cell(reopened_sheet_id, 2, 2)
                .expect("B2")
                .value,
            office_common::CellValue::Bool(true)
        );
    }

    #[test]
    fn rejects_set_range_values_for_read_only_workbook() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: true,
            })
            .expect("open workbook");
        let sheet_id = runtime.worksheets(workbook).expect("worksheets")[0].id;

        let result = runtime.set_range_values(SetRangeValuesSpec {
            workbook,
            range: RangeRef::single_cell(WorkbookId(1), sheet_id, 2, 1),
            values: OmArray::scalar(OmValue::Text("blocked".to_string())),
        });

        assert!(result.is_err());
    }

    #[test]
    fn rejects_ranges_for_a_different_workbook_id() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let sheet_id = runtime.worksheets(workbook).expect("worksheets")[0].id;

        let get_result = runtime.get_range_values(GetRangeValuesSpec {
            workbook,
            range: RangeRef::single_cell(WorkbookId(999), sheet_id, 1, 1),
        });
        assert!(get_result.is_err());

        let set_result = runtime.set_range_values(SetRangeValuesSpec {
            workbook,
            range: RangeRef::single_cell(WorkbookId(999), sheet_id, 1, 1),
            values: OmArray::scalar(OmValue::Number(10.0)),
        });
        assert!(set_result.is_err());
    }

    #[test]
    fn rejects_zero_based_range_coordinates() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let sheet_id = runtime.worksheets(workbook).expect("worksheets")[0].id;

        let result = runtime.get_range_values(GetRangeValuesSpec {
            workbook,
            range: RangeRef::single_rect(
                WorkbookId(1),
                sheet_id,
                Rect {
                    row_first: 0,
                    row_last: 1,
                    col_first: 1,
                    col_last: 1,
                },
            ),
        });

        assert!(result.is_err());
    }

    #[test]
    fn idempotent_set_range_values_keeps_worksheet_clean() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let sheet_id = runtime.worksheets(workbook).expect("worksheets")[0].id;

        runtime
            .set_range_values(SetRangeValuesSpec {
                workbook,
                range: RangeRef::single_cell(WorkbookId(1), sheet_id, 1, 1),
                values: OmArray::scalar(OmValue::Number(42.0)),
            })
            .expect("set same value");

        let worksheet = runtime
            .workbook_state(workbook)
            .expect("state")
            .worksheet_data_for_sheet(sheet_id)
            .expect("worksheet data");
        assert!(!worksheet.dirty);
    }

    #[test]
    fn rejects_format_hint_mismatch_on_open() {
        let mut runtime = ExcelRuntime::new();
        let result = runtime.open_workbook(OpenWorkbookSpec {
            bytes: synthetic_workbook_bytes(),
            format_hint: Some(FileFormat::Xlsm),
            profile: ExcelProfile::Excel365,
            read_only: false,
        });

        assert!(result.is_err());
    }

    #[test]
    fn rejects_unsupported_save_conversion() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");

        let result = runtime.save_workbook(
            workbook,
            SaveWorkbookSpec {
                format: FileFormat::Xlsm,
                profile: ExcelProfile::Excel365,
                lossless: true,
            },
        );

        assert!(result.is_err());
    }

    #[test]
    fn supports_format_rejects_strict_xlsx() {
        assert!(supports_format(FileFormat::Xlsx));
        assert!(!supports_format(FileFormat::StrictXlsx));
    }

    #[test]
    fn closed_workbook_handle_is_rejected_after_close() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");

        runtime.close_workbook(workbook).expect("close workbook");

        assert!(runtime.workbook_model(workbook).is_err());
        assert!(
            runtime
                .get_range_values(GetRangeValuesSpec {
                    workbook,
                    range: RangeRef::single_cell(WorkbookId(1), office_common::SheetId(1), 1, 1),
                })
                .is_err()
        );
    }

    #[test]
    fn closing_active_workbook_falls_back_to_previous_open_workbook() {
        let mut runtime = ExcelRuntime::new();
        let first = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open first workbook");
        let second = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open second workbook");

        runtime
            .close_workbook(second)
            .expect("close second workbook");

        let active = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveWorkbook", &[])
                .expect("active workbook"),
        );
        assert_eq!(active, first.0);
    }

    #[test]
    fn closing_workbook_twice_returns_stale_handle_error() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");

        runtime.close_workbook(workbook).expect("close workbook");
        let error = runtime
            .close_workbook(workbook)
            .expect_err("closing stale workbook should fail");

        assert_eq!(error.code, OmErrorCode::InvalidState);
    }

    #[test]
    fn root_application_dispatch_exposes_active_workbook_sheet_and_workbooks() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");

        let application = runtime.root_application();
        let active_workbook = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveWorkbook", &[])
                .expect("ActiveWorkbook"),
        );
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let active_cell = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveCell", &[])
                .expect("ActiveCell"),
        );
        let selection = expect_object_handle(
            runtime
                .dispatch_get(application, "Selection", &[])
                .expect("Selection"),
        );
        let workbooks = expect_object_handle(
            runtime
                .dispatch_get(application, "Workbooks", &[])
                .expect("Workbooks"),
        );
        let worksheets = expect_object_handle(
            runtime
                .dispatch_get(active_workbook, "Worksheets", &[])
                .expect("Worksheets"),
        );
        let workbook_item = expect_object_handle(
            runtime
                .dispatch_invoke(workbooks, "Item", &[OmValue::Number(1.0)])
                .expect("Workbooks.Item(1)"),
        );
        let worksheet_item = expect_object_handle(
            runtime
                .dispatch_invoke(worksheets, "Item", &[OmValue::Number(1.0)])
                .expect("Worksheets.Item(1)"),
        );

        assert_eq!(workbook.0, active_workbook);
        assert_eq!(active_workbook, workbook_item);
        assert_ne!(application, active_sheet);
        assert_ne!(application, workbooks);
        assert_ne!(active_workbook, worksheets);
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_sheet, "Name", &[])
                    .expect("sheet name")
            ),
            "Sheet1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(worksheet_item, "Name", &[])
                    .expect("worksheet item name")
            ),
            "Sheet1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_cell, "Address", &[])
                    .expect("active cell address")
            ),
            "$A$1"
        );
        let active_cell_parent = expect_object_handle(
            runtime
                .dispatch_get(active_cell, "Parent", &[])
                .expect("active cell parent"),
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_cell_parent, "Name", &[])
                    .expect("active cell parent name")
            ),
            expect_text(
                runtime
                    .dispatch_get(active_sheet, "Name", &[])
                    .expect("active sheet name")
            )
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(selection, "Address", &[])
                    .expect("selection address")
            ),
            "$A$1"
        );

        let updated_selection = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("B1:C1".to_string())])
                .expect("Range(B1:C1)"),
        );
        let selection_after_range = expect_object_handle(
            runtime
                .dispatch_get(application, "Selection", &[])
                .expect("Selection after Range"),
        );
        let active_cell_after_range = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveCell", &[])
                .expect("ActiveCell after Range"),
        );
        let selection_after_cells = expect_object_handle(
            runtime
                .dispatch_invoke(
                    active_sheet,
                    "Cells",
                    &[OmValue::Number(2.0), OmValue::Number(2.0)],
                )
                .expect("Cells(2, 2)"),
        );

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(updated_selection, "Address", &[])
                    .expect("updated selection address")
            ),
            "$B$1:$C$1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(selection_after_range, "Address", &[])
                    .expect("Selection after Range address")
            ),
            "$B$1:$C$1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_cell_after_range, "Address", &[])
                    .expect("ActiveCell after Range address")
            ),
            "$B$1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(selection_after_cells, "Address", &[])
                    .expect("Cells(2, 2) address")
            ),
            "$B$2"
        );
        let selection_after_cells_handle = expect_object_handle(
            runtime
                .dispatch_get(application, "Selection", &[])
                .expect("Selection after Cells"),
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(selection_after_cells_handle, "Address", &[])
                    .expect("Selection after Cells address")
            ),
            "$B$2"
        );
    }

    #[test]
    fn application_calculate_full_rebuild_dispatch_is_a_noop_entrypoint() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let application = runtime.root_application();
        let selection_before = expect_object_handle(
            runtime
                .dispatch_get(application, "Selection", &[])
                .expect("Selection before rebuild"),
        );

        assert!(matches!(
            runtime
                .dispatch_invoke(application, "CalculateFullRebuild", &[])
                .expect("Application.CalculateFullRebuild"),
            OmValue::Empty
        ));

        let selection_after = expect_object_handle(
            runtime
                .dispatch_get(application, "Selection", &[])
                .expect("Selection after rebuild"),
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(selection_before, "Address", &[])
                    .expect("Selection before rebuild address")
            ),
            expect_text(
                runtime
                    .dispatch_get(selection_after, "Address", &[])
                    .expect("Selection after rebuild address")
            )
        );
        assert_eq!(
            runtime
                .dispatch_invoke(application, "CalculateFullRebuild", &[OmValue::Bool(true)],)
                .expect_err("CalculateFullRebuild arguments should be rejected")
                .code,
            OmErrorCode::InvalidArgument
        );
    }

    #[test]
    fn worksheet_activate_switches_active_workbook_and_selection() {
        let mut runtime = ExcelRuntime::new();
        let workbook1 = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook1");
        let workbook2 = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook2");
        let worksheets1 = expect_object_handle(
            runtime
                .dispatch_get(workbook1.0, "Worksheets", &[])
                .expect("Workbook1.Worksheets"),
        );
        let worksheet1 = expect_object_handle(
            runtime
                .dispatch_invoke(worksheets1, "Item", &[OmValue::Number(1.0)])
                .expect("Workbook1.Worksheets.Item(1)"),
        );
        let worksheets2 = expect_object_handle(
            runtime
                .dispatch_get(workbook2.0, "Worksheets", &[])
                .expect("Workbook2.Worksheets"),
        );
        let worksheet2 = expect_object_handle(
            runtime
                .dispatch_invoke(worksheets2, "Item", &[OmValue::Number(1.0)])
                .expect("Workbook2.Worksheets.Item(1)"),
        );

        runtime
            .dispatch_set(
                worksheet1,
                "Name",
                OmValue::Text("FirstSheet".to_string()),
                &[],
            )
            .expect("rename workbook1 sheet");
        runtime
            .dispatch_set(
                worksheet2,
                "Name",
                OmValue::Text("SecondSheet".to_string()),
                &[],
            )
            .expect("rename workbook2 sheet");

        assert_eq!(
            expect_object_handle(
                runtime
                    .dispatch_get(runtime.root_application(), "ActiveWorkbook", &[])
                    .expect("ActiveWorkbook before Activate")
            ),
            workbook2.0
        );
        let active_sheet_before_activate = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet before Activate"),
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_sheet_before_activate, "Name", &[])
                    .expect("ActiveSheet before Activate name")
            ),
            "SecondSheet"
        );

        assert!(matches!(
            runtime
                .dispatch_invoke(worksheet1, "Activate", &[])
                .expect("Worksheet.Activate"),
            OmValue::Empty
        ));

        let application = runtime.root_application();
        let active_workbook = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveWorkbook", &[])
                .expect("ActiveWorkbook after Activate"),
        );
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet after Activate"),
        );
        let selection = expect_object_handle(
            runtime
                .dispatch_get(application, "Selection", &[])
                .expect("Selection after Activate"),
        );
        let active_cell = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveCell", &[])
                .expect("ActiveCell after Activate"),
        );

        assert_eq!(active_workbook, workbook1.0);
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_sheet, "Name", &[])
                    .expect("ActiveSheet after Activate name")
            ),
            "FirstSheet"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(selection, "Address", &[])
                    .expect("Selection after Activate address")
            ),
            "$A$1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_cell, "Address", &[])
                    .expect("ActiveCell after Activate address")
            ),
            "$A$1"
        );
        assert_eq!(
            runtime
                .dispatch_invoke(worksheet1, "Activate", &[OmValue::Bool(true)])
                .expect_err("Worksheet.Activate args should be rejected")
                .code,
            OmErrorCode::InvalidArgument
        );
    }

    #[test]
    fn range_select_switches_active_workbook_and_selection() {
        let mut runtime = ExcelRuntime::new();
        let workbook1 = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook1");
        let workbook2 = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook2");
        let worksheets1 = expect_object_handle(
            runtime
                .dispatch_get(workbook1.0, "Worksheets", &[])
                .expect("Workbook1.Worksheets"),
        );
        let worksheet1 = expect_object_handle(
            runtime
                .dispatch_invoke(worksheets1, "Item", &[OmValue::Number(1.0)])
                .expect("Workbook1.Worksheets.Item(1)"),
        );
        let worksheets2 = expect_object_handle(
            runtime
                .dispatch_get(workbook2.0, "Worksheets", &[])
                .expect("Workbook2.Worksheets"),
        );
        let worksheet2 = expect_object_handle(
            runtime
                .dispatch_invoke(worksheets2, "Item", &[OmValue::Number(1.0)])
                .expect("Workbook2.Worksheets.Item(1)"),
        );

        runtime
            .dispatch_set(
                worksheet1,
                "Name",
                OmValue::Text("FirstSheet".to_string()),
                &[],
            )
            .expect("rename workbook1 sheet");
        runtime
            .dispatch_set(
                worksheet2,
                "Name",
                OmValue::Text("SecondSheet".to_string()),
                &[],
            )
            .expect("rename workbook2 sheet");

        let range = expect_object_handle(
            runtime
                .dispatch_invoke(worksheet1, "Range", &[OmValue::Text("B1:C1".to_string())])
                .expect("Workbook1.Range(B1:C1)"),
        );
        assert_eq!(
            expect_object_handle(
                runtime
                    .dispatch_get(runtime.root_application(), "ActiveWorkbook", &[])
                    .expect("ActiveWorkbook before Select")
            ),
            workbook2.0
        );
        let active_sheet_before_select = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet before Select"),
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_sheet_before_select, "Name", &[])
                    .expect("ActiveSheet before Select name")
            ),
            "SecondSheet"
        );

        assert!(matches!(
            runtime
                .dispatch_invoke(range, "Select", &[])
                .expect("Range.Select"),
            OmValue::Empty
        ));

        let application = runtime.root_application();
        let active_workbook = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveWorkbook", &[])
                .expect("ActiveWorkbook after Select"),
        );
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet after Select"),
        );
        let selection = expect_object_handle(
            runtime
                .dispatch_get(application, "Selection", &[])
                .expect("Selection after Select"),
        );
        let active_cell = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveCell", &[])
                .expect("ActiveCell after Select"),
        );

        assert_eq!(active_workbook, workbook1.0);
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_sheet, "Name", &[])
                    .expect("ActiveSheet after Select name")
            ),
            "FirstSheet"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(selection, "Address", &[])
                    .expect("Selection after Select address")
            ),
            "$B$1:$C$1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_cell, "Address", &[])
                    .expect("ActiveCell after Select address")
            ),
            "$B$1"
        );
        assert_eq!(
            runtime
                .dispatch_invoke(range, "Select", &[OmValue::Bool(true)])
                .expect_err("Range.Select args should be rejected")
                .code,
            OmErrorCode::InvalidArgument
        );
    }

    #[test]
    fn blank_workbook_bytes_load_as_single_empty_sheet() {
        let runtime = ExcelRuntime::new();
        let loaded = runtime
            .codec
            .load(&blank_workbook_bytes(), LoadOptions::default())
            .expect("load blank workbook");

        assert_eq!(loaded.detected_format, FileFormat::Xlsx);
        assert_eq!(loaded.state.worksheets.len(), 1);
        assert_eq!(loaded.state.worksheets[0].name, "Sheet1");
        assert!(
            loaded
                .state
                .worksheet_data_for_sheet(loaded.state.worksheets[0].id)
                .expect("worksheet data")
                .cells
                .is_empty()
        );
    }

    #[test]
    fn worksheet_range_dispatch_normalizes_lowercase_absolute_and_reversed_a1_refs() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );

        let range = expect_object_handle(
            runtime
                .dispatch_invoke(
                    active_sheet,
                    "Range",
                    &[
                        OmValue::Text("$b$2".to_string()),
                        OmValue::Text("$a$1".to_string()),
                    ],
                )
                .expect("Range($b$2, $a$1)"),
        );

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(range, "Address", &[])
                    .expect("Address")
            ),
            "$A$1:$B$2"
        );
        assert_eq!(
            expect_number(runtime.dispatch_get(range, "Count", &[]).expect("Count")),
            4.0
        );
        assert_eq!(
            expect_number(runtime.dispatch_get(range, "Row", &[]).expect("Row")),
            1.0
        );
        assert_eq!(
            expect_number(runtime.dispatch_get(range, "Column", &[]).expect("Column")),
            1.0
        );
    }

    #[test]
    fn range_dispatch_rows_and_columns_project_count_over_same_rect() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1:B2".to_string())])
                .expect("Range(A1:B2)"),
        );
        let rows = expect_object_handle(
            runtime
                .dispatch_get(range, "Rows", &[])
                .expect("Range.Rows"),
        );
        let columns = expect_object_handle(
            runtime
                .dispatch_get(range, "Columns", &[])
                .expect("Range.Columns"),
        );

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(rows, "Address", &[])
                    .expect("Rows.Address")
            ),
            "$A$1:$B$2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(columns, "Address", &[])
                    .expect("Columns.Address")
            ),
            "$A$1:$B$2"
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(range, "Count", &[])
                    .expect("Range.Count")
            ),
            4.0
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(rows, "Count", &[])
                    .expect("Rows.Count")
            ),
            2.0
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(columns, "Count", &[])
                    .expect("Columns.Count")
            ),
            2.0
        );
        let rows_columns = expect_object_handle(
            runtime
                .dispatch_get(rows, "Columns", &[])
                .expect("Rows.Columns"),
        );
        let columns_rows = expect_object_handle(
            runtime
                .dispatch_get(columns, "Rows", &[])
                .expect("Columns.Rows"),
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(rows_columns, "Count", &[])
                    .expect("Rows.Columns.Count")
            ),
            2.0
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(columns_rows, "Count", &[])
                    .expect("Columns.Rows.Count")
            ),
            2.0
        );
        let rows_parent = expect_object_handle(
            runtime
                .dispatch_get(rows, "Parent", &[])
                .expect("Rows.Parent"),
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(rows_parent, "Name", &[])
                    .expect("Rows.Parent.Name")
            ),
            expect_text(
                runtime
                    .dispatch_get(active_sheet, "Name", &[])
                    .expect("ActiveSheet.Name")
            )
        );
    }

    #[test]
    fn range_item_dispatch_supports_2d_ranges_and_row_column_projections() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1:B2".to_string())])
                .expect("Range(A1:B2)"),
        );
        let rows = expect_object_handle(
            runtime
                .dispatch_get(range, "Rows", &[])
                .expect("Range.Rows"),
        );
        let columns = expect_object_handle(
            runtime
                .dispatch_get(range, "Columns", &[])
                .expect("Range.Columns"),
        );
        let second_cell = expect_object_handle(
            runtime
                .dispatch_invoke(range, "Item", &[OmValue::Number(2.0), OmValue::Number(2.0)])
                .expect("Range.Item(2, 2)"),
        );
        let linear_second_cell = expect_object_handle(
            runtime
                .dispatch_invoke(range, "Item", &[OmValue::Number(2.0)])
                .expect("Range.Item(2)"),
        );
        let second_row = expect_object_handle(
            runtime
                .dispatch_invoke(rows, "Item", &[OmValue::Number(2.0)])
                .expect("Rows.Item(2)"),
        );
        let second_column = expect_object_handle(
            runtime
                .dispatch_invoke(columns, "Item", &[OmValue::Number(2.0)])
                .expect("Columns.Item(2)"),
        );
        let single_row_range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1:B1".to_string())])
                .expect("Range(A1:B1)"),
        );
        let single_row_item = expect_object_handle(
            runtime
                .dispatch_invoke(single_row_range, "Item", &[OmValue::Number(2.0)])
                .expect("Range(A1:B1).Item(2)"),
        );

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(second_cell, "Address", &[])
                    .expect("Range.Item(2, 2).Address")
            ),
            "$B$2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(linear_second_cell, "Address", &[])
                    .expect("Range.Item(2).Address")
            ),
            "$B$1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(second_row, "Address", &[])
                    .expect("Rows.Item(2).Address")
            ),
            "$A$2:$B$2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(second_column, "Address", &[])
                    .expect("Columns.Item(2).Address")
            ),
            "$B$1:$B$2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(single_row_item, "Address", &[])
                    .expect("Range(A1:B1).Item(2).Address")
            ),
            "$B$1"
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(second_cell, "Count", &[])
                    .expect("Range.Item(2, 2).Count")
            ),
            1.0
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(second_row, "Count", &[])
                    .expect("Rows.Item(2).Count")
            ),
            2.0
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(second_column, "Count", &[])
                    .expect("Columns.Item(2).Count")
            ),
            2.0
        );

        assert_eq!(
            runtime
                .dispatch_invoke(range, "Item", &[OmValue::Number(5.0)])
                .expect_err("Range.Item(5) should be out of bounds")
                .code,
            OmErrorCode::InvalidArgument
        );
        assert_eq!(
            runtime
                .dispatch_invoke(rows, "Item", &[OmValue::Number(3.0)])
                .expect_err("Rows.Item(3) should be out of bounds")
                .code,
            OmErrorCode::InvalidArgument
        );
    }

    #[test]
    fn range_dispatch_current_region_expands_contiguous_region_and_preserves_blank_seed() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let formula_cell = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("B1".to_string())])
                .expect("Range(B1)"),
        );
        let partial_region = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1:B1".to_string())])
                .expect("Range(A1:B1)"),
        );
        let formula_current_region = expect_object_handle(
            runtime
                .dispatch_get(formula_cell, "CurrentRegion", &[])
                .expect("B1.CurrentRegion"),
        );
        let partial_current_region = expect_object_handle(
            runtime
                .dispatch_get(partial_region, "CurrentRegion", &[])
                .expect("A1:B1.CurrentRegion"),
        );

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(formula_current_region, "Address", &[])
                    .expect("B1.CurrentRegion.Address")
            ),
            "$A$1:$C$1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(partial_current_region, "Address", &[])
                    .expect("A1:B1.CurrentRegion.Address")
            ),
            "$A$1:$C$1"
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(formula_current_region, "Count", &[])
                    .expect("B1.CurrentRegion.Count")
            ),
            3.0
        );

        let blank_workbook = runtime.create_workbook().expect("blank workbook");
        let blank_worksheets = expect_object_handle(
            runtime
                .dispatch_get(blank_workbook.0, "Worksheets", &[])
                .expect("blank worksheets"),
        );
        let blank_sheet = expect_object_handle(
            runtime
                .dispatch_invoke(blank_worksheets, "Item", &[OmValue::Number(1.0)])
                .expect("blank worksheet"),
        );
        let blank_seed = expect_object_handle(
            runtime
                .dispatch_invoke(blank_sheet, "Range", &[OmValue::Text("B2".to_string())])
                .expect("blank Range(B2)"),
        );
        let blank_current_region = expect_object_handle(
            runtime
                .dispatch_get(blank_seed, "CurrentRegion", &[])
                .expect("blank B2.CurrentRegion"),
        );

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(blank_current_region, "Address", &[])
                    .expect("blank B2.CurrentRegion.Address")
            ),
            "$B$2"
        );
    }

    #[test]
    fn range_dispatch_cells_restores_cell_view_over_row_and_column_projections() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1:B2".to_string())])
                .expect("Range(A1:B2)"),
        );
        let rows = expect_object_handle(
            runtime
                .dispatch_get(range, "Rows", &[])
                .expect("Range.Rows"),
        );
        let columns = expect_object_handle(
            runtime
                .dispatch_get(range, "Columns", &[])
                .expect("Range.Columns"),
        );
        let rows_cells = expect_object_handle(
            runtime
                .dispatch_get(rows, "Cells", &[])
                .expect("Rows.Cells"),
        );
        let columns_cells = expect_object_handle(
            runtime
                .dispatch_get(columns, "Cells", &[])
                .expect("Columns.Cells"),
        );
        let rows_cells_item = expect_object_handle(
            runtime
                .dispatch_invoke(rows_cells, "Item", &[OmValue::Number(3.0)])
                .expect("Rows.Cells.Item(3)"),
        );

        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(rows, "Count", &[])
                    .expect("Rows.Count")
            ),
            2.0
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(rows_cells, "Count", &[])
                    .expect("Rows.Cells.Count")
            ),
            4.0
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(columns_cells, "Count", &[])
                    .expect("Columns.Cells.Count")
            ),
            4.0
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(rows_cells, "Address", &[])
                    .expect("Rows.Cells.Address")
            ),
            "$A$1:$B$2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(columns_cells, "Address", &[])
                    .expect("Columns.Cells.Address")
            ),
            "$A$1:$B$2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(rows_cells_item, "Address", &[])
                    .expect("Rows.Cells.Item(3).Address")
            ),
            "$A$2"
        );
    }

    #[test]
    fn range_offset_and_resize_dispatch_transform_rects_and_preserve_projection_views() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("B2:C3".to_string())])
                .expect("Range(B2:C3)"),
        );
        let rows = expect_object_handle(
            runtime
                .dispatch_get(range, "Rows", &[])
                .expect("Range.Rows"),
        );
        let unchanged = expect_object_handle(
            runtime
                .dispatch_invoke(range, "Offset", &[])
                .expect("Range.Offset()"),
        );
        let shifted = expect_object_handle(
            runtime
                .dispatch_invoke(
                    range,
                    "Offset",
                    &[OmValue::Number(-1.0), OmValue::Number(-1.0)],
                )
                .expect("Range.Offset(-1, -1)"),
        );
        let column_shifted = expect_object_handle(
            runtime
                .dispatch_invoke(range, "Offset", &[OmValue::Missing, OmValue::Number(1.0)])
                .expect("Range.Offset(, 1)"),
        );
        let rows_shifted = expect_object_handle(
            runtime
                .dispatch_invoke(rows, "Offset", &[OmValue::Number(1.0)])
                .expect("Rows.Offset(1)"),
        );
        let resized = expect_object_handle(
            runtime
                .dispatch_invoke(range, "Resize", &[OmValue::Number(3.0)])
                .expect("Range.Resize(3)"),
        );
        let narrowed = expect_object_handle(
            runtime
                .dispatch_invoke(range, "Resize", &[OmValue::Missing, OmValue::Number(1.0)])
                .expect("Range.Resize(, 1)"),
        );
        let rows_resized = expect_object_handle(
            runtime
                .dispatch_invoke(rows, "Resize", &[OmValue::Number(3.0)])
                .expect("Rows.Resize(3)"),
        );

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(unchanged, "Address", &[])
                    .expect("Range.Offset().Address")
            ),
            "$B$2:$C$3"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(shifted, "Address", &[])
                    .expect("Range.Offset(-1, -1).Address")
            ),
            "$A$1:$B$2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(column_shifted, "Address", &[])
                    .expect("Range.Offset(, 1).Address")
            ),
            "$C$2:$D$3"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(rows_shifted, "Address", &[])
                    .expect("Rows.Offset(1).Address")
            ),
            "$B$3:$C$4"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(resized, "Address", &[])
                    .expect("Range.Resize(3).Address")
            ),
            "$B$2:$C$4"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(narrowed, "Address", &[])
                    .expect("Range.Resize(, 1).Address")
            ),
            "$B$2:$B$3"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(rows_resized, "Address", &[])
                    .expect("Rows.Resize(3).Address")
            ),
            "$B$2:$C$4"
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(rows_shifted, "Count", &[])
                    .expect("Rows.Offset(1).Count")
            ),
            2.0
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(resized, "Count", &[])
                    .expect("Range.Resize(3).Count")
            ),
            6.0
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(rows_resized, "Count", &[])
                    .expect("Rows.Resize(3).Count")
            ),
            3.0
        );
        assert_eq!(
            runtime
                .dispatch_invoke(range, "Offset", &[OmValue::Number(-2.0)])
                .expect_err("Range.Offset(-2) should be out of bounds")
                .code,
            OmErrorCode::InvalidArgument
        );
        assert_eq!(
            runtime
                .dispatch_invoke(range, "Resize", &[OmValue::Number(0.0)])
                .expect_err("Range.Resize(0) should be rejected")
                .code,
            OmErrorCode::InvalidArgument
        );
    }

    #[test]
    fn range_dispatch_has_formula_reports_uniform_and_mixed_formula_state() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let formula_cell = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("B1".to_string())])
                .expect("Range(B1)"),
        );
        let plain_cell = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1".to_string())])
                .expect("Range(A1)"),
        );
        let mixed_row = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1:C1".to_string())])
                .expect("Range(A1:C1)"),
        );

        assert!(expect_bool(
            runtime
                .dispatch_get(formula_cell, "HasFormula", &[])
                .expect("B1.HasFormula")
        ));
        assert!(!expect_bool(
            runtime
                .dispatch_get(plain_cell, "HasFormula", &[])
                .expect("A1.HasFormula")
        ));
        assert!(matches!(
            runtime
                .dispatch_get(mixed_row, "HasFormula", &[])
                .expect("A1:C1.HasFormula"),
            OmValue::Null
        ));
    }

    #[test]
    fn range_dispatch_text_returns_display_text_for_scalar_and_uniform_ranges() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let number_cell = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1".to_string())])
                .expect("Range(A1)"),
        );
        let formula_cell = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("B1".to_string())])
                .expect("Range(B1)"),
        );
        let mixed_row = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1:C1".to_string())])
                .expect("Range(A1:C1)"),
        );
        let uniform_range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A2:B2".to_string())])
                .expect("Range(A2:B2)"),
        );
        let blank_cell = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("C3".to_string())])
                .expect("Range(C3)"),
        );

        runtime
            .dispatch_set(
                uniform_range,
                "Value2",
                OmValue::Text("same".to_string()),
                &[],
            )
            .expect("Range(A2:B2).Value2");

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(number_cell, "Text", &[])
                    .expect("A1.Text")
            ),
            "42"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(formula_cell, "Text", &[])
                    .expect("B1.Text")
            ),
            "SHARED"
        );
        assert!(matches!(
            runtime
                .dispatch_get(mixed_row, "Text", &[])
                .expect("A1:C1.Text"),
            OmValue::Null
        ));
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(uniform_range, "Text", &[])
                    .expect("A2:B2.Text")
            ),
            "same"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(blank_cell, "Text", &[])
                    .expect("C3.Text")
            ),
            ""
        );
    }

    #[test]
    fn range_address_dispatch_accepts_optional_absolute_flags() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1:B2".to_string())])
                .expect("Range(A1:B2)"),
        );

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(range, "Address", &[])
                    .expect("Address()")
            ),
            "$A$1:$B$2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(range, "Address", &[OmValue::Bool(false)])
                    .expect("Address(false)")
            ),
            "$A1:$B2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(
                        range,
                        "Address",
                        &[OmValue::Bool(false), OmValue::Bool(false)]
                    )
                    .expect("Address(false, false)")
            ),
            "A1:B2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(range, "Address", &[OmValue::Missing, OmValue::Bool(false)])
                    .expect("Address(, false)")
            ),
            "A$1:B$2"
        );
        assert_eq!(
            runtime
                .dispatch_get(range, "Address", &[OmValue::Number(1.0)])
                .expect_err("Address(1) should be rejected")
                .code,
            OmErrorCode::TypeMismatch
        );
        assert_eq!(
            runtime
                .dispatch_get(
                    range,
                    "Address",
                    &[
                        OmValue::Bool(true),
                        OmValue::Bool(true),
                        OmValue::Bool(true)
                    ],
                )
                .expect_err("Address with too many args should be rejected")
                .code,
            OmErrorCode::InvalidArgument
        );
    }

    #[test]
    fn range_dispatch_entire_row_and_column_expand_to_full_sheet_axes() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("B2:C3".to_string())])
                .expect("Range(B2:C3)"),
        );
        let entire_row = expect_object_handle(
            runtime
                .dispatch_get(range, "EntireRow", &[])
                .expect("Range.EntireRow"),
        );
        let entire_column = expect_object_handle(
            runtime
                .dispatch_get(range, "EntireColumn", &[])
                .expect("Range.EntireColumn"),
        );
        let entire_row_rows = expect_object_handle(
            runtime
                .dispatch_get(entire_row, "Rows", &[])
                .expect("EntireRow.Rows"),
        );
        let entire_column_columns = expect_object_handle(
            runtime
                .dispatch_get(entire_column, "Columns", &[])
                .expect("EntireColumn.Columns"),
        );

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(entire_row, "Address", &[])
                    .expect("EntireRow.Address")
            ),
            "$A$2:$XFD$3"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(entire_column, "Address", &[])
                    .expect("EntireColumn.Address")
            ),
            "$B$1:$C$1048576"
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(entire_row, "Count", &[])
                    .expect("EntireRow.Count")
            ),
            32768.0
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(entire_column, "Count", &[])
                    .expect("EntireColumn.Count")
            ),
            2097152.0
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(entire_row_rows, "Count", &[])
                    .expect("EntireRow.Rows.Count")
            ),
            2.0
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(entire_column_columns, "Count", &[])
                    .expect("EntireColumn.Columns.Count")
            ),
            2.0
        );
    }

    #[test]
    fn worksheet_rows_and_columns_dispatch_expose_lazy_axis_handles() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let rows = expect_object_handle(
            runtime
                .dispatch_get(active_sheet, "Rows", &[])
                .expect("Rows"),
        );
        let second_row = expect_object_handle(
            runtime
                .dispatch_get(active_sheet, "Rows", &[OmValue::Number(2.0)])
                .expect("Rows(2)"),
        );
        let rows_item = expect_object_handle(
            runtime
                .dispatch_invoke(rows, "Item", &[OmValue::Number(2.0)])
                .expect("Rows.Item(2)"),
        );
        let text_rows = expect_object_handle(
            runtime
                .dispatch_get(active_sheet, "Rows", &[OmValue::Text("2:3".to_string())])
                .expect("Rows(\"2:3\")"),
        );
        let columns = expect_object_handle(
            runtime
                .dispatch_get(active_sheet, "Columns", &[])
                .expect("Columns"),
        );
        let second_column = expect_object_handle(
            runtime
                .dispatch_get(active_sheet, "Columns", &[OmValue::Number(2.0)])
                .expect("Columns(2)"),
        );
        let text_column = expect_object_handle(
            runtime
                .dispatch_get(active_sheet, "Columns", &[OmValue::Text("B".to_string())])
                .expect("Columns(\"B\")"),
        );
        let text_columns = expect_object_handle(
            runtime
                .dispatch_get(active_sheet, "Columns", &[OmValue::Text("B:C".to_string())])
                .expect("Columns(\"B:C\")"),
        );
        let columns_item = expect_object_handle(
            runtime
                .dispatch_invoke(columns, "Item", &[OmValue::Number(2.0)])
                .expect("Columns.Item(2)"),
        );

        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(rows, "Count", &[])
                    .expect("Rows.Count")
            ),
            1048576.0
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(columns, "Count", &[])
                    .expect("Columns.Count")
            ),
            16384.0
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(second_row, "Address", &[])
                    .expect("Rows(2).Address")
            ),
            "$A$2:$XFD$2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(rows_item, "Address", &[])
                    .expect("Rows.Item(2).Address")
            ),
            "$A$2:$XFD$2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(text_rows, "Address", &[])
                    .expect("Rows(\"2:3\").Address")
            ),
            "$A$2:$XFD$3"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(second_column, "Address", &[])
                    .expect("Columns(2).Address")
            ),
            "$B$1:$B$1048576"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(text_column, "Address", &[])
                    .expect("Columns(\"B\").Address")
            ),
            "$B$1:$B$1048576"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(text_columns, "Address", &[])
                    .expect("Columns(\"B:C\").Address")
            ),
            "$B$1:$C$1048576"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(columns_item, "Address", &[])
                    .expect("Columns.Item(2).Address")
            ),
            "$B$1:$B$1048576"
        );
        assert_eq!(
            runtime
                .dispatch_get(active_sheet, "Rows", &[OmValue::Number(0.0)])
                .expect_err("Rows(0) should be rejected")
                .code,
            OmErrorCode::InvalidArgument
        );
        assert_eq!(
            runtime
                .dispatch_get(active_sheet, "Columns", &[OmValue::Text("XFE".to_string())])
                .expect_err("Columns(\"XFE\") should be rejected")
                .code,
            OmErrorCode::InvalidArgument
        );
        assert_eq!(
            runtime
                .dispatch_get(active_sheet, "Rows", &[OmValue::Text("0:1".to_string())])
                .expect_err("Rows(\"0:1\") should be rejected")
                .code,
            OmErrorCode::InvalidArgument
        );
    }

    #[test]
    fn worksheet_cells_dispatch_supports_full_sheet_cell_view_and_indexed_cells() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let full_cells = expect_object_handle(
            runtime
                .dispatch_get(active_sheet, "Cells", &[])
                .expect("Cells"),
        );
        let indexed_cell = expect_object_handle(
            runtime
                .dispatch_get(
                    active_sheet,
                    "Cells",
                    &[OmValue::Number(2.0), OmValue::Number(2.0)],
                )
                .expect("Cells(2, 2)"),
        );
        let full_cells_item = expect_object_handle(
            runtime
                .dispatch_invoke(full_cells, "Item", &[OmValue::Number(2.0)])
                .expect("Cells.Item(2)"),
        );
        let full_cells_rows = expect_object_handle(
            runtime
                .dispatch_get(full_cells, "Rows", &[])
                .expect("Cells.Rows"),
        );
        let full_cells_columns = expect_object_handle(
            runtime
                .dispatch_get(full_cells, "Columns", &[])
                .expect("Cells.Columns"),
        );

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(full_cells, "Address", &[])
                    .expect("Cells.Address")
            ),
            "$A$1:$XFD$1048576"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(full_cells_item, "Address", &[])
                    .expect("Cells.Item(2).Address")
            ),
            "$B$1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(indexed_cell, "Address", &[])
                    .expect("Cells(2, 2).Address")
            ),
            "$B$2"
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(full_cells_rows, "Count", &[])
                    .expect("Cells.Rows.Count")
            ),
            1048576.0
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(full_cells_columns, "Count", &[])
                    .expect("Cells.Columns.Count")
            ),
            16384.0
        );
    }

    #[test]
    fn worksheet_cells_dispatch_accepts_excel_column_labels() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let get_cell = expect_object_handle(
            runtime
                .dispatch_get(
                    active_sheet,
                    "Cells",
                    &[OmValue::Number(2.0), OmValue::Text("B".to_string())],
                )
                .expect("Cells(2, \"B\")"),
        );
        let invoke_cell = expect_object_handle(
            runtime
                .dispatch_invoke(
                    active_sheet,
                    "Cells",
                    &[OmValue::Number(3.0), OmValue::Text("$C".to_string())],
                )
                .expect("Cells(3, \"$C\")"),
        );

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(get_cell, "Address", &[])
                    .expect("Cells(2, \"B\").Address")
            ),
            "$B$2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(invoke_cell, "Address", &[])
                    .expect("Cells(3, \"$C\").Address")
            ),
            "$C$3"
        );
        assert_eq!(
            runtime
                .dispatch_get(
                    active_sheet,
                    "Cells",
                    &[OmValue::Number(1.0), OmValue::Text("XFE".to_string())],
                )
                .expect_err("Cells(1, \"XFE\") should be rejected")
                .code,
            OmErrorCode::InvalidArgument
        );
    }

    #[test]
    fn collection_item_dispatch_reports_argument_and_lookup_errors() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let application = runtime.root_application();
        let workbooks = expect_object_handle(
            runtime
                .dispatch_get(application, "Workbooks", &[])
                .expect("Workbooks"),
        );
        let active_workbook = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveWorkbook", &[])
                .expect("ActiveWorkbook"),
        );
        let worksheets = expect_object_handle(
            runtime
                .dispatch_get(active_workbook, "Worksheets", &[])
                .expect("Worksheets"),
        );

        assert_eq!(
            runtime
                .dispatch_invoke(workbooks, "Item", &[OmValue::Bool(true)])
                .expect_err("type mismatch")
                .code,
            OmErrorCode::TypeMismatch
        );
        assert_eq!(
            runtime
                .dispatch_invoke(workbooks, "Item", &[OmValue::Number(0.0)])
                .expect_err("invalid index")
                .code,
            OmErrorCode::InvalidArgument
        );
        assert_eq!(
            runtime
                .dispatch_invoke(workbooks, "Item", &[OmValue::Text("Missing".to_string())])
                .expect_err("missing workbook")
                .code,
            OmErrorCode::NotFound
        );
        assert_eq!(
            runtime
                .dispatch_invoke(worksheets, "Item", &[])
                .expect_err("missing worksheet arg")
                .code,
            OmErrorCode::InvalidArgument
        );
        assert_eq!(
            runtime
                .dispatch_invoke(worksheets, "Item", &[OmValue::Text("Missing".to_string())])
                .expect_err("missing worksheet")
                .code,
            OmErrorCode::NotFound
        );
    }

    #[test]
    fn workbooks_add_dispatch_creates_blank_unsaved_workbook_and_updates_active_context() {
        let mut runtime = ExcelRuntime::new();
        let unique = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("system clock")
            .as_nanos();
        let base_dir = std::env::temp_dir().join(format!("ootd-workbooks-add-{unique}"));
        fs::create_dir_all(&base_dir).expect("create add fixture dir");
        let target_path = base_dir.join("created.xlsx");

        let application = runtime.root_application();
        let workbooks = expect_object_handle(
            runtime
                .dispatch_get(application, "Workbooks", &[])
                .expect("Workbooks"),
        );
        let workbook = expect_object_handle(
            runtime
                .dispatch_invoke(workbooks, "Add", &[])
                .expect("Workbooks.Add"),
        );
        let active_workbook = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveWorkbook", &[])
                .expect("ActiveWorkbook"),
        );
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let used_range = expect_object_handle(
            runtime
                .dispatch_get(active_sheet, "UsedRange", &[])
                .expect("UsedRange"),
        );

        assert_eq!(workbook, active_workbook);
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(workbooks, "Count", &[])
                    .expect("Workbooks.Count")
            ),
            1.0
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(workbook, "Name", &[])
                    .expect("Workbook.Name")
            ),
            "Book1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(workbook, "FullName", &[])
                    .expect("Workbook.FullName")
            ),
            "Book1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(used_range, "Address", &[])
                    .expect("UsedRange.Address")
            ),
            "$A$1"
        );
        assert!(expect_bool(
            runtime
                .dispatch_get(workbook, "Saved", &[])
                .expect("Workbook.Saved on add")
        ));

        runtime
            .dispatch_set(
                used_range,
                "Value",
                OmValue::Text("created".to_string()),
                &[],
            )
            .expect("Range.Value");
        assert!(!expect_bool(
            runtime
                .dispatch_get(workbook, "Saved", &[])
                .expect("Workbook.Saved after edit")
        ));

        runtime
            .dispatch_invoke(
                workbook,
                "SaveAs",
                &[OmValue::Text(target_path.to_string_lossy().into_owned())],
            )
            .expect("Workbook.SaveAs");
        assert!(expect_bool(
            runtime
                .dispatch_get(workbook, "Saved", &[])
                .expect("Workbook.Saved after SaveAs")
        ));

        let reopened = ExcelRuntime::new()
            .codec
            .load(
                &fs::read(&target_path).expect("read created workbook"),
                LoadOptions::default(),
            )
            .expect("reload created workbook");
        assert_eq!(reopened.state.worksheets[0].name, "Sheet1");
        assert_eq!(
            reopened
                .state
                .cell(reopened.state.worksheets[0].id, 1, 1)
                .map(|cell| cell.value.clone()),
            Some(CellValue::Text("created".to_string()))
        );

        fs::remove_dir_all(&base_dir).expect("cleanup add fixture");
    }

    #[test]
    fn workbook_dispatch_methods_validate_argument_counts() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");

        assert_eq!(
            runtime
                .dispatch_invoke(workbook.0, "Save", &[OmValue::Bool(true)])
                .expect_err("Save args")
                .code,
            OmErrorCode::InvalidArgument
        );
        assert_eq!(
            runtime
                .dispatch_invoke(
                    workbook.0,
                    "Close",
                    &[OmValue::Bool(true), OmValue::Bool(false)],
                )
                .expect_err("Close args")
                .code,
            OmErrorCode::InvalidArgument
        );
    }

    #[test]
    fn range_dispatch_get_and_set_reject_index_arguments() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1".to_string())])
                .expect("Range(A1)"),
        );

        assert_eq!(
            runtime
                .dispatch_get(range, "Value2", &[OmValue::Number(1.0)])
                .expect_err("Range.Value2 args")
                .code,
            OmErrorCode::InvalidArgument
        );
        assert_eq!(
            runtime
                .dispatch_set(
                    range,
                    "Value2",
                    OmValue::Text("edited".to_string()),
                    &[OmValue::Number(1.0)],
                )
                .expect_err("Range.Value2 set args")
                .code,
            OmErrorCode::InvalidArgument
        );
    }

    #[test]
    fn range_dispatch_scalar_value2_broadcasts_over_multi_cell_rect() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let sheet_id = runtime.worksheets(workbook).expect("worksheets")[0].id;
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1:B2".to_string())])
                .expect("Range(A1:B2)"),
        );

        runtime
            .dispatch_set(range, "Value2", OmValue::Text("fill".to_string()), &[])
            .expect("set scalar Value2");

        let OmValue::Array(array) = runtime
            .dispatch_get(range, "Value2", &[])
            .expect("get scalar-broadcast array")
        else {
            panic!("expected OmValue::Array");
        };

        assert_eq!(array.rows, 2);
        assert_eq!(array.cols, 2);
        assert_eq!(array.values, vec![OmValue::Text("fill".to_string()); 4]);

        let updated = runtime
            .get_range_values(GetRangeValuesSpec {
                workbook,
                range: RangeRef::single_rect(
                    WorkbookId(1),
                    sheet_id,
                    Rect {
                        row_first: 1,
                        row_last: 2,
                        col_first: 1,
                        col_last: 2,
                    },
                ),
            })
            .expect("updated range");
        assert_eq!(updated.values, array.values);
    }

    #[test]
    fn worksheet_dispatch_rejects_invalid_range_and_cells_arguments() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );

        let invalid_range = runtime.dispatch_invoke(
            active_sheet,
            "Range",
            &[OmValue::Text("A1:B2:C3".to_string())],
        );
        assert_eq!(
            invalid_range.expect_err("invalid range should fail").code,
            OmErrorCode::Parse
        );

        let invalid_cells_type = runtime.dispatch_invoke(
            active_sheet,
            "Cells",
            &[OmValue::Text("1".to_string()), OmValue::Number(1.0)],
        );
        assert_eq!(
            invalid_cells_type
                .expect_err("non-numeric Cells row should fail")
                .code,
            OmErrorCode::TypeMismatch
        );

        let invalid_cells_index = runtime.dispatch_invoke(
            active_sheet,
            "Cells",
            &[OmValue::Number(0.0), OmValue::Number(1.0)],
        );
        assert_eq!(
            invalid_cells_index
                .expect_err("zero-based Cells row should fail")
                .code,
            OmErrorCode::InvalidArgument
        );
    }

    #[test]
    fn worksheet_range_and_cells_dispatch_round_trip_value2_and_address() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let application = runtime.root_application();
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );

        let range_a1 = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1".to_string())])
                .expect("Range(A1)"),
        );
        let cell_a1 = expect_object_handle(
            runtime
                .dispatch_invoke(
                    active_sheet,
                    "Cells",
                    &[OmValue::Number(1.0), OmValue::Number(1.0)],
                )
                .expect("Cells(1, 1)"),
        );

        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(range_a1, "Value2", &[])
                    .expect("A1 Value2")
            ),
            42.0
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(range_a1, "Address", &[])
                    .expect("A1 Address")
            ),
            "$A$1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(cell_a1, "Address", &[])
                    .expect("Cells(1, 1) Address")
            ),
            "$A$1"
        );

        runtime
            .dispatch_set(
                range_a1,
                "Value2",
                OmValue::Text("dispatch-edited".to_string()),
                &[],
            )
            .expect("set A1 Value2");

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(range_a1, "Value2", &[])
                    .expect("edited A1 Value2")
            ),
            "dispatch-edited"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(cell_a1, "Value2", &[])
                    .expect("Cells(1, 1) Value2")
            ),
            "dispatch-edited"
        );
    }

    #[test]
    fn worksheet_range_dispatch_accepts_range_handles_and_mixed_variant_endpoints() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let application = runtime.root_application();
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let active_workbook = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveWorkbook", &[])
                .expect("ActiveWorkbook"),
        );
        let cell_a1 = expect_object_handle(
            runtime
                .dispatch_invoke(
                    active_sheet,
                    "Cells",
                    &[OmValue::Number(1.0), OmValue::Number(1.0)],
                )
                .expect("Cells(1, 1)"),
        );
        let cell_b2 = expect_object_handle(
            runtime
                .dispatch_invoke(
                    active_sheet,
                    "Cells",
                    &[OmValue::Number(2.0), OmValue::Number(2.0)],
                )
                .expect("Cells(2, 2)"),
        );
        let header_range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1:B1".to_string())])
                .expect("Range(A1:B1)"),
        );
        let object_range = expect_object_handle(
            runtime
                .dispatch_invoke(
                    active_sheet,
                    "Range",
                    &[OmValue::Object(cell_b2), OmValue::Object(cell_a1)],
                )
                .expect("Range(cell_b2, cell_a1)"),
        );
        let single_object_range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Object(header_range)])
                .expect("Range(header_range)"),
        );
        let mixed_range = expect_object_handle(
            runtime
                .dispatch_invoke(
                    active_sheet,
                    "Range",
                    &[OmValue::Object(header_range), OmValue::Object(cell_b2)],
                )
                .expect("Range(header_range, cell_b2)"),
        );
        let text_object_range = expect_object_handle(
            runtime
                .dispatch_invoke(
                    active_sheet,
                    "Range",
                    &[OmValue::Text("A1".to_string()), OmValue::Object(cell_b2)],
                )
                .expect("Range(\"A1\", cell_b2)"),
        );
        let second_workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open second workbook");
        let foreign_worksheets = expect_object_handle(
            runtime
                .dispatch_get(second_workbook.0, "Worksheets", &[])
                .expect("second workbook worksheets"),
        );
        let foreign_sheet = expect_object_handle(
            runtime
                .dispatch_invoke(foreign_worksheets, "Item", &[OmValue::Number(1.0)])
                .expect("second workbook worksheet"),
        );
        let foreign_cell = expect_object_handle(
            runtime
                .dispatch_invoke(
                    foreign_sheet,
                    "Cells",
                    &[OmValue::Number(1.0), OmValue::Number(1.0)],
                )
                .expect("foreign Cells(1, 1)"),
        );

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(object_range, "Address", &[])
                    .expect("object_range Address")
            ),
            "$A$1:$B$2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(single_object_range, "Address", &[])
                    .expect("single_object_range Address")
            ),
            "$A$1:$B$1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(mixed_range, "Address", &[])
                    .expect("mixed_range Address")
            ),
            "$A$1:$B$2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(text_object_range, "Address", &[])
                    .expect("text_object_range Address")
            ),
            "$A$1:$B$2"
        );
        assert_eq!(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Object(active_workbook)])
                .expect_err("Range(workbook) should fail")
                .code,
            OmErrorCode::TypeMismatch
        );
        assert_eq!(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Object(foreign_cell)])
                .expect_err("Range(foreign_cell) should fail")
                .code,
            OmErrorCode::InvalidArgument
        );
    }

    #[test]
    fn range_dispatch_round_trips_array_value2() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let sheet_id = runtime.worksheets(workbook).expect("worksheets")[0].id;
        let application = runtime.root_application();
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1:B2".to_string())])
                .expect("Range(A1:B2)"),
        );

        runtime
            .dispatch_set(
                range,
                "Value2",
                OmValue::Array(
                    OmArray::new(
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
                ),
                &[],
            )
            .expect("set array Value2");

        let value = runtime
            .dispatch_get(range, "Value2", &[])
            .expect("get array Value2");
        let OmValue::Array(array) = value else {
            panic!("expected OmValue::Array");
        };

        assert_eq!(array.rows, 2);
        assert_eq!(array.cols, 2);
        assert_eq!(array.values[0], OmValue::Number(1.0));
        assert_eq!(array.values[1], OmValue::Text("two".to_string()));
        assert_eq!(array.values[2], OmValue::Bool(true));
        assert_eq!(array.values[3], OmValue::Empty);

        let updated = runtime
            .get_range_values(GetRangeValuesSpec {
                workbook,
                range: RangeRef::single_rect(
                    WorkbookId(1),
                    sheet_id,
                    Rect {
                        row_first: 1,
                        row_last: 2,
                        col_first: 1,
                        col_last: 2,
                    },
                ),
            })
            .expect("updated range");
        assert_eq!(updated.values, array.values);
    }

    #[test]
    fn range_dispatch_value_alias_matches_value2_semantics() {
        let mut runtime = ExcelRuntime::new();
        runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1".to_string())])
                .expect("Range(A1)"),
        );

        runtime
            .dispatch_set(range, "Value", OmValue::Text("alias".to_string()), &[])
            .expect("set Value");
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(range, "Value2", &[])
                    .expect("Value2 after Value set")
            ),
            "alias"
        );

        runtime
            .dispatch_set(range, "Value2", OmValue::Text("updated".to_string()), &[])
            .expect("set Value2");
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(range, "Value", &[])
                    .expect("Value after Value2 set")
            ),
            "updated"
        );
    }

    #[test]
    fn range_dispatch_formula_reads_formula_text_and_persists_formula_sets() {
        let mut runtime = ExcelRuntime::new();
        let unique = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("system clock")
            .as_nanos();
        let base_dir = std::env::temp_dir().join(format!("ootd-range-formula-{unique}"));
        fs::create_dir_all(&base_dir).expect("create formula fixture dir");
        let source_path = base_dir.join("source.xlsx");
        let target_path = base_dir.join("formula-target.xlsx");
        fs::write(&source_path, synthetic_workbook_bytes()).expect("write source workbook");

        let workbooks = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "Workbooks", &[])
                .expect("Workbooks"),
        );
        let workbook = expect_object_handle(
            runtime
                .dispatch_invoke(
                    workbooks,
                    "Open",
                    &[OmValue::Text(source_path.to_string_lossy().into_owned())],
                )
                .expect("Workbooks.Open"),
        );
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let formula_cell = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("B1".to_string())])
                .expect("Range(B1)"),
        );
        let constant_cell = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1".to_string())])
                .expect("Range(A1)"),
        );
        let new_formula_cell = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("C2".to_string())])
                .expect("Range(C2)"),
        );

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(formula_cell, "Formula", &[])
                    .expect("B1 Formula")
            ),
            r#"=UPPER("shared")"#
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(constant_cell, "Formula", &[])
                    .expect("A1 Formula")
            ),
            42.0
        );

        runtime
            .dispatch_set(
                new_formula_cell,
                "Formula",
                OmValue::Text("=SUM(A1:B1)".to_string()),
                &[],
            )
            .expect("set C2 Formula");
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(new_formula_cell, "Formula", &[])
                    .expect("C2 Formula")
            ),
            "=SUM(A1:B1)"
        );
        assert_eq!(
            runtime
                .dispatch_get(new_formula_cell, "Value", &[])
                .expect("C2 Value after Formula"),
            OmValue::Empty
        );
        assert!(!expect_bool(
            runtime
                .dispatch_get(workbook, "Saved", &[])
                .expect("Workbook.Saved after Formula set")
        ));

        runtime
            .dispatch_invoke(
                workbook,
                "SaveAs",
                &[OmValue::Text(target_path.to_string_lossy().into_owned())],
            )
            .expect("Workbook.SaveAs");

        let reopened = ExcelRuntime::new()
            .codec
            .load(
                &fs::read(&target_path).expect("read formula workbook"),
                LoadOptions::default(),
            )
            .expect("reload formula workbook");
        let reopened_sheet_id = reopened.state.worksheets[0].id;
        let reopened_formula = reopened
            .state
            .cell(reopened_sheet_id, 2, 3)
            .expect("C2 reopened");
        assert_eq!(
            reopened_formula.formula.as_ref().expect("C2 formula").text,
            "SUM(A1:B1)"
        );
        assert_eq!(reopened_formula.value, CellValue::Blank);

        fs::remove_dir_all(&base_dir).expect("cleanup formula fixture");
    }

    #[test]
    fn range_clear_contents_clears_values_and_formulas_without_changing_selection() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let application = runtime.root_application();
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1:C1".to_string())])
                .expect("Range(A1:C1)"),
        );
        let formula_cell = expect_object_handle(
            runtime
                .dispatch_invoke(range, "Item", &[OmValue::Number(1.0), OmValue::Number(2.0)])
                .expect("Range(A1:C1).Item(1, 2)"),
        );

        assert!(matches!(
            runtime
                .dispatch_invoke(range, "ClearContents", &[])
                .expect("Range.ClearContents"),
            OmValue::Empty
        ));

        let selection = expect_object_handle(
            runtime
                .dispatch_get(application, "Selection", &[])
                .expect("Selection after ClearContents"),
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(selection, "Address", &[])
                    .expect("Selection address after ClearContents")
            ),
            "$A$1:$C$1"
        );
        assert_eq!(
            runtime
                .dispatch_get(range, "Text", &[])
                .expect("Range(A1:C1).Text after ClearContents"),
            OmValue::Text(String::new())
        );
        assert_eq!(
            runtime
                .dispatch_get(formula_cell, "Value", &[])
                .expect("B1 Value after ClearContents"),
            OmValue::Empty
        );
        assert_eq!(
            runtime
                .dispatch_get(formula_cell, "Formula", &[])
                .expect("B1 Formula after ClearContents"),
            OmValue::Empty
        );
        assert_eq!(
            runtime
                .dispatch_get(formula_cell, "HasFormula", &[])
                .expect("B1 HasFormula after ClearContents"),
            OmValue::Bool(false)
        );
        assert!(!expect_bool(
            runtime
                .dispatch_get(workbook.0, "Saved", &[])
                .expect("Workbook.Saved after ClearContents")
        ));
        assert_eq!(
            runtime
                .dispatch_invoke(range, "ClearContents", &[OmValue::Bool(true)])
                .expect_err("Range.ClearContents args should be rejected")
                .code,
            OmErrorCode::InvalidArgument
        );

        let saved = runtime
            .save_workbook(
                workbook,
                SaveWorkbookSpec {
                    format: FileFormat::Xlsx,
                    profile: ExcelProfile::Excel365,
                    lossless: true,
                },
            )
            .expect("save workbook after ClearContents");
        let reopened = ExcelRuntime::new()
            .codec
            .load(&saved, LoadOptions::default())
            .expect("reload cleared workbook");
        let sheet_id = reopened.state.worksheets[0].id;
        assert!(reopened.state.cell(sheet_id, 1, 1).is_none());
        assert!(reopened.state.cell(sheet_id, 1, 2).is_none());
        assert!(reopened.state.cell(sheet_id, 1, 3).is_none());
    }

    #[test]
    fn application_range_dispatch_targets_the_active_sheet() {
        let mut runtime = ExcelRuntime::new();
        let workbook1 = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook1");
        let workbook2 = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook2");
        let worksheets1 = expect_object_handle(
            runtime
                .dispatch_get(workbook1.0, "Worksheets", &[])
                .expect("Workbook1.Worksheets"),
        );
        let worksheet1 = expect_object_handle(
            runtime
                .dispatch_invoke(worksheets1, "Item", &[OmValue::Number(1.0)])
                .expect("Workbook1.Worksheets.Item(1)"),
        );
        let worksheets2 = expect_object_handle(
            runtime
                .dispatch_get(workbook2.0, "Worksheets", &[])
                .expect("Workbook2.Worksheets"),
        );
        let worksheet2 = expect_object_handle(
            runtime
                .dispatch_invoke(worksheets2, "Item", &[OmValue::Number(1.0)])
                .expect("Workbook2.Worksheets.Item(1)"),
        );

        runtime
            .dispatch_set(
                worksheet1,
                "Name",
                OmValue::Text("FirstSheet".to_string()),
                &[],
            )
            .expect("rename workbook1 sheet");
        runtime
            .dispatch_set(
                worksheet2,
                "Name",
                OmValue::Text("SecondSheet".to_string()),
                &[],
            )
            .expect("rename workbook2 sheet");
        runtime
            .dispatch_invoke(worksheet1, "Activate", &[])
            .expect("Worksheet.Activate");

        let application = runtime.root_application();
        let application_range = expect_object_handle(
            runtime
                .dispatch_invoke(application, "Range", &[OmValue::Text("B1:C1".to_string())])
                .expect("Application.Range(B1:C1)"),
        );
        let active_workbook = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveWorkbook", &[])
                .expect("ActiveWorkbook after Application.Range"),
        );
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet after Application.Range"),
        );
        let active_cell = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveCell", &[])
                .expect("ActiveCell after Application.Range"),
        );
        let selection = expect_object_handle(
            runtime
                .dispatch_get(application, "Selection", &[])
                .expect("Selection after Application.Range"),
        );
        let range_parent = expect_object_handle(
            runtime
                .dispatch_get(application_range, "Parent", &[])
                .expect("Application.Range parent"),
        );

        assert_eq!(active_workbook, workbook1.0);
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(application_range, "Address", &[])
                    .expect("Application.Range address")
            ),
            "$B$1:$C$1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_sheet, "Name", &[])
                    .expect("ActiveSheet after Application.Range name")
            ),
            "FirstSheet"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(range_parent, "Name", &[])
                    .expect("Application.Range parent name")
            ),
            "FirstSheet"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(selection, "Address", &[])
                    .expect("Selection after Application.Range address")
            ),
            "$B$1:$C$1"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_cell, "Address", &[])
                    .expect("ActiveCell after Application.Range address")
            ),
            "$B$1"
        );

        let mut empty_runtime = ExcelRuntime::new();
        assert_eq!(
            empty_runtime
                .dispatch_invoke(
                    empty_runtime.root_application(),
                    "Range",
                    &[OmValue::Text("A1".to_string())],
                )
                .expect_err("Application.Range without active workbook")
                .code,
            OmErrorCode::InvalidState
        );
    }

    #[test]
    fn application_cells_dispatch_targets_the_active_sheet_and_supports_index_args() {
        let mut runtime = ExcelRuntime::new();
        let workbook1 = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook1");
        let workbook2 = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook2");
        let worksheets1 = expect_object_handle(
            runtime
                .dispatch_get(workbook1.0, "Worksheets", &[])
                .expect("Workbook1.Worksheets"),
        );
        let worksheet1 = expect_object_handle(
            runtime
                .dispatch_invoke(worksheets1, "Item", &[OmValue::Number(1.0)])
                .expect("Workbook1.Worksheets.Item(1)"),
        );
        let worksheets2 = expect_object_handle(
            runtime
                .dispatch_get(workbook2.0, "Worksheets", &[])
                .expect("Workbook2.Worksheets"),
        );
        let worksheet2 = expect_object_handle(
            runtime
                .dispatch_invoke(worksheets2, "Item", &[OmValue::Number(1.0)])
                .expect("Workbook2.Worksheets.Item(1)"),
        );

        runtime
            .dispatch_set(
                worksheet1,
                "Name",
                OmValue::Text("FirstSheet".to_string()),
                &[],
            )
            .expect("rename workbook1 sheet");
        runtime
            .dispatch_set(
                worksheet2,
                "Name",
                OmValue::Text("SecondSheet".to_string()),
                &[],
            )
            .expect("rename workbook2 sheet");
        runtime
            .dispatch_invoke(worksheet1, "Activate", &[])
            .expect("Worksheet.Activate");

        let application = runtime.root_application();
        let full_cells = expect_object_handle(
            runtime
                .dispatch_get(application, "Cells", &[])
                .expect("Application.Cells"),
        );
        let indexed_cell = expect_object_handle(
            runtime
                .dispatch_get(
                    application,
                    "Cells",
                    &[OmValue::Number(2.0), OmValue::Text("B".to_string())],
                )
                .expect("Application.Cells(2, \"B\")"),
        );
        let active_workbook = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveWorkbook", &[])
                .expect("ActiveWorkbook after Application.Cells"),
        );
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet after Application.Cells"),
        );
        let active_cell = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveCell", &[])
                .expect("ActiveCell after Application.Cells"),
        );
        let selection = expect_object_handle(
            runtime
                .dispatch_get(application, "Selection", &[])
                .expect("Selection after Application.Cells"),
        );

        assert_eq!(active_workbook, workbook1.0);
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(full_cells, "Address", &[])
                    .expect("Application.Cells address")
            ),
            "$A$1:$XFD$1048576"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(indexed_cell, "Address", &[])
                    .expect("Application.Cells(2, \"B\") address")
            ),
            "$B$2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_sheet, "Name", &[])
                    .expect("ActiveSheet after Application.Cells name")
            ),
            "FirstSheet"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(selection, "Address", &[])
                    .expect("Selection after Application.Cells address")
            ),
            "$B$2"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_cell, "Address", &[])
                    .expect("ActiveCell after Application.Cells address")
            ),
            "$B$2"
        );

        let mut empty_runtime = ExcelRuntime::new();
        let empty_application = empty_runtime.root_application();
        assert_eq!(
            empty_runtime
                .dispatch_get(empty_application, "Cells", &[])
                .expect("Application.Cells without active workbook"),
            OmValue::Empty
        );
    }

    #[test]
    fn workbook_close_rejects_stale_object_handles() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let application = runtime.root_application();
        let active_workbook = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveWorkbook", &[])
                .expect("ActiveWorkbook"),
        );
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let range_a1 = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1".to_string())])
                .expect("Range(A1)"),
        );

        runtime
            .dispatch_invoke(active_workbook, "Save", &[])
            .expect("Save");
        runtime
            .dispatch_invoke(active_workbook, "Close", &[])
            .expect("Close");

        assert_eq!(
            runtime
                .dispatch_get(application, "ActiveWorkbook", &[])
                .expect("empty ActiveWorkbook"),
            OmValue::Empty
        );
        assert_eq!(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("empty ActiveSheet"),
            OmValue::Empty
        );
        assert_eq!(
            runtime
                .dispatch_get(application, "ActiveCell", &[])
                .expect("empty ActiveCell"),
            OmValue::Empty
        );
        assert_eq!(
            runtime
                .dispatch_get(application, "Selection", &[])
                .expect("empty Selection"),
            OmValue::Empty
        );

        let workbook_name = runtime.dispatch_get(active_workbook, "Name", &[]);
        assert_eq!(
            workbook_name.expect_err("stale workbook should fail").code,
            OmErrorCode::InvalidState
        );

        let sheet_name = runtime.dispatch_get(active_sheet, "Name", &[]);
        assert_eq!(
            sheet_name.expect_err("stale worksheet should fail").code,
            OmErrorCode::InvalidState
        );

        let range_value = runtime.dispatch_get(range_a1, "Value2", &[]);
        assert_eq!(
            range_value.expect_err("stale range should fail").code,
            OmErrorCode::InvalidState
        );

        assert!(runtime.workbook_model(workbook).is_err());
    }

    #[test]
    fn unsupported_or_missing_member_is_rejected() {
        let mut runtime = ExcelRuntime::new();
        let _workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let application = runtime.root_application();
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );

        let unsupported = runtime.dispatch_get(active_sheet, "Range", &[]);
        assert_eq!(
            unsupported
                .expect_err("unsupported surface should fail")
                .code,
            OmErrorCode::Unsupported
        );

        let missing = runtime.dispatch_get(active_sheet, "DefinitelyNotAMember", &[]);
        assert_eq!(
            missing.expect_err("missing member should fail").code,
            OmErrorCode::NotFound
        );
    }

    #[test]
    fn workbooks_open_dispatch_reads_workbook_from_filesystem() {
        let mut runtime = ExcelRuntime::new();
        let path = std::env::temp_dir().join(format!(
            "ootd-runtime-open-{}-{}.xlsx",
            std::process::id(),
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .expect("time")
                .as_nanos()
        ));
        fs::write(&path, synthetic_workbook_bytes()).expect("write workbook fixture");

        let workbooks = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "Workbooks", &[])
                .expect("Workbooks"),
        );
        let opened = expect_object_handle(
            runtime
                .dispatch_invoke(
                    workbooks,
                    "Open",
                    &[OmValue::Text(path.to_string_lossy().into_owned())],
                )
                .expect("Workbooks.Open"),
        );

        let active_workbook = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveWorkbook", &[])
                .expect("ActiveWorkbook"),
        );
        let workbook_item = expect_object_handle(
            runtime
                .dispatch_invoke(
                    workbooks,
                    "Item",
                    &[OmValue::Text(
                        path.file_name()
                            .and_then(|value| value.to_str())
                            .expect("filename")
                            .to_string(),
                    )],
                )
                .expect("Workbooks.Item(filename)"),
        );

        assert_eq!(opened, active_workbook);
        assert_eq!(opened, workbook_item);

        fs::remove_file(path).expect("cleanup fixture");
    }

    #[test]
    fn workbook_save_dispatch_persists_changes_back_to_opened_path() {
        let mut runtime = ExcelRuntime::new();
        let path = std::env::temp_dir().join(format!(
            "ootd-runtime-save-{}-{}.xlsx",
            std::process::id(),
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .expect("time")
                .as_nanos()
        ));
        fs::write(&path, synthetic_workbook_bytes()).expect("write workbook fixture");

        let workbooks = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "Workbooks", &[])
                .expect("Workbooks"),
        );
        let workbook = expect_object_handle(
            runtime
                .dispatch_invoke(
                    workbooks,
                    "Open",
                    &[OmValue::Text(path.to_string_lossy().into_owned())],
                )
                .expect("Workbooks.Open"),
        );
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );

        runtime
            .dispatch_set(
                active_sheet,
                "Name",
                OmValue::Text("SavedName".to_string()),
                &[],
            )
            .expect("rename worksheet");
        runtime
            .dispatch_invoke(workbook, "Save", &[])
            .expect("Workbook.Save");

        let reopened = ExcelRuntime::new()
            .codec
            .load(
                &fs::read(&path).expect("read saved workbook"),
                office_common::LoadOptions::default(),
            )
            .expect("reload saved workbook");
        assert_eq!(reopened.state.worksheets[0].name, "SavedName");

        fs::remove_file(path).expect("cleanup fixture");
    }

    #[test]
    fn worksheet_name_dispatch_renames_sheet_and_persists_on_save() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let application = runtime.root_application();
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );

        runtime
            .dispatch_set(
                active_sheet,
                "Name",
                OmValue::Text("Renamed".to_string()),
                &[],
            )
            .expect("rename worksheet");

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_sheet, "Name", &[])
                    .expect("renamed sheet name")
            ),
            "Renamed"
        );

        let worksheets = expect_object_handle(
            runtime
                .dispatch_get(workbook.0, "Worksheets", &[])
                .expect("Worksheets"),
        );
        let renamed_sheet = expect_object_handle(
            runtime
                .dispatch_invoke(worksheets, "Item", &[OmValue::Text("Renamed".to_string())])
                .expect("Worksheets.Item(Renamed)"),
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(renamed_sheet, "Name", &[])
                    .expect("renamed sheet lookup")
            ),
            "Renamed"
        );

        let saved = runtime
            .save_workbook(
                workbook,
                SaveWorkbookSpec {
                    format: FileFormat::Xlsx,
                    profile: ExcelProfile::Excel365,
                    lossless: true,
                },
            )
            .expect("save workbook");
        let reopened = ExcelRuntime::new()
            .codec
            .load(&saved, office_common::LoadOptions::default())
            .expect("reopen saved workbook");

        assert_eq!(reopened.state.worksheets[0].name, "Renamed");
    }

    #[test]
    fn worksheet_name_dispatch_rejects_invalid_names_and_read_only_workbooks() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: true,
            })
            .expect("open workbook");
        let application = runtime.root_application();
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );

        let read_only_error = runtime
            .dispatch_set(
                active_sheet,
                "Name",
                OmValue::Text("Blocked".to_string()),
                &[],
            )
            .expect_err("read-only rename should fail");
        assert_eq!(read_only_error.code, OmErrorCode::InvalidState);
        assert!(expect_bool(
            runtime
                .dispatch_get(workbook.0, "ReadOnly", &[])
                .expect("Workbook.ReadOnly")
        ));

        runtime
            .close_workbook(workbook)
            .expect("close read-only workbook");

        let writable = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open writable workbook");
        let writable_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );

        let invalid_name = runtime
            .dispatch_set(
                writable_sheet,
                "Name",
                OmValue::Text("Bad/Name".to_string()),
                &[],
            )
            .expect_err("invalid rename should fail");
        assert_eq!(invalid_name.code, OmErrorCode::InvalidArgument);

        runtime
            .close_workbook(writable)
            .expect("close writable workbook");
    }

    #[test]
    fn workbook_name_and_worksheet_index_dispatch_report_runtime_metadata() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let application = runtime.root_application();
        let active_workbook = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveWorkbook", &[])
                .expect("ActiveWorkbook"),
        );
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );

        assert_eq!(workbook.0, active_workbook);
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_workbook, "Name", &[])
                    .expect("Workbook.Name")
            ),
            "Workbook"
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_workbook, "Path", &[])
                    .expect("Workbook.Path")
            ),
            ""
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(active_workbook, "FullName", &[])
                    .expect("Workbook.FullName")
            ),
            "Workbook"
        );
        assert!(!expect_bool(
            runtime
                .dispatch_get(active_workbook, "ReadOnly", &[])
                .expect("Workbook.ReadOnly")
        ));
        assert!(expect_bool(
            runtime
                .dispatch_get(active_workbook, "Saved", &[])
                .expect("Workbook.Saved")
        ));
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(active_sheet, "Index", &[])
                    .expect("Worksheet.Index")
            ),
            1.0
        );
    }

    #[test]
    fn parent_dispatch_walks_back_up_object_graph() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let application = runtime.root_application();
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(application, "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let range = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("B2".to_string())])
                .expect("Range(B2)"),
        );

        assert_eq!(
            expect_object_handle(
                runtime
                    .dispatch_get(workbook.0, "Parent", &[])
                    .expect("Workbook.Parent")
            ),
            application
        );
        assert_eq!(
            expect_object_handle(
                runtime
                    .dispatch_get(active_sheet, "Parent", &[])
                    .expect("Worksheet.Parent")
            ),
            workbook.0
        );
        let range_parent = expect_object_handle(
            runtime
                .dispatch_get(range, "Parent", &[])
                .expect("Range.Parent"),
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(range_parent, "Name", &[])
                    .expect("Range.Parent.Name")
            ),
            expect_text(
                runtime
                    .dispatch_get(active_sheet, "Name", &[])
                    .expect("ActiveSheet.Name")
            )
        );
        assert_eq!(
            expect_object_handle(
                runtime
                    .dispatch_get(range_parent, "Parent", &[])
                    .expect("Range.Parent.Parent")
            ),
            workbook.0
        );
    }

    #[test]
    fn workbook_save_as_dispatch_writes_new_path_and_updates_workbook_name() {
        let mut runtime = ExcelRuntime::new();
        let unique = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("system clock")
            .as_nanos();
        let base_dir = std::env::temp_dir().join(format!("ootd-saveas-{unique}"));
        let source_dir = base_dir.join("source");
        let target_dir = base_dir.join("target");
        fs::create_dir_all(&source_dir).expect("create source dir");
        fs::create_dir_all(&target_dir).expect("create target dir");
        let source_path = source_dir.join("source.xlsx");
        let target_path = target_dir.join("target.xlsx");
        fs::write(&source_path, synthetic_workbook_bytes()).expect("write source workbook");

        let workbooks = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "Workbooks", &[])
                .expect("Workbooks"),
        );
        let workbook = expect_object_handle(
            runtime
                .dispatch_invoke(
                    workbooks,
                    "Open",
                    &[OmValue::Text(source_path.to_string_lossy().into_owned())],
                )
                .expect("Workbooks.Open"),
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(workbook, "Path", &[])
                    .expect("Workbook.Path after Open")
            ),
            source_dir.to_string_lossy()
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(workbook, "FullName", &[])
                    .expect("Workbook.FullName after Open")
            ),
            source_path.to_string_lossy()
        );
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );

        runtime
            .dispatch_set(
                active_sheet,
                "Name",
                OmValue::Text("SavedAsName".to_string()),
                &[],
            )
            .expect("rename worksheet before SaveAs");
        runtime
            .dispatch_invoke(
                workbook,
                "SaveAs",
                &[OmValue::Text(target_path.to_string_lossy().into_owned())],
            )
            .expect("Workbook.SaveAs");

        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(workbook, "Name", &[])
                    .expect("Workbook.Name after SaveAs")
            ),
            target_path
                .file_name()
                .and_then(|value| value.to_str())
                .expect("target file name")
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(workbook, "Path", &[])
                    .expect("Workbook.Path after SaveAs")
            ),
            target_dir.to_string_lossy()
        );
        assert_eq!(
            expect_text(
                runtime
                    .dispatch_get(workbook, "FullName", &[])
                    .expect("Workbook.FullName after SaveAs")
            ),
            target_path.to_string_lossy()
        );

        let reopened_after_save_as = ExcelRuntime::new()
            .codec
            .load(
                &fs::read(&target_path).expect("read SaveAs target"),
                office_common::LoadOptions::default(),
            )
            .expect("reload SaveAs target");
        assert_eq!(
            reopened_after_save_as.state.worksheets[0].name,
            "SavedAsName"
        );

        runtime
            .dispatch_set(
                active_sheet,
                "Name",
                OmValue::Text("SavedAfterSaveAs".to_string()),
                &[],
            )
            .expect("rename worksheet after SaveAs");
        runtime
            .dispatch_invoke(workbook, "Save", &[])
            .expect("Workbook.Save after SaveAs");

        let reopened_after_save = ExcelRuntime::new()
            .codec
            .load(
                &fs::read(&target_path).expect("read target after Save"),
                office_common::LoadOptions::default(),
            )
            .expect("reload target after Save");
        assert_eq!(
            reopened_after_save.state.worksheets[0].name,
            "SavedAfterSaveAs"
        );

        fs::remove_dir_all(&base_dir).expect("cleanup SaveAs fixture");
    }

    #[test]
    fn workbook_saved_dispatch_tracks_dirty_state_across_mutations_and_saves() {
        let mut runtime = ExcelRuntime::new();
        let unique = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("system clock")
            .as_nanos();
        let base_dir = std::env::temp_dir().join(format!("ootd-saved-state-{unique}"));
        let source_dir = base_dir.join("source");
        let target_dir = base_dir.join("target");
        fs::create_dir_all(&source_dir).expect("create source dir");
        fs::create_dir_all(&target_dir).expect("create target dir");
        let source_path = source_dir.join("source.xlsx");
        let target_path = target_dir.join("saved-target.xlsx");
        fs::write(&source_path, synthetic_workbook_bytes()).expect("write source workbook");

        let workbooks = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "Workbooks", &[])
                .expect("Workbooks"),
        );
        let workbook = expect_object_handle(
            runtime
                .dispatch_invoke(
                    workbooks,
                    "Open",
                    &[OmValue::Text(source_path.to_string_lossy().into_owned())],
                )
                .expect("Workbooks.Open"),
        );
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );
        let first_cell = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1".to_string())])
                .expect("Range(A1)"),
        );

        assert!(expect_bool(
            runtime
                .dispatch_get(workbook, "Saved", &[])
                .expect("Workbook.Saved on open")
        ));

        runtime
            .dispatch_set(workbook, "Saved", OmValue::Bool(false), &[])
            .expect("Workbook.Saved = false");
        assert!(!expect_bool(
            runtime
                .dispatch_get(workbook, "Saved", &[])
                .expect("Workbook.Saved after false")
        ));

        runtime
            .dispatch_set(workbook, "Saved", OmValue::Bool(true), &[])
            .expect("Workbook.Saved = true");
        assert!(expect_bool(
            runtime
                .dispatch_get(workbook, "Saved", &[])
                .expect("Workbook.Saved after true")
        ));

        runtime
            .dispatch_set(
                active_sheet,
                "Name",
                OmValue::Text("DirtyRename".to_string()),
                &[],
            )
            .expect("rename worksheet");
        assert!(!expect_bool(
            runtime
                .dispatch_get(workbook, "Saved", &[])
                .expect("Workbook.Saved after rename")
        ));

        runtime
            .dispatch_set(workbook, "Saved", OmValue::Bool(true), &[])
            .expect("Workbook.Saved = true after rename");
        assert!(expect_bool(
            runtime
                .dispatch_get(workbook, "Saved", &[])
                .expect("Workbook.Saved after reset")
        ));

        runtime
            .dispatch_set(
                first_cell,
                "Value",
                OmValue::Text("dirty value".to_string()),
                &[],
            )
            .expect("Range.Value");
        assert!(!expect_bool(
            runtime
                .dispatch_get(workbook, "Saved", &[])
                .expect("Workbook.Saved after cell edit")
        ));

        runtime
            .dispatch_invoke(workbook, "Save", &[])
            .expect("Workbook.Save");
        assert!(expect_bool(
            runtime
                .dispatch_get(workbook, "Saved", &[])
                .expect("Workbook.Saved after Save")
        ));

        runtime
            .dispatch_set(
                active_sheet,
                "Name",
                OmValue::Text("SavedAfterSaveAs".to_string()),
                &[],
            )
            .expect("rename worksheet before SaveAs");
        assert!(!expect_bool(
            runtime
                .dispatch_get(workbook, "Saved", &[])
                .expect("Workbook.Saved before SaveAs")
        ));

        runtime
            .dispatch_invoke(
                workbook,
                "SaveAs",
                &[OmValue::Text(target_path.to_string_lossy().into_owned())],
            )
            .expect("Workbook.SaveAs");
        assert!(expect_bool(
            runtime
                .dispatch_get(workbook, "Saved", &[])
                .expect("Workbook.Saved after SaveAs")
        ));

        fs::remove_dir_all(&base_dir).expect("cleanup Saved fixture");
    }

    #[test]
    fn workbook_close_dispatch_with_save_changes_persists_and_invalidates_handle() {
        let mut runtime = ExcelRuntime::new();
        let unique = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("system clock")
            .as_nanos();
        let source_path =
            std::env::temp_dir().join(format!("ootd-close-save-changes-{unique}.xlsx"));
        fs::write(&source_path, synthetic_workbook_bytes()).expect("write source workbook");

        let workbooks = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "Workbooks", &[])
                .expect("Workbooks"),
        );
        let workbook = expect_object_handle(
            runtime
                .dispatch_invoke(
                    workbooks,
                    "Open",
                    &[OmValue::Text(source_path.to_string_lossy().into_owned())],
                )
                .expect("Workbooks.Open"),
        );
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );

        runtime
            .dispatch_set(
                active_sheet,
                "Name",
                OmValue::Text("ClosedSaved".to_string()),
                &[],
            )
            .expect("rename worksheet before close");
        runtime
            .dispatch_invoke(workbook, "Close", &[OmValue::Bool(true)])
            .expect("Workbook.Close(true)");

        let stale_error = runtime
            .dispatch_get(workbook, "Name", &[])
            .expect_err("closed workbook handle should be stale");
        assert_eq!(stale_error.code, OmErrorCode::InvalidState);

        let reopened = ExcelRuntime::new()
            .codec
            .load(
                &fs::read(&source_path).expect("read closed workbook"),
                office_common::LoadOptions::default(),
            )
            .expect("reload closed workbook");
        assert_eq!(reopened.state.worksheets[0].name, "ClosedSaved");

        fs::remove_file(&source_path).expect("cleanup source fixture");
    }

    #[test]
    fn workbook_close_dispatch_without_save_changes_leaves_source_file_untouched() {
        let mut runtime = ExcelRuntime::new();
        let unique = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("system clock")
            .as_nanos();
        let source_path =
            std::env::temp_dir().join(format!("ootd-close-without-save-{unique}.xlsx"));
        fs::write(&source_path, synthetic_workbook_bytes()).expect("write source workbook");

        let workbooks = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "Workbooks", &[])
                .expect("Workbooks"),
        );
        let workbook = expect_object_handle(
            runtime
                .dispatch_invoke(
                    workbooks,
                    "Open",
                    &[OmValue::Text(source_path.to_string_lossy().into_owned())],
                )
                .expect("Workbooks.Open"),
        );
        let active_sheet = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
                .expect("ActiveSheet"),
        );

        runtime
            .dispatch_set(
                active_sheet,
                "Name",
                OmValue::Text("DiscardedRename".to_string()),
                &[],
            )
            .expect("rename worksheet before close");
        runtime
            .dispatch_invoke(workbook, "Close", &[OmValue::Bool(false)])
            .expect("Workbook.Close(false)");

        let stale_error = runtime
            .dispatch_get(workbook, "Name", &[])
            .expect_err("closed workbook handle should be stale");
        assert_eq!(stale_error.code, OmErrorCode::InvalidState);

        let reopened = ExcelRuntime::new()
            .codec
            .load(
                &fs::read(&source_path).expect("read closed workbook"),
                office_common::LoadOptions::default(),
            )
            .expect("reload closed workbook");
        assert_eq!(reopened.state.worksheets[0].name, "Sheet1");

        fs::remove_file(&source_path).expect("cleanup source fixture");
    }

    #[test]
    fn workbooks_and_worksheets_count_dispatch_report_collection_sizes() {
        let mut runtime = ExcelRuntime::new();
        let first = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open first workbook");
        let second = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open second workbook");
        let workbooks = expect_object_handle(
            runtime
                .dispatch_get(runtime.root_application(), "Workbooks", &[])
                .expect("Workbooks"),
        );
        let worksheets = expect_object_handle(
            runtime
                .dispatch_get(first.0, "Worksheets", &[])
                .expect("Workbook.Worksheets"),
        );

        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(workbooks, "Count", &[])
                    .expect("Workbooks.Count")
            ),
            2.0
        );
        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(worksheets, "Count", &[])
                    .expect("Worksheets.Count")
            ),
            1.0
        );

        runtime
            .close_workbook(second)
            .expect("close second workbook");

        assert_eq!(
            expect_number(
                runtime
                    .dispatch_get(workbooks, "Count", &[])
                    .expect("Workbooks.Count after close")
            ),
            1.0
        );
    }

    #[test]
    fn collection_parent_dispatch_returns_owning_objects() {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open workbook");
        let application = runtime.root_application();
        let workbooks = expect_object_handle(
            runtime
                .dispatch_get(application, "Workbooks", &[])
                .expect("Workbooks"),
        );
        let worksheets = expect_object_handle(
            runtime
                .dispatch_get(workbook.0, "Worksheets", &[])
                .expect("Workbook.Worksheets"),
        );

        assert_eq!(
            expect_object_handle(
                runtime
                    .dispatch_get(workbooks, "Parent", &[])
                    .expect("Workbooks.Parent")
            ),
            application
        );
        assert_eq!(
            expect_object_handle(
                runtime
                    .dispatch_get(worksheets, "Parent", &[])
                    .expect("Worksheets.Parent")
            ),
            workbook.0
        );
    }

    fn synthetic_workbook_bytes() -> Vec<u8> {
        let package = OpcPackage::new(vec![
            OpcPart {
                name: "[Content_Types].xml".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"#
                    .to_vec(),
            },
            OpcPart {
                name: "_rels/.rels".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#
                    .to_vec(),
            },
            OpcPart {
                name: "xl/workbook.xml".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#
                    .to_vec(),
            },
            OpcPart {
                name: "xl/_rels/workbook.xml.rels".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#
                    .to_vec(),
            },
            OpcPart {
                name: "xl/sharedStrings.xml".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><t>shared</t></si>
</sst>"#
                    .to_vec(),
            },
            OpcPart {
                name: "xl/worksheets/sheet1.xml".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C1"/>
  <sheetData>
    <row r="1">
      <c r="A1"><v>42</v></c>
      <c r="B1" t="str"><f>UPPER("shared")</f><v>SHARED</v></c>
      <c r="C1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#
                    .to_vec(),
            },
            OpcPart {
                name: "customXml/item1.xml".to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes: br#"<payload preserved="true"/>"#.to_vec(),
            },
        ]);

        package.to_bytes().expect("package bytes")
    }
}
