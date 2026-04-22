use excel_model::WorkbookState;
use excel_xlsx::{LoadedXlsxWorkbook, XlsxCodec};
use office_codegen::{OmFocusSurfaceRegistry, build_focus_surface_registry_from_json};
use office_common::{
    FileFormat, GetRangeValuesSpec, LoadOptions, ObjectHandle, OmArray, OmError, OmErrorCode,
    OmResult, OmValue, OpaquePart, OpenWorkbookSpec, RangeHandle, RangeRef, Rect, SaveOptions,
    SaveWorkbookSpec, SetRangeValuesSpec, SheetId, WorkbookHandle, WorkbookId, WorkbookModel,
    WorksheetHandle, WorksheetModel,
};
use office_idl::{AccessMode, SupportState};
use std::collections::{BTreeMap, BTreeSet};

const ROOT_APPLICATION_HANDLE_VALUE: u64 = 0;
const FIRST_DYNAMIC_OBJECT_HANDLE_VALUE: u64 = 1_000_000;
const PINNED_OM_TEMPLATE_JSON: &str = include_str!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../specs/pinned/office_idl_excel_om.template.json"
));

#[derive(Debug)]
struct RuntimeWorkbook {
    loaded: LoadedXlsxWorkbook,
    read_only: bool,
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
    },
}

#[derive(Debug)]
pub struct ExcelRuntime {
    codec: XlsxCodec,
    dispatch_registry: OmFocusSurfaceRegistry,
    root_application: ObjectHandle,
    next_handle: u64,
    next_object_handle: u64,
    active_workbook: Option<WorkbookHandle>,
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
            active_workbook: None,
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

        let handle_value = self.next_handle;
        self.next_handle += 1;
        let workbook_id = WorkbookId(handle_value);
        loaded.state.assign_workbook_id(workbook_id);
        let workbook_handle = WorkbookHandle(ObjectHandle(handle_value));

        self.workbooks.insert(
            handle_value,
            RuntimeWorkbook {
                loaded,
                read_only: spec.read_only,
            },
        );
        self.objects.insert(
            handle_value,
            RuntimeObjectKind::Workbook {
                workbook: workbook_handle,
            },
        );
        self.active_workbook = Some(workbook_handle);

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
            } => self.dispatch_get_range(workbook, sheet_id, rect, member, args),
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
            RuntimeObjectKind::Range {
                workbook,
                sheet_id,
                rect,
            } => self.dispatch_set_range(workbook, sheet_id, rect, member, value, args),
            RuntimeObjectKind::Application
            | RuntimeObjectKind::WorkbooksCollection
            | RuntimeObjectKind::Workbook { .. }
            | RuntimeObjectKind::WorksheetsCollection { .. }
            | RuntimeObjectKind::Worksheet { .. } => Err(OmError::unsupported(format!(
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
            RuntimeObjectKind::Application => Err(OmError::unsupported(format!(
                "member {member} cannot be invoked on Application"
            ))),
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
            RuntimeObjectKind::Range { .. } => Err(OmError::unsupported(format!(
                "member {member} cannot be invoked on Range"
            ))),
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
        if !args.is_empty() {
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
                let sheet_id = self
                    .runtime_workbook(active_workbook)?
                    .loaded
                    .state
                    .worksheets
                    .first()
                    .map(|worksheet| worksheet.id)
                    .ok_or_else(|| {
                        OmError::new(OmErrorCode::NotFound, "workbook has no worksheets")
                    })?;
                Ok(OmValue::Object(
                    self.register_worksheet_handle(active_workbook, sheet_id).0,
                ))
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
                let (row, col) = parse_cells_args(args)?;
                Ok(OmValue::Object(
                    self.register_range_handle(workbook, sheet_id, Rect::single_cell(row, col))
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
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "Range.{member} does not accept arguments"
            )));
        }

        match member {
            "Value2" => {
                let array = self.get_range_values(GetRangeValuesSpec {
                    workbook,
                    range: self.range_ref(workbook, sheet_id, rect)?,
                })?;
                if array.rows == 1 && array.cols == 1 {
                    Ok(array.values.into_iter().next().unwrap_or(OmValue::Empty))
                } else {
                    Ok(OmValue::Array(array))
                }
            }
            "Address" => Ok(OmValue::Text(format_rect_address(rect))),
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
        member: &str,
        value: OmValue,
        args: &[OmValue],
    ) -> OmResult<()> {
        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "Range.{member} does not accept index arguments"
            )));
        }

        match member {
            "Value2" => {
                let values = match value {
                    OmValue::Array(array) => array,
                    scalar => OmArray::new(
                        rect.height() as usize,
                        rect.width() as usize,
                        vec![scalar; rect.height() as usize * rect.width() as usize],
                    )?,
                };
                self.set_range_values(SetRangeValuesSpec {
                    workbook,
                    range: self.range_ref(workbook, sheet_id, rect)?,
                    values,
                })
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
                let _ = self.save_workbook(
                    workbook,
                    SaveWorkbookSpec {
                        format,
                        profile: office_common::ExcelProfile::Excel365,
                        lossless: true,
                    },
                )?;
                Ok(OmValue::Empty)
            }
            "Close" => {
                if args.len() > 1 {
                    return Err(OmError::invalid_argument(
                        "Workbook.Close accepts at most one save_changes argument",
                    ));
                }
                self.close_workbook(workbook)?;
                Ok(OmValue::Empty)
            }
            _ => Err(OmError::unsupported(format!(
                "Workbook.{member} is not implemented as a method"
            ))),
        }
    }

    fn dispatch_invoke_workbooks(&mut self, member: &str, args: &[OmValue]) -> OmResult<OmValue> {
        match member {
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
                let rect = parse_range_args(args)?;
                Ok(OmValue::Object(
                    self.register_range_handle(workbook, sheet_id, rect).0,
                ))
            }
            "Cells" => {
                let (row, col) = parse_cells_args(args)?;
                Ok(OmValue::Object(
                    self.register_range_handle(workbook, sheet_id, Rect::single_cell(row, col))
                        .0,
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
        RangeHandle(self.register_object(RuntimeObjectKind::Range {
            workbook,
            sheet_id,
            rect,
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
}

pub fn supports_format(format: FileFormat) -> bool {
    matches!(
        format,
        FileFormat::Xlsx | FileFormat::Xlsm | FileFormat::Xltx | FileFormat::Xltm
    )
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

fn parse_range_args(args: &[OmValue]) -> OmResult<Rect> {
    match args {
        [OmValue::Text(a1)] => parse_rect_a1(a1),
        [OmValue::Text(start), OmValue::Text(end)] => {
            let start = parse_cell_a1(start)?;
            let end = parse_cell_a1(end)?;
            Ok(Rect {
                row_first: start.0.min(end.0),
                row_last: start.0.max(end.0),
                col_first: start.1.min(end.1),
                col_last: start.1.max(end.1),
            })
        }
        _ => Err(OmError::invalid_argument(
            "Worksheet.Range expects one A1 reference or two A1 cell references",
        )),
    }
}

fn parse_cells_args(args: &[OmValue]) -> OmResult<(u32, u32)> {
    if args.len() != 2 {
        return Err(OmError::invalid_argument(
            "Worksheet.Cells expects row and column arguments",
        ));
    }
    Ok((
        coerce_u32_arg(&args[0], "Worksheet.Cells row")?,
        coerce_u32_arg(&args[1], "Worksheet.Cells column")?,
    ))
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

fn format_rect_address(rect: Rect) -> String {
    if rect.row_first == rect.row_last && rect.col_first == rect.col_last {
        format!("${}$${}", column_to_letters(rect.col_first), rect.row_first).replace("$$", "$")
    } else {
        format!(
            "${}$${}:${}$${}",
            column_to_letters(rect.col_first),
            rect.row_first,
            column_to_letters(rect.col_last),
            rect.row_last
        )
        .replace("$$", "$")
    }
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
    use super::{ExcelRuntime, supports_format};
    use office_common::{
        ExcelProfile, FileFormat, GetRangeValuesSpec, ObjectHandle, OmArray, OmErrorCode, OmValue,
        OpenWorkbookSpec, RangeRef, Rect, SaveWorkbookSpec, SetRangeValuesSpec, WorkbookId,
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

        runtime.close_workbook(second).expect("close second workbook");

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
            expect_text(runtime.dispatch_get(range, "Address", &[]).expect("Address")),
            "$A$1:$B$2"
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
