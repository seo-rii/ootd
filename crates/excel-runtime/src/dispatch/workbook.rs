use super::super::{
    ExcelRuntime, RuntimeNamesScope, RuntimeObjectKind, RuntimeSheetCollectionKind,
    RuntimeSheetTemplate, XL_OPEN_XML_STRICT_WORKBOOK, XL_OPEN_XML_TEMPLATE,
    XL_OPEN_XML_TEMPLATE_MACRO_ENABLED, XL_OPEN_XML_WORKBOOK, XL_OPEN_XML_WORKBOOK_MACRO_ENABLED,
    file_format_from_path, file_format_to_excel_value, om_value_is_omitted,
    validate_check_spelling_args, validate_export_as_fixed_format_args, validate_optional_bool_arg,
    validate_optional_integer_arg, validate_optional_text_arg, validate_print_out_args,
    validate_print_preview_args,
};
use office_common::{
    ExcelProfile, FileFormat, ObjectHandle, OmError, OmErrorCode, OmResult, OmValue,
    OpenWorkbookSpec, SaveWorkbookSpec, WorkbookHandle,
};
use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};

impl ExcelRuntime {
    pub(crate) fn dispatch_get_workbook(
        &mut self,
        workbook: WorkbookHandle,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Workbook", member, false)?;
        if !args.is_empty() && !matches!(member, "Worksheets" | "Sheets" | "Charts" | "Names") {
            return Err(OmError::invalid_argument(format!(
                "Workbook.{member} does not accept index arguments"
            )));
        }

        match member {
            "Name" => Ok(OmValue::Text(
                self.workbook_model(workbook)?.display_name.clone(),
            )),
            "Parent" => Ok(OmValue::Object(self.root_application())),
            "Application" => Ok(OmValue::Object(self.root_application())),
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
            "FileFormat" => Ok(OmValue::Number(f64::from(file_format_to_excel_value(
                self.workbook_model(workbook)?.format,
            )))),
            "Date1904" => Ok(OmValue::Bool(self.workbook_model(workbook)?.date1904)),
            "IsAddin" => Ok(OmValue::Bool(self.workbook_model(workbook)?.is_addin)),
            "HasVBProject" => Ok(OmValue::Bool(
                self.runtime_workbook(workbook)?
                    .loaded
                    .package
                    .contains("xl/vbaProject.bin"),
            )),
            "ReadOnly" => Ok(OmValue::Bool(self.runtime_workbook(workbook)?.read_only)),
            "Saved" => Ok(OmValue::Bool(
                self.runtime_workbook(workbook)?.saved_for_prompt(),
            )),
            "Worksheets" | "Sheets" | "Charts" => {
                let handle = self.register_object(RuntimeObjectKind::WorksheetsCollection {
                    workbook,
                    kind: match member {
                        "Worksheets" => RuntimeSheetCollectionKind::Worksheets,
                        "Sheets" => RuntimeSheetCollectionKind::Sheets,
                        "Charts" => RuntimeSheetCollectionKind::Charts,
                        _ => unreachable!("sheet collection member"),
                    },
                });
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            "Names" => {
                let handle = self.register_names_handle(workbook, RuntimeNamesScope::Workbook);
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            "ActiveSheet" => {
                let sheet_id = self.active_sheet_id(workbook)?;
                Ok(OmValue::Object(
                    self.register_sheet_object_handle(workbook, sheet_id)?,
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "Workbook.{member} is not implemented as a property"
            ))),
        }
    }

    pub(crate) fn dispatch_get_workbooks(
        &mut self,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Workbooks", member, false)?;

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
            "Application" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Workbooks.Application does not accept arguments",
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

    pub(crate) fn dispatch_invoke_workbook(
        &mut self,
        workbook: WorkbookHandle,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Workbook", member, false)?;

        match member {
            "Activate" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Workbook.Activate does not accept arguments",
                    ));
                }
                let selection = self.default_selection(workbook)?;
                self.set_selection(workbook, selection.sheet_id, selection.rect);
                Ok(OmValue::Empty)
            }
            "Save" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Workbook.Save does not accept arguments",
                    ));
                }
                let runtime = self.runtime_workbook(workbook)?;
                let path = runtime.source_path.clone().ok_or_else(|| {
                    OmError::invalid_state(
                        "Workbook.Save requires a source path; use Workbook.SaveAs",
                    )
                })?;
                if runtime.read_only {
                    return Err(OmError::invalid_state(
                        "Workbook.Save cannot overwrite a read-only workbook; use Workbook.SaveAs or Workbook.SaveCopyAs with a new filename",
                    ));
                }
                let format = self.workbook_model(workbook)?.format;
                let prepared = self.prepare_workbook_save(
                    workbook,
                    SaveWorkbookSpec {
                        format,
                        profile: ExcelProfile::Excel365,
                        lossless: true,
                    },
                )?;
                fs::write(&path, &prepared.bytes).map_err(|error| {
                    OmError::new(
                        OmErrorCode::Io,
                        format!("failed to write workbook {}: {error}", path.display()),
                    )
                })?;
                self.commit_workbook_save_baseline(workbook, prepared.next_loaded, None, None)?;
                Ok(OmValue::Empty)
            }
            "SaveAs" => {
                if args.is_empty() || args.len() > 12 {
                    return Err(OmError::invalid_argument(
                        "Workbook.SaveAs expects Filename and up to 11 optional arguments",
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
                let format = match args.get(1) {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => {
                        file_format_from_path(&path)
                            .unwrap_or(self.workbook_model(workbook)?.format)
                    }
                    Some(OmValue::Number(format)) => {
                        if !format.is_finite()
                            || format.fract() != 0.0
                            || *format < i32::MIN as f64
                            || *format > i32::MAX as f64
                        {
                            return Err(OmError::invalid_argument(
                                "Workbook.SaveAs FileFormat expects an integral XlFileFormat value",
                            ));
                        }
                        match *format as i32 {
                            XL_OPEN_XML_WORKBOOK => FileFormat::Xlsx,
                            XL_OPEN_XML_WORKBOOK_MACRO_ENABLED => FileFormat::Xlsm,
                            XL_OPEN_XML_TEMPLATE => FileFormat::Xltx,
                            XL_OPEN_XML_TEMPLATE_MACRO_ENABLED => FileFormat::Xltm,
                            XL_OPEN_XML_STRICT_WORKBOOK => FileFormat::StrictXlsx,
                            other => {
                                return Err(OmError::unsupported(format!(
                                    "Workbook.SaveAs FileFormat {other} is not implemented",
                                )));
                            }
                        }
                    }
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Workbook.SaveAs FileFormat expects a numeric XlFileFormat value when provided",
                        ));
                    }
                };
                validate_optional_text_arg(args, 2, "Workbook.SaveAs Password")?;
                validate_optional_text_arg(args, 3, "Workbook.SaveAs WriteResPassword")?;
                validate_optional_bool_arg(args, 4, "Workbook.SaveAs ReadOnlyRecommended")?;
                validate_optional_bool_arg(args, 5, "Workbook.SaveAs CreateBackup")?;
                validate_optional_integer_arg(args, 6, "Workbook.SaveAs AccessMode")?;
                validate_optional_integer_arg(args, 7, "Workbook.SaveAs ConflictResolution")?;
                validate_optional_bool_arg(args, 8, "Workbook.SaveAs AddToMru")?;
                validate_optional_integer_arg(args, 9, "Workbook.SaveAs TextCodepage")?;
                validate_optional_integer_arg(args, 10, "Workbook.SaveAs TextVisualLayout")?;
                validate_optional_bool_arg(args, 11, "Workbook.SaveAs Local")?;
                let read_only = self.runtime_workbook(workbook)?.read_only;
                let prepared = self.prepare_workbook_save(
                    workbook,
                    SaveWorkbookSpec {
                        format,
                        profile: ExcelProfile::Excel365,
                        lossless: true,
                    },
                )?;
                if read_only {
                    let write_result = fs::OpenOptions::new()
                        .write(true)
                        .create_new(true)
                        .open(&path)
                        .and_then(|mut file| file.write_all(&prepared.bytes));
                    write_result.map_err(|error| {
                        if error.kind() == std::io::ErrorKind::AlreadyExists {
                            OmError::invalid_state(
                                "Workbook.SaveAs for a read-only workbook requires a new filename",
                            )
                        } else {
                            OmError::new(
                                OmErrorCode::Io,
                                format!(
                                    "failed to write read-only workbook {}: {error}",
                                    path.display()
                                ),
                            )
                        }
                    })?;
                } else {
                    fs::write(&path, &prepared.bytes).map_err(|error| {
                        OmError::new(
                            OmErrorCode::Io,
                            format!("failed to write workbook {}: {error}", path.display()),
                        )
                    })?;
                }
                let display_name = path
                    .file_name()
                    .and_then(|value| value.to_str())
                    .filter(|value| !value.is_empty())
                    .map(str::to_owned);
                self.commit_workbook_save_baseline(
                    workbook,
                    prepared.next_loaded,
                    Some(path),
                    display_name,
                )?;
                if read_only {
                    self.runtime_workbook_mut(workbook)?.read_only = false;
                }
                Ok(OmValue::Empty)
            }
            "SaveCopyAs" => {
                if args.len() != 1 {
                    return Err(OmError::invalid_argument(
                        "Workbook.SaveCopyAs expects a single filename argument",
                    ));
                }
                let path = match &args[0] {
                    OmValue::Text(path) => PathBuf::from(path),
                    _ => {
                        return Err(OmError::type_mismatch(
                            "Workbook.SaveCopyAs expects a string filename",
                        ));
                    }
                };
                let format = self.workbook_model(workbook)?.format;
                let read_only = self.runtime_workbook(workbook)?.read_only;
                let bytes = self.save_workbook(
                    workbook,
                    SaveWorkbookSpec {
                        format,
                        profile: ExcelProfile::Excel365,
                        lossless: true,
                    },
                )?;
                if read_only {
                    let write_result = fs::OpenOptions::new()
                        .write(true)
                        .create_new(true)
                        .open(&path)
                        .and_then(|mut file| file.write_all(&bytes));
                    write_result.map_err(|error| {
                        if error.kind() == std::io::ErrorKind::AlreadyExists {
                            OmError::invalid_state(
                                "Workbook.SaveCopyAs for a read-only workbook requires a new filename",
                            )
                        } else {
                            OmError::new(
                                OmErrorCode::Io,
                                format!(
                                    "failed to write read-only workbook copy {}: {error}",
                                    path.display()
                                ),
                            )
                        }
                    })?;
                } else {
                    fs::write(&path, &bytes).map_err(|error| {
                        OmError::new(
                            OmErrorCode::Io,
                            format!("failed to write workbook copy {}: {error}", path.display()),
                        )
                    })?;
                }
                Ok(OmValue::Empty)
            }
            "RefreshAll" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Workbook.RefreshAll does not accept arguments",
                    ));
                }
                Ok(OmValue::Empty)
            }
            "CheckSpelling" => {
                validate_check_spelling_args(args, "Workbook")?;
                self.runtime_workbook(workbook)?;
                Ok(OmValue::Empty)
            }
            "ExportAsFixedFormat" => {
                validate_export_as_fixed_format_args(args, "Workbook")?;
                self.runtime_workbook(workbook)?;
                Ok(OmValue::Empty)
            }
            "PrintPreview" => {
                validate_print_preview_args(args, "Workbook")?;
                self.runtime_workbook(workbook)?;
                Ok(OmValue::Empty)
            }
            "PrintOut" => {
                validate_print_out_args(args, "Workbook")?;
                self.runtime_workbook(workbook)?;
                Ok(OmValue::Empty)
            }
            "Close" => {
                if args.len() > 3 {
                    return Err(OmError::invalid_argument(
                        "Workbook.Close accepts at most SaveChanges, Filename, and RouteWorkbook arguments",
                    ));
                }
                let save_changes = match args.first() {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => None,
                    Some(OmValue::Bool(save_changes)) => Some(*save_changes),
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Workbook.Close SaveChanges expects a boolean when provided",
                        ));
                    }
                };
                let filename = match args.get(1) {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => None,
                    Some(OmValue::Text(path)) => Some(PathBuf::from(path)),
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Workbook.Close Filename expects a string when provided",
                        ));
                    }
                };
                match args.get(2) {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => {}
                    Some(OmValue::Bool(_)) => {}
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Workbook.Close RouteWorkbook expects a boolean when provided",
                        ));
                    }
                }
                let should_save = match save_changes {
                    Some(save_changes) => save_changes,
                    None => {
                        if self.runtime_workbook(workbook)?.prompt_dirty && self.display_alerts {
                            return Err(OmError::invalid_state(
                                "Workbook.Close requires an explicit SaveChanges value because DisplayAlerts is enabled and no prompt callback is configured",
                            ));
                        }
                        false
                    }
                };
                if should_save {
                    let runtime = self.runtime_workbook(workbook)?;
                    let source_path = runtime.source_path.clone();
                    let read_only = runtime.read_only;
                    if read_only && filename.is_none() {
                        return Err(OmError::invalid_state(
                            "Workbook.Close cannot overwrite a read-only workbook; provide a new Filename",
                        ));
                    }
                    let save_path = filename.clone().or(source_path).ok_or_else(|| {
                        OmError::invalid_state(
                            "Workbook.Close with SaveChanges=true requires a Filename or source path",
                        )
                    })?;
                    let filename_format = file_format_from_path(&save_path);
                    let format = match filename_format {
                        Some(format) => format,
                        None => self.workbook_model(workbook)?.format,
                    };
                    let bytes = self.save_workbook(
                        workbook,
                        SaveWorkbookSpec {
                            format,
                            profile: ExcelProfile::Excel365,
                            lossless: true,
                        },
                    )?;
                    if read_only {
                        let write_result = fs::OpenOptions::new()
                            .write(true)
                            .create_new(true)
                            .open(&save_path)
                            .and_then(|mut file| file.write_all(&bytes));
                        write_result.map_err(|error| {
                            if error.kind() == std::io::ErrorKind::AlreadyExists {
                                OmError::invalid_state(
                                    "Workbook.Close for a read-only workbook requires a new Filename",
                                )
                            } else {
                                OmError::new(
                                    OmErrorCode::Io,
                                    format!(
                                        "failed to write read-only workbook {}: {error}",
                                        save_path.display()
                                    ),
                                )
                            }
                        })?;
                    } else {
                        fs::write(&save_path, &bytes).map_err(|error| {
                            OmError::new(
                                OmErrorCode::Io,
                                format!(
                                    "failed to write workbook {}: {error}",
                                    save_path.display()
                                ),
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

    pub(crate) fn dispatch_invoke_workbooks(
        &mut self,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Workbooks", member, false)?;

        match member {
            "Add" => {
                if args.len() > 1 {
                    return Err(OmError::invalid_argument(
                        "Workbooks.Add accepts at most a single Template argument",
                    ));
                }
                let workbook = match args.first() {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => {
                        self.create_workbook()?
                    }
                    Some(OmValue::Number(template)) => {
                        let template = RuntimeSheetTemplate::from_xl_wba_template(*template)
                            .ok_or_else(|| {
                                OmError::unsupported(
                                    "Workbooks.Add supports numeric XlWBATemplate values xlWBATWorksheet, xlWBATChart, xlWBATExcel4MacroSheet, and xlWBATExcel4IntlMacroSheet",
                                )
                            })?;
                        self.create_workbook_from_template_kind(template)?
                    }
                    Some(OmValue::Text(path)) => {
                        let workbook_name = self.allocate_created_workbook_name();
                        self.open_detached_workbook_from_path(path, Some(workbook_name))?
                    }
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Workbooks.Add Template expects an XlWBATemplate numeric value or template path when provided",
                        ));
                    }
                };
                Ok(OmValue::Object(workbook.0))
            }
            "Open" => {
                if args.is_empty() || args.len() > 15 {
                    return Err(OmError::invalid_argument(
                        "Workbooks.Open expects Filename and up to 14 optional arguments",
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
                match args.get(1) {
                    None => {}
                    Some(value) if om_value_is_omitted(value) => {}
                    Some(OmValue::Bool(_)) => {}
                    Some(OmValue::Number(value)) => {
                        if !value.is_finite()
                            || value.fract() != 0.0
                            || *value < i32::MIN as f64
                            || *value > i32::MAX as f64
                        {
                            return Err(OmError::invalid_argument(
                                "Workbooks.Open UpdateLinks expects an integer numeric value when provided",
                            ));
                        }
                    }
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Workbooks.Open UpdateLinks expects a boolean or numeric value when provided",
                        ));
                    }
                }
                let read_only = match args.get(2) {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => false,
                    Some(OmValue::Bool(read_only)) => *read_only,
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Workbooks.Open ReadOnly expects a boolean value when provided",
                        ));
                    }
                };
                validate_optional_integer_arg(args, 3, "Workbooks.Open Format")?;
                validate_optional_text_arg(args, 4, "Workbooks.Open Password")?;
                validate_optional_text_arg(args, 5, "Workbooks.Open WriteResPassword")?;
                validate_optional_bool_arg(args, 6, "Workbooks.Open IgnoreReadOnlyRecommended")?;
                validate_optional_integer_arg(args, 7, "Workbooks.Open Origin")?;
                validate_optional_text_arg(args, 8, "Workbooks.Open Delimiter")?;
                validate_optional_bool_arg(args, 9, "Workbooks.Open Editable")?;
                validate_optional_bool_arg(args, 10, "Workbooks.Open Notify")?;
                validate_optional_integer_arg(args, 11, "Workbooks.Open Converter")?;
                validate_optional_bool_arg(args, 12, "Workbooks.Open AddToMru")?;
                validate_optional_bool_arg(args, 13, "Workbooks.Open Local")?;
                validate_optional_integer_arg(args, 14, "Workbooks.Open CorruptLoad")?;
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
                            read_only,
                        },
                        display_name,
                        Some(PathBuf::from(path)),
                    )?
                    .0,
                ))
            }
            "Item" => self.resolve_workbook_item(args),
            "Close" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Workbooks.Close does not accept arguments",
                    ));
                }
                let workbook_handles = self
                    .workbooks
                    .keys()
                    .copied()
                    .map(|id| WorkbookHandle(ObjectHandle(id)))
                    .collect::<Vec<_>>();
                for workbook in workbook_handles {
                    self.close_workbook(workbook)?;
                }
                Ok(OmValue::Empty)
            }
            _ => Err(OmError::unsupported(format!(
                "Workbooks.{member} is not implemented as a method"
            ))),
        }
    }
}
