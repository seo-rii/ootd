use super::*;
use crate::{
    XL_COPY, XL_CUT, XL_PASTE_COLUMN_WIDTHS, XL_PASTE_COMMENTS, XL_PASTE_SPECIAL_OPERATION_ADD,
    XL_PASTE_SPECIAL_OPERATION_DIVIDE, XL_PASTE_VALIDATION, XL_PASTE_VALUES,
};

fn open_clean_workbook(runtime: &mut ExcelRuntime) -> WorkbookHandle {
    runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: synthetic_workbook_bytes(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open synthetic workbook")
}

fn worksheet_handle(runtime: &mut ExcelRuntime, workbook: WorkbookHandle) -> ObjectHandle {
    let worksheets = expect_object_handle(
        runtime
            .dispatch_get(workbook.0, "Worksheets", &[])
            .expect("Workbook.Worksheets"),
    );
    expect_object_handle(
        runtime
            .dispatch_invoke(worksheets, "Item", &[OmValue::Number(1.0)])
            .expect("Worksheets.Item(1)"),
    )
}

fn range_handle(
    runtime: &mut ExcelRuntime,
    worksheet: ObjectHandle,
    address: &str,
) -> ObjectHandle {
    expect_object_handle(
        runtime
            .dispatch_invoke(worksheet, "Range", &[OmValue::Text(address.to_string())])
            .unwrap_or_else(|error| panic!("Range({address}): {error:?}")),
    )
}

fn set_row_values(
    runtime: &mut ExcelRuntime,
    range: ObjectHandle,
    values: Vec<OmValue>,
    label: &str,
) {
    runtime
        .dispatch_set(
            range,
            "Value2",
            OmValue::Array(
                OmArray::new(1, values.len(), values)
                    .unwrap_or_else(|error| panic!("{label}: array: {error:?}")),
            ),
            &[],
        )
        .unwrap_or_else(|error| panic!("{label}: set values: {error:?}"));
}

fn commit_workbook_baseline(runtime: &mut ExcelRuntime, workbook: WorkbookHandle, label: &str) {
    let mut bytes = Vec::new();
    runtime
        .save_workbook_to_writer(
            workbook,
            SaveWorkbookSpec {
                format: FileFormat::Xlsx,
                profile: ExcelProfile::Excel365,
                lossless: true,
            },
            &mut bytes,
        )
        .unwrap_or_else(|error| panic!("{label}: commit baseline: {error:?}"));
    assert!(!bytes.is_empty(), "{label}: baseline bytes");
    assert_eq!(
        runtime
            .workbook_dirty_domains(workbook)
            .unwrap_or_else(|error| panic!("{label}: clean dirty domains: {error:?}")),
        WorkbookDirtyDomains::default(),
        "{label}: clean baseline",
    );
}

fn assert_semantic_dirty_domains(runtime: &ExcelRuntime, workbook: WorkbookHandle, label: &str) {
    assert_eq!(
        runtime
            .workbook_dirty_domains(workbook)
            .unwrap_or_else(|error| panic!("{label}: dirty domains: {error:?}")),
        WorkbookDirtyDomains {
            prompt_dirty: true,
            semantic_dirty: true,
            serialization_dirty: true,
            ..WorkbookDirtyDomains::default()
        },
        "{label}: dirty domains",
    );
}

fn assert_custom_paste_failure_preserves_workbooks_and_session(
    runtime: &mut ExcelRuntime,
    source_workbook: WorkbookHandle,
    destination_workbook: WorkbookHandle,
    source: ObjectHandle,
    destination: ObjectHandle,
    clipboard_member: &str,
    paste_args: &[OmValue],
    expected_code: OmErrorCode,
    expected_message: &str,
    label: &str,
) {
    runtime
        .dispatch_invoke(
            source,
            "Find",
            &[OmValue::Text("custom paste session marker".to_string())],
        )
        .unwrap_or_else(|error| panic!("{label}: seed Find state: {error:?}"));
    runtime
        .dispatch_invoke(source, clipboard_member, &[])
        .unwrap_or_else(|error| panic!("{label}: arm clipboard: {error:?}"));

    let expected_mode = if clipboard_member == "Copy" {
        XL_COPY
    } else {
        XL_CUT
    };
    assert!(
        runtime.find_state.is_some(),
        "{label}: Find state precondition"
    );
    let clipboard = runtime
        .clipboard
        .as_ref()
        .unwrap_or_else(|| panic!("{label}: clipboard precondition"));
    assert_eq!(clipboard.mode, expected_mode, "{label}: clipboard mode");
    assert_eq!(
        clipboard.workbook, source_workbook,
        "{label}: clipboard workbook",
    );
    assert_eq!(
        runtime.cut_copy_mode,
        Some(expected_mode),
        "{label}: CutCopyMode",
    );

    let source_before = runtime_workbook_persistence_snapshot(runtime, source_workbook);
    let destination_before = runtime_workbook_persistence_snapshot(runtime, destination_workbook);
    let source_dirty_before = runtime
        .workbook_dirty_domains(source_workbook)
        .unwrap_or_else(|error| panic!("{label}: source dirty domains: {error:?}"));
    let destination_dirty_before = runtime
        .workbook_dirty_domains(destination_workbook)
        .unwrap_or_else(|error| panic!("{label}: destination dirty domains: {error:?}"));
    let session_before = runtime_session_mutation_snapshot(runtime);

    let error = match runtime.dispatch_invoke(destination, "PasteSpecial", paste_args) {
        Ok(value) => panic!("{label}: custom PasteSpecial succeeded: {value:?}"),
        Err(error) => error,
    };

    assert_eq!(error.code, expected_code, "{label}: {error:?}");
    assert!(
        error.message.contains(expected_message),
        "{label}: {error:?}",
    );
    assert_eq!(
        runtime_workbook_persistence_snapshot(runtime, source_workbook),
        source_before,
        "{label}: source workbook",
    );
    assert_eq!(
        runtime_workbook_persistence_snapshot(runtime, destination_workbook),
        destination_before,
        "{label}: destination workbook",
    );
    assert_eq!(
        runtime
            .workbook_dirty_domains(source_workbook)
            .unwrap_or_else(|error| panic!("{label}: source dirty domains after: {error:?}")),
        source_dirty_before,
        "{label}: source dirty domains",
    );
    assert_eq!(
        runtime
            .workbook_dirty_domains(destination_workbook)
            .unwrap_or_else(|error| panic!("{label}: destination dirty domains after: {error:?}")),
        destination_dirty_before,
        "{label}: destination dirty domains",
    );
    assert_eq!(
        runtime_session_mutation_snapshot(runtime),
        session_before,
        "{label}: session",
    );
}

#[test]
fn range_metadata_only_paste_special_is_fail_closed_and_atomic() {
    for (paste_type, paste_name) in [
        (XL_PASTE_COMMENTS, "xlPasteComments"),
        (XL_PASTE_VALIDATION, "xlPasteValidation"),
        (XL_PASTE_COLUMN_WIDTHS, "xlPasteColumnWidths"),
    ] {
        for clipboard_member in ["Copy", "Cut"] {
            let label = format!("{clipboard_member} {paste_name}");
            let mut runtime = ExcelRuntime::new();
            let workbook = open_clean_workbook(&mut runtime);
            let worksheet = worksheet_handle(&mut runtime, workbook);
            let source = range_handle(&mut runtime, worksheet, "A60:B60");
            let destination = range_handle(&mut runtime, worksheet, "D60:E60");
            set_row_values(
                &mut runtime,
                source,
                vec![OmValue::Number(1.0), OmValue::Number(2.0)],
                &label,
            );
            set_row_values(
                &mut runtime,
                destination,
                vec![OmValue::Number(3.0), OmValue::Number(4.0)],
                &label,
            );
            commit_workbook_baseline(&mut runtime, workbook, &label);

            assert_custom_paste_failure_preserves_workbooks_and_session(
                &mut runtime,
                workbook,
                workbook,
                source,
                destination,
                clipboard_member,
                &[OmValue::Number(f64::from(paste_type))],
                OmErrorCode::Unsupported,
                paste_name,
                &label,
            );
        }
    }
}

#[test]
fn range_metadata_only_paste_special_is_stably_unsupported_for_read_only_destination() {
    for (paste_type, paste_name) in [
        (XL_PASTE_COMMENTS, "xlPasteComments"),
        (XL_PASTE_VALIDATION, "xlPasteValidation"),
        (XL_PASTE_COLUMN_WIDTHS, "xlPasteColumnWidths"),
    ] {
        for clipboard_member in ["Copy", "Cut"] {
            let label = format!("read-only {clipboard_member} {paste_name}");
            let mut runtime = ExcelRuntime::new();
            let source_workbook = open_clean_workbook(&mut runtime);
            let source_worksheet = worksheet_handle(&mut runtime, source_workbook);
            let source = range_handle(&mut runtime, source_worksheet, "A60:B60");
            set_row_values(
                &mut runtime,
                source,
                vec![OmValue::Number(1.0), OmValue::Number(2.0)],
                &label,
            );
            commit_workbook_baseline(&mut runtime, source_workbook, &label);

            let destination_workbook = runtime
                .open_workbook(OpenWorkbookSpec {
                    bytes: synthetic_workbook_bytes(),
                    format_hint: Some(FileFormat::Xlsx),
                    profile: ExcelProfile::Excel365,
                    read_only: true,
                })
                .unwrap_or_else(|error| panic!("{label}: open read-only workbook: {error:?}"));
            let destination_worksheet = worksheet_handle(&mut runtime, destination_workbook);
            let destination = range_handle(&mut runtime, destination_worksheet, "D60:E60");

            assert_custom_paste_failure_preserves_workbooks_and_session(
                &mut runtime,
                source_workbook,
                destination_workbook,
                source,
                destination,
                clipboard_member,
                &[OmValue::Number(f64::from(paste_type))],
                OmErrorCode::Unsupported,
                paste_name,
                &label,
            );
        }
    }
}

#[test]
fn range_metadata_only_paste_special_reports_capability_before_clipboard_state() {
    for (paste_type, paste_name) in [
        (XL_PASTE_COMMENTS, "xlPasteComments"),
        (XL_PASTE_VALIDATION, "xlPasteValidation"),
        (XL_PASTE_COLUMN_WIDTHS, "xlPasteColumnWidths"),
    ] {
        let mut runtime = ExcelRuntime::new();
        let workbook = open_clean_workbook(&mut runtime);
        let worksheet = worksheet_handle(&mut runtime, workbook);
        let destination = range_handle(&mut runtime, worksheet, "D60:E60");
        let workbook_before = runtime_workbook_persistence_snapshot(&runtime, workbook);
        let dirty_before = runtime
            .workbook_dirty_domains(workbook)
            .unwrap_or_else(|error| panic!("{paste_name}: dirty domains: {error:?}"));
        let session_before = runtime_session_mutation_snapshot(&runtime);

        let error = runtime
            .dispatch_invoke(
                destination,
                "PasteSpecial",
                &[OmValue::Number(f64::from(paste_type))],
            )
            .expect_err("metadata-only PasteSpecial should report unsupported capability");

        assert_eq!(error.code, OmErrorCode::Unsupported, "{paste_name}");
        assert_eq!(
            error.message,
            format!("Range.PasteSpecial Paste {paste_name} is not implemented"),
        );
        assert_eq!(
            runtime_workbook_persistence_snapshot(&runtime, workbook),
            workbook_before,
            "{paste_name}: workbook",
        );
        assert_eq!(
            runtime
                .workbook_dirty_domains(workbook)
                .unwrap_or_else(|error| panic!("{paste_name}: dirty domains after: {error:?}")),
            dirty_before,
            "{paste_name}: dirty domains",
        );
        assert_eq!(
            runtime_session_mutation_snapshot(&runtime),
            session_before,
            "{paste_name}: session",
        );
    }
}

#[test]
fn range_custom_paste_special_late_type_errors_are_atomic() {
    for (clipboard_member, source_values, destination_values, expected_message, label) in [
        (
            "Copy",
            vec![OmValue::Number(2.0), OmValue::Text("bad".to_string())],
            vec![OmValue::Number(10.0), OmValue::Number(20.0)],
            "numeric source values",
            "late source coercion",
        ),
        (
            "Cut",
            vec![OmValue::Number(2.0), OmValue::Number(3.0)],
            vec![OmValue::Number(10.0), OmValue::Text("bad".to_string())],
            "numeric destination values",
            "late destination coercion",
        ),
    ] {
        let mut runtime = ExcelRuntime::new();
        let workbook = open_clean_workbook(&mut runtime);
        let worksheet = worksheet_handle(&mut runtime, workbook);
        let source = range_handle(&mut runtime, worksheet, "A20:B20");
        let destination = range_handle(&mut runtime, worksheet, "D20:E20");
        set_row_values(&mut runtime, source, source_values, label);
        set_row_values(&mut runtime, destination, destination_values, label);
        commit_workbook_baseline(&mut runtime, workbook, label);

        assert_custom_paste_failure_preserves_workbooks_and_session(
            &mut runtime,
            workbook,
            workbook,
            source,
            destination,
            clipboard_member,
            &[
                OmValue::Number(f64::from(XL_PASTE_VALUES)),
                OmValue::Number(f64::from(XL_PASTE_SPECIAL_OPERATION_ADD)),
            ],
            OmErrorCode::TypeMismatch,
            expected_message,
            label,
        );
    }
}

#[test]
fn range_custom_paste_special_cross_workbook_late_divide_by_zero_is_atomic() {
    let mut runtime = ExcelRuntime::new();
    let source_workbook = open_clean_workbook(&mut runtime);
    let destination_workbook = open_clean_workbook(&mut runtime);
    let source_sheet = worksheet_handle(&mut runtime, source_workbook);
    let destination_sheet = worksheet_handle(&mut runtime, destination_workbook);
    let source = range_handle(&mut runtime, source_sheet, "A20:B20");
    let destination = range_handle(&mut runtime, destination_sheet, "D20:E20");
    set_row_values(
        &mut runtime,
        source,
        vec![OmValue::Number(2.0), OmValue::Number(0.0)],
        "divide source",
    );
    set_row_values(
        &mut runtime,
        destination,
        vec![OmValue::Number(10.0), OmValue::Number(20.0)],
        "divide destination",
    );
    commit_workbook_baseline(&mut runtime, source_workbook, "divide source");
    commit_workbook_baseline(&mut runtime, destination_workbook, "divide destination");

    assert_custom_paste_failure_preserves_workbooks_and_session(
        &mut runtime,
        source_workbook,
        destination_workbook,
        source,
        destination,
        "Cut",
        &[
            OmValue::Number(f64::from(XL_PASTE_VALUES)),
            OmValue::Number(f64::from(XL_PASTE_SPECIAL_OPERATION_DIVIDE)),
        ],
        OmErrorCode::InvalidArgument,
        "cannot divide by zero",
        "cross-workbook late divide by zero",
    );
}

#[test]
fn range_custom_paste_special_rejects_spill_sources_atomically() {
    {
        let (mut runtime, workbook, worksheet, _) = runtime_with_sequence_spill();
        let normal = range_handle(&mut runtime, worksheet, "I10");
        runtime
            .dispatch_set(normal, "Value2", OmValue::Number(9.0), &[])
            .expect("seed normal source before spill anchor");
        let source = range_handle(&mut runtime, worksheet, "I10:J10");
        let destination = range_handle(&mut runtime, worksheet, "A20");

        assert_custom_paste_failure_preserves_workbooks_and_session(
            &mut runtime,
            workbook,
            workbook,
            source,
            destination,
            "Copy",
            &[OmValue::Number(f64::from(XL_PASTE_VALUES))],
            OmErrorCode::InvalidState,
            "R10C10",
            "copy spill-anchor source",
        );
    }

    {
        let (mut runtime, source_workbook, source_sheet, _) = runtime_with_sequence_spill();
        let destination_workbook = open_clean_workbook(&mut runtime);
        let destination_sheet = worksheet_handle(&mut runtime, destination_workbook);
        let normal = range_handle(&mut runtime, source_sheet, "I11");
        runtime
            .dispatch_set(normal, "Value2", OmValue::Number(9.0), &[])
            .expect("seed normal source before spill child");
        let source = range_handle(&mut runtime, source_sheet, "I11:J11");
        let destination = range_handle(&mut runtime, destination_sheet, "A20");

        assert_custom_paste_failure_preserves_workbooks_and_session(
            &mut runtime,
            source_workbook,
            destination_workbook,
            source,
            destination,
            "Cut",
            &[OmValue::Number(f64::from(XL_PASTE_VALUES))],
            OmErrorCode::InvalidState,
            "R11C10",
            "cross-workbook cut spill-child source",
        );
    }
}

#[test]
fn range_custom_paste_special_rejects_spill_destinations_atomically() {
    {
        let (mut runtime, workbook, worksheet, _) = runtime_with_sequence_spill();
        let source = range_handle(&mut runtime, worksheet, "A20:B20");
        set_row_values(
            &mut runtime,
            source,
            vec![OmValue::Number(9.0), OmValue::Number(8.0)],
            "spill-anchor destination source",
        );
        let destination = range_handle(&mut runtime, worksheet, "I10");

        assert_custom_paste_failure_preserves_workbooks_and_session(
            &mut runtime,
            workbook,
            workbook,
            source,
            destination,
            "Copy",
            &[OmValue::Number(f64::from(XL_PASTE_VALUES))],
            OmErrorCode::InvalidState,
            "R10C10",
            "copy spill-anchor destination",
        );
    }

    {
        let (mut runtime, destination_workbook, destination_sheet, _) =
            runtime_with_sequence_spill();
        let source_workbook = open_clean_workbook(&mut runtime);
        let source_sheet = worksheet_handle(&mut runtime, source_workbook);
        let source = range_handle(&mut runtime, source_sheet, "A20:B20");
        set_row_values(
            &mut runtime,
            source,
            vec![OmValue::Number(9.0), OmValue::Number(8.0)],
            "spill-child destination source",
        );
        let destination = range_handle(&mut runtime, destination_sheet, "I11");

        assert_custom_paste_failure_preserves_workbooks_and_session(
            &mut runtime,
            source_workbook,
            destination_workbook,
            source,
            destination,
            "Cut",
            &[OmValue::Number(f64::from(XL_PASTE_VALUES))],
            OmErrorCode::InvalidState,
            "R11C10",
            "cross-workbook cut spill-child destination",
        );
    }

    {
        let (mut runtime, workbook, destination_sheet, _) = runtime_with_sequence_spill();
        let worksheets = expect_object_handle(
            runtime
                .dispatch_get(workbook.0, "Worksheets", &[])
                .expect("Workbook.Worksheets for cross-sheet spill failure"),
        );
        let source_sheet = expect_object_handle(
            runtime
                .dispatch_invoke(worksheets, "Add", &[])
                .expect("add cross-sheet spill source"),
        );
        let source = range_handle(&mut runtime, source_sheet, "A20:B20");
        set_row_values(
            &mut runtime,
            source,
            vec![OmValue::Number(9.0), OmValue::Number(8.0)],
            "cross-sheet spill destination source",
        );
        let destination = range_handle(&mut runtime, destination_sheet, "I10");

        assert_custom_paste_failure_preserves_workbooks_and_session(
            &mut runtime,
            workbook,
            workbook,
            source,
            destination,
            "Cut",
            &[OmValue::Number(f64::from(XL_PASTE_VALUES))],
            OmErrorCode::InvalidState,
            "R10C10",
            "cross-sheet cut spill-anchor destination",
        );
    }

    {
        let (mut runtime, destination_workbook, destination_sheet, _) =
            runtime_with_sequence_spill();
        let source_workbook = open_clean_workbook(&mut runtime);
        let source_sheet = worksheet_handle(&mut runtime, source_workbook);
        let source = range_handle(&mut runtime, source_sheet, "A20");
        runtime
            .dispatch_set(source, "Value2", OmValue::Number(3.0), &[])
            .expect("seed unchanged spill-child source");
        let destination = range_handle(&mut runtime, destination_sheet, "J11");
        let source_value = runtime
            .dispatch_get(source, "Value2", &[])
            .expect("unchanged source value");
        let destination_value = runtime
            .dispatch_get(destination, "Value2", &[])
            .expect("unchanged spill-child value");
        assert_eq!(source_value, OmValue::Number(3.0));
        assert_eq!(destination_value, source_value);

        assert_custom_paste_failure_preserves_workbooks_and_session(
            &mut runtime,
            source_workbook,
            destination_workbook,
            source,
            destination,
            "Copy",
            &[OmValue::Number(f64::from(XL_PASTE_VALUES))],
            OmErrorCode::InvalidState,
            "R11C10",
            "unchanged spill-child destination",
        );
    }
}

#[test]
fn range_custom_paste_special_skip_blanks_leaves_spill_destination_untouched() {
    let (mut runtime, workbook, worksheet, sheet_id) = runtime_with_sequence_spill();
    let source = range_handle(&mut runtime, worksheet, "A20:B20");
    set_row_values(
        &mut runtime,
        source,
        vec![OmValue::Number(9.0), OmValue::Empty],
        "skip-blanks source",
    );
    let destination = range_handle(&mut runtime, worksheet, "I10");
    let spill_before = {
        let state = runtime
            .workbook_state(workbook)
            .expect("workbook state before");
        let data = state
            .worksheet_data_for_sheet(sheet_id)
            .expect("worksheet data before");
        (
            [(10, 10), (10, 11), (11, 10), (11, 11)]
                .map(|key| (key, data.cells.get(&key).cloned())),
            data.dynamic_array_formulas.clone(),
            data.spill_ranges.clone(),
            data.spill_owners.clone(),
        )
    };

    runtime
        .dispatch_invoke(source, "Copy", &[])
        .expect("arm copy clipboard");
    runtime
        .dispatch_invoke(
            destination,
            "PasteSpecial",
            &[
                OmValue::Number(f64::from(XL_PASTE_VALUES)),
                OmValue::Missing,
                OmValue::Bool(true),
            ],
        )
        .expect("custom PasteSpecial SkipBlanks next to spill");

    let i10 = range_handle(&mut runtime, worksheet, "I10");
    assert_eq!(
        runtime
            .dispatch_get(i10, "Value2", &[])
            .expect("I10 after SkipBlanks"),
        OmValue::Number(9.0),
    );
    let state = runtime.workbook_state(workbook).expect("workbook state");
    let data = state
        .worksheet_data_for_sheet(sheet_id)
        .expect("worksheet data");
    let spill_after = (
        [(10, 10), (10, 11), (11, 10), (11, 11)].map(|key| (key, data.cells.get(&key).cloned())),
        data.dynamic_array_formulas.clone(),
        data.spill_ranges.clone(),
        data.spill_owners.clone(),
    );
    assert_eq!(spill_after, spill_before);
    state.validate_for_save().expect("valid spill state");
    assert!(runtime.clipboard.is_none());
    assert_eq!(runtime.cut_copy_mode, None);
}

#[test]
fn range_custom_paste_special_cut_preserves_same_sheet_overlap_semantics() {
    let mut runtime = ExcelRuntime::new();
    let workbook = open_clean_workbook(&mut runtime);
    let worksheet = worksheet_handle(&mut runtime, workbook);
    let source = range_handle(&mut runtime, worksheet, "A1:B1");
    let destination = range_handle(&mut runtime, worksheet, "B1");
    runtime
        .dispatch_invoke(
            source,
            "Find",
            &[OmValue::Text("overlap marker".to_string())],
        )
        .expect("seed Find state");
    runtime
        .dispatch_invoke(source, "Cut", &[])
        .expect("arm cut clipboard");

    runtime
        .dispatch_invoke(
            destination,
            "PasteSpecial",
            &[OmValue::Number(f64::from(XL_PASTE_VALUES))],
        )
        .expect("custom cut PasteSpecial overlap");

    for (address, expected) in [
        ("A1", OmValue::Empty),
        ("B1", OmValue::Number(42.0)),
        ("C1", OmValue::Text("SHARED".to_string())),
    ] {
        let cell = range_handle(&mut runtime, worksheet, address);
        assert_eq!(
            runtime
                .dispatch_get(cell, "Value2", &[])
                .unwrap_or_else(|error| panic!("{address}: Value2: {error:?}")),
            expected,
            "{address}",
        );
        assert_eq!(
            runtime
                .dispatch_get(cell, "HasFormula", &[])
                .unwrap_or_else(|error| panic!("{address}: HasFormula: {error:?}")),
            OmValue::Bool(false),
            "{address}",
        );
    }
    assert_semantic_dirty_domains(&runtime, workbook, "same-sheet overlap");
    let state = runtime.workbook_state(workbook).expect("same-sheet state");
    let data = state
        .worksheet_data_for_sheet(state.worksheets()[0].id)
        .expect("same-sheet data");
    assert!(data.dirty);
    for key in [(1, 1), (1, 2), (1, 3)] {
        assert!(
            data.dirty_cells.contains(&key),
            "missing dirty cell {key:?}"
        );
    }
    assert!(runtime.clipboard.is_none());
    assert!(runtime.find_state.is_none());
    assert_eq!(runtime.cut_copy_mode, None);
}

#[test]
fn range_custom_paste_special_cut_cross_sheet_commits_both_sides() {
    let mut runtime = ExcelRuntime::new();
    let workbook = open_clean_workbook(&mut runtime);
    let destination_sheet = worksheet_handle(&mut runtime, workbook);
    let destination_sheet_id = runtime
        .workbook_state(workbook)
        .expect("workbook state")
        .worksheets()[0]
        .id;
    let worksheets = expect_object_handle(
        runtime
            .dispatch_get(workbook.0, "Worksheets", &[])
            .expect("Workbook.Worksheets"),
    );
    let source_sheet = expect_object_handle(
        runtime
            .dispatch_invoke(worksheets, "Add", &[])
            .expect("Worksheets.Add source sheet"),
    );
    let source_sheet_id = match runtime
        .runtime_object(source_sheet)
        .expect("source worksheet runtime object")
    {
        RuntimeObjectKind::Worksheet { sheet_id, .. } => sheet_id,
        _ => panic!("source handle must remain a worksheet"),
    };
    let source = range_handle(&mut runtime, source_sheet, "A20:B20");
    let destination = range_handle(&mut runtime, destination_sheet, "D20");
    set_row_values(
        &mut runtime,
        source,
        vec![OmValue::Number(7.0), OmValue::Number(8.0)],
        "cross-sheet source",
    );
    commit_workbook_baseline(&mut runtime, workbook, "cross-sheet setup");

    runtime
        .dispatch_invoke(source, "Cut", &[])
        .expect("arm cross-sheet cut clipboard");
    runtime
        .dispatch_invoke(
            destination,
            "PasteSpecial",
            &[OmValue::Number(f64::from(XL_PASTE_VALUES))],
        )
        .expect("custom cross-sheet cut PasteSpecial");

    for (worksheet, address, expected) in [
        (source_sheet, "A20", OmValue::Empty),
        (source_sheet, "B20", OmValue::Empty),
        (destination_sheet, "D20", OmValue::Number(7.0)),
        (destination_sheet, "E20", OmValue::Number(8.0)),
    ] {
        let cell = range_handle(&mut runtime, worksheet, address);
        assert_eq!(
            runtime
                .dispatch_get(cell, "Value2", &[])
                .unwrap_or_else(|error| panic!("{address}: Value2: {error:?}")),
            expected,
            "{address}",
        );
    }
    assert_semantic_dirty_domains(&runtime, workbook, "cross-sheet cut");
    let state = runtime.workbook_state(workbook).expect("cross-sheet state");
    let source_data = state
        .worksheet_data_for_sheet(source_sheet_id)
        .expect("cross-sheet source data");
    let destination_data = state
        .worksheet_data_for_sheet(destination_sheet_id)
        .expect("cross-sheet destination data");
    assert!(source_data.dirty);
    assert!(destination_data.dirty);
    for key in [(20, 1), (20, 2)] {
        assert!(source_data.dirty_cells.contains(&key));
    }
    for key in [(20, 4), (20, 5)] {
        assert!(destination_data.dirty_cells.contains(&key));
    }
    assert!(runtime.clipboard.is_none());
    assert_eq!(runtime.cut_copy_mode, None);
}

#[test]
fn range_custom_paste_special_cut_cross_workbook_commits_both() {
    let mut runtime = ExcelRuntime::new();
    let source_workbook = open_clean_workbook(&mut runtime);
    let destination_workbook = open_clean_workbook(&mut runtime);
    let source_sheet = worksheet_handle(&mut runtime, source_workbook);
    let destination_sheet = worksheet_handle(&mut runtime, destination_workbook);
    let source = range_handle(&mut runtime, source_sheet, "A1:B1");
    let destination = range_handle(&mut runtime, destination_sheet, "C3");
    runtime
        .dispatch_invoke(
            source,
            "Find",
            &[OmValue::Text("cross-workbook marker".to_string())],
        )
        .expect("seed Find state");
    runtime
        .dispatch_invoke(source, "Cut", &[])
        .expect("arm cut clipboard");

    runtime
        .dispatch_invoke(
            destination,
            "PasteSpecial",
            &[OmValue::Number(f64::from(XL_PASTE_VALUES))],
        )
        .expect("custom cross-workbook cut PasteSpecial");

    for (worksheet, address, expected) in [
        (source_sheet, "A1", OmValue::Empty),
        (source_sheet, "B1", OmValue::Empty),
        (destination_sheet, "C3", OmValue::Number(42.0)),
        (destination_sheet, "D3", OmValue::Text("SHARED".to_string())),
    ] {
        let cell = range_handle(&mut runtime, worksheet, address);
        assert_eq!(
            runtime
                .dispatch_get(cell, "Value2", &[])
                .unwrap_or_else(|error| panic!("{address}: Value2: {error:?}")),
            expected,
            "{address}",
        );
        assert_eq!(
            runtime
                .dispatch_get(cell, "HasFormula", &[])
                .unwrap_or_else(|error| panic!("{address}: HasFormula: {error:?}")),
            OmValue::Bool(false),
            "{address}",
        );
    }
    assert_semantic_dirty_domains(&runtime, source_workbook, "cross-workbook source");
    assert_semantic_dirty_domains(&runtime, destination_workbook, "cross-workbook destination");
    let source_state = runtime
        .workbook_state(source_workbook)
        .expect("cross-workbook source state");
    let source_data = source_state
        .worksheet_data_for_sheet(source_state.worksheets()[0].id)
        .expect("cross-workbook source data");
    assert!(source_data.dirty);
    for key in [(1, 1), (1, 2)] {
        assert!(source_data.dirty_cells.contains(&key));
    }
    let destination_state = runtime
        .workbook_state(destination_workbook)
        .expect("cross-workbook destination state");
    let destination_data = destination_state
        .worksheet_data_for_sheet(destination_state.worksheets()[0].id)
        .expect("cross-workbook destination data");
    assert!(destination_data.dirty);
    for key in [(3, 3), (3, 4)] {
        assert!(destination_data.dirty_cells.contains(&key));
    }
    assert!(runtime.clipboard.is_none());
    assert!(runtime.find_state.is_none());
    assert_eq!(runtime.cut_copy_mode, None);
}
