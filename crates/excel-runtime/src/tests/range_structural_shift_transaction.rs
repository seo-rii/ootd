use super::*;
use crate::{XL_COPY, XL_SHIFT_DOWN, XL_SHIFT_TO_LEFT, XL_SHIFT_TO_RIGHT, XL_SHIFT_UP};

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

fn set_number(runtime: &mut ExcelRuntime, worksheet: ObjectHandle, address: &str, value: f64) {
    let range = range_handle(runtime, worksheet, address);
    runtime
        .dispatch_set(range, "Value2", OmValue::Number(value), &[])
        .unwrap_or_else(|error| panic!("{address}.Value2: {error:?}"));
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

fn assert_structural_failure_is_atomic(
    runtime: &mut ExcelRuntime,
    workbook: WorkbookHandle,
    target: ObjectHandle,
    member: &str,
    shift: i32,
    expected_code: OmErrorCode,
    expected_message_fragments: &[&str],
    label: &str,
) {
    runtime
        .dispatch_invoke(
            target,
            "Find",
            &[OmValue::Text("structural session marker".to_string())],
        )
        .unwrap_or_else(|error| panic!("{label}: seed Find state: {error:?}"));
    runtime
        .dispatch_invoke(target, "Copy", &[])
        .unwrap_or_else(|error| panic!("{label}: arm clipboard: {error:?}"));
    assert!(
        runtime.find_state.is_some(),
        "{label}: Find state precondition"
    );
    assert_eq!(runtime.cut_copy_mode, Some(XL_COPY), "{label}");
    assert!(
        runtime.clipboard.is_some(),
        "{label}: clipboard precondition"
    );

    let workbook_before = runtime_workbook_persistence_snapshot(runtime, workbook);
    let dirty_before = runtime
        .workbook_dirty_domains(workbook)
        .unwrap_or_else(|error| panic!("{label}: dirty domains before: {error:?}"));
    let session_before = runtime_session_mutation_snapshot(runtime);

    let error = match runtime.dispatch_invoke(target, member, &[OmValue::Number(f64::from(shift))])
    {
        Ok(value) => panic!("{label}: Range.{member} unexpectedly succeeded: {value:?}"),
        Err(error) => error,
    };

    assert_eq!(error.code, expected_code, "{label}: {error:?}");
    for fragment in expected_message_fragments {
        assert!(error.message.contains(fragment), "{label}: {error:?}");
    }
    assert_eq!(
        runtime_workbook_persistence_snapshot(runtime, workbook),
        workbook_before,
        "{label}: workbook",
    );
    assert_eq!(
        runtime
            .workbook_dirty_domains(workbook)
            .unwrap_or_else(|error| panic!("{label}: dirty domains after: {error:?}")),
        dirty_before,
        "{label}: dirty domains",
    );
    assert_eq!(
        runtime_session_mutation_snapshot(runtime),
        session_before,
        "{label}: runtime session",
    );
}

#[test]
fn range_insert_late_lane_overflow_is_atomic() {
    for (target_address, shift, seeds, message_fragment, label) in [
        (
            "D1:E1",
            XL_SHIFT_DOWN,
            [("D2", 11.0), ("E1048576", 22.0)],
            "rows",
            "late column row overflow",
        ),
        (
            "A20:A21",
            XL_SHIFT_TO_RIGHT,
            [("B20", 11.0), ("XFD21", 22.0)],
            "columns",
            "late row column overflow",
        ),
    ] {
        let mut runtime = ExcelRuntime::new();
        let workbook = open_clean_workbook(&mut runtime);
        let worksheet = worksheet_handle(&mut runtime, workbook);
        for (address, value) in seeds {
            set_number(&mut runtime, worksheet, address, value);
        }
        commit_workbook_baseline(&mut runtime, workbook, label);
        let target = range_handle(&mut runtime, worksheet, target_address);

        assert_structural_failure_is_atomic(
            &mut runtime,
            workbook,
            target,
            "Insert",
            shift,
            OmErrorCode::InvalidArgument,
            &[message_fragment],
            label,
        );
    }
}

#[test]
fn range_insert_delete_reject_spill_corridors_atomically() {
    for (member, target_address, shift, normal_address, expected_message_fragments, label) in [
        (
            "Insert",
            "I9:J9",
            XL_SHIFT_DOWN,
            "I10",
            &["R10C10"][..],
            "insert down through spill anchor",
        ),
        (
            "Delete",
            "I11:J11",
            XL_SHIFT_UP,
            "I12",
            &["R11C10", "R10C10"][..],
            "delete up through spill child",
        ),
        (
            "Insert",
            "K9:K10",
            XL_SHIFT_TO_RIGHT,
            "L9",
            &["R10C11", "R10C10"][..],
            "insert right through spill child",
        ),
        (
            "Delete",
            "I9:I10",
            XL_SHIFT_TO_LEFT,
            "J9",
            &["R10C10"][..],
            "delete left through spill anchor",
        ),
    ] {
        let (mut runtime, workbook, worksheet, _) = runtime_with_sequence_spill();
        set_number(&mut runtime, worksheet, normal_address, 99.0);
        commit_workbook_baseline(&mut runtime, workbook, label);
        let target = range_handle(&mut runtime, worksheet, target_address);

        assert_structural_failure_is_atomic(
            &mut runtime,
            workbook,
            target,
            member,
            shift,
            OmErrorCode::InvalidState,
            expected_message_fragments,
            label,
        );
    }
}

#[test]
fn range_insert_rejects_unmaterialized_spill_intersection() {
    let (mut runtime, workbook, worksheet, sheet_id) = runtime_with_sequence_spill();
    {
        let worksheet_data = runtime
            .runtime_workbook_mut(workbook)
            .expect("runtime workbook")
            .loaded
            .state
            .worksheet_data_for_sheet_mut(sheet_id)
            .expect("worksheet data");
        for key in [(10, 11), (11, 11)] {
            worksheet_data.cells.remove(&key);
            worksheet_data.spill_owners.remove(&key);
        }
    }
    set_number(&mut runtime, worksheet, "K20", 17.0);
    commit_workbook_baseline(&mut runtime, workbook, "unmaterialized spill child");
    let target = range_handle(&mut runtime, worksheet, "K10");

    assert_structural_failure_is_atomic(
        &mut runtime,
        workbook,
        target,
        "Insert",
        XL_SHIFT_DOWN,
        OmErrorCode::InvalidState,
        &["R10C11", "R10C10"],
        "unmaterialized spill child",
    );
}

#[test]
fn range_structural_shifts_fail_closed_for_reference_formulas_atomically() {
    for (member, target_address, shift, formula_address, formula_text, formula_cell, label) in [
        (
            "Insert",
            "A1",
            XL_SHIFT_DOWN,
            "M50",
            "=A2",
            "R50C13",
            "insert with external formula owner",
        ),
        (
            "Delete",
            "A1",
            XL_SHIFT_TO_LEFT,
            "M1",
            "=A1",
            "R1C13",
            "delete with moved formula owner",
        ),
    ] {
        let mut runtime = ExcelRuntime::new();
        let workbook = open_clean_workbook(&mut runtime);
        let worksheet = worksheet_handle(&mut runtime, workbook);
        let formula = range_handle(&mut runtime, worksheet, formula_address);
        runtime
            .dispatch_set(
                formula,
                "Formula",
                OmValue::Text(formula_text.to_string()),
                &[],
            )
            .unwrap_or_else(|error| panic!("{label}: seed formula: {error:?}"));
        commit_workbook_baseline(&mut runtime, workbook, label);
        let target = range_handle(&mut runtime, worksheet, target_address);

        assert_structural_failure_is_atomic(
            &mut runtime,
            workbook,
            target,
            member,
            shift,
            OmErrorCode::Unsupported,
            &["structural formula retarget", formula_cell],
            label,
        );
    }
}

#[test]
fn range_structural_shifts_fail_closed_for_reference_defined_names_atomically() {
    for (
        member,
        target_address,
        shift,
        worksheet_scope,
        defined_name,
        refers_to,
        is_r1c1,
        scope_fragment,
        label,
    ) in [
        (
            "Insert",
            "A1",
            XL_SHIFT_DOWN,
            false,
            "WorkbookShiftOwner",
            "=Sheet1!$M$50",
            false,
            "workbook",
            "insert with external workbook name owner",
        ),
        (
            "Delete",
            "A1",
            XL_SHIFT_TO_LEFT,
            true,
            "WorksheetShiftOwner",
            "=Sheet1!R1C1",
            true,
            "worksheet 1",
            "delete with moved worksheet name owner",
        ),
    ] {
        let mut runtime = ExcelRuntime::new();
        let workbook = open_clean_workbook(&mut runtime);
        let worksheet = worksheet_handle(&mut runtime, workbook);
        let names_owner = if worksheet_scope {
            worksheet
        } else {
            workbook.0
        };
        let names = expect_object_handle(
            runtime
                .dispatch_get(names_owner, "Names", &[])
                .unwrap_or_else(|error| panic!("{label}: Names: {error:?}")),
        );
        let add_args = if is_r1c1 {
            vec![
                OmValue::Text(defined_name.to_string()),
                OmValue::Missing,
                OmValue::Missing,
                OmValue::Missing,
                OmValue::Missing,
                OmValue::Missing,
                OmValue::Missing,
                OmValue::Missing,
                OmValue::Missing,
                OmValue::Text(refers_to.to_string()),
            ]
        } else {
            vec![
                OmValue::Text(defined_name.to_string()),
                OmValue::Text(refers_to.to_string()),
            ]
        };
        runtime
            .dispatch_invoke(names, "Add", &add_args)
            .unwrap_or_else(|error| panic!("{label}: Names.Add: {error:?}"));
        commit_workbook_baseline(&mut runtime, workbook, label);
        let target = range_handle(&mut runtime, worksheet, target_address);

        assert_structural_failure_is_atomic(
            &mut runtime,
            workbook,
            target,
            member,
            shift,
            OmErrorCode::Unsupported,
            &[
                "structural defined-name retarget",
                defined_name,
                scope_fragment,
            ],
            label,
        );
    }
}

#[test]
fn range_structural_shifts_reject_intersecting_merged_cells_atomically() {
    let mut package = OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("base package");
    let source_xml = String::from_utf8(
        package
            .part("xl/worksheets/sheet1.xml")
            .expect("source worksheet")
            .bytes
            .clone(),
    )
    .expect("worksheet utf8");
    let merged_xml = source_xml.replace(
        "</sheetData>",
        "</sheetData>\n  <mergeCells count=\"1\"><mergeCell ref=\"D4:E5\"/></mergeCells>",
    );
    assert_ne!(merged_xml, source_xml, "merge fixture replacement");
    package
        .replace_part_bytes("xl/worksheets/sheet1.xml", merged_xml.into_bytes())
        .expect("replace worksheet");
    let input = package.to_bytes().expect("merged workbook bytes");

    for (member, target_address, shift, label) in [
        (
            "Insert",
            "D1:E1",
            XL_SHIFT_DOWN,
            "insert corridor through merged range",
        ),
        (
            "Delete",
            "D4:E4",
            XL_SHIFT_UP,
            "delete corridor through merged range",
        ),
    ] {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: input.clone(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .unwrap_or_else(|error| panic!("{label}: open: {error:?}"));
        let worksheet = worksheet_handle(&mut runtime, workbook);
        let target = range_handle(&mut runtime, worksheet, target_address);

        assert_structural_failure_is_atomic(
            &mut runtime,
            workbook,
            target,
            member,
            shift,
            OmErrorCode::Unsupported,
            &[
                "structural merged-cell retarget",
                "worksheet 1",
                "R4C4:R5C5",
            ],
            label,
        );
    }

    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: input,
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open non-intersecting merge fixture");
    let worksheet = worksheet_handle(&mut runtime, workbook);
    let target = range_handle(&mut runtime, worksheet, "A1");
    runtime
        .dispatch_invoke(
            target,
            "Insert",
            &[OmValue::Number(f64::from(XL_SHIFT_DOWN))],
        )
        .expect("non-intersecting merged range must remain eligible");

    let mut saved = Vec::new();
    runtime
        .save_workbook_to_writer(
            workbook,
            SaveWorkbookSpec {
                format: FileFormat::Xlsx,
                profile: ExcelProfile::Excel365,
                lossless: true,
            },
            &mut saved,
        )
        .expect("save non-intersecting merged range shift");
    let reopened = runtime
        .codec
        .load(&saved, LoadOptions::default())
        .expect("reopen non-intersecting merged range shift");
    let reopened_sheet_id = reopened.state.worksheets()[0].id;
    assert_eq!(
        reopened
            .state
            .worksheet_data_for_sheet(reopened_sheet_id)
            .expect("reopened worksheet data")
            .structural_owners
            .merged_ranges,
        vec![Rect {
            row_first: 4,
            row_last: 5,
            col_first: 4,
            col_last: 5,
        }],
    );
}

#[test]
fn range_structural_multilane_insert_commits_and_reopens() {
    let mut runtime = ExcelRuntime::new();
    let workbook = open_clean_workbook(&mut runtime);
    let worksheet = worksheet_handle(&mut runtime, workbook);
    let names = expect_object_handle(
        runtime
            .dispatch_get(workbook.0, "Names", &[])
            .expect("Workbook.Names"),
    );
    runtime
        .dispatch_invoke(
            names,
            "Add",
            &[
                OmValue::Text("ConstantShiftOwner".to_string()),
                OmValue::Text("=42".to_string()),
            ],
        )
        .expect("Names.Add constant owner");
    let target = range_handle(&mut runtime, worksheet, "A1:B1");
    runtime
        .dispatch_invoke(target, "Find", &[OmValue::Text("shared".to_string())])
        .expect("seed Find state");
    runtime
        .dispatch_invoke(target, "Copy", &[])
        .expect("arm clipboard");

    runtime
        .dispatch_invoke(
            target,
            "Insert",
            &[OmValue::Number(f64::from(XL_SHIFT_DOWN))],
        )
        .expect("A1:B1.Insert(xlShiftDown)");

    assert!(runtime.find_state.is_none());
    assert!(runtime.cut_copy_mode.is_none());
    assert!(runtime.clipboard.is_none());
    assert_eq!(
        runtime
            .workbook_dirty_domains(workbook)
            .expect("dirty domains after Insert"),
        WorkbookDirtyDomains {
            prompt_dirty: true,
            semantic_dirty: true,
            serialization_dirty: true,
            ..WorkbookDirtyDomains::default()
        },
    );
    runtime
        .workbook_state(workbook)
        .expect("workbook state after Insert")
        .validate_for_save()
        .expect("valid state after Insert");

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
        .expect("save inserted workbook");
    let reopened = runtime
        .codec
        .load(&bytes, LoadOptions::default())
        .expect("reopen inserted workbook");
    let reopened_sheet_id = reopened.state.worksheets()[0].id;
    assert!(
        reopened.state.cell(reopened_sheet_id, 1, 1).is_none(),
        "A1 should be vacated",
    );
    assert!(
        reopened.state.cell(reopened_sheet_id, 1, 2).is_none(),
        "B1 should be vacated",
    );
    assert_eq!(
        reopened
            .state
            .cell(reopened_sheet_id, 2, 1)
            .expect("A2 after reopen")
            .value,
        CellValue::Number(42.0),
    );
    let formula_cell = reopened
        .state
        .cell(reopened_sheet_id, 2, 2)
        .expect("B2 after reopen");
    assert_eq!(formula_cell.value, CellValue::Text("SHARED".to_string()));
    assert_eq!(
        formula_cell.formula.as_ref().expect("B2 formula").text,
        r#"UPPER("shared")"#,
    );
    assert_eq!(
        reopened
            .state
            .lookup_name_in_scope(office_common::NameScope::Workbook, "ConstantShiftOwner")
            .expect("constant defined name after reopen")
            .refers_to
            .text,
        "42",
    );
}
