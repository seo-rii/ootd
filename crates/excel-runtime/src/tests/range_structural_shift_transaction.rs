use super::*;
use crate::{XL_COPY, XL_SHIFT_DOWN, XL_SHIFT_TO_LEFT, XL_SHIFT_TO_RIGHT, XL_SHIFT_UP};
use office_common::DrawingAnchor;

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
fn range_structural_shifts_reject_intersecting_data_validations_atomically() {
    let mut package = OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("base package");
    let source_xml = String::from_utf8(
        package
            .part("xl/worksheets/sheet1.xml")
            .expect("source worksheet")
            .bytes
            .clone(),
    )
    .expect("worksheet utf8");
    let validation_xml = source_xml.replace(
        "</sheetData>",
        "</sheetData>\n  <dataValidations count=\"1\"><dataValidation type=\"whole\" sqref=\"D4:E5 F8\"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>",
    );
    assert_ne!(
        validation_xml, source_xml,
        "data-validation fixture replacement"
    );
    package
        .replace_part_bytes("xl/worksheets/sheet1.xml", validation_xml.into_bytes())
        .expect("replace worksheet");
    let input = package.to_bytes().expect("data-validation workbook bytes");

    for (member, target_address, shift, label) in [
        (
            "Insert",
            "D1:E1",
            XL_SHIFT_DOWN,
            "insert corridor through data validation",
        ),
        (
            "Delete",
            "D4:E4",
            XL_SHIFT_UP,
            "delete corridor through data validation",
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
                "structural data-validation retarget",
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
        .expect("open non-intersecting data-validation fixture");
    let worksheet = worksheet_handle(&mut runtime, workbook);
    let target = range_handle(&mut runtime, worksheet, "A1");
    runtime
        .dispatch_invoke(
            target,
            "Insert",
            &[OmValue::Number(f64::from(XL_SHIFT_DOWN))],
        )
        .expect("non-intersecting data validation must remain eligible");

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
        .expect("save non-intersecting data-validation shift");
    let reopened = runtime
        .codec
        .load(&saved, LoadOptions::default())
        .expect("reopen non-intersecting data-validation shift");
    let reopened_sheet_id = reopened.state.worksheets()[0].id;
    assert_eq!(
        reopened
            .state
            .worksheet_data_for_sheet(reopened_sheet_id)
            .expect("reopened worksheet data")
            .structural_owners
            .data_validation_ranges,
        vec![
            Rect {
                row_first: 4,
                row_last: 5,
                col_first: 4,
                col_last: 5,
            },
            Rect::single_cell(8, 6),
        ],
    );
    let reopened_package = OpcPackage::from_bytes(&saved).expect("reopened package");
    let reopened_xml = std::str::from_utf8(
        &reopened_package
            .part("xl/worksheets/sheet1.xml")
            .expect("reopened worksheet")
            .bytes,
    )
    .expect("reopened worksheet utf8");
    assert!(reopened_xml.contains(r#"sqref="D4:E5 F8""#));
}

#[test]
fn range_structural_shifts_reject_data_validation_formula_owners_atomically() {
    let mut package = OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("base package");
    let source_xml = String::from_utf8(
        package
            .part("xl/worksheets/sheet1.xml")
            .expect("source worksheet")
            .bytes
            .clone(),
    )
    .expect("worksheet utf8");
    let validation_xml = source_xml.replace(
        "</sheetData>",
        "</sheetData>\n  <dataValidations count=\"1\"><dataValidation type=\"custom\" sqref=\"D4:E5\"><formula1>=$A$1&gt;0</formula1></dataValidation></dataValidations>",
    );
    assert_ne!(
        validation_xml, source_xml,
        "data-validation formula fixture replacement"
    );
    package
        .replace_part_bytes("xl/worksheets/sheet1.xml", validation_xml.into_bytes())
        .expect("replace worksheet");
    let input = package
        .to_bytes()
        .expect("data-validation formula workbook bytes");

    for (member, shift, label) in [
        (
            "Insert",
            XL_SHIFT_DOWN,
            "insert moves a data-validation formula precedent",
        ),
        (
            "Delete",
            XL_SHIFT_UP,
            "delete moves a data-validation formula precedent",
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
        let target = range_handle(&mut runtime, worksheet, "A1");

        assert_structural_failure_is_atomic(
            &mut runtime,
            workbook,
            target,
            member,
            shift,
            OmErrorCode::Unsupported,
            &[
                "structural data-validation formula retarget",
                "worksheet 1",
                "formula 1",
            ],
            label,
        );
    }
}

#[test]
fn range_structural_shifts_inventory_x14_data_validation_owners() {
    let mut package = OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("base package");
    let source_xml = String::from_utf8(
        package
            .part("xl/worksheets/sheet1.xml")
            .expect("source worksheet")
            .bytes
            .clone(),
    )
    .expect("worksheet utf8");
    let range_xml = source_xml.replace(
        "</worksheet>",
        r#"  <extLst><ext uri="{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}"><x14:dataValidations xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" count="1"><x14:dataValidation><x14:formula1><xm:f>1</xm:f></x14:formula1><xm:sqref>D4:E5 F8</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>
</worksheet>"#,
    );
    assert_ne!(range_xml, source_xml, "x14 range fixture replacement");
    package
        .replace_part_bytes("xl/worksheets/sheet1.xml", range_xml.into_bytes())
        .expect("replace worksheet");
    let range_input = package.to_bytes().expect("x14 range workbook bytes");

    for (member, target_address, shift, label) in [
        (
            "Insert",
            "D1:E1",
            XL_SHIFT_DOWN,
            "insert corridor through x14 validation",
        ),
        (
            "Delete",
            "D4:E4",
            XL_SHIFT_UP,
            "delete corridor through x14 validation",
        ),
    ] {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: range_input.clone(),
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
                "structural data-validation retarget",
                "worksheet 1",
                "R4C4:R5C5",
            ],
            label,
        );
    }

    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: range_input,
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open non-intersecting x14 validation fixture");
    let worksheet = worksheet_handle(&mut runtime, workbook);
    let target = range_handle(&mut runtime, worksheet, "A1");
    runtime
        .dispatch_invoke(
            target,
            "Insert",
            &[OmValue::Number(f64::from(XL_SHIFT_DOWN))],
        )
        .expect("non-intersecting x14 validation must remain eligible");
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
        .expect("save non-intersecting x14 validation shift");
    let reopened = runtime
        .codec
        .load(&saved, LoadOptions::default())
        .expect("reopen non-intersecting x14 validation shift");
    let reopened_sheet_id = reopened.state.worksheets()[0].id;
    let reopened_owners = &reopened
        .state
        .worksheet_data_for_sheet(reopened_sheet_id)
        .expect("reopened worksheet data")
        .structural_owners;
    assert_eq!(
        reopened_owners.data_validation_ranges,
        vec![
            Rect {
                row_first: 4,
                row_last: 5,
                col_first: 4,
                col_last: 5,
            },
            Rect::single_cell(8, 6),
        ],
    );
    assert_eq!(
        reopened_owners.data_validation_formulas,
        vec!["1".to_string()],
    );

    let mut formula_package =
        OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("formula base package");
    let formula_xml = source_xml.replace(
        "</worksheet>",
        r#"  <extLst><ext uri="{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}"><x14:dataValidations xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" count="1"><x14:dataValidation type="custom"><x14:formula1><xm:f>=$A$1&gt;0</xm:f></x14:formula1><xm:sqref>D4:E5</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>
</worksheet>"#,
    );
    formula_package
        .replace_part_bytes("xl/worksheets/sheet1.xml", formula_xml.into_bytes())
        .expect("replace formula worksheet");
    let formula_input = formula_package
        .to_bytes()
        .expect("x14 formula workbook bytes");
    let mut formula_runtime = ExcelRuntime::new();
    let formula_workbook = formula_runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: formula_input,
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open x14 formula fixture");
    let formula_worksheet = worksheet_handle(&mut formula_runtime, formula_workbook);
    let formula_target = range_handle(&mut formula_runtime, formula_worksheet, "A1");
    assert_structural_failure_is_atomic(
        &mut formula_runtime,
        formula_workbook,
        formula_target,
        "Insert",
        XL_SHIFT_DOWN,
        OmErrorCode::Unsupported,
        &[
            "structural data-validation formula retarget",
            "worksheet 1",
            "formula 1",
        ],
        "x14 validation formula owner",
    );
}

#[test]
fn range_structural_shifts_use_resolved_table_range_owners() {
    let mut package = OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("base package");
    let content_types = String::from_utf8(
        package
            .part("[Content_Types].xml")
            .expect("content types")
            .bytes
            .clone(),
    )
    .expect("content types utf8")
    .replace(
        "</Types>",
        "  <Override PartName=\"/xl/tables/table1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>\n</Types>",
    );
    package
        .replace_part_bytes("[Content_Types].xml", content_types.into_bytes())
        .expect("replace table content types");
    let source_xml = String::from_utf8(
        package
            .part("xl/worksheets/sheet1.xml")
            .expect("source worksheet")
            .bytes
            .clone(),
    )
    .expect("worksheet utf8");
    let table_xml = source_xml.replace(
        "</worksheet>",
        r#"  <tableParts xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="1"><tablePart r:id="rIdTable1"/></tableParts>
</worksheet>"#,
    );
    assert_ne!(table_xml, source_xml, "table fixture replacement");
    package
        .replace_part_bytes("xl/worksheets/sheet1.xml", table_xml.into_bytes())
        .expect("replace worksheet");
    package
        .add_part(OpcPart {
            name: "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
            content_type: None,
            compression: CompressionMethod::Stored,
            bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>"#
                .to_vec(),
        })
        .expect("add worksheet table relationship");
    package
        .add_part(OpcPart {
            name: "xl/tables/table1.xml".to_string(),
            content_type: None,
            compression: CompressionMethod::Stored,
            bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="D4:E5" totalsRowShown="0">
  <autoFilter ref="D4:E5"/>
  <tableColumns count="2"><tableColumn id="1" name="Left"/><tableColumn id="2" name="Right"/></tableColumns>
</table>"#
                .to_vec(),
        })
        .expect("add table part");
    let input = package.to_bytes().expect("table workbook bytes");

    let mut allowed_runtime = ExcelRuntime::new();
    let allowed_workbook = allowed_runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: input.clone(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open non-intersecting table fixture");
    let allowed_worksheet = worksheet_handle(&mut allowed_runtime, allowed_workbook);
    let allowed_target = range_handle(&mut allowed_runtime, allowed_worksheet, "A1");
    allowed_runtime
        .dispatch_invoke(
            allowed_target,
            "Insert",
            &[OmValue::Number(f64::from(XL_SHIFT_DOWN))],
        )
        .expect("non-intersecting table insert must remain eligible");
    let mut saved = Vec::new();
    allowed_runtime
        .save_workbook_to_writer(
            allowed_workbook,
            SaveWorkbookSpec {
                format: FileFormat::Xlsx,
                profile: ExcelProfile::Excel365,
                lossless: true,
            },
            &mut saved,
        )
        .expect("save non-intersecting table insert");
    let reopened = allowed_runtime
        .codec
        .load(&saved, LoadOptions::default())
        .expect("reopen non-intersecting table insert");
    assert_eq!(
        reopened
            .package
            .part("xl/tables/table1.xml")
            .expect("reopened table part")
            .bytes,
        package
            .part("xl/tables/table1.xml")
            .expect("source table part")
            .bytes,
    );
    let reopened_sheet_id = reopened.state.worksheets()[0].id;
    let reopened_owners = &reopened
        .state
        .worksheet_data_for_sheet(reopened_sheet_id)
        .expect("reopened table worksheet data")
        .structural_owners;
    assert_eq!(
        reopened_owners.table_relationship_ids,
        vec!["rIdTable1".to_string()],
    );
    assert_eq!(reopened_owners.table_owners.len(), 1);
    assert_eq!(reopened_owners.table_owners[0].relationship_id, "rIdTable1");
    assert_eq!(
        reopened_owners.table_owners[0].part_uri,
        "xl/tables/table1.xml"
    );
    assert_eq!(
        reopened_owners.table_owners[0].range,
        Rect {
            row_first: 4,
            row_last: 5,
            col_first: 4,
            col_last: 5,
        },
    );
    assert!(reopened_owners.table_owners[0].formulas.is_empty());

    let mut blocked_runtime = ExcelRuntime::new();
    let blocked_workbook = blocked_runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: input.clone(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open intersecting table fixture");
    let blocked_worksheet = worksheet_handle(&mut blocked_runtime, blocked_workbook);
    let blocked_target = range_handle(&mut blocked_runtime, blocked_worksheet, "D4:E4");
    assert_structural_failure_is_atomic(
        &mut blocked_runtime,
        blocked_workbook,
        blocked_target,
        "Delete",
        XL_SHIFT_UP,
        OmErrorCode::Unsupported,
        &[
            "structural table range retarget",
            "worksheet 1",
            "rIdTable1",
            "xl/tables/table1.xml",
            "R4C4:R5C5",
        ],
        "intersecting delete with resolved table owner",
    );

    let mut formula_package = OpcPackage::from_bytes(&input).expect("formula table package");
    let formula_table_xml = String::from_utf8(
        formula_package
            .part("xl/tables/table1.xml")
            .expect("formula table part")
            .bytes
            .clone(),
    )
    .expect("formula table utf8")
    .replace(
        r#"<tableColumn id="1" name="Left"/>"#,
        r#"<tableColumn id="1" name="Left"><calculatedColumnFormula>=$A$1+1</calculatedColumnFormula></tableColumn>"#,
    );
    formula_package
        .replace_part_bytes("xl/tables/table1.xml", formula_table_xml.into_bytes())
        .expect("replace formula table part");
    let formula_input = formula_package
        .to_bytes()
        .expect("formula table workbook bytes");
    let mut formula_runtime = ExcelRuntime::new();
    let formula_workbook = formula_runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: formula_input,
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open formula table fixture");
    let formula_worksheet = worksheet_handle(&mut formula_runtime, formula_workbook);
    let formula_target = range_handle(&mut formula_runtime, formula_worksheet, "A1");
    assert_structural_failure_is_atomic(
        &mut formula_runtime,
        formula_workbook,
        formula_target,
        "Insert",
        XL_SHIFT_DOWN,
        OmErrorCode::Unsupported,
        &[
            "structural table formula retarget",
            "worksheet 1",
            "rIdTable1",
            "xl/tables/table1.xml",
            "formula 1",
        ],
        "table formula owner",
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

#[test]
fn range_structural_shifts_inventory_row_and_column_metadata_owners() {
    let mut package = OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("base package");
    let worksheet_xml = String::from_utf8(
        package
            .part("xl/worksheets/sheet1.xml")
            .expect("worksheet part")
            .bytes
            .clone(),
    )
    .expect("worksheet XML")
    .replace(
        r#"  <dimension ref="A1:C1"/>"#,
        r#"  <dimension ref="A1:G4"/>
  <cols><col min="4" max="5" width="12" customWidth="1"/></cols>"#,
    )
    .replace(
        "    </row>\n  </sheetData>",
        r#"      <c r="F1"><v>6</v></c>
    </row>
    <row r="4" ht="24" customHeight="1"><extLst><ext uri="urn:row"><payload preserved="true"/></ext></extLst></row>
  </sheetData>"#,
    );
    package
        .replace_part_bytes("xl/worksheets/sheet1.xml", worksheet_xml.into_bytes())
        .expect("replace worksheet metadata fixture");
    let input = package.to_bytes().expect("worksheet metadata bytes");

    let mut allowed_runtime = ExcelRuntime::new();
    let allowed_workbook = allowed_runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: input.clone(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open non-intersecting metadata fixture");
    let allowed_worksheet = worksheet_handle(&mut allowed_runtime, allowed_workbook);
    let allowed_target = range_handle(&mut allowed_runtime, allowed_worksheet, "F1");
    allowed_runtime
        .dispatch_invoke(
            allowed_target,
            "Insert",
            &[OmValue::Number(f64::from(XL_SHIFT_TO_RIGHT))],
        )
        .expect("non-intersecting metadata insert must remain eligible");
    let mut saved = Vec::new();
    allowed_runtime
        .save_workbook_to_writer(
            allowed_workbook,
            SaveWorkbookSpec {
                format: FileFormat::Xlsx,
                profile: ExcelProfile::Excel365,
                lossless: true,
            },
            &mut saved,
        )
        .expect("save non-intersecting metadata insert");
    let reopened = allowed_runtime
        .codec
        .load(&saved, LoadOptions::default())
        .expect("reopen non-intersecting metadata insert");
    let reopened_sheet_id = reopened.state.worksheets()[0].id;
    assert_eq!(
        reopened
            .state
            .cell(reopened_sheet_id, 1, 7)
            .expect("G1 after reopen")
            .value,
        CellValue::Number(6.0),
    );
    let reopened_owners = &reopened
        .state
        .worksheet_data_for_sheet(reopened_sheet_id)
        .expect("reopened metadata worksheet")
        .structural_owners;
    assert_eq!(
        reopened_owners.row_metadata_ranges,
        vec![Rect {
            row_first: 4,
            row_last: 4,
            col_first: 1,
            col_last: ExcelLimits::MAX_COLUMN_INDEX,
        }],
    );
    assert_eq!(
        reopened_owners.column_metadata_ranges,
        vec![Rect {
            row_first: 1,
            row_last: ExcelLimits::MAX_ROW_INDEX,
            col_first: 4,
            col_last: 5,
        }],
    );
    let saved_package = OpcPackage::from_bytes(&saved).expect("saved metadata package");
    let saved_worksheet_xml = String::from_utf8(
        saved_package
            .part("xl/worksheets/sheet1.xml")
            .expect("saved metadata worksheet")
            .bytes
            .clone(),
    )
    .expect("saved metadata worksheet XML");
    assert!(
        saved_worksheet_xml
            .contains(r#"<cols><col min="4" max="5" width="12" customWidth="1"/></cols>"#,)
    );
    assert!(saved_worksheet_xml.contains(
        r#"<row r="4" ht="24" customHeight="1"><extLst><ext uri="urn:row"><payload preserved="true"/></ext></extLst></row>"#,
    ));

    for (target_address, shift, expected_message_fragments, label) in [
        (
            "A1",
            XL_SHIFT_DOWN,
            ["structural row metadata retarget", "worksheet 1", "row 4"],
            "row metadata corridor",
        ),
        (
            "A1",
            XL_SHIFT_TO_RIGHT,
            [
                "structural column metadata retarget",
                "worksheet 1",
                "columns C4:C5",
            ],
            "column metadata corridor",
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
            "Insert",
            shift,
            OmErrorCode::Unsupported,
            &expected_message_fragments,
            label,
        );
    }
}

#[test]
fn range_structural_shifts_inventory_chart_source_owners() {
    let input = synthetic_workbook_with_embedded_chart_bytes();

    let mut allowed_runtime = ExcelRuntime::new();
    let allowed_workbook = allowed_runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: input.clone(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open non-intersecting chart source fixture");
    let allowed_worksheet = worksheet_handle(&mut allowed_runtime, allowed_workbook);
    set_number(&mut allowed_runtime, allowed_worksheet, "F1", 6.0);
    let allowed_target = range_handle(&mut allowed_runtime, allowed_worksheet, "F1");
    allowed_runtime
        .dispatch_invoke(
            allowed_target,
            "Insert",
            &[OmValue::Number(f64::from(XL_SHIFT_TO_RIGHT))],
        )
        .expect("non-intersecting chart source insert must remain eligible");
    let mut saved = Vec::new();
    allowed_runtime
        .save_workbook_to_writer(
            allowed_workbook,
            SaveWorkbookSpec {
                format: FileFormat::Xlsx,
                profile: ExcelProfile::Excel365,
                lossless: true,
            },
            &mut saved,
        )
        .expect("save non-intersecting chart source insert");
    let reopened = allowed_runtime
        .codec
        .load(&saved, LoadOptions::default())
        .expect("reopen non-intersecting chart source insert");
    let reopened_sheet_id = reopened.state.worksheets()[0].id;
    assert_eq!(
        reopened
            .state
            .cell(reopened_sheet_id, 1, 7)
            .expect("G1 after reopen")
            .value,
        CellValue::Number(6.0),
    );
    let reopened_chart = reopened
        .state
        .charts()
        .values()
        .next()
        .expect("reopened chart");
    let reopened_series = reopened_chart.series.first().expect("reopened series");
    assert_eq!(
        reopened_series
            .name
            .as_ref()
            .expect("reopened series name")
            .raw
            .text,
        "Sheet1!$C$1",
    );
    assert_eq!(
        reopened_series
            .x_values
            .as_ref()
            .expect("reopened x-values")
            .raw
            .text,
        "Sheet1!$A$1:$B$1",
    );
    assert_eq!(
        reopened_series
            .values
            .as_ref()
            .expect("reopened values")
            .raw
            .text,
        "Sheet1!$A$1:$C$1",
    );

    for (member, shift) in [("Insert", XL_SHIFT_DOWN), ("Delete", XL_SHIFT_UP)] {
        let mut blocked_runtime = ExcelRuntime::new();
        let blocked_workbook = blocked_runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: input.clone(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .unwrap_or_else(|error| panic!("open intersecting chart source fixture: {error:?}"));
        let blocked_worksheet = worksheet_handle(&mut blocked_runtime, blocked_workbook);
        let blocked_target = range_handle(&mut blocked_runtime, blocked_worksheet, "A1");
        assert_structural_failure_is_atomic(
            &mut blocked_runtime,
            blocked_workbook,
            blocked_target,
            member,
            shift,
            OmErrorCode::Unsupported,
            &[
                "structural chart source retarget",
                "series 1",
                "x-values",
                "worksheet 1",
                "range R1C1:R1C2",
            ],
            &format!("chart source corridor {member}"),
        );
    }
}

#[test]
fn range_structural_shifts_inventory_drawing_anchor_owners() {
    let mut package = OpcPackage::from_bytes(&synthetic_workbook_with_embedded_chart_bytes())
        .expect("embedded chart package");
    let drawing_xml = String::from_utf8(
        package
            .part("xl/drawings/drawing1.xml")
            .expect("drawing part")
            .bytes
            .clone(),
    )
    .expect("drawing XML")
    .replace(
        r#"<xdr:absoluteAnchor ar:tag="keep" xmlns:ar="urn:anchor-root">"#,
        r#"<xdr:twoCellAnchor editAs="twoCell" ar:tag="keep" xmlns:ar="urn:anchor-root">"#,
    )
    .replace(
        r#"<xdr:pos x="25400" y="38100" pg:tag="keep" xmlns:pg="urn:pos"/>"#,
        r#"<xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"#,
    )
    .replace(
        r#"<xdr:ext cx="1270000" cy="635000" eg:tag="keep" xmlns:eg="urn:ext"/>"#,
        r#"<xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"#,
    )
    .replace("</xdr:absoluteAnchor>", "</xdr:twoCellAnchor>");
    assert!(drawing_xml.contains(r#"<xdr:twoCellAnchor editAs="twoCell""#));
    package
        .replace_part_bytes("xl/drawings/drawing1.xml", drawing_xml.into_bytes())
        .expect("replace drawing anchor fixture");
    let input = package.to_bytes().expect("two-cell drawing workbook bytes");

    let mut allowed_runtime = ExcelRuntime::new();
    let allowed_workbook = allowed_runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: input.clone(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open non-intersecting drawing anchor fixture");
    let allowed_worksheet = worksheet_handle(&mut allowed_runtime, allowed_workbook);
    set_number(&mut allowed_runtime, allowed_worksheet, "F1", 6.0);
    let allowed_target = range_handle(&mut allowed_runtime, allowed_worksheet, "F1");
    allowed_runtime
        .dispatch_invoke(
            allowed_target,
            "Insert",
            &[OmValue::Number(f64::from(XL_SHIFT_TO_RIGHT))],
        )
        .expect("non-intersecting drawing anchor insert must remain eligible");
    let mut saved = Vec::new();
    allowed_runtime
        .save_workbook_to_writer(
            allowed_workbook,
            SaveWorkbookSpec {
                format: FileFormat::Xlsx,
                profile: ExcelProfile::Excel365,
                lossless: true,
            },
            &mut saved,
        )
        .expect("save non-intersecting drawing anchor insert");
    let reopened = allowed_runtime
        .codec
        .load(&saved, LoadOptions::default())
        .expect("reopen non-intersecting drawing anchor insert");
    let reopened_sheet_id = reopened.state.worksheets()[0].id;
    assert_eq!(
        reopened
            .state
            .cell(reopened_sheet_id, 1, 7)
            .expect("G1 after reopen")
            .value,
        CellValue::Number(6.0),
    );
    let reopened_drawing = reopened
        .state
        .drawings()
        .values()
        .next()
        .expect("reopened drawing");
    let DrawingObjectModel::ChartFrame(reopened_chart_object) = &reopened_drawing.objects[0] else {
        panic!("expected reopened chart frame");
    };
    let Some(DrawingAnchor::TwoCell(reopened_anchor)) = reopened_chart_object.anchor.as_ref()
    else {
        panic!("expected reopened two-cell anchor");
    };
    assert_eq!(reopened_anchor.from.row_zero_based, 3);
    assert_eq!(reopened_anchor.from.col_zero_based, 3);
    assert_eq!(reopened_anchor.to.row_zero_based, 4);
    assert_eq!(reopened_anchor.to.col_zero_based, 4);
    let saved_package = OpcPackage::from_bytes(&saved).expect("saved drawing package");
    let saved_drawing_xml = String::from_utf8(
        saved_package
            .part("xl/drawings/drawing1.xml")
            .expect("saved drawing part")
            .bytes
            .clone(),
    )
    .expect("saved drawing XML");
    assert!(saved_drawing_xml.contains(r#"<xdr:twoCellAnchor editAs="twoCell""#));
    assert!(saved_drawing_xml.contains("<xdr:col>3</xdr:col>"));
    assert!(saved_drawing_xml.contains("<xdr:row>4</xdr:row>"));

    for (member, shift) in [("Insert", XL_SHIFT_DOWN), ("Delete", XL_SHIFT_UP)] {
        let mut blocked_runtime = ExcelRuntime::new();
        let blocked_workbook = blocked_runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: input.clone(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .unwrap_or_else(|error| panic!("open intersecting drawing fixture: {error:?}"));
        let blocked_worksheet = worksheet_handle(&mut blocked_runtime, blocked_workbook);
        let blocked_target = range_handle(&mut blocked_runtime, blocked_worksheet, "D1:E1");
        assert_structural_failure_is_atomic(
            &mut blocked_runtime,
            blocked_workbook,
            blocked_target,
            member,
            shift,
            OmErrorCode::Unsupported,
            &[
                "structural drawing anchor retarget",
                "worksheet 1",
                "range R4C4:R5C5",
            ],
            &format!("drawing anchor corridor {member}"),
        );
    }

    let mut opaque_runtime = ExcelRuntime::new();
    let opaque_workbook = opaque_runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: synthetic_workbook_with_embedded_chart_and_raw_shape_bytes(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open opaque drawing anchor fixture");
    let opaque_worksheet = worksheet_handle(&mut opaque_runtime, opaque_workbook);
    let opaque_target = range_handle(&mut opaque_runtime, opaque_worksheet, "J20");
    assert_structural_failure_is_atomic(
        &mut opaque_runtime,
        opaque_workbook,
        opaque_target,
        "Insert",
        XL_SHIFT_DOWN,
        OmErrorCode::Unsupported,
        &["structural drawing anchor retarget", "opaque anchor"],
        "opaque drawing anchor",
    );
}
