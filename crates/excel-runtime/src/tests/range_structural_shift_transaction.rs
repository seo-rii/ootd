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
fn partial_corridor_shifts_retarget_formulas_inside_the_band() {
    for (
        member,
        target_address,
        direction,
        formula_address,
        formula_text,
        moved_address,
        moved_formula,
    ) in [
        (
            "Insert",
            "A1",
            XL_SHIFT_DOWN,
            "M50",
            "=A2+B2",
            "M50",
            "=A3+B2",
        ),
        (
            "Delete",
            "A1",
            XL_SHIFT_TO_LEFT,
            "M1",
            "=A1+M2",
            "L1",
            "=#REF!+M2",
        ),
        (
            "Insert",
            "B2:C3",
            XL_SHIFT_DOWN,
            "E1",
            "=SUM(B2:C5)+SUM(A2:C5)",
            "E1",
            "=SUM(B4:C7)+SUM(A2:C5)",
        ),
    ] {
        let label = format!("{member} {target_address}");
        let mut runtime = ExcelRuntime::new();
        let workbook = open_clean_workbook(&mut runtime);
        let worksheet = worksheet_handle(&mut runtime, workbook);
        set_formula(&mut runtime, worksheet, formula_address, formula_text);
        shift(&mut runtime, worksheet, target_address, member, direction);
        assert_eq!(
            formula_of(&mut runtime, worksheet, moved_address),
            OmValue::Text(moved_formula.to_string()),
            "{label}",
        );
    }
}

#[test]
fn structural_shifts_retarget_r1c1_formulas() {
    let mut runtime = ExcelRuntime::new();
    let workbook = open_clean_workbook(&mut runtime);
    let worksheet = worksheet_handle(&mut runtime, workbook);
    for (address, formula) in [("D10", "=R[-6]C+R1C1"), ("B3", "=R[5]C[1]")] {
        let range = range_handle(&mut runtime, worksheet, address);
        runtime
            .dispatch_set(
                range,
                "FormulaR1C1",
                OmValue::Text(formula.to_string()),
                &[],
            )
            .unwrap_or_else(|error| panic!("{address}.FormulaR1C1: {error:?}"));
    }

    shift(&mut runtime, worksheet, "A5:XFD6", "Insert", XL_SHIFT_DOWN);
    shift(
        &mut runtime,
        worksheet,
        "A1:A20",
        "Insert",
        XL_SHIFT_TO_RIGHT,
    );

    // D10 lands on E12 and still reads D4 (now E4); B3 lands on C3 and reads C8 (now D10).
    for (address, formula) in [("E12", "=R[-8]C+R1C2"), ("C3", "=R[7]C[1]")] {
        let range = range_handle(&mut runtime, worksheet, address);
        assert_eq!(
            runtime
                .dispatch_get(range, "FormulaR1C1", &[])
                .unwrap_or_else(|error| panic!("{address}.FormulaR1C1: {error:?}")),
            OmValue::Text(formula.to_string()),
            "{address}",
        );
    }
}

#[test]
fn range_structural_shifts_fail_closed_for_reference_formulas_atomically() {
    for (member, target_address, shift, formula_address, formula_text, formula_cell, label) in [
        (
            "Insert",
            "A1",
            XL_SHIFT_DOWN,
            "M50",
            "=SUM(Sheet1:Sheet1!A2)",
            "R50C13",
            "insert with a 3D formula owner",
        ),
        (
            "Delete",
            "A1",
            XL_SHIFT_TO_LEFT,
            "M1",
            "=SUM(Sheet1:Sheet1!A1)",
            "R1C13",
            "delete with a moved 3D formula owner",
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
            "=Sheet1!A50",
            false,
            "workbook",
            "insert with a relative workbook name owner",
        ),
        (
            "Delete",
            "A1",
            XL_SHIFT_TO_LEFT,
            true,
            "WorksheetShiftOwner",
            "=Sheet1!RC[12]",
            true,
            "worksheet 1",
            "delete with a relative R1C1 worksheet name owner",
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
            "E1:F1",
            XL_SHIFT_DOWN,
            "insert band cutting through merged range",
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

    // A corridor beside the merge leaves it; a band spanning its columns moves it.
    for (target_address, expected_rows) in [("A1", (4, 5)), ("D1:E1", (5, 6))] {
        let mut runtime = ExcelRuntime::new();
        let workbook = runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: input.clone(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open merge fixture");
        let worksheet = worksheet_handle(&mut runtime, workbook);
        let target = range_handle(&mut runtime, worksheet, target_address);
        runtime
            .dispatch_invoke(
                target,
                "Insert",
                &[OmValue::Number(f64::from(XL_SHIFT_DOWN))],
            )
            .unwrap_or_else(|error| {
                panic!("{target_address}: merged range must follow: {error:?}")
            });

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
            .expect("save merged range shift");
        let reopened = runtime
            .codec
            .load(&saved, LoadOptions::default())
            .expect("reopen merged range shift");
        let reopened_sheet_id = reopened.state.worksheets()[0].id;
        assert_eq!(
            reopened
                .state
                .worksheet_data_for_sheet(reopened_sheet_id)
                .expect("reopened worksheet data")
                .structural_owners
                .merged_ranges,
            vec![Rect {
                row_first: expected_rows.0,
                row_last: expected_rows.1,
                col_first: 4,
                col_last: 5,
            }],
            "{target_address}",
        );
    }
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
            "E1:F1",
            XL_SHIFT_DOWN,
            "insert band cutting through data validation",
        ),
        (
            "Delete",
            "E4:F4",
            XL_SHIFT_UP,
            "delete band cutting through data validation",
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
        "</sheetData>\n  <dataValidations count=\"1\"><dataValidation type=\"custom\" sqref=\"D4:E5\"><formula1>=Sheet1!$A$1&gt;0</formula1></dataValidation></dataValidations>",
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
            "E1:F1",
            XL_SHIFT_DOWN,
            "insert band cutting through x14 validation",
        ),
        (
            "Delete",
            "E4:F4",
            XL_SHIFT_UP,
            "delete band cutting through x14 validation",
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
        r#"  <extLst><ext uri="{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}"><x14:dataValidations xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" count="1"><x14:dataValidation type="custom"><x14:formula1><xm:f>=Sheet1!$A$1&gt;0</xm:f></x14:formula1><xm:sqref>D4:E5</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>
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

    // A partial corridor moves cells, not whole rows or columns, so row heights and column widths
    // stay where they are.
    for (shift, label) in [
        (XL_SHIFT_DOWN, "row metadata beside a band"),
        (XL_SHIFT_TO_RIGHT, "column metadata beside a band"),
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
        runtime
            .dispatch_invoke(target, "Insert", &[OmValue::Number(f64::from(shift))])
            .unwrap_or_else(|error| panic!("{label}: {error:?}"));
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
            .unwrap_or_else(|error| panic!("{label}: save: {error:?}"));
        let saved_xml = String::from_utf8(
            OpcPackage::from_bytes(&saved)
                .expect("saved package")
                .part("xl/worksheets/sheet1.xml")
                .expect("saved worksheet")
                .bytes
                .clone(),
        )
        .expect("saved worksheet XML");
        assert!(
            saved_xml.contains(r#"<col min="4" max="5" width="12" customWidth="1"/>"#),
            "{label}: {saved_xml}",
        );
        assert!(
            saved_xml.contains(r#"<row r="4" ht="24" customHeight="1">"#),
            "{label}: {saved_xml}",
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

    // A band that holds a whole series source moves it; one that deletes a whole source refuses.
    let mut moved_runtime = ExcelRuntime::new();
    let moved_workbook = moved_runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: input.clone(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open chart source band fixture");
    let moved_worksheet = worksheet_handle(&mut moved_runtime, moved_workbook);
    shift(
        &mut moved_runtime,
        moved_worksheet,
        "A1:C1",
        "Insert",
        XL_SHIFT_DOWN,
    );
    let moved_series = moved_runtime
        .runtime_workbook_mut(moved_workbook)
        .expect("runtime workbook")
        .loaded
        .state
        .charts()
        .values()
        .next()
        .expect("chart")
        .series[0]
        .clone();
    assert_eq!(moved_series.name.expect("name").raw.text, "Sheet1!$C$2");
    assert_eq!(
        moved_series.x_values.expect("x-values").raw.text,
        "Sheet1!$A$2:$B$2"
    );
    assert_eq!(
        moved_series.values.expect("values").raw.text,
        "Sheet1!$A$2:$C$2"
    );

    let mut blocked_runtime = ExcelRuntime::new();
    let blocked_workbook = blocked_runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: input.clone(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .unwrap_or_else(|error| panic!("open deleted chart source fixture: {error:?}"));
    let blocked_worksheet = worksheet_handle(&mut blocked_runtime, blocked_workbook);
    let blocked_target = range_handle(&mut blocked_runtime, blocked_worksheet, "A1:B1");
    assert_structural_failure_is_atomic(
        &mut blocked_runtime,
        blocked_workbook,
        blocked_target,
        "Delete",
        XL_SHIFT_UP,
        OmErrorCode::Unsupported,
        &[
            "structural chart source retarget",
            "series 1",
            "x-values",
            "range on deleted cells",
        ],
        "band deleting a whole chart source",
    );
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
    let opaque_target = range_handle(&mut opaque_runtime, opaque_worksheet, "F1");
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

fn formula_of(runtime: &mut ExcelRuntime, worksheet: ObjectHandle, address: &str) -> OmValue {
    let range = range_handle(runtime, worksheet, address);
    runtime
        .dispatch_get(range, "Formula", &[])
        .unwrap_or_else(|error| panic!("{address}.Formula: {error:?}"))
}

fn value_of(runtime: &mut ExcelRuntime, worksheet: ObjectHandle, address: &str) -> OmValue {
    let range = range_handle(runtime, worksheet, address);
    runtime
        .dispatch_get(range, "Value2", &[])
        .unwrap_or_else(|error| panic!("{address}.Value2: {error:?}"))
}

fn shift(
    runtime: &mut ExcelRuntime,
    worksheet: ObjectHandle,
    address: &str,
    member: &str,
    direction: i32,
) {
    let target = range_handle(runtime, worksheet, address);
    runtime
        .dispatch_invoke(target, member, &[OmValue::Number(f64::from(direction))])
        .unwrap_or_else(|error| panic!("{address}.{member}: {error:?}"));
}

fn set_formula(runtime: &mut ExcelRuntime, worksheet: ObjectHandle, address: &str, formula: &str) {
    let range = range_handle(runtime, worksheet, address);
    runtime
        .dispatch_set(range, "Formula", OmValue::Text(formula.to_string()), &[])
        .unwrap_or_else(|error| panic!("{address}.Formula: {error:?}"));
}

#[test]
fn whole_row_and_column_shifts_retarget_formulas_and_names() {
    let mut runtime = ExcelRuntime::new();
    let workbook = open_clean_workbook(&mut runtime);
    let worksheet = worksheet_handle(&mut runtime, workbook);
    let sheet_name = runtime.worksheets(workbook).expect("worksheets")[0]
        .name
        .clone();
    for row in 1..=5 {
        set_number(&mut runtime, worksheet, &format!("E{row}"), f64::from(row));
    }
    set_formula(&mut runtime, worksheet, "F1", "=SUM(E1:E5)");
    set_formula(&mut runtime, worksheet, "G1", "=E3*10");
    set_formula(&mut runtime, worksheet, "H1", "=SUM(Block)");
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
                OmValue::Text("Block".to_string()),
                OmValue::Text(format!("={sheet_name}!$E$2:$E$4")),
            ],
        )
        .expect("Names.Add Block");

    shift(&mut runtime, worksheet, "A2:XFD2", "Insert", XL_SHIFT_DOWN);
    for (address, formula) in [
        ("F1", "=SUM(E1:E6)"),
        ("G1", "=E4*10"),
        ("H1", "=SUM(Block)"),
    ] {
        assert_eq!(
            formula_of(&mut runtime, worksheet, address),
            OmValue::Text(formula.to_string()),
            "after row insert {address}"
        );
    }
    let block = expect_object_handle(
        runtime
            .dispatch_invoke(names, "Item", &[OmValue::Text("Block".to_string())])
            .expect("Names.Item(Block)"),
    );
    assert_eq!(
        runtime
            .dispatch_get(block, "RefersTo", &[])
            .expect("RefersTo"),
        OmValue::Text(format!("={sheet_name}!$E$3:$E$5"))
    );
    runtime
        .calculate_workbook_with_report(workbook)
        .expect("calculate after insert");
    assert_eq!(
        value_of(&mut runtime, worksheet, "F1"),
        OmValue::Number(15.0)
    );
    assert_eq!(
        value_of(&mut runtime, worksheet, "G1"),
        OmValue::Number(30.0)
    );
    assert_eq!(
        value_of(&mut runtime, worksheet, "H1"),
        OmValue::Number(9.0)
    );

    shift(&mut runtime, worksheet, "A3:XFD4", "Delete", XL_SHIFT_UP);
    assert_eq!(
        formula_of(&mut runtime, worksheet, "F1"),
        OmValue::Text("=SUM(E1:E4)".to_string())
    );
    assert_eq!(
        formula_of(&mut runtime, worksheet, "G1"),
        OmValue::Text("=#REF!*10".to_string())
    );
    assert_eq!(
        runtime
            .dispatch_get(block, "RefersTo", &[])
            .expect("RefersTo"),
        OmValue::Text(format!("={sheet_name}!$E$3:$E$3"))
    );

    shift(
        &mut runtime,
        worksheet,
        "A1:A1048576",
        "Insert",
        XL_SHIFT_TO_RIGHT,
    );
    assert_eq!(
        formula_of(&mut runtime, worksheet, "G1"),
        OmValue::Text("=SUM(F1:F4)".to_string()),
        "the SUM formula moved from F1 to G1 and now reads column F"
    );
    runtime
        .calculate_workbook_with_report(workbook)
        .expect("calculate after delete and column insert");
    assert_eq!(
        value_of(&mut runtime, worksheet, "G1"),
        OmValue::Number(1.0 + 4.0 + 5.0)
    );
    assert_eq!(
        value_of(&mut runtime, worksheet, "H1"),
        OmValue::from(CellValue::Error(CellError::Ref))
    );
    assert_eq!(
        value_of(&mut runtime, worksheet, "I1"),
        OmValue::Number(4.0)
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
        .expect("save retargeted workbook");
    let sheet_xml = String::from_utf8(
        OpcPackage::from_bytes(&saved)
            .expect("saved package")
            .part("xl/worksheets/sheet1.xml")
            .expect("sheet")
            .bytes
            .clone(),
    )
    .expect("sheet utf-8");
    assert!(sheet_xml.contains("<f>SUM(F1:F4)</f>"), "{sheet_xml}");
}

const STRUCTURED_SHEET: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:G12"/><sheetViews><sheetView workbookViewId="0"><selection activeCell="D10" sqref="D10"/></sheetView></sheetViews><cols><col min="3" max="3" width="20" customWidth="1"/></cols><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="10"><c r="D10"><v>10</v></c></row><row r="12" ht="30" customHeight="1"><c r="G12"><v>12</v></c></row></sheetData><mergeCells count="2"><mergeCell ref="B1:C1"/><mergeCell ref="D10:E11"/></mergeCells><conditionalFormatting sqref="F10:F12"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="D10:D20"><formula1>"a,b"</formula1></dataValidation></dataValidations><hyperlinks><hyperlink ref="G12" location="Sheet1!A1" display="home"/></hyperlinks></worksheet>"#;

fn saved_sheet_xml(runtime: &ExcelRuntime, workbook: WorkbookHandle) -> (Vec<u8>, String) {
    let saved = runtime
        .save_workbook(
            workbook,
            SaveWorkbookSpec {
                format: FileFormat::Xlsx,
                profile: ExcelProfile::Excel365,
                lossless: true,
            },
        )
        .expect("save structured workbook");
    let sheet = String::from_utf8(
        OpcPackage::from_bytes(&saved)
            .expect("saved package")
            .part("xl/worksheets/sheet1.xml")
            .expect("sheet")
            .bytes
            .clone(),
    )
    .expect("sheet utf-8");
    (saved, sheet)
}

#[test]
fn whole_row_and_column_shifts_move_worksheet_structure() {
    let mut package = OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("package");
    package
        .replace_part_bytes(
            "xl/worksheets/sheet1.xml",
            STRUCTURED_SHEET.as_bytes().to_vec(),
        )
        .expect("replace sheet");
    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: package.to_bytes().expect("bytes"),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open structured workbook");
    let worksheet = worksheet_handle(&mut runtime, workbook);

    shift(&mut runtime, worksheet, "A5:XFD6", "Insert", XL_SHIFT_DOWN);
    let (_, sheet) = saved_sheet_xml(&runtime, workbook);
    for expected in [
        r#"<selection activeCell="D12" sqref="D12"/>"#,
        r#"<col min="3" max="3" width="20" customWidth="1"/>"#,
        r#"<row r="14" ht="30" customHeight="1"><c r="G14"><v>12</v></c></row>"#,
        r#"<mergeCell ref="B1:C1"/><mergeCell ref="D12:E13"/>"#,
        r#"<conditionalFormatting sqref="F12:F14">"#,
        r#"<dataValidation type="list" allowBlank="1" sqref="D12:D22"><formula1>"a,b"</formula1>"#,
        r#"<hyperlink ref="G14" location="Sheet1!A1" display="home"/>"#,
    ] {
        assert!(
            sheet.contains(expected),
            "after row insert, missing {expected} in:\n{sheet}"
        );
    }

    shift(
        &mut runtime,
        worksheet,
        "B1:B1048576",
        "Insert",
        XL_SHIFT_TO_RIGHT,
    );
    let (saved, sheet) = saved_sheet_xml(&runtime, workbook);
    for expected in [
        r#"<selection activeCell="E12" sqref="E12"/>"#,
        r#"<col min="4" max="4" width="20" customWidth="1"/>"#,
        r#"<c r="C1"><v>2</v></c>"#,
        r#"<mergeCell ref="C1:D1"/><mergeCell ref="E12:F13"/>"#,
        r#"<conditionalFormatting sqref="G12:G14">"#,
        r#"sqref="E12:E22""#,
        r#"<hyperlink ref="H14""#,
    ] {
        assert!(
            sheet.contains(expected),
            "after column insert, missing {expected} in:\n{sheet}"
        );
    }

    let target = range_handle(&mut runtime, worksheet, "A13:XFD13");
    let error = runtime
        .dispatch_invoke(target, "Delete", &[OmValue::Number(f64::from(XL_SHIFT_UP))])
        .expect_err("deleting through a merged range fails closed");
    assert_eq!(error.code, OmErrorCode::Unsupported);
    assert!(error.message.contains("merged-cell"), "{error:?}");

    let mut reopened = ExcelRuntime::new();
    let reopened_workbook = reopened
        .open_workbook(OpenWorkbookSpec {
            bytes: saved,
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("reopen structured workbook");
    let reopened_sheet = worksheet_handle(&mut reopened, reopened_workbook);
    shift(
        &mut reopened,
        reopened_sheet,
        "A1:XFD2",
        "Delete",
        XL_SHIFT_UP,
    );
    let (_, sheet) = saved_sheet_xml(&reopened, reopened_workbook);
    for expected in [
        r#"<mergeCell ref="E10:F11"/>"#,
        r#"<hyperlink ref="H12""#,
        r#"<row r="12" ht="30" customHeight="1">"#,
        r#"<col min="4" max="4" width="20" customWidth="1"/>"#,
    ] {
        assert!(
            sheet.contains(expected),
            "after reopen and delete, missing {expected} in:\n{sheet}"
        );
    }
    assert!(
        !sheet.contains("C1:D1"),
        "the merge on deleted row 1 is removed:\n{sheet}"
    );
}

#[test]
fn partial_corridor_shifts_move_worksheet_structure_inside_the_band() {
    let sheet_xml = STRUCTURED_SHEET
        .replace(r#"sqref="F10:F12""#, r#"sqref="E10:G12""#)
        .replace(r#"<hyperlink ref="G12""#, r#"<hyperlink ref="G12:H12""#);
    let mut package = OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("package");
    package
        .replace_part_bytes("xl/worksheets/sheet1.xml", sheet_xml.into_bytes())
        .expect("replace sheet");
    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: package.to_bytes().expect("bytes"),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open structured workbook");
    let worksheet = worksheet_handle(&mut runtime, workbook);

    // Cells D5:F6 shift down: everything inside columns D:F from row 5 moves two rows, the
    // conditional format the band cuts splits, and rows and columns keep their metadata.
    shift(&mut runtime, worksheet, "D5:F6", "Insert", XL_SHIFT_DOWN);
    let (saved, sheet) = saved_sheet_xml(&runtime, workbook);
    for expected in [
        r#"<selection activeCell="D12" sqref="D12"/>"#,
        r#"<col min="3" max="3" width="20" customWidth="1"/>"#,
        r#"<row r="12" ht="30" customHeight="1"><c r="D12"><v>10</v></c><c r="G12"><v>12</v></c></row>"#,
        r#"<mergeCell ref="B1:C1"/><mergeCell ref="D12:E13"/>"#,
        r#"<conditionalFormatting sqref="E12:F14 G10:G12">"#,
        r#"sqref="D12:D22""#,
        r#"<hyperlink ref="G12:H12""#,
    ] {
        assert!(
            sheet.contains(expected),
            "after band insert, missing {expected} in:\n{sheet}"
        );
    }

    // A band through part of the hyperlink's range would split it.
    let target = range_handle(&mut runtime, worksheet, "G1");
    assert_structural_failure_is_atomic(
        &mut runtime,
        workbook,
        target,
        "Insert",
        XL_SHIFT_DOWN,
        OmErrorCode::Unsupported,
        &["part of range G12:H12"],
        "band cutting a hyperlink",
    );

    let mut reopened = ExcelRuntime::new();
    let reopened_workbook = reopened
        .open_workbook(OpenWorkbookSpec {
            bytes: saved,
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("reopen band-shifted workbook");
    let reopened_sheet = worksheet_handle(&mut reopened, reopened_workbook);
    shift(
        &mut reopened,
        reopened_sheet,
        "D1:E2",
        "Delete",
        XL_SHIFT_UP,
    );
    let (_, sheet) = saved_sheet_xml(&reopened, reopened_workbook);
    for expected in [
        r#"<mergeCell ref="D10:E11"/>"#,
        r#"<c r="D10"><v>10</v></c>"#,
        r#"sqref="D10:D20""#,
    ] {
        assert!(
            sheet.contains(expected),
            "after reopen and band delete, missing {expected} in:\n{sheet}"
        );
    }
}

const RULE_SHEET: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><dimension ref="A1:A4"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData><conditionalFormatting sqref="B2:B4"><cfRule type="expression" dxfId="0" priority="1"><formula>$A2&gt;$A$1</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="custom" sqref="C2:C4"><formula1>$A2&gt;0</formula1></dataValidation></dataValidations><extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression" priority="2" id="{00000000-0000-0000-0000-000000000001}"><xm:f>$A$1&lt;3</xm:f><x14:dxf/></x14:cfRule><xm:sqref>D2:D4</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext><ext uri="{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}"><x14:dataValidations count="1"><x14:dataValidation type="custom"><x14:formula1><xm:f>$A$2</xm:f></x14:formula1><xm:sqref>E2:E4</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>"#;

#[test]
fn structural_shifts_retarget_conditional_format_and_validation_formulas() {
    let mut package = OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("package");
    package
        .replace_part_bytes("xl/worksheets/sheet1.xml", RULE_SHEET.as_bytes().to_vec())
        .expect("replace sheet");
    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: package.to_bytes().expect("bytes"),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open rule workbook");
    let worksheet = worksheet_handle(&mut runtime, workbook);

    // Ranges and rule formulas move together, so relative rules keep reading the same cells.
    shift(&mut runtime, worksheet, "A1:XFD1", "Insert", XL_SHIFT_DOWN);
    let (saved, sheet) = saved_sheet_xml(&runtime, workbook);
    for expected in [
        r#"<conditionalFormatting sqref="B3:B5"><cfRule type="expression" dxfId="0" priority="1"><formula>$A3&gt;$A$2</formula>"#,
        r#"<dataValidation type="custom" sqref="C3:C5"><formula1>$A3&gt;0</formula1>"#,
        r#"<xm:f>$A$2&lt;3</xm:f><x14:dxf/></x14:cfRule><xm:sqref>D3:D5</xm:sqref>"#,
        r#"<xm:f>$A$3</xm:f></x14:formula1><xm:sqref>E3:E5</xm:sqref>"#,
    ] {
        assert!(
            sheet.contains(expected),
            "after row insert, missing {expected} in:\n{sheet}"
        );
    }

    let reopened = runtime
        .codec
        .load(&saved, LoadOptions::default())
        .expect("reopen rule workbook");
    let owners = &reopened
        .state
        .worksheet_data_for_sheet(reopened.state.worksheets()[0].id)
        .expect("reopened worksheet data")
        .structural_owners;
    assert_eq!(
        owners
            .conditional_formats
            .iter()
            .map(|format| (
                format.ranges.clone(),
                format.formulas.clone(),
                format.extension
            ))
            .collect::<Vec<_>>(),
        vec![
            (
                vec![Rect {
                    row_first: 3,
                    row_last: 5,
                    col_first: 2,
                    col_last: 2,
                }],
                vec!["$A3>$A$2".to_string()],
                false,
            ),
            (
                vec![Rect {
                    row_first: 3,
                    row_last: 5,
                    col_first: 4,
                    col_last: 4,
                }],
                vec!["$A$2<3".to_string()],
                true,
            ),
        ],
    );

    // A band through part of a conditional format with relative rules would give its parts
    // different anchors.
    let target = range_handle(&mut runtime, worksheet, "B4");
    assert_structural_failure_is_atomic(
        &mut runtime,
        workbook,
        target,
        "Insert",
        XL_SHIFT_TO_RIGHT,
        OmErrorCode::Unsupported,
        &["structural conditional-format retarget", "R3C2:R5C2"],
        "band cutting a conditional format",
    );
}

#[test]
fn whole_row_shifts_move_comments_notes_and_hyperlinks() {
    let mut package =
        OpcPackage::from_bytes(&synthetic_comment_workbook_bytes()).expect("comment package");
    package
        .replace_part_bytes(
            "xl/drawings/vmlDrawing1.vml",
            br#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="_x0000_s1025"><x:ClientData ObjectType="Note"><x:Anchor>1, 15, 0, 2, 3, 15, 4, 16</x:Anchor><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>"#
                .to_vec(),
        )
        .expect("replace VML");
    // A threaded comment pairs with the legacy comment on A1.
    const THREADED_TYPE: &str = "application/vnd.ms-excel.threadedcomments+xml";
    for (part_name, from, to) in [
        (
            "[Content_Types].xml",
            "</Types>",
            r#"<Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/></Types>"#,
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            "</Relationships>",
            r#"<Relationship Id="rIdThreaded1" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/></Relationships>"#,
        ),
    ] {
        let xml = String::from_utf8(package.part(part_name).expect(part_name).bytes.clone())
            .expect("utf-8")
            .replace(from, to);
        package
            .replace_part_bytes(part_name, xml.into_bytes())
            .expect(part_name);
    }
    package
        .add_part(OpcPart {
            name: "xl/threadedComments/threadedComment1.xml".to_string(),
            content_type: Some(THREADED_TYPE.to_string()),
            compression: CompressionMethod::Deflated,
            bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="A1" dT="2026-01-01T00:00:00.00" personId="{00000000-0000-0000-0000-000000000001}" id="{00000000-0000-0000-0000-000000000002}"><text>Note</text></threadedComment></ThreadedComments>"#
                .to_vec(),
        })
        .expect("add threaded comment part");
    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: package.to_bytes().expect("bytes"),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open comment workbook");
    let worksheet = worksheet_handle(&mut runtime, workbook);

    shift(&mut runtime, worksheet, "A1:XFD2", "Insert", XL_SHIFT_DOWN);
    let (saved, sheet) = saved_sheet_xml(&runtime, workbook);
    assert!(
        sheet.contains(r#"<hyperlink ref="C3" r:id="rId1"/>"#),
        "{sheet}"
    );
    let saved_package = OpcPackage::from_bytes(&saved).expect("saved package");
    let part = |name: &str| {
        String::from_utf8(saved_package.part(name).expect(name).bytes.clone()).expect("utf-8")
    };
    let comments = part("xl/comments1.xml");
    assert!(
        comments.contains(r#"<comment ref="A3" authorId="0">"#),
        "{comments}"
    );
    let threaded = part("xl/threadedComments/threadedComment1.xml");
    assert!(
        threaded.contains(r#"<threadedComment ref="A3" "#),
        "{threaded}"
    );
    let vml = part("xl/drawings/vmlDrawing1.vml");
    assert!(
        vml.contains(
            "<x:Anchor>1, 15, 2, 2, 3, 15, 6, 16</x:Anchor><x:Row>2</x:Row><x:Column>0</x:Column>"
        ),
        "{vml}"
    );

    let target = range_handle(&mut runtime, worksheet, "A3:XFD3");
    let error = runtime
        .dispatch_invoke(target, "Delete", &[OmValue::Number(f64::from(XL_SHIFT_UP))])
        .expect_err("deleting a cell with a threaded comment fails closed");
    assert_eq!(error.code, OmErrorCode::Unsupported);
    assert!(
        error
            .message
            .contains("structural threaded comment removal"),
        "{error:?}"
    );

    let mut reopened = ExcelRuntime::new();
    let reopened_workbook = reopened
        .open_workbook(OpenWorkbookSpec {
            bytes: saved,
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("reopen comment workbook");
    let reopened_sheet = worksheet_handle(&mut reopened, reopened_workbook);
    shift(
        &mut reopened,
        reopened_sheet,
        "A1:XFD1",
        "Delete",
        XL_SHIFT_UP,
    );
    let (saved, _) = saved_sheet_xml(&reopened, reopened_workbook);
    let comments = String::from_utf8(
        OpcPackage::from_bytes(&saved)
            .expect("package")
            .part("xl/comments1.xml")
            .expect("comments")
            .bytes
            .clone(),
    )
    .expect("utf-8");
    assert!(
        comments.contains(r#"<comment ref="A2" authorId="0">"#),
        "{comments}"
    );
}

#[test]
fn structural_deletes_remove_comments_on_deleted_cells() {
    let mut package =
        OpcPackage::from_bytes(&synthetic_comment_workbook_bytes()).expect("comment package");
    package
        .replace_part_bytes(
            "xl/comments1.xml",
            br#"<?xml version="1.0" encoding="UTF-8"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Codex</author></authors><commentList><comment ref="A1" authorId="0"><text><t>Remove me</t></text></comment><comment ref="B5" authorId="0"><text><t>Keep me</t></text></comment></commentList></comments>"#
                .to_vec(),
        )
        .expect("replace comments");
    package
        .replace_part_bytes(
            "xl/drawings/vmlDrawing1.vml",
            br#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shapetype id="_x0000_t202"/><v:shape id="_x0000_s1025"><x:ClientData ObjectType="Note"><x:Anchor>1, 15, 0, 2, 3, 15, 4, 16</x:Anchor><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape><v:shape id="_x0000_s1026"><x:ClientData ObjectType="Note"><x:Anchor>1, 15, 4, 2, 3, 15, 8, 16</x:Anchor><x:Row>4</x:Row><x:Column>1</x:Column></x:ClientData></v:shape></xml>"#
                .to_vec(),
        )
        .expect("replace VML");
    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: package.to_bytes().expect("bytes"),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open comment workbook");
    let worksheet = worksheet_handle(&mut runtime, workbook);

    // Deleting column A removes the comment on A1 and its note; the comment on B5 moves to A5.
    shift(
        &mut runtime,
        worksheet,
        "A1:A1048576",
        "Delete",
        XL_SHIFT_TO_LEFT,
    );
    let (saved, _) = saved_sheet_xml(&runtime, workbook);
    let saved_package = OpcPackage::from_bytes(&saved).expect("saved package");
    let part = |name: &str| {
        String::from_utf8(saved_package.part(name).expect(name).bytes.clone()).expect("utf-8")
    };
    let comments = part("xl/comments1.xml");
    assert!(!comments.contains("Remove me"), "{comments}");
    assert!(
        comments.contains(r#"<comment ref="A5" authorId="0"><text><t>Keep me</t>"#),
        "{comments}"
    );
    let vml = part("xl/drawings/vmlDrawing1.vml");
    assert!(!vml.contains("_x0000_s1025"), "{vml}");
    assert!(vml.contains(r#"<v:shapetype id="_x0000_t202"/>"#), "{vml}");
    assert!(
        vml.contains(
            "<x:Anchor>0, 15, 4, 2, 2, 15, 8, 16</x:Anchor><x:Row>4</x:Row><x:Column>0</x:Column>"
        ),
        "{vml}"
    );

    // Removing the last comment would leave an empty comments part.
    let target = range_handle(&mut runtime, worksheet, "A1:A1048576");
    assert_structural_failure_is_atomic(
        &mut runtime,
        workbook,
        target,
        "Delete",
        XL_SHIFT_TO_LEFT,
        OmErrorCode::Unsupported,
        &["removal of every comment"],
        "deleting the last comment",
    );
}

#[test]
fn whole_row_shifts_move_chart_sources_and_frames() {
    let mut package = OpcPackage::from_bytes(&synthetic_workbook_with_embedded_chart_bytes())
        .expect("chart package");
    let drawing = String::from_utf8(
        package
            .part("xl/drawings/drawing1.xml")
            .expect("drawing")
            .bytes
            .clone(),
    )
    .expect("drawing utf-8");
    let start = drawing.find("<xdr:absoluteAnchor").expect("anchor start");
    let frame = drawing.find("<xdr:graphicFrame").expect("frame start");
    let end = drawing.find("</xdr:absoluteAnchor>").expect("anchor end");
    let drawing = format!(
        "{}<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>{}</xdr:twoCellAnchor>{}",
        &drawing[..start],
        &drawing[frame..end],
        &drawing[end + "</xdr:absoluteAnchor>".len()..],
    );
    package
        .replace_part_bytes("xl/drawings/drawing1.xml", drawing.into_bytes())
        .expect("replace drawing");
    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: package.to_bytes().expect("bytes"),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open chart workbook");
    let worksheet = worksheet_handle(&mut runtime, workbook);

    shift(&mut runtime, worksheet, "A1:XFD2", "Insert", XL_SHIFT_DOWN);
    let (saved, _) = saved_sheet_xml(&runtime, workbook);
    let saved_package = OpcPackage::from_bytes(&saved).expect("saved package");
    let part = |name: &str| {
        String::from_utf8(saved_package.part(name).expect(name).bytes.clone()).expect("utf-8")
    };
    let chart = part("xl/charts/chart1.xml");
    for expected in [
        "<c:f>Sheet1!$C$3</c:f>",
        "<c:f>Sheet1!$A$3:$B$3</c:f>",
        "<c:f>Sheet1!$A$3:$C$3</c:f>",
    ] {
        assert!(chart.contains(expected), "missing {expected} in:\n{chart}");
    }
    assert!(chart.contains("Revenue Trend"), "the chart keeps its title");
    let drawing = part("xl/drawings/drawing1.xml");
    for expected in [
        "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row>",
        "<xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row>",
    ] {
        assert!(
            drawing.contains(expected),
            "missing {expected} in:\n{drawing}"
        );
    }
    assert!(
        drawing.contains("Embedded Revenue Chart"),
        "the frame keeps its properties"
    );

    let mut reopened = ExcelRuntime::new();
    reopened
        .open_workbook(OpenWorkbookSpec {
            bytes: saved,
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("reopen chart workbook");
}

fn table_fixture_bytes(table_xml: &str) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("base package");
    let text = |package: &OpcPackage, name: &str| {
        String::from_utf8(package.part(name).expect(name).bytes.clone()).expect("utf-8")
    };
    let content_types = text(&package, "[Content_Types].xml").replace(
        "</Types>",
        "<Override PartName=\"/xl/tables/table1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/></Types>",
    );
    package
        .replace_part_bytes("[Content_Types].xml", content_types.into_bytes())
        .expect("content types");
    let sheet = text(&package, "xl/worksheets/sheet1.xml").replace(
        "</worksheet>",
        r#"<tableParts xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="1"><tablePart r:id="rIdTable1"/></tableParts></worksheet>"#,
    );
    package
        .replace_part_bytes("xl/worksheets/sheet1.xml", sheet.into_bytes())
        .expect("worksheet");
    for (name, bytes) in [
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>"#.to_vec(),
        ),
        ("xl/tables/table1.xml", table_xml.as_bytes().to_vec()),
    ] {
        package
            .add_part(OpcPart {
                name: name.to_string(),
                content_type: None,
                compression: CompressionMethod::Stored,
                bytes,
            })
            .expect("add part");
    }
    package.to_bytes().expect("table workbook bytes")
}

#[test]
fn whole_row_and_column_shifts_move_tables() {
    let bytes = table_fixture_bytes(
        r#"<?xml version="1.0" encoding="UTF-8"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="D4:E6" totalsRowShown="0"><autoFilter ref="D4:E6"/><tableColumns count="2"><tableColumn id="1" name="Left"/><tableColumn id="2" name="Right"/></tableColumns></table>"#,
    );
    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes,
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open table workbook");
    let worksheet = worksheet_handle(&mut runtime, workbook);
    let table_xml = |runtime: &ExcelRuntime| {
        let (saved, _) = saved_sheet_xml(runtime, workbook);
        let package = OpcPackage::from_bytes(&saved).expect("saved package");
        let table = String::from_utf8(
            package
                .part("xl/tables/table1.xml")
                .expect("table")
                .bytes
                .clone(),
        )
        .expect("utf-8");
        (saved, table)
    };

    shift(&mut runtime, worksheet, "A5:XFD5", "Insert", XL_SHIFT_DOWN);
    shift(
        &mut runtime,
        worksheet,
        "A1:A1048576",
        "Insert",
        XL_SHIFT_TO_RIGHT,
    );
    let (_, table) = table_xml(&runtime);
    assert!(table.contains(r#"ref="E4:F7""#), "{table}");
    assert!(table.contains(r#"<autoFilter ref="E4:F7"/>"#), "{table}");
    assert!(
        table.contains(r#"<tableColumn id="2" name="Right"/>"#),
        "{table}"
    );

    let target = range_handle(&mut runtime, worksheet, "A4:XFD4");
    let error = runtime
        .dispatch_invoke(target, "Delete", &[OmValue::Number(f64::from(XL_SHIFT_UP))])
        .expect_err("deleting the header row");
    assert_eq!(error.code, OmErrorCode::Unsupported);
    assert!(error.message.contains("table range"), "{error:?}");

    shift(&mut runtime, worksheet, "A5:XFD6", "Delete", XL_SHIFT_UP);
    let (saved, table) = table_xml(&runtime);
    assert!(table.contains(r#"ref="E4:F5""#), "{table}");
    let reopened = runtime
        .codec
        .load(&saved, LoadOptions::default())
        .expect("reopen table workbook");
    let sheet_id = reopened.state.worksheets()[0].id;
    assert_eq!(
        reopened
            .state
            .worksheet_data_for_sheet(sheet_id)
            .expect("worksheet")
            .structural_owners
            .table_owners[0]
            .range,
        Rect {
            row_first: 4,
            row_last: 5,
            col_first: 5,
            col_last: 6,
        }
    );
}

#[test]
fn whole_column_edits_inside_tables_add_and_remove_table_columns() {
    let bytes = table_fixture_bytes(
        r#"<?xml version="1.0" encoding="UTF-8"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="D4:F6" totalsRowShown="0"><autoFilter ref="D4:F6"><filterColumn colId="2"><filters><filter val="x"/></filters></filterColumn></autoFilter><tableColumns count="3"><tableColumn id="1" name="Left"/><tableColumn id="2" name="Mid"/><tableColumn id="3" name="Right"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>"#,
    );
    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes,
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open table workbook");
    let worksheet = worksheet_handle(&mut runtime, workbook);
    let table_xml = |runtime: &ExcelRuntime| {
        let (saved, _) = saved_sheet_xml(runtime, workbook);
        let table = String::from_utf8(
            OpcPackage::from_bytes(&saved)
                .expect("saved package")
                .part("xl/tables/table1.xml")
                .expect("table")
                .bytes
                .clone(),
        )
        .expect("utf-8");
        (saved, table)
    };

    // A column inserted inside the table becomes a new table column with a unique name.
    shift(
        &mut runtime,
        worksheet,
        "E1:E1048576",
        "Insert",
        XL_SHIFT_TO_RIGHT,
    );
    let (saved, table) = table_xml(&runtime);
    for expected in [
        r#"ref="D4:G6""#,
        r#"<autoFilter ref="D4:G6"><filterColumn colId="3">"#,
        r#"<tableColumns count="4"><tableColumn id="1" name="Left"/><tableColumn id="4" name="Column1"/><tableColumn id="2" name="Mid"/><tableColumn id="3" name="Right"/></tableColumns>"#,
    ] {
        assert!(table.contains(expected), "missing {expected} in {table}");
    }
    assert_eq!(
        value_of(&mut runtime, worksheet, "E4"),
        OmValue::Text("Column1".to_string())
    );
    let reopened = runtime
        .codec
        .load(&saved, LoadOptions::default())
        .expect("reopen table workbook");
    assert_eq!(
        reopened
            .state
            .worksheet_data_for_sheet(reopened.state.worksheets()[0].id)
            .expect("worksheet")
            .structural_owners
            .table_owners[0]
            .column_names,
        ["Left", "Column1", "Mid", "Right"],
    );

    // Deleting table columns removes them, and the filter on a deleted column goes with it.
    shift(
        &mut runtime,
        worksheet,
        "F1:G1048576",
        "Delete",
        XL_SHIFT_TO_LEFT,
    );
    let (_, table) = table_xml(&runtime);
    for expected in [
        r#"ref="D4:E6""#,
        r#"<autoFilter ref="D4:E6"></autoFilter>"#,
        r#"<tableColumns count="2"><tableColumn id="1" name="Left"/><tableColumn id="4" name="Column1"/></tableColumns>"#,
    ] {
        assert!(table.contains(expected), "missing {expected} in {table}");
    }

    // Deleting a column that a structured reference may name would leave it dangling.
    set_formula(&mut runtime, worksheet, "J1", "=SUM(Table1[Left])");
    let target = range_handle(&mut runtime, worksheet, "E1:E1048576");
    assert_structural_failure_is_atomic(
        &mut runtime,
        workbook,
        target,
        "Delete",
        XL_SHIFT_TO_LEFT,
        OmErrorCode::Unsupported,
        &["table column removal", "Table1"],
        "deleting a referenced table column",
    );
}

#[test]
fn partial_corridors_beside_shapes_leave_them_in_place() {
    let open = |runtime: &mut ExcelRuntime| {
        runtime
            .open_workbook(OpenWorkbookSpec {
                bytes: synthetic_workbook_with_embedded_chart_and_raw_shape_bytes(),
                format_hint: Some(FileFormat::Xlsx),
                profile: ExcelProfile::Excel365,
                read_only: false,
            })
            .expect("open shape workbook")
    };
    // A band in columns A:B does not reach the shape anchored in column F.
    let mut runtime = ExcelRuntime::new();
    let workbook = open(&mut runtime);
    let worksheet = worksheet_handle(&mut runtime, workbook);
    shift(&mut runtime, worksheet, "A1:B1", "Insert", XL_SHIFT_DOWN);
    let (saved, _) = saved_sheet_xml(&runtime, workbook);
    let drawing = String::from_utf8(
        OpcPackage::from_bytes(&saved)
            .expect("saved package")
            .part("xl/drawings/drawing1.xml")
            .expect("drawing")
            .bytes
            .clone(),
    )
    .expect("utf-8");
    assert!(
        drawing.contains(
            "<xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>"
        ),
        "the shape stays put: {drawing}"
    );

    // A band through the shape's cells would need the shape split from its cells.
    let mut runtime = ExcelRuntime::new();
    let workbook = open(&mut runtime);
    let worksheet = worksheet_handle(&mut runtime, workbook);
    let target = range_handle(&mut runtime, worksheet, "F1");
    assert_structural_failure_is_atomic(
        &mut runtime,
        workbook,
        target,
        "Insert",
        XL_SHIFT_DOWN,
        OmErrorCode::Unsupported,
        &["structural drawing anchor retarget", "opaque anchor"],
        "band reaching a shape",
    );
}

#[test]
fn whole_row_shifts_move_shapes_and_keep_absolute_objects() {
    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: synthetic_workbook_with_embedded_chart_and_raw_shape_bytes(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open shape workbook");
    let worksheet = worksheet_handle(&mut runtime, workbook);
    shift(&mut runtime, worksheet, "A1:XFD2", "Insert", XL_SHIFT_DOWN);
    let (saved, _) = saved_sheet_xml(&runtime, workbook);
    let saved_package = OpcPackage::from_bytes(&saved).expect("saved package");
    let drawing = String::from_utf8(
        saved_package
            .part("xl/drawings/drawing1.xml")
            .expect("drawing")
            .bytes
            .clone(),
    )
    .expect("utf-8");
    assert!(
        drawing.contains("<xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"),
        "the shape moves with its row: {drawing}"
    );
    assert!(
        drawing.contains("<a:t>Keep me</a:t>"),
        "the shape is preserved: {drawing}"
    );
    assert!(
        drawing.contains(r#"<xdr:pos x="25400" y="38100""#),
        "the absolute chart stays put: {drawing}"
    );
    let chart = String::from_utf8(
        saved_package
            .part("xl/charts/chart1.xml")
            .expect("chart")
            .bytes
            .clone(),
    )
    .expect("utf-8");
    assert!(chart.contains("<c:f>Sheet1!$A$3:$C$3</c:f>"), "{chart}");
    let mut reopened = ExcelRuntime::new();
    reopened
        .open_workbook(OpenWorkbookSpec {
            bytes: saved,
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("reopen shape workbook");
}
