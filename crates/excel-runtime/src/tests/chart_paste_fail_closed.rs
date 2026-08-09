use super::*;
use crate::{
    XL_COPY, XL_CUT, XL_PASTE_ALL, XL_PASTE_ALL_EXCEPT_BORDERS,
    XL_PASTE_ALL_MERGING_CONDITIONAL_FORMATS, XL_PASTE_ALL_USING_SOURCE_THEME,
    XL_PASTE_COLUMN_WIDTHS, XL_PASTE_COMMENTS, XL_PASTE_FORMATS, XL_PASTE_FORMULAS,
    XL_PASTE_FORMULAS_AND_NUMBER_FORMATS, XL_PASTE_VALIDATION, XL_PASTE_VALUES,
    XL_PASTE_VALUES_AND_NUMBER_FORMATS,
};

fn unsupported_chart_paste_types() -> [(i32, &'static str); 9] {
    [
        (XL_PASTE_ALL_EXCEPT_BORDERS, "xlPasteAllExceptBorders"),
        (
            XL_PASTE_ALL_MERGING_CONDITIONAL_FORMATS,
            "xlPasteAllMergingConditionalFormats",
        ),
        (
            XL_PASTE_ALL_USING_SOURCE_THEME,
            "xlPasteAllUsingSourceTheme",
        ),
        (
            XL_PASTE_FORMULAS_AND_NUMBER_FORMATS,
            "xlPasteFormulasAndNumberFormats",
        ),
        (
            XL_PASTE_VALUES_AND_NUMBER_FORMATS,
            "xlPasteValuesAndNumberFormats",
        ),
        (XL_PASTE_FORMATS, "xlPasteFormats"),
        (XL_PASTE_COMMENTS, "xlPasteComments"),
        (XL_PASTE_VALIDATION, "xlPasteValidation"),
        (XL_PASTE_COLUMN_WIDTHS, "xlPasteColumnWidths"),
    ]
}

fn open_chart_workbook(
    runtime: &mut ExcelRuntime,
    read_only: bool,
) -> (WorkbookHandle, ObjectHandle, ObjectHandle) {
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: synthetic_workbook_with_embedded_chart_bytes(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only,
        })
        .expect("open workbook with embedded chart");
    let worksheet = expect_object_handle(
        runtime
            .dispatch_get(workbook.0, "Worksheets", &[OmValue::Number(1.0)])
            .expect("Workbook.Worksheets(1)"),
    );
    let chart_objects = expect_object_handle(
        runtime
            .dispatch_get(worksheet, "ChartObjects", &[])
            .expect("Worksheet.ChartObjects"),
    );
    let chart_object = expect_object_handle(
        runtime
            .dispatch_invoke(chart_objects, "Item", &[OmValue::Number(1.0)])
            .expect("ChartObjects.Item(1)"),
    );
    let chart = expect_object_handle(
        runtime
            .dispatch_get(chart_object, "Chart", &[])
            .expect("ChartObject.Chart"),
    );
    (workbook, worksheet, chart)
}

fn seed_chart_paste_source(
    runtime: &mut ExcelRuntime,
    worksheet: ObjectHandle,
    label: &str,
) -> ObjectHandle {
    let source = expect_object_handle(
        runtime
            .dispatch_invoke(worksheet, "Range", &[OmValue::Text("B2:C4".to_string())])
            .unwrap_or_else(|error| panic!("{label}: Worksheet.Range(B2:C4): {error:?}")),
    );
    runtime
        .dispatch_set(
            source,
            "Value2",
            OmValue::Array(
                OmArray::new(
                    3,
                    2,
                    vec![
                        OmValue::Number(10.0),
                        OmValue::Number(20.0),
                        OmValue::Number(30.0),
                        OmValue::Number(40.0),
                        OmValue::Number(50.0),
                        OmValue::Number(60.0),
                    ],
                )
                .unwrap_or_else(|error| panic!("{label}: source array: {error:?}")),
            ),
            &[],
        )
        .unwrap_or_else(|error| panic!("{label}: seed source values: {error:?}"));
    source
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

fn assert_chart_paste_failure_preserves_workbooks_and_session(
    runtime: &mut ExcelRuntime,
    source_workbook: WorkbookHandle,
    destination_workbook: WorkbookHandle,
    source: ObjectHandle,
    chart: ObjectHandle,
    clipboard_member: &str,
    paste_type: i32,
    paste_name: &str,
    label: &str,
) {
    runtime
        .dispatch_invoke(
            source,
            "Find",
            &[OmValue::Text("chart paste session marker".to_string())],
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
    assert_eq!(runtime.cut_copy_mode, Some(expected_mode), "{label}");
    let clipboard = runtime
        .clipboard
        .as_ref()
        .unwrap_or_else(|| panic!("{label}: clipboard precondition"));
    assert_eq!(clipboard.mode, expected_mode, "{label}: clipboard mode");
    assert_eq!(
        clipboard.workbook, source_workbook,
        "{label}: clipboard workbook",
    );
    assert!(
        runtime.find_state.is_some(),
        "{label}: Find state precondition"
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

    let error = runtime
        .dispatch_invoke(chart, "Paste", &[OmValue::Number(f64::from(paste_type))])
        .expect_err("unmodeled-format Chart.Paste should fail closed");

    assert_eq!(error.code, OmErrorCode::Unsupported, "{label}: {error:?}");
    assert_eq!(
        error.message,
        format!("Chart.Paste Type {paste_name} is not implemented"),
        "{label}",
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
fn chart_unmodeled_format_paste_is_fail_closed_and_atomic() {
    for (paste_type, paste_name) in unsupported_chart_paste_types() {
        for clipboard_member in ["Copy", "Cut"] {
            let label = format!("{clipboard_member} {paste_name}");
            let mut runtime = ExcelRuntime::new();
            let (workbook, worksheet, chart) = open_chart_workbook(&mut runtime, false);
            let source = seed_chart_paste_source(&mut runtime, worksheet, &label);
            commit_workbook_baseline(&mut runtime, workbook, &label);

            assert_chart_paste_failure_preserves_workbooks_and_session(
                &mut runtime,
                workbook,
                workbook,
                source,
                chart,
                clipboard_member,
                paste_type,
                paste_name,
                &label,
            );
        }
    }
}

#[test]
fn chart_unmodeled_format_paste_is_stably_unsupported_for_read_only_destination() {
    for (paste_type, paste_name) in unsupported_chart_paste_types() {
        for clipboard_member in ["Copy", "Cut"] {
            let label = format!("read-only {clipboard_member} {paste_name}");
            let mut runtime = ExcelRuntime::new();
            let source_workbook = runtime
                .open_workbook(OpenWorkbookSpec {
                    bytes: synthetic_workbook_bytes(),
                    format_hint: Some(FileFormat::Xlsx),
                    profile: ExcelProfile::Excel365,
                    read_only: false,
                })
                .unwrap_or_else(|error| panic!("{label}: open source workbook: {error:?}"));
            let source_worksheet = expect_object_handle(
                runtime
                    .dispatch_get(source_workbook.0, "Worksheets", &[OmValue::Number(1.0)])
                    .unwrap_or_else(|error| panic!("{label}: source Worksheets(1): {error:?}")),
            );
            let source = seed_chart_paste_source(&mut runtime, source_worksheet, &label);
            commit_workbook_baseline(&mut runtime, source_workbook, &label);
            let (destination_workbook, _, chart) = open_chart_workbook(&mut runtime, true);

            assert_chart_paste_failure_preserves_workbooks_and_session(
                &mut runtime,
                source_workbook,
                destination_workbook,
                source,
                chart,
                clipboard_member,
                paste_type,
                paste_name,
                &label,
            );
        }
    }
}

#[test]
fn chart_unmodeled_format_paste_reports_capability_before_clipboard_state() {
    for (paste_type, paste_name) in unsupported_chart_paste_types() {
        let mut runtime = ExcelRuntime::new();
        let (workbook, _, chart) = open_chart_workbook(&mut runtime, false);
        let workbook_before = runtime_workbook_persistence_snapshot(&runtime, workbook);
        let dirty_before = runtime
            .workbook_dirty_domains(workbook)
            .unwrap_or_else(|error| panic!("{paste_name}: dirty domains: {error:?}"));
        let session_before = runtime_session_mutation_snapshot(&runtime);

        let error = runtime
            .dispatch_invoke(chart, "Paste", &[OmValue::Number(f64::from(paste_type))])
            .expect_err("unmodeled-format Chart.Paste should report unsupported capability");

        assert_eq!(error.code, OmErrorCode::Unsupported, "{paste_name}");
        assert_eq!(
            error.message,
            format!("Chart.Paste Type {paste_name} is not implemented"),
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
fn chart_cut_paste_is_fail_closed_and_atomic() {
    for (paste_type, paste_name) in [
        (XL_PASTE_ALL, "xlPasteAll"),
        (XL_PASTE_FORMULAS, "xlPasteFormulas"),
        (XL_PASTE_VALUES, "xlPasteValues"),
    ] {
        let label = format!("Cut {paste_name}");
        let mut runtime = ExcelRuntime::new();
        let (workbook, worksheet, chart) = open_chart_workbook(&mut runtime, false);
        let source = seed_chart_paste_source(&mut runtime, worksheet, &label);
        commit_workbook_baseline(&mut runtime, workbook, &label);
        runtime
            .dispatch_invoke(
                source,
                "Find",
                &[OmValue::Text("chart cut paste marker".to_string())],
            )
            .unwrap_or_else(|error| panic!("{label}: seed Find state: {error:?}"));
        runtime
            .dispatch_invoke(source, "Cut", &[])
            .unwrap_or_else(|error| panic!("{label}: arm Cut clipboard: {error:?}"));

        let workbook_before = runtime_workbook_persistence_snapshot(&runtime, workbook);
        let dirty_before = runtime
            .workbook_dirty_domains(workbook)
            .unwrap_or_else(|error| panic!("{label}: dirty domains: {error:?}"));
        let session_before = runtime_session_mutation_snapshot(&runtime);

        let error = runtime
            .dispatch_invoke(chart, "Paste", &[OmValue::Number(f64::from(paste_type))])
            .expect_err("Chart.Paste must not silently consume a Cut clipboard");

        assert_eq!(error.code, OmErrorCode::Unsupported, "{label}: {error:?}");
        assert_eq!(
            error.message, "Chart.Paste does not support a Cut range clipboard",
            "{label}",
        );
        assert_eq!(
            runtime_workbook_persistence_snapshot(&runtime, workbook),
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
            runtime_session_mutation_snapshot(&runtime),
            session_before,
            "{label}: session",
        );
    }
}

#[test]
fn chart_cut_paste_reports_cut_capability_before_destination_mutability() {
    let label = "read-only cross-workbook Cut xlPasteValues";
    let mut runtime = ExcelRuntime::new();
    let source_workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: synthetic_workbook_bytes(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .unwrap_or_else(|error| panic!("{label}: open source workbook: {error:?}"));
    let source_worksheet = expect_object_handle(
        runtime
            .dispatch_get(source_workbook.0, "Worksheets", &[OmValue::Number(1.0)])
            .unwrap_or_else(|error| panic!("{label}: source Worksheets(1): {error:?}")),
    );
    let source = seed_chart_paste_source(&mut runtime, source_worksheet, label);
    commit_workbook_baseline(&mut runtime, source_workbook, label);
    let (destination_workbook, _, chart) = open_chart_workbook(&mut runtime, true);
    runtime
        .dispatch_invoke(source, "Cut", &[])
        .unwrap_or_else(|error| panic!("{label}: arm Cut clipboard: {error:?}"));

    let source_before = runtime_workbook_persistence_snapshot(&runtime, source_workbook);
    let destination_before = runtime_workbook_persistence_snapshot(&runtime, destination_workbook);
    let session_before = runtime_session_mutation_snapshot(&runtime);

    let error = runtime
        .dispatch_invoke(
            chart,
            "Paste",
            &[OmValue::Number(f64::from(XL_PASTE_VALUES))],
        )
        .expect_err("Chart.Paste Cut must fail before destination mutation checks");

    assert_eq!(error.code, OmErrorCode::Unsupported, "{label}: {error:?}");
    assert_eq!(
        error.message, "Chart.Paste does not support a Cut range clipboard",
        "{label}",
    );
    assert_eq!(
        runtime_workbook_persistence_snapshot(&runtime, source_workbook),
        source_before,
        "{label}: source workbook",
    );
    assert_eq!(
        runtime_workbook_persistence_snapshot(&runtime, destination_workbook),
        destination_before,
        "{label}: destination workbook",
    );
    assert_eq!(
        runtime_session_mutation_snapshot(&runtime),
        session_before,
        "{label}: session",
    );
}
