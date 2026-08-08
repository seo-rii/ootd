use super::*;
use crate::{XL_COPY, XL_PASTE_FORMULAS, XL_PASTE_VALUES};

fn open_chart_workbook(runtime: &mut ExcelRuntime) -> (WorkbookHandle, ObjectHandle, ObjectHandle) {
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: synthetic_workbook_with_embedded_chart_bytes(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
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

fn assert_success_does_not_retain_internal_handles(paste_type: i32, paste_name: &str) {
    let label = format!("Copy {paste_name}");
    let mut runtime = ExcelRuntime::new();
    let (_, worksheet, chart) = open_chart_workbook(&mut runtime);
    let source = seed_chart_paste_source(&mut runtime, worksheet, &label);
    let series_collection = expect_object_handle(
        runtime
            .dispatch_get(chart, "SeriesCollection", &[])
            .unwrap_or_else(|error| panic!("{label}: Chart.SeriesCollection: {error:?}")),
    );
    let original_series = expect_object_handle(
        runtime
            .dispatch_invoke(series_collection, "Item", &[OmValue::Number(1.0)])
            .unwrap_or_else(|error| panic!("{label}: SeriesCollection.Item(1): {error:?}")),
    );

    runtime
        .dispatch_invoke(
            source,
            "Find",
            &[OmValue::Text("chart paste handle marker".to_string())],
        )
        .unwrap_or_else(|error| panic!("{label}: seed Find state: {error:?}"));
    runtime
        .dispatch_invoke(source, "Copy", &[])
        .unwrap_or_else(|error| panic!("{label}: arm clipboard: {error:?}"));
    assert_eq!(runtime.cut_copy_mode, Some(XL_COPY), "{label}");

    let session_before = runtime_session_mutation_snapshot(&runtime);
    let mut expected_object_handles = session_before.object_handles.clone();
    expected_object_handles.retain(|object_handle| *object_handle != original_series.0);
    let mut expected_stale_object_handles = session_before.stale_object_handles.clone();
    expected_stale_object_handles.push(original_series.0);
    expected_stale_object_handles.sort_unstable();

    assert_eq!(
        runtime
            .dispatch_invoke(chart, "Paste", &[OmValue::Number(f64::from(paste_type))])
            .unwrap_or_else(|error| panic!("{label}: Chart.Paste: {error:?}")),
        OmValue::Empty,
        "{label}",
    );

    let session_after = runtime_session_mutation_snapshot(&runtime);
    assert_eq!(
        session_after.object_handles, expected_object_handles,
        "{label}: object registry",
    );
    assert_eq!(
        session_after.stale_object_handles, expected_stale_object_handles,
        "{label}: stale object registry",
    );
    assert_eq!(
        session_after.next_object_handle, session_before.next_object_handle,
        "{label}: object allocator",
    );
    assert_eq!(runtime.cut_copy_mode, None, "{label}: CutCopyMode");
    assert!(runtime.clipboard.is_none(), "{label}: clipboard");
    assert!(runtime.find_state.is_none(), "{label}: Find state");
    assert_eq!(
        runtime
            .dispatch_get(original_series, "Values", &[])
            .expect_err("the replaced Series handle should be stale")
            .code,
        OmErrorCode::InvalidState,
        "{label}: original Series handle",
    );
}

#[test]
fn chart_values_paste_success_does_not_retain_internal_handles() {
    assert_success_does_not_retain_internal_handles(XL_PASTE_VALUES, "xlPasteValues");
}

#[test]
fn chart_formula_paste_success_does_not_retain_internal_handles() {
    assert_success_does_not_retain_internal_handles(XL_PASTE_FORMULAS, "xlPasteFormulas");
}

#[test]
fn chart_formula_paste_cross_workbook_failure_does_not_retain_internal_handles() {
    let label = "cross-workbook Copy xlPasteFormulas";
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
    let (destination_workbook, _, chart) = open_chart_workbook(&mut runtime);

    runtime
        .dispatch_invoke(
            source,
            "Find",
            &[OmValue::Text("chart paste failure marker".to_string())],
        )
        .unwrap_or_else(|error| panic!("{label}: seed Find state: {error:?}"));
    runtime
        .dispatch_invoke(source, "Copy", &[])
        .unwrap_or_else(|error| panic!("{label}: arm clipboard: {error:?}"));

    let source_before = runtime_workbook_persistence_snapshot(&runtime, source_workbook);
    let destination_before = runtime_workbook_persistence_snapshot(&runtime, destination_workbook);
    let source_dirty_before = runtime
        .workbook_dirty_domains(source_workbook)
        .unwrap_or_else(|error| panic!("{label}: source dirty domains: {error:?}"));
    let destination_dirty_before = runtime
        .workbook_dirty_domains(destination_workbook)
        .unwrap_or_else(|error| panic!("{label}: destination dirty domains: {error:?}"));
    let session_before = runtime_session_mutation_snapshot(&runtime);

    let error = runtime
        .dispatch_invoke(
            chart,
            "Paste",
            &[OmValue::Number(f64::from(XL_PASTE_FORMULAS))],
        )
        .expect_err("cross-workbook Chart.Paste should fail");

    assert_eq!(error.code, OmErrorCode::Unsupported, "{label}: {error:?}");
    assert_eq!(
        error.message, "Chart.SetSourceData cross-workbook ranges are not supported",
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
        runtime
            .workbook_dirty_domains(source_workbook)
            .unwrap_or_else(|error| panic!("{label}: source dirty domains after: {error:?}")),
        source_dirty_before,
        "{label}: source dirty domains",
    );
    assert_eq!(
        runtime
            .workbook_dirty_domains(destination_workbook)
            .unwrap_or_else(|error| {
                panic!("{label}: destination dirty domains after: {error:?}")
            }),
        destination_dirty_before,
        "{label}: destination dirty domains",
    );
    assert_eq!(
        runtime_session_mutation_snapshot(&runtime),
        session_before,
        "{label}: runtime session",
    );
}
