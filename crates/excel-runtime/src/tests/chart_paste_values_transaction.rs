use super::*;
use crate::{XL_COPY, XL_PASTE_VALUES};

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

fn seed_non_finite_chart_paste_source(
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
    // Deliberately bypass the public finite-number boundary to fault-inject the later
    // Series.Values validation path exercised by this transaction rollback test.
    let (workbook, sheet_id) = match runtime
        .runtime_object(worksheet)
        .unwrap_or_else(|error| panic!("{label}: worksheet object: {error:?}"))
    {
        RuntimeObjectKind::Worksheet { workbook, sheet_id } => (workbook, sheet_id),
        other => panic!("{label}: expected worksheet object, got {other:?}"),
    };
    runtime
        .runtime_workbook_mut(workbook)
        .unwrap_or_else(|error| panic!("{label}: runtime workbook: {error:?}"))
        .loaded
        .state
        .worksheet_data_for_sheet_mut(sheet_id)
        .unwrap_or_else(|error| panic!("{label}: worksheet data: {error:?}"))
        .cells
        .get_mut(&(3, 3))
        .expect("seeded C3 source cell")
        .value = CellValue::Number(f64::INFINITY);
    source
}

fn assert_values_paste_failure_is_atomic(
    runtime: &mut ExcelRuntime,
    source_workbook: WorkbookHandle,
    destination_workbook: WorkbookHandle,
    source: ObjectHandle,
    chart: ObjectHandle,
    label: &str,
) {
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
    let original_series_values = runtime
        .dispatch_get(original_series, "Values", &[])
        .unwrap_or_else(|error| panic!("{label}: original Series.Values: {error:?}"));

    runtime
        .dispatch_invoke(
            source,
            "Find",
            &[OmValue::Text("chart paste rollback marker".to_string())],
        )
        .unwrap_or_else(|error| panic!("{label}: seed Find state: {error:?}"));
    runtime
        .dispatch_invoke(source, "Copy", &[])
        .unwrap_or_else(|error| panic!("{label}: arm clipboard: {error:?}"));

    assert_eq!(runtime.cut_copy_mode, Some(XL_COPY), "{label}");
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
        .dispatch_invoke(
            chart,
            "Paste",
            &[OmValue::Number(f64::from(XL_PASTE_VALUES))],
        )
        .expect_err("Chart.Paste values should reject non-finite source values");

    assert_eq!(
        error.code,
        OmErrorCode::InvalidArgument,
        "{label}: {error:?}"
    );
    assert_eq!(
        error.message, "Series.Values array values must be finite numbers",
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
            .unwrap_or_else(|error| {
                panic!("{label}: destination dirty domains after: {error:?}")
            }),
        destination_dirty_before,
        "{label}: destination dirty domains",
    );
    assert_eq!(
        runtime_session_mutation_snapshot(runtime),
        session_before,
        "{label}: runtime session",
    );
    assert_eq!(
        runtime
            .dispatch_get(original_series, "Values", &[])
            .unwrap_or_else(|error| panic!("{label}: original Series handle: {error:?}")),
        original_series_values,
        "{label}: original Series handle and values",
    );
}

#[test]
fn chart_values_paste_late_validation_failure_is_atomic_in_same_workbook() {
    let label = "same-workbook Copy xlPasteValues";
    let mut runtime = ExcelRuntime::new();
    let (workbook, worksheet, chart) = open_chart_workbook(&mut runtime);
    let source = seed_non_finite_chart_paste_source(&mut runtime, worksheet, label);

    assert_values_paste_failure_is_atomic(&mut runtime, workbook, workbook, source, chart, label);
}

#[test]
fn chart_values_paste_late_validation_failure_is_atomic_across_workbooks() {
    let label = "cross-workbook Copy xlPasteValues";
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
    let source = seed_non_finite_chart_paste_source(&mut runtime, source_worksheet, label);
    let (destination_workbook, _, chart) = open_chart_workbook(&mut runtime);

    assert_values_paste_failure_is_atomic(
        &mut runtime,
        source_workbook,
        destination_workbook,
        source,
        chart,
        label,
    );
}
