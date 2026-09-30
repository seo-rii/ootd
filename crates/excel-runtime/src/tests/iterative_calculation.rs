use super::*;
use excel_xlsx::WorkbookIteration;

fn open_runtime() -> (ExcelRuntime, WorkbookHandle, ObjectHandle) {
    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: synthetic_workbook_bytes(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open iterative workbook");
    let sheet = expect_object_handle(
        runtime
            .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
            .expect("ActiveSheet"),
    );
    (runtime, workbook, sheet)
}

fn range(runtime: &mut ExcelRuntime, sheet: ObjectHandle, address: &str) -> ObjectHandle {
    expect_object_handle(
        runtime
            .dispatch_invoke(sheet, "Range", &[OmValue::Text(address.to_string())])
            .unwrap_or_else(|error| panic!("Range({address}): {error:?}")),
    )
}

fn set_formula(runtime: &mut ExcelRuntime, sheet: ObjectHandle, address: &str, formula: &str) {
    let target = range(runtime, sheet, address);
    runtime
        .dispatch_set(target, "Formula", OmValue::Text(formula.to_string()), &[])
        .unwrap_or_else(|error| panic!("set {address}: {error:?}"));
}

fn value(runtime: &mut ExcelRuntime, sheet: ObjectHandle, address: &str) -> OmValue {
    let target = range(runtime, sheet, address);
    runtime
        .dispatch_get(target, "Value2", &[])
        .unwrap_or_else(|error| panic!("{address} Value2: {error:?}"))
}

fn iteration(max_iterations: u32, max_change: f64) -> WorkbookIteration {
    WorkbookIteration {
        enabled: true,
        max_iterations,
        max_change,
    }
}

#[test]
fn iterative_calculation_converges_circular_references() {
    let (mut runtime, workbook, sheet) = open_runtime();
    runtime
        .set_workbook_iteration(workbook, iteration(100, 0.000_001))
        .expect("enable iteration");
    set_formula(&mut runtime, sheet, "M1", "=M2/2+10");
    set_formula(&mut runtime, sheet, "M2", "=M1");

    let report = runtime
        .calculate_workbook_with_report(workbook)
        .expect("iterative calculation");
    assert!(report.circular.is_empty(), "{:?}", report.circular);
    for address in ["M1", "M2"] {
        let OmValue::Number(number) = value(&mut runtime, sheet, address) else {
            panic!("{address} is not numeric");
        };
        assert!((number - 20.0).abs() < 0.000_01, "{address} = {number}");
    }
}

#[test]
fn iterative_calculation_stops_after_the_maximum_iteration_count() {
    let (mut runtime, workbook, sheet) = open_runtime();
    runtime
        .set_workbook_iteration(workbook, iteration(10, 0.001))
        .expect("enable iteration");
    set_formula(&mut runtime, sheet, "M1", "=M1+1");
    runtime
        .calculate_workbook_with_report(workbook)
        .expect("first iterative calculation");
    assert_eq!(value(&mut runtime, sheet, "M1"), OmValue::Number(10.0));
    runtime
        .calculate_workbook_with_report(workbook)
        .expect("second iterative calculation");
    assert_eq!(value(&mut runtime, sheet, "M1"), OmValue::Number(20.0));
}

#[test]
fn circular_references_stay_errors_without_iteration() {
    let (mut runtime, workbook, sheet) = open_runtime();
    assert_eq!(
        runtime.workbook_iteration(workbook).expect("iteration"),
        WorkbookIteration::default()
    );
    set_formula(&mut runtime, sheet, "M1", "=M2/2+10");
    set_formula(&mut runtime, sheet, "M2", "=M1");
    let report = runtime
        .calculate_workbook_with_report(workbook)
        .expect("non-iterative calculation");
    assert_eq!(report.circular.len(), 2);
    assert_eq!(
        value(&mut runtime, sheet, "M1"),
        OmValue::from(CellValue::Error(CellError::Calc))
    );
}

#[test]
fn iteration_settings_persist_through_calc_pr() {
    let (mut runtime, workbook, _) = open_runtime();
    runtime
        .set_workbook_iteration(workbook, iteration(50, 0.01))
        .expect("enable iteration");
    let invalid = runtime
        .set_workbook_iteration(workbook, iteration(0, 0.01))
        .expect_err("zero iterations are rejected");
    assert_eq!(invalid.code, OmErrorCode::InvalidArgument);

    let saved = runtime
        .save_workbook(
            workbook,
            SaveWorkbookSpec {
                format: FileFormat::Xlsx,
                profile: ExcelProfile::Excel365,
                lossless: true,
            },
        )
        .expect("save iteration settings");
    let package = OpcPackage::from_bytes(&saved).expect("saved package");
    let workbook_xml = String::from_utf8(
        package
            .part("xl/workbook.xml")
            .expect("workbook")
            .bytes
            .clone(),
    )
    .expect("workbook utf-8");
    assert!(
        workbook_xml.contains(r#"iterate="1" iterateCount="50" iterateDelta="0.01""#),
        "{workbook_xml}"
    );

    let mut reopened = ExcelRuntime::new();
    let reopened_workbook = reopened
        .open_workbook(OpenWorkbookSpec {
            bytes: saved,
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("reopen iteration settings");
    assert_eq!(
        reopened
            .workbook_iteration(reopened_workbook)
            .expect("reopened iteration"),
        iteration(50, 0.01)
    );
}
