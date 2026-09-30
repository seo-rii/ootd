use super::*;

fn open_runtime() -> (ExcelRuntime, WorkbookHandle, ObjectHandle) {
    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: synthetic_workbook_bytes(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open workbook");
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

#[test]
fn implicit_intersection_selects_the_cell_in_the_formula_row_or_column() {
    let (mut runtime, workbook, sheet) = open_runtime();
    for (address, value) in [
        ("P2", 10.0),
        ("P3", 20.0),
        ("P4", 30.0),
        ("Q6", 7.0),
        ("R6", 8.0),
    ] {
        let cell = range(&mut runtime, sheet, address);
        runtime
            .dispatch_set(cell, "Value2", OmValue::Number(value), &[])
            .unwrap_or_else(|error| panic!("seed {address}: {error:?}"));
    }
    let cases = [
        ("S3", "=@P2:P4*2", OmValue::Number(40.0)),
        ("S4", "=@P2:P4+@P3", OmValue::Number(50.0)),
        ("R1", "=@Q6:S6", OmValue::Number(8.0)),
        (
            "S9",
            "=@P2:P4",
            OmValue::from(CellValue::Error(CellError::Value)),
        ),
        ("S10", "=IFERROR(@P2:P4,-1)", OmValue::Number(-1.0)),
        ("S11", "=@SUM(P2:P4)", OmValue::Number(60.0)),
        ("S12", "=IFERROR(#N/A,5)", OmValue::Number(5.0)),
    ];
    for (address, formula, _) in &cases {
        let cell = range(&mut runtime, sheet, address);
        runtime
            .dispatch_set(cell, "Formula", OmValue::Text(formula.to_string()), &[])
            .unwrap_or_else(|error| panic!("set {address}: {error:?}"));
    }
    runtime
        .calculate_workbook_with_report(workbook)
        .expect("calculate implicit intersections");
    for (address, formula, expected) in cases {
        let cell = range(&mut runtime, sheet, address);
        assert_eq!(
            runtime.dispatch_get(cell, "Value2", &[]).expect("value"),
            expected,
            "{address} {formula}"
        );
        assert_eq!(
            expect_text(runtime.dispatch_get(cell, "Formula", &[]).expect("formula")),
            formula,
            "{address} keeps its typed formula"
        );
    }
}
