use super::*;
use office_common::{FormulaSource, NameScope, NameValidationMode};

fn open_runtime() -> (ExcelRuntime, WorkbookHandle, ObjectHandle, SheetId) {
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
    let sheet_id = runtime.worksheets(workbook).expect("worksheets")[0].id;
    (runtime, workbook, sheet, sheet_id)
}

fn add_name(
    runtime: &mut ExcelRuntime,
    workbook: WorkbookHandle,
    scope: NameScope,
    name: &str,
    text: &str,
) {
    runtime
        .runtime_workbook_mut(workbook)
        .expect("runtime workbook")
        .loaded
        .state
        .defined_names_mut()
        .add(
            scope,
            name,
            FormulaSource {
                text: text.to_string(),
                is_r1c1: false,
            },
            NameValidationMode::StrictExcel,
        )
        .unwrap_or_else(|error| panic!("add {name}: {error:?}"));
}

fn range(runtime: &mut ExcelRuntime, sheet: ObjectHandle, address: &str) -> ObjectHandle {
    expect_object_handle(
        runtime
            .dispatch_invoke(sheet, "Range", &[OmValue::Text(address.to_string())])
            .unwrap_or_else(|error| panic!("Range({address}): {error:?}")),
    )
}

#[test]
fn defined_names_evaluate_constants_ranges_and_scopes() {
    let (mut runtime, workbook, sheet, sheet_id) = open_runtime();
    let sheet_name = runtime.worksheets(workbook).expect("worksheets")[0]
        .name
        .clone();
    for (address, value) in [("P1", 10.0), ("P2", 20.0), ("P3", 30.0)] {
        let cell = range(&mut runtime, sheet, address);
        runtime
            .dispatch_set(cell, "Value2", OmValue::Number(value), &[])
            .expect("seed value");
    }
    add_name(&mut runtime, workbook, NameScope::Workbook, "Rate", "0.5");
    add_name(
        &mut runtime,
        workbook,
        NameScope::Workbook,
        "Label",
        "\"total\"",
    );
    add_name(
        &mut runtime,
        workbook,
        NameScope::Workbook,
        "Doubled",
        "Rate*2",
    );
    add_name(
        &mut runtime,
        workbook,
        NameScope::Workbook,
        "Values",
        &format!("{sheet_name}!$P$1:$P$3"),
    );
    add_name(
        &mut runtime,
        workbook,
        NameScope::Workbook,
        "Corners",
        &format!("{sheet_name}!$P$1,{sheet_name}!$P$3"),
    );
    add_name(&mut runtime, workbook, NameScope::Workbook, "Scoped", "1");
    add_name(
        &mut runtime,
        workbook,
        NameScope::Worksheet(sheet_id),
        "Scoped",
        "2",
    );
    add_name(
        &mut runtime,
        workbook,
        NameScope::Workbook,
        "Triple",
        "LAMBDA(x,x*3)",
    );
    add_name(
        &mut runtime,
        workbook,
        NameScope::Workbook,
        "Constants",
        "{1,2,3}",
    );

    let cases = [
        ("R1", "=P1*Rate", OmValue::Number(5.0)),
        ("R2", "=Label&\"!\"", OmValue::Text("total!".to_string())),
        ("R3", "=Doubled", OmValue::Number(1.0)),
        ("R4", "=SUM(Values)", OmValue::Number(60.0)),
        ("R5", "=SUM(Corners)", OmValue::Number(40.0)),
        ("R6", "=Scoped", OmValue::Number(2.0)),
        ("R7", "=Triple(4)", OmValue::Number(12.0)),
        ("R8", "=SUM(Constants)", OmValue::Number(6.0)),
        ("R9", "=ROWS(Values)", OmValue::Number(3.0)),
        ("R10", "=INDEX(Values,2)", OmValue::Number(20.0)),
        (
            "R11",
            "=P1&\" / \"&P2+1",
            OmValue::Text("10 / 21".to_string()),
        ),
        ("R12", "=LEN(\"ab\"&Label)", OmValue::Number(7.0)),
        (
            "R13",
            "=IF(P1>5,\"big \"&P1,\"small\")",
            OmValue::Text("big 10".to_string()),
        ),
        ("R14", "=SUM({1,2;3,4},P1)", OmValue::Number(20.0)),
        ("R15", "=AVERAGE({2,\"a\",TRUE,4})", OmValue::Number(3.0)),
        ("R16", "=MAX(-1,{-5,-2})", OmValue::Number(-1.0)),
    ];
    for (address, formula, _) in &cases {
        let cell = range(&mut runtime, sheet, address);
        runtime
            .dispatch_set(cell, "Formula", OmValue::Text(formula.to_string()), &[])
            .unwrap_or_else(|error| panic!("set {address}: {error:?}"));
    }
    runtime
        .calculate_workbook_with_report(workbook)
        .expect("calculate defined names");
    let mut failures = Vec::new();
    for (address, formula, expected) in cases {
        let cell = range(&mut runtime, sheet, address);
        let actual = runtime.dispatch_get(cell, "Value2", &[]).expect("value");
        if actual != expected {
            failures.push(format!(
                "{address} {formula}: expected {expected:?}, got {actual:?}"
            ));
        }
    }
    assert!(failures.is_empty(), "{failures:#?}");
}
