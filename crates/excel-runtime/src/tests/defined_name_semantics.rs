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
        // Text passed straight to an aggregate is coerced as VALUE coerces it.
        ("R17", "=SUM(\"3\",P1&\"\")", OmValue::Number(13.0)),
        (
            "R18",
            "=SUM(\"x\"&1)",
            OmValue::from(CellValue::Error(CellError::Value)),
        ),
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

#[test]
fn relative_name_references_resolve_at_the_calling_cell() {
    let (mut runtime, workbook, sheet, _) = open_runtime();
    let sheet_name = runtime.worksheets(workbook).expect("worksheets")[0]
        .name
        .clone();
    for row in 1..=4 {
        let cell = range(&mut runtime, sheet, &format!("P{row}"));
        runtime
            .dispatch_set(cell, "Value2", OmValue::Number(f64::from(row * 10)), &[])
            .expect("seed value");
    }
    // As Excel writes them: relative to A1, so `A1048576` is "the row above" in the same column
    // and `XFD1` is "the column to the left".
    add_name(
        &mut runtime,
        workbook,
        NameScope::Workbook,
        "Above",
        &format!("{sheet_name}!A1048576"),
    );
    add_name(
        &mut runtime,
        workbook,
        NameScope::Workbook,
        "Left",
        &format!("{sheet_name}!XFD1"),
    );
    add_name(
        &mut runtime,
        workbook,
        NameScope::Workbook,
        "Fixed",
        &format!("{sheet_name}!$P$1"),
    );
    add_name(
        &mut runtime,
        workbook,
        NameScope::Workbook,
        "RunToHere",
        &format!("{sheet_name}!$P$1:$P1048576"),
    );

    let cases = [
        ("P3", "=Above+1", 21.0),
        ("Q2", "=Left*2", 40.0),
        ("Q4", "=Fixed+Left", 50.0),
        ("P9", "=SUM(RunToHere)", 91.0),
    ];
    for (address, formula, _) in &cases {
        let cell = range(&mut runtime, sheet, address);
        runtime
            .dispatch_set(cell, "Formula", OmValue::Text(formula.to_string()), &[])
            .unwrap_or_else(|error| panic!("set {address}: {error:?}"));
    }
    runtime
        .calculate_workbook_with_report(workbook)
        .expect("calculate relative names");
    for (address, formula, expected) in cases {
        let cell = range(&mut runtime, sheet, address);
        assert_eq!(
            runtime.dispatch_get(cell, "Value2", &[]).expect("value"),
            OmValue::Number(expected),
            "{address} {formula}"
        );
    }
}

#[test]
fn names_add_and_refers_to_read_relative_text_at_the_active_cell() {
    let (mut runtime, workbook, sheet, _) = open_runtime();
    let sheet_name = runtime.worksheets(workbook).expect("worksheets")[0]
        .name
        .clone();
    for row in 1..=4 {
        let cell = range(&mut runtime, sheet, &format!("P{row}"));
        runtime
            .dispatch_set(cell, "Value2", OmValue::Number(f64::from(row * 10)), &[])
            .expect("seed value");
    }
    let select = |runtime: &mut ExcelRuntime, address: &str| {
        let cell = range(runtime, sheet, address);
        runtime
            .dispatch_invoke(cell, "Select", &[])
            .unwrap_or_else(|error| panic!("select {address}: {error:?}"));
    };
    let stored = |runtime: &ExcelRuntime, name: &str| {
        runtime
            .runtime_workbook(workbook)
            .expect("runtime workbook")
            .loaded
            .state
            .defined_names
            .lookup_in_scope(NameScope::Workbook, name)
            .expect("defined name")
            .refers_to
            .text
            .clone()
    };
    let names = expect_object_handle(
        runtime
            .dispatch_get(workbook.0, "Names", &[])
            .expect("Workbook.Names"),
    );

    // Typed at B2, `B1` is "the row above" and `$P$1` stays put.
    select(&mut runtime, "B2");
    let above = expect_object_handle(
        runtime
            .dispatch_invoke(
                names,
                "Add",
                &[
                    OmValue::Text("Above".to_string()),
                    OmValue::Text(format!("={sheet_name}!B1+{sheet_name}!$P$1")),
                ],
            )
            .expect("Names.Add"),
    );
    assert_eq!(
        stored(&runtime, "Above"),
        format!("{sheet_name}!A1048576+{sheet_name}!$P$1")
    );
    assert_eq!(
        runtime
            .dispatch_get(above, "RefersTo", &[])
            .expect("RefersTo"),
        OmValue::Text(format!("={sheet_name}!B1+{sheet_name}!$P$1"))
    );

    // Read from D5, the same name reads as the row above D5.
    select(&mut runtime, "D5");
    assert_eq!(
        runtime
            .dispatch_get(above, "RefersTo", &[])
            .expect("RefersTo"),
        OmValue::Text(format!("={sheet_name}!D4+{sheet_name}!$P$1"))
    );
    assert_eq!(
        runtime
            .dispatch_get(above, "RefersToR1C1", &[])
            .expect("RefersToR1C1"),
        OmValue::Text(format!("={sheet_name}!R[-1]C+{sheet_name}!R1C16"))
    );

    // Assigned at D5, `C5` is "the column to the left".
    runtime
        .dispatch_set(
            above,
            "RefersTo",
            OmValue::Text(format!("={sheet_name}!C5")),
            &[],
        )
        .expect("set RefersTo");
    assert_eq!(stored(&runtime, "Above"), format!("{sheet_name}!XFD1"));

    let cell = range(&mut runtime, sheet, "Q3");
    runtime
        .dispatch_set(cell, "Formula", OmValue::Text("=Above+1".to_string()), &[])
        .expect("set formula");
    runtime
        .calculate_workbook_with_report(workbook)
        .expect("calculate");
    assert_eq!(
        runtime.dispatch_get(cell, "Value2", &[]).expect("value"),
        OmValue::Number(31.0)
    );
}

#[test]
fn formulas_use_built_in_names_without_their_prefix() {
    let (mut runtime, workbook, sheet, sheet_id) = open_runtime();
    let sheet_name = runtime.worksheets(workbook).expect("worksheets")[0]
        .name
        .clone();
    for row in 1..=3 {
        let cell = range(&mut runtime, sheet, &format!("P{row}"));
        runtime
            .dispatch_set(cell, "Value2", OmValue::Number(f64::from(row * 10)), &[])
            .expect("seed value");
    }
    // Excel stores a print area as the sheet-scoped `_xlnm.Print_Area` and formulas spell it
    // `Print_Area`.
    runtime
        .runtime_workbook_mut(workbook)
        .expect("runtime workbook")
        .loaded
        .state
        .defined_names_mut()
        .add(
            NameScope::Worksheet(sheet_id),
            "_xlnm.Print_Area",
            FormulaSource {
                text: format!("{sheet_name}!$P$1:$P$3"),
                is_r1c1: false,
            },
            NameValidationMode::PreserveLoadedInvalid,
        )
        .expect("add print area");

    let cases = [
        ("R1", "=ROWS(Print_Area)".to_string(), 3.0),
        ("R2", "=SUM(print_area)".to_string(), 60.0),
        ("R3", format!("=SUM({sheet_name}!Print_Area)"), 60.0),
        ("R4", "=INDEX(_xlnm.Print_Area,2)".to_string(), 20.0),
    ];
    for (address, formula, _) in &cases {
        let cell = range(&mut runtime, sheet, address);
        runtime
            .dispatch_set(cell, "Formula", OmValue::Text(formula.clone()), &[])
            .unwrap_or_else(|error| panic!("set {address}: {error:?}"));
    }
    runtime
        .calculate_workbook_with_report(workbook)
        .expect("calculate built-in names");
    for (address, formula, expected) in cases {
        let cell = range(&mut runtime, sheet, address);
        assert_eq!(
            runtime.dispatch_get(cell, "Value2", &[]).expect("value"),
            OmValue::Number(expected),
            "{address} {formula}"
        );
    }
}

#[test]
fn external_workbook_references_read_cached_link_values() {
    // The first link caches sheet `Cached` of `external.xlsx` with A1 = 42.
    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: external_data_fixture::package_with_external_data(&synthetic_workbook_bytes()),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open workbook with external links");
    let sheet = expect_object_handle(
        runtime
            .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
            .expect("ActiveSheet"),
    );
    add_name(
        &mut runtime,
        workbook,
        NameScope::Workbook,
        "Linked",
        "[1]Cached!$A$1",
    );

    let cases = [
        ("R1", "=[1]Cached!A1*2", 84.0),
        ("R2", "='[1]Cached'!A1+1", 43.0),
        ("R3", "=SUM([1]Cached!A1:B2)", 42.0),
        ("R4", "=[external.xlsx]Cached!A1", 42.0),
        ("R5", "=[1]Cached!B5+1", 1.0),
        ("R6", "=Linked+1", 43.0),
    ];
    for (address, formula, _) in &cases {
        let cell = range(&mut runtime, sheet, address);
        runtime
            .dispatch_set(cell, "Formula", OmValue::Text(formula.to_string()), &[])
            .unwrap_or_else(|error| panic!("set {address}: {error:?}"));
    }
    runtime
        .calculate_workbook_with_report(workbook)
        .expect("calculate external references");
    let mut failures = Vec::new();
    for (address, formula, expected) in cases {
        let cell = range(&mut runtime, sheet, address);
        let actual = runtime.dispatch_get(cell, "Value2", &[]).expect("value");
        if actual != OmValue::Number(expected) {
            failures.push(format!(
                "{address} {formula}: expected {expected}, got {actual:?}"
            ));
        }
    }
    assert!(failures.is_empty(), "{failures:#?}");
}
