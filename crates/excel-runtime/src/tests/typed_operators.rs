use super::*;

/// Operators over typed values, as Excel evaluates them. P1 holds 10, P2 the text "abc", P3 the
/// text "3", and P9 is blank.
#[test]
fn operators_coerce_and_compare_typed_values() {
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
    let range = |runtime: &mut ExcelRuntime, address: &str| {
        expect_object_handle(
            runtime
                .dispatch_invoke(sheet, "Range", &[OmValue::Text(address.to_string())])
                .unwrap_or_else(|error| panic!("Range({address}): {error:?}")),
        )
    };
    for (address, value) in [
        ("P1", OmValue::Number(10.0)),
        ("P2", OmValue::Text("abc".to_string())),
        ("P3", OmValue::Text("3".to_string())),
    ] {
        let cell = range(&mut runtime, address);
        runtime
            .dispatch_set(cell, "Value2", value, &[])
            .expect("seed value");
    }
    let number = OmValue::Number;
    let logical = OmValue::Bool;
    let text = |value: &str| OmValue::Text(value.to_string());
    let error = |value| OmValue::from(CellValue::Error(value));
    let cases = [
        // Comparisons return logicals, compare text without case, and order numbers before
        // text before logicals.
        ("=1=1", logical(true)),
        ("=P1>5", logical(true)),
        ("=1<>2", logical(true)),
        ("=\"a\"=\"A\"", logical(true)),
        ("=\"a\"<\"B\"", logical(true)),
        ("=P2=\"ABC\"", logical(true)),
        ("=\"x\">1", logical(true)),
        ("=TRUE>1", logical(true)),
        ("=TRUE>\"z\"", logical(true)),
        // A blank compares as the empty value of the other operand's type.
        ("=P9=0", logical(true)),
        ("=P9=\"\"", logical(true)),
        ("=P9=FALSE", logical(true)),
        ("=#N/A=1", error(CellError::NA)),
        // Arithmetic coerces logicals, blanks, and numeric text.
        ("=TRUE+1", number(2.0)),
        ("=\"3\"+1", number(4.0)),
        ("=P3*2", number(6.0)),
        ("=P9+1", number(1.0)),
        ("=\"x\"+1", error(CellError::Value)),
        ("=(P1>5)*2", number(2.0)),
        ("=1/0", error(CellError::Div0)),
        // `^` and `%`, with negation binding tighter than `^`.
        ("=2^3", number(8.0)),
        ("=2*3^2", number(18.0)),
        ("=2^3^2", number(64.0)),
        ("=-2^2", number(4.0)),
        ("=2^-1", number(0.5)),
        ("=0^0", error(CellError::Num)),
        ("=0^-1", error(CellError::Div0)),
        ("=(-8)^(1/3)", error(CellError::Num)),
        ("=50%", number(0.5)),
        ("=10%*P1", number(1.0)),
        // `&` joins text forms, binding looser than arithmetic.
        ("=1&2", text("12")),
        ("=(1&2)+1", number(13.0)),
        ("=P2&1", text("abc1")),
        ("=\"v\"&TRUE", text("vTRUE")),
        ("=P1+1&\"x\"", text("11x")),
        ("=P1&\"\"=\"10\"", logical(true)),
        // Logical functions return logicals; unary plus keeps text.
        ("=NOT(P1>5)", logical(false)),
        ("=AND(TRUE,1)", logical(true)),
        ("=ISNUMBER(P1)", logical(true)),
        ("=+P2", text("abc")),
        ("=-P1", number(-10.0)),
        ("=IF(P1>5,\"y\",\"n\")", text("y")),
    ];
    for (index, (formula, _)) in cases.iter().enumerate() {
        let cell = range(&mut runtime, &format!("R{}", index + 1));
        runtime
            .dispatch_set(cell, "Formula", OmValue::Text(formula.to_string()), &[])
            .unwrap_or_else(|error| panic!("set {formula}: {error:?}"));
    }
    runtime
        .calculate_workbook_with_report(workbook)
        .expect("calculate operators");
    let mut failures = Vec::new();
    for (index, (formula, expected)) in cases.into_iter().enumerate() {
        let cell = range(&mut runtime, &format!("R{}", index + 1));
        let actual = runtime.dispatch_get(cell, "Value2", &[]).expect("value");
        if actual != expected {
            failures.push(format!("{formula}: expected {expected:?}, got {actual:?}"));
        }
    }
    assert!(failures.is_empty(), "{failures:#?}");
}
