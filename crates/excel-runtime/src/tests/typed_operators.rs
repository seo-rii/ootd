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
        // Function arguments use the same operators and coercions.
        ("=SUM(2^3,\"3\")", number(11.0)),
        ("=ABS(-2^2)", number(4.0)),
        ("=IF(\"a\"=\"A\",\"y\",\"n\")", text("y")),
        ("=ROUND(10%*P1+0.25,1)", number(1.3)),
        ("=LEN(1=1)", number(4.0)),
        ("=UPPER(\"a\"&1=1)", text("FALSE")),
        ("=CONCAT(\"x\",2^2)", text("x4")),
        ("=IFERROR(1/0,\"z\")", text("z")),
        ("=ISERROR(P2+1)", logical(true)),
        // Errors are values: they stop at the operator or function that catches them.
        ("=ISERROR(\"abc\")", logical(false)),
        ("=ISERROR(1/0+5)", logical(true)),
        ("=IFERROR(SUM(1/0,2),0)", number(0.0)),
        ("=IFERROR(P2*2,\"bad\")", text("bad")),
        ("=IFNA(#N/A,1)", number(1.0)),
        ("=#N/A&\"x\"", error(CellError::NA)),
        ("=1/0>#N/A", error(CellError::Div0)),
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

/// Arrays from ranges, constants, and array functions, with operators applied element by
/// element. P1:P3 holds 1, 2, 3 and Q1:Q3 holds 10, 20, 30.
#[test]
fn operators_apply_element_by_element_to_arrays() {
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
        ("P1", 1.0),
        ("P2", 2.0),
        ("P3", 3.0),
        ("Q1", 10.0),
        ("Q2", 20.0),
        ("Q3", 30.0),
    ] {
        let cell = range(&mut runtime, address);
        runtime
            .dispatch_set(cell, "Value2", OmValue::Number(value), &[])
            .expect("seed value");
    }
    let number = OmValue::Number;
    let logical = OmValue::Bool;
    let text = |value: &str| OmValue::Text(value.to_string());
    let cases = [
        ("=SUM((P1:P3>1)*1)", number(2.0)),
        ("=SUMPRODUCT((P1:P3>1)*Q1:Q3)", number(50.0)),
        ("=SUMPRODUCT(--(P1:P3>1))", number(2.0)),
        ("=SUM(--(P1:P3>1))", number(2.0)),
        ("=SUM(P1:P3*Q1:Q3)", number(140.0)),
        ("=SUM(P1:P3*2)", number(12.0)),
        ("=SUM({1,2,3}*2)", number(12.0)),
        ("=MAX(P1:P3-Q1:Q3)", number(-9.0)),
        ("=AVERAGE(P1:P3^2)", number(14.0 / 3.0)),
        ("=SUM(IF(P1:P3>1,Q1:Q3,0))", number(50.0)),
        ("=COUNT(IF(P1:P3>1,1,\"\"))", number(2.0)),
        ("=AND(P1:P3>0)", logical(true)),
        ("=OR(P1:P3>2)", logical(true)),
        ("=CONCAT(P1:P3&\"x\")", text("1x2x3x")),
        ("=TEXTJOIN(\",\",TRUE,P1:P3*2)", text("2,4,6")),
        ("=SUM(SEQUENCE(3)*P1:P3)", number(14.0)),
        // A single row or column stretches across the other operand; a larger array's extra
        // positions are #N/A.
        ("=SUM({1,2;3,4}*{10,100})", number(640.0)),
        ("=SUM(P1:P3*{1,2})", number(18.0)),
        (
            "=SUM(P1:P3*{1;2})",
            OmValue::from(CellValue::Error(CellError::NA)),
        ),
        // Where one value is wanted, a computed array shows its top-left element.
        ("=SUM(P1:P3)*{2,3}", number(12.0)),
    ];
    for (index, (formula, _)) in cases.iter().enumerate() {
        let cell = range(&mut runtime, &format!("R{}", index + 1));
        runtime
            .dispatch_set(cell, "Formula", OmValue::Text(formula.to_string()), &[])
            .unwrap_or_else(|error| panic!("set {formula}: {error:?}"));
    }
    // A dynamic-array formula spills the array an expression computes.
    let spill = range(&mut runtime, "T1");
    runtime
        .dispatch_set(
            spill,
            "Formula2",
            OmValue::Text("=P1:P3*2".to_string()),
            &[],
        )
        .expect("set spilling formula");
    runtime
        .calculate_workbook_with_report(workbook)
        .expect("calculate arrays");
    let mut failures = Vec::new();
    for (index, (formula, expected)) in cases.into_iter().enumerate() {
        let cell = range(&mut runtime, &format!("R{}", index + 1));
        let actual = runtime.dispatch_get(cell, "Value2", &[]).expect("value");
        if actual != expected {
            failures.push(format!("{formula}: expected {expected:?}, got {actual:?}"));
        }
    }
    for (address, expected) in [("T1", 2.0), ("T2", 4.0), ("T3", 6.0)] {
        let cell = range(&mut runtime, address);
        let actual = runtime.dispatch_get(cell, "Value2", &[]).expect("value");
        if actual != OmValue::Number(expected) {
            failures.push(format!(
                "spill {address}: expected {expected}, got {actual:?}"
            ));
        }
    }
    assert!(failures.is_empty(), "{failures:#?}");
}
