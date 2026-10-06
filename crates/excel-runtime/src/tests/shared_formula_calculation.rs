use super::*;

const SHARED_FORMULA_SHEET: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c><c r="C1"><f t="shared" ref="C1:D2" si="1">$A$1+A1</f><v>2</v></c><c r="D1"><f t="shared" si="1"/><v>3</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c><c r="C2"><f t="shared" si="1"/><v>3</v></c></row><row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row></sheetData></worksheet>"#;

fn open_shared_formula_workbook(bytes: Vec<u8>) -> (ExcelRuntime, WorkbookHandle, ObjectHandle) {
    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes,
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open shared formula workbook");
    let active_sheet = expect_object_handle(
        runtime
            .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
            .expect("ActiveSheet"),
    );
    (runtime, workbook, active_sheet)
}

fn shared_formula_workbook_bytes() -> Vec<u8> {
    let mut package =
        OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("base shared formula package");
    package
        .replace_part_bytes(
            "xl/worksheets/sheet1.xml",
            SHARED_FORMULA_SHEET.as_bytes().to_vec(),
        )
        .expect("replace shared formula worksheet");
    package.to_bytes().expect("shared formula package bytes")
}

fn range(runtime: &mut ExcelRuntime, sheet: ObjectHandle, address: &str) -> ObjectHandle {
    expect_object_handle(
        runtime
            .dispatch_invoke(sheet, "Range", &[OmValue::Text(address.to_string())])
            .unwrap_or_else(|error| panic!("Range({address}): {error:?}")),
    )
}

fn value2(runtime: &mut ExcelRuntime, sheet: ObjectHandle, address: &str) -> f64 {
    let cell = range(runtime, sheet, address);
    expect_number(
        runtime
            .dispatch_get(cell, "Value2", &[])
            .unwrap_or_else(|error| panic!("{address} Value2: {error:?}")),
    )
}

fn formula(runtime: &mut ExcelRuntime, sheet: ObjectHandle, address: &str) -> String {
    let cell = range(runtime, sheet, address);
    expect_text(
        runtime
            .dispatch_get(cell, "Formula", &[])
            .unwrap_or_else(|error| panic!("{address} Formula: {error:?}")),
    )
}

#[test]
fn shared_formula_children_expose_and_calculate_their_translated_formula() {
    let (mut runtime, workbook, sheet) =
        open_shared_formula_workbook(shared_formula_workbook_bytes());
    for (address, expected) in [
        ("B1", "=A1*2"),
        ("B2", "=A2*2"),
        ("B3", "=A3*2"),
        ("C1", "=$A$1+A1"),
        ("D1", "=$A$1+B1"),
        ("C2", "=$A$1+A2"),
    ] {
        assert_eq!(formula(&mut runtime, sheet, address), expected, "{address}");
    }
    assert!(
        expect_bool(
            runtime
                .dispatch_get(workbook.0, "Saved", &[])
                .expect("Workbook.Saved after open")
        ),
        "expanding shared children must not dirty a freshly opened workbook"
    );

    for (address, value) in [("A1", 10.0), ("A2", 20.0), ("A3", 30.0)] {
        let cell = range(&mut runtime, sheet, address);
        runtime
            .dispatch_set(cell, "Value2", OmValue::Number(value), &[])
            .unwrap_or_else(|error| panic!("set {address}: {error:?}"));
    }
    runtime
        .dispatch_invoke(runtime.root_application(), "Calculate", &[])
        .expect("calculate shared formulas");
    for (address, expected) in [
        ("B1", 20.0),
        ("B2", 40.0),
        ("B3", 60.0),
        ("C1", 20.0),
        ("C2", 30.0),
        ("D1", 30.0),
    ] {
        assert_eq!(value2(&mut runtime, sheet, address), expected, "{address}");
    }

    let saved = runtime
        .save_workbook(
            workbook,
            SaveWorkbookSpec {
                format: FileFormat::Xlsx,
                profile: ExcelProfile::Excel365,
                lossless: true,
            },
        )
        .expect("save calculated shared formulas");
    let saved_package = OpcPackage::from_bytes(&saved).expect("saved package");
    let sheet_xml = String::from_utf8(
        saved_package
            .part("xl/worksheets/sheet1.xml")
            .expect("saved worksheet")
            .bytes
            .clone(),
    )
    .expect("saved worksheet utf-8");
    for expected in [
        r#"<c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>20</v></c>"#,
        r#"<c r="B2"><f t="shared" si="0"/><v>40</v></c>"#,
        r#"<c r="B3"><f t="shared" si="0"/><v>60</v></c>"#,
        r#"<c r="C2"><f t="shared" si="1"/><v>30</v></c>"#,
        r#"<c r="D1"><f t="shared" si="1"/><v>30</v></c>"#,
    ] {
        assert!(
            sheet_xml.contains(expected),
            "missing {expected} in:\n{sheet_xml}"
        );
    }

    let (mut reopened, _, reopened_sheet) = open_shared_formula_workbook(saved);
    assert_eq!(formula(&mut reopened, reopened_sheet, "B3"), "=A3*2");
    assert_eq!(value2(&mut reopened, reopened_sheet, "B3"), 60.0);
}

#[test]
fn shared_formula_children_reject_runtime_formula_edits() {
    let (mut runtime, workbook, sheet) =
        open_shared_formula_workbook(shared_formula_workbook_bytes());
    let before = runtime_workbook_persistence_snapshot(&runtime, workbook);
    let child = range(&mut runtime, sheet, "B2");
    let error = runtime
        .dispatch_set(child, "Formula", OmValue::Text("=1".to_string()), &[])
        .expect_err("shared formula child edit must fail closed");
    assert_eq!(error.code, OmErrorCode::InvalidState);
    assert!(
        error.message.contains("shared formula child R2C2"),
        "{error:?}"
    );
    assert_eq!(
        runtime_workbook_persistence_snapshot(&runtime, workbook),
        before
    );
}

fn insert_rows(runtime: &mut ExcelRuntime, sheet: ObjectHandle, address: &str) {
    let target = range(runtime, sheet, address);
    runtime
        .dispatch_invoke(
            target,
            "Insert",
            &[OmValue::Number(f64::from(crate::XL_SHIFT_DOWN))],
        )
        .unwrap_or_else(|error| panic!("{address}.Insert: {error:?}"));
}

fn saved_sheet(runtime: &ExcelRuntime, workbook: WorkbookHandle) -> String {
    let saved = runtime
        .save_workbook(
            workbook,
            SaveWorkbookSpec {
                format: FileFormat::Xlsx,
                profile: ExcelProfile::Excel365,
                lossless: true,
            },
        )
        .expect("save shared formula workbook");
    String::from_utf8(
        OpcPackage::from_bytes(&saved)
            .expect("saved package")
            .part("xl/worksheets/sheet1.xml")
            .expect("sheet")
            .bytes
            .clone(),
    )
    .expect("sheet utf-8")
}

#[test]
fn whole_row_shifts_unshare_moved_shared_formulas() {
    let (mut runtime, workbook, sheet) =
        open_shared_formula_workbook(shared_formula_workbook_bytes());
    insert_rows(&mut runtime, sheet, "A2:XFD2");
    for (address, expected) in [
        ("B1", "=A1*2"),
        ("B3", "=A3*2"),
        ("B4", "=A4*2"),
        ("C1", "=$A$1+A1"),
        ("D1", "=$A$1+B1"),
        ("C3", "=$A$1+A3"),
    ] {
        assert_eq!(formula(&mut runtime, sheet, address), expected, "{address}");
    }
    runtime
        .calculate_workbook_with_report(workbook)
        .expect("calculate unshared formulas");
    assert_eq!(value2(&mut runtime, sheet, "B4"), 6.0);
    assert_eq!(value2(&mut runtime, sheet, "C3"), 3.0);
    let sheet_xml = saved_sheet(&runtime, workbook);
    assert!(!sheet_xml.contains(r#"t="shared""#), "{sheet_xml}");
    for expected in [
        r#"<c r="B3"><f>A3*2</f>"#,
        r#"<c r="B4"><f>A4*2</f>"#,
        r#"<c r="D1"><f>$A$1+B1</f>"#,
    ] {
        assert!(
            sheet_xml.contains(expected),
            "missing {expected} in:\n{sheet_xml}"
        );
    }
}

#[test]
fn whole_row_shifts_keep_untouched_shared_formulas_shared() {
    let (mut runtime, workbook, sheet) =
        open_shared_formula_workbook(shared_formula_workbook_bytes());
    insert_rows(&mut runtime, sheet, "A10:XFD10");
    let sheet_xml = saved_sheet(&runtime, workbook);
    assert!(
        sheet_xml.contains(r#"<f t="shared" ref="B1:B3" si="0">A1*2</f>"#),
        "{sheet_xml}"
    );
    assert!(
        sheet_xml.contains(r#"<f t="shared" si="0"/>"#),
        "{sheet_xml}"
    );
}
