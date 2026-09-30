use super::*;

const LEGACY_ARRAY_SHEET: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>3</v></c><c r="J1"><f t="array" ref="J1:J2">SORT(A1:A2,1,-1)</f><v>2</v></c><c r="K1"><f t="array" ref="K1:L2">SUM(A1:A2)</f><v>3</v></c><c r="L1"><v>3</v></c><c r="M1"><f t="array" ref="M1:M3">SORT(A1:A2)</f><v>1</v></c><c r="N1"><f t="array" ref="N1:O2">TRANSPOSE(A1:A2)</f><v>1</v></c><c r="O1"><v>2</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="J2"><v>1</v></c><c r="K2"><v>3</v></c><c r="L2"><v>3</v></c><c r="M2"><v>2</v></c><c r="N2"><v>1</v></c><c r="O2"><v>2</v></c></row><row r="3"><c r="M3"><v>0</v></c></row></sheetData></worksheet>"#;

fn open_legacy_array_workbook() -> (ExcelRuntime, WorkbookHandle, ObjectHandle) {
    let mut package =
        OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("base legacy array package");
    package
        .replace_part_bytes(
            "xl/worksheets/sheet1.xml",
            LEGACY_ARRAY_SHEET.as_bytes().to_vec(),
        )
        .expect("replace legacy array worksheet");
    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: package.to_bytes().expect("legacy array package bytes"),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open legacy array workbook");
    let sheet = expect_object_handle(
        runtime
            .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
            .expect("ActiveSheet"),
    );
    (runtime, workbook, sheet)
}

fn cell_value(runtime: &mut ExcelRuntime, sheet: ObjectHandle, address: &str) -> OmValue {
    let cell = expect_object_handle(
        runtime
            .dispatch_invoke(sheet, "Range", &[OmValue::Text(address.to_string())])
            .unwrap_or_else(|error| panic!("Range({address}): {error:?}")),
    );
    runtime
        .dispatch_get(cell, "Value2", &[])
        .unwrap_or_else(|error| panic!("{address} Value2: {error:?}"))
}

#[test]
fn legacy_array_formulas_recalculate_into_their_fixed_range() {
    let (mut runtime, workbook, sheet) = open_legacy_array_workbook();
    let a2 = expect_object_handle(
        runtime
            .dispatch_invoke(sheet, "Range", &[OmValue::Text("A2".to_string())])
            .expect("Range(A2)"),
    );
    runtime
        .dispatch_set(a2, "Value2", OmValue::Number(5.0), &[])
        .expect("edit legacy array precedent");
    runtime
        .calculate_workbook_with_report(workbook)
        .expect("calculate legacy arrays");

    let number = OmValue::Number;
    let not_available = OmValue::from(CellValue::Error(CellError::NA));
    for (address, expected) in [
        ("J1", number(5.0)),
        ("J2", number(1.0)),
        ("K1", number(6.0)),
        ("L1", number(6.0)),
        ("K2", number(6.0)),
        ("L2", number(6.0)),
        ("M1", number(1.0)),
        ("M2", number(5.0)),
        ("M3", not_available.clone()),
        ("N1", number(1.0)),
        ("O1", number(5.0)),
        ("N2", number(1.0)),
        ("O2", number(5.0)),
    ] {
        assert_eq!(
            cell_value(&mut runtime, sheet, address),
            expected,
            "{address}"
        );
    }

    let worksheet = runtime
        .runtime_workbook(workbook)
        .expect("runtime workbook")
        .loaded
        .state
        .worksheet_data_for_sheet(runtime.worksheets(workbook).expect("worksheets")[0].id)
        .expect("worksheet");
    assert!(
        worksheet.spill_ranges.is_empty(),
        "legacy arrays never spill"
    );
    assert_eq!(worksheet.formula_groups.len(), 4);
}
