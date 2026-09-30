use super::*;
use office_common::{OmArray, RangeRef};

const SHEET_PART: &str = "xl/worksheets/sheet1.xml";
const FILE_CELLS: [&str; 4] = [
    r#"<c r="M20"><f>_xlfn.XLOOKUP(A1,B1:B2,C1:C2)</f><v>0</v></c>"#,
    r#"<c r="N20"><f>_xlfn.LET(_xlpm.x,2,_xlpm.x*3)</f><v>6</v></c>"#,
    r#"<c r="O20"><f>_xlfn.SINGLE(A1:A3)</f><v>0</v></c>"#,
    r#"<c r="P20"><f>_xlfn.VENDORFUNC(A1)</f><v>0</v></c>"#,
];

fn grammar_workbook_bytes() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("package");
    let sheet = String::from_utf8(package.part(SHEET_PART).expect("sheet").bytes.clone())
        .expect("sheet utf-8")
        .replace(
            "  </sheetData>",
            &format!(
                "    <row r=\"20\">{}</row>\n  </sheetData>",
                FILE_CELLS.concat()
            ),
        );
    package
        .replace_part_bytes(SHEET_PART, sheet.into_bytes())
        .expect("replace sheet");
    package.to_bytes().expect("grammar workbook bytes")
}

fn saved_sheet(saved: &[u8]) -> String {
    String::from_utf8(
        OpcPackage::from_bytes(saved)
            .expect("saved package")
            .part(SHEET_PART)
            .expect("sheet")
            .bytes
            .clone(),
    )
    .expect("sheet utf-8")
}

fn formula_text(loaded: &crate::LoadedXlsxWorkbook, col: u32) -> Option<String> {
    loaded
        .state
        .worksheet_data_for_sheet(SheetId(1))
        .expect("worksheet")
        .cells
        .get(&(20, col))
        .and_then(|cell| cell.formula.as_ref())
        .map(|formula| formula.text.clone())
}

#[test]
fn file_formula_grammar_loads_as_typed_grammar_and_saves_back_exactly() {
    let bytes = grammar_workbook_bytes();
    let mut loaded = XlsxCodec
        .load(&bytes, CommonLoadOptions::default())
        .expect("load file-grammar formulas");
    for (col, expected) in [
        (13, "XLOOKUP(A1,B1:B2,C1:C2)"),
        (14, "LET(x,2,x*3)"),
        (15, "@A1:A3"),
        (16, "_xlfn.VENDORFUNC(A1)"),
    ] {
        assert_eq!(
            formula_text(&loaded, col).as_deref(),
            Some(expected),
            "column {col}"
        );
    }

    let untouched = XlsxCodec
        .save(&loaded, CommonSaveOptions::default())
        .expect("no-op save");
    assert_eq!(saved_sheet(&untouched), saved_sheet(&bytes));

    let worksheet = loaded
        .state
        .worksheet_data_for_sheet_mut(SheetId(1))
        .expect("worksheet");
    worksheet.dirty = true;
    worksheet.dirty_cells.extend((13..=16).map(|col| (20, col)));
    let touched = XlsxCodec
        .save(&loaded, CommonSaveOptions::default())
        .expect("touched save");
    let sheet = saved_sheet(&touched);
    for cell in FILE_CELLS {
        assert!(
            sheet.contains(cell),
            "touched save must keep {cell} in:\n{sheet}"
        );
    }
}

#[test]
fn typed_formulas_save_with_file_grammar_prefixes() {
    let mut loaded = XlsxCodec
        .load(&synthetic_workbook_bytes(), CommonLoadOptions::default())
        .expect("load workbook");
    for (col, formula) in [
        (13, "=SORT(A1:A3)"),
        (14, "=MAP(A1:A3,LAMBDA(v,v*2))"),
        (15, "=SUM(@A1:A3,1)"),
        (16, "=IFERROR(XMATCH(1,A1:A3),0)"),
    ] {
        loaded
            .state
            .set_range_formulas(
                &RangeRef::single_cell(WorkbookId(0), SheetId(1), 20, col),
                &OmArray::scalar(OmValue::Text(formula.to_string())),
            )
            .unwrap_or_else(|error| panic!("{formula}: {error:?}"));
    }
    let saved = XlsxCodec
        .save(&loaded, CommonSaveOptions::default())
        .expect("save typed formulas");
    let sheet = saved_sheet(&saved);
    for expected in [
        "<f>_xlfn._xlws.SORT(A1:A3)</f>",
        "<f>_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2))</f>",
        "<f>SUM(_xlfn.SINGLE(A1:A3),1)</f>",
        "<f>IFERROR(_xlfn.XMATCH(1,A1:A3),0)</f>",
    ] {
        assert!(sheet.contains(expected), "missing {expected} in:\n{sheet}");
    }

    let reopened = XlsxCodec
        .load(&saved, CommonLoadOptions::default())
        .expect("reopen typed formulas");
    assert_eq!(
        formula_text(&reopened, 14).as_deref(),
        Some("MAP(A1:A3,LAMBDA(v,v*2))")
    );
    assert_eq!(
        formula_text(&reopened, 15).as_deref(),
        Some("SUM(@A1:A3,1)")
    );
}
