use super::*;

const SPREADSHEET_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const SHEET_PART: &str = "xl/worksheets/sheet1.xml";

fn saved_part_text(saved: &[u8], part_name: &str) -> String {
    let package = OpcPackage::from_bytes(saved).expect("saved package");
    String::from_utf8(
        package
            .part(part_name)
            .unwrap_or_else(|| panic!("saved part {part_name}"))
            .bytes
            .clone(),
    )
    .expect("saved part utf-8")
}

#[test]
fn entity_references_in_formula_and_value_text_load_exactly() {
    let cells = parse_worksheet_cells(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="str"><f>"a"&amp;"b"&lt;&gt;"c"</f><v>a&amp;b&lt;c&#33;&#x21;&quot;&apos;</v></c>
      <c r="B1"><v>&#49;.5</v></c>
      <c r="C1" t="str"><v><![CDATA[x&y]]>&amp;z</v></c>
    </row>
  </sheetData>
</worksheet>"#,
        &[],
        SPREADSHEET_NAMESPACE,
        "/xl/worksheets/sheet1.xml",
    )
    .expect("worksheet cells with entity references")
    .cells;

    assert_eq!(
        cells.get(&(1, 1)),
        Some(&CellData {
            value: CellValue::Text("a&b<c!!\"'".to_string()),
            formula: Some(FormulaSource {
                text: r#""a"&"b"<>"c""#.to_string(),
                is_r1c1: false,
            }),
            style_id: None,
        })
    );
    assert_eq!(
        cells.get(&(1, 2)).map(|cell| &cell.value),
        Some(&CellValue::Number(1.5))
    );
    assert_eq!(
        cells.get(&(1, 3)).map(|cell| &cell.value),
        Some(&CellValue::Text("x&y&z".to_string()))
    );
}

#[test]
fn unknown_entity_reference_in_cell_text_fails_closed() {
    let error = parse_worksheet_cells(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="3"><c r="B3" t="str"><v>a&nbsp;b</v></c></row></sheetData>
</worksheet>"#,
        &[],
        SPREADSHEET_NAMESPACE,
        "/xl/worksheets/sheet1.xml",
    )
    .expect_err("unknown entity reference must fail closed");

    assert_eq!(error.code, OmErrorCode::Parse);
    assert_eq!(
        error.message,
        "/xl/worksheets/sheet1.xml: unknown XML entity reference: &nbsp;"
    );
}

#[test]
fn entity_bearing_formula_and_value_text_round_trip_through_dirty_save() {
    let codec = XlsxCodec;
    let mut package =
        OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("synthetic package");
    package
        .replace_part_bytes(
            SHEET_PART,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="str"><f>"a"&amp;"b"&lt;&gt;"c"</f><v>a&amp;b&lt;&gt;c</v></c>
      <c r="B1"><f>A1&amp;"x"</f><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#
                .to_vec(),
        )
        .expect("replace worksheet part");
    let mut loaded = codec
        .load(
            &package.to_bytes().expect("fixture bytes"),
            CommonLoadOptions::default(),
        )
        .expect("load entity-bearing cells");
    let worksheet = loaded
        .state
        .worksheet_data_for_sheet_mut(SheetId(1))
        .expect("worksheet data");
    assert_eq!(
        worksheet
            .cells
            .get(&(1, 2))
            .and_then(|cell| cell.formula.as_ref())
            .map(|formula| formula.text.as_str()),
        Some(r#"A1&"x""#)
    );
    worksheet.dirty = true;
    worksheet.dirty_cells.extend([(1, 1), (1, 2)]);

    let saved = codec
        .save(&loaded, CommonSaveOptions::default())
        .expect("save dirty entity-bearing cells");
    let saved_sheet = saved_part_text(&saved, SHEET_PART);
    assert!(
        saved_sheet.contains(
            r#"<c r="A1" t="str"><f>"a"&amp;"b"&lt;&gt;"c"</f><v>a&amp;b&lt;&gt;c</v></c>"#
        ),
        "{saved_sheet}"
    );
    assert!(
        saved_sheet.contains(r#"<c r="B1"><f>A1&amp;"x"</f><v>1</v></c>"#),
        "{saved_sheet}"
    );

    let reopened = codec
        .load(&saved, CommonLoadOptions::default())
        .expect("reopen entity-bearing cells");
    let worksheet = reopened
        .state
        .worksheet_data_for_sheet(SheetId(1))
        .expect("reopened worksheet");
    assert_eq!(
        worksheet.cells.get(&(1, 1)),
        Some(&CellData {
            value: CellValue::Text("a&b<>c".to_string()),
            formula: Some(FormulaSource {
                text: r#""a"&"b"<>"c""#.to_string(),
                is_r1c1: false,
            }),
            style_id: None,
        })
    );
    assert_eq!(
        worksheet.cells.get(&(1, 2)),
        Some(&CellData {
            value: CellValue::Number(1.0),
            formula: Some(FormulaSource {
                text: r#"A1&"x""#.to_string(),
                is_r1c1: false,
            }),
            style_id: None,
        })
    );
}
