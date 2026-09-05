use super::*;

const SPREADSHEET_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const SHEET_PART: &str = "xl/worksheets/sheet1.xml";
const SHARED_STRINGS_PART: &str = "xl/sharedStrings.xml";

const MATRIX_SHARED_STRINGS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>shared</t></si>
  <si><r><rPr><b/></rPr><t>rich</t></r></si>
</sst>"#;

/// One worksheet row that exercises every supported cell value channel. Each entry keeps its
/// exact source XML so no-op and unrelated-edit saves can be compared byte for byte.
const MATRIX_CELLS: [&str; 26] = [
    r#"<c r="A1"/>"#,
    r#"<c r="B1" s="1"/>"#,
    r#"<c r="C1" t="b"><v>1</v></c>"#,
    r#"<c r="D1" t="b"><v>false</v></c>"#,
    r#"<c r="E1"><v>1E+20</v></c>"#,
    r#"<c r="F1" t="n"><v>.5</v></c>"#,
    r#"<c r="G1" t="e"><v>#DIV/0!</v></c>"#,
    r#"<c r="H1" t="e"><v>#VENDOR!</v></c>"#,
    r#"<c r="I1" t="s"><v>0</v></c>"#,
    r#"<c r="J1" t="str"><v>plain</v></c>"#,
    r#"<c r="K1" t="str"><v></v></c>"#,
    r#"<c r="L1" t="inlineStr"><is><t>inline</t></is></c>"#,
    r#"<c r="M1" t="d"><v>2026-09-03</v></c>"#,
    r#"<c r="N1"><f>1+1</f><v>2</v></c>"#,
    r#"<c r="O1" t="str"><f>"a"&amp;"b"</f><v>ab</v></c>"#,
    r#"<c r="P1" t="str"><f>""</f><v></v></c>"#,
    r#"<c r="Q1" t="b"><f>1=1</f><v>1</v></c>"#,
    r#"<c r="R1" t="e"><f>1/0</f><v>#DIV/0!</v></c>"#,
    r#"<c r="S1"><f>SUM(A1:B1)</f></c>"#,
    r#"<c r="T1" t="str"><f>A1</f></c>"#,
    r#"<c r="U1" t="b"></c>"#,
    r#"<c r="V1" t="n" s="1"></c>"#,
    r#"<c r="W1" t="s"/>"#,
    r#"<c r="X1" t="s"><v>1</v></c>"#,
    r#"<c r="Y1" s="1"><f>1</f><v>1</v></c>"#,
    r#"<c r="Z1"><v>0.1</v></c>"#,
];

/// The canonical form every matrix cell takes when it is rewritten with an unchanged value.
/// `None` marks a coordinate that is not materialized and therefore disappears from the row.
const MATRIX_TOUCHED_CELLS: [Option<&str>; 26] = [
    None,
    Some(r#"<c r="B1" s="1"/>"#),
    Some(r#"<c r="C1" t="b"><v>1</v></c>"#),
    Some(r#"<c r="D1" t="b"><v>0</v></c>"#),
    Some(r#"<c r="E1"><v>100000000000000000000</v></c>"#),
    Some(r#"<c r="F1"><v>0.5</v></c>"#),
    Some(r#"<c r="G1" t="e"><v>#DIV/0!</v></c>"#),
    Some(r#"<c r="H1" t="e"><v>#VENDOR!</v></c>"#),
    Some(r#"<c r="I1" t="inlineStr"><is><t>shared</t></is></c>"#),
    Some(r#"<c r="J1" t="inlineStr"><is><t>plain</t></is></c>"#),
    Some(r#"<c r="K1" t="inlineStr"><is><t></t></is></c>"#),
    Some(r#"<c r="L1" t="inlineStr"><is><t>inline</t></is></c>"#),
    Some(r#"<c r="M1" t="d"><v>2026-09-03</v></c>"#),
    Some(r#"<c r="N1"><f>1+1</f><v>2</v></c>"#),
    Some(r#"<c r="O1" t="str"><f>"a"&amp;"b"</f><v>ab</v></c>"#),
    Some(r#"<c r="P1" t="str"><f>""</f><v></v></c>"#),
    Some(r#"<c r="Q1" t="b"><f>1=1</f><v>1</v></c>"#),
    Some(r#"<c r="R1" t="e"><f>1/0</f><v>#DIV/0!</v></c>"#),
    Some(r#"<c r="S1"><f>SUM(A1:B1)</f></c>"#),
    Some(r#"<c r="T1"><f>A1</f></c>"#),
    None,
    Some(r#"<c r="V1" s="1"/>"#),
    None,
    Some(r#"<c r="X1" t="inlineStr">"#),
    Some(r#"<c r="Y1" s="1"><f>1</f><v>1</v></c>"#),
    Some(r#"<c r="Z1"><v>0.1</v></c>"#),
];

const RICH_TEXT_COLUMN: u32 = 24;

fn matrix_sheet_xml() -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<worksheet xmlns=\"{SPREADSHEET_NAMESPACE}\">\n  <sheetData>\n    <row r=\"1\">{}</row>\n  </sheetData>\n</worksheet>",
        MATRIX_CELLS.concat()
    )
}

fn matrix_workbook_bytes() -> Vec<u8> {
    let mut package =
        OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("synthetic package");
    package
        .replace_part_bytes(
            SHARED_STRINGS_PART,
            MATRIX_SHARED_STRINGS.as_bytes().to_vec(),
        )
        .expect("replace shared strings part");
    package
        .replace_part_bytes(SHEET_PART, matrix_sheet_xml().into_bytes())
        .expect("replace worksheet part");
    package.to_bytes().expect("matrix workbook bytes")
}

fn constant(value: CellValue) -> Option<CellData> {
    Some(CellData {
        value,
        formula: None,
        style_id: None,
    })
}

fn formula(value: CellValue, text: &str) -> Option<CellData> {
    Some(CellData {
        value,
        formula: Some(FormulaSource {
            text: text.to_string(),
            is_r1c1: false,
        }),
        style_id: None,
    })
}

/// The typed model every matrix coordinate must produce on load and again after any save that
/// keeps its value. The rich-text column is checked separately because its preserved payload
/// legitimately changes representation when it is materialized inline.
fn expected_matrix_cells() -> Vec<(u32, Option<CellData>)> {
    vec![
        (1, None),
        (
            2,
            Some(CellData {
                value: CellValue::Blank,
                formula: None,
                style_id: Some(StyleId(1)),
            }),
        ),
        (3, constant(CellValue::Bool(true))),
        (4, constant(CellValue::Bool(false))),
        (5, constant(CellValue::Number(1e20))),
        (6, constant(CellValue::Number(0.5))),
        (7, constant(CellValue::Error(CellError::Div0))),
        (
            8,
            constant(CellValue::Error(CellError::UnknownLexical(
                "#VENDOR!".to_string(),
            ))),
        ),
        (9, constant(CellValue::Text("shared".to_string()))),
        (10, constant(CellValue::Text("plain".to_string()))),
        (11, constant(CellValue::Text(String::new()))),
        (12, constant(CellValue::Text("inline".to_string()))),
        (
            13,
            constant(CellValue::IsoDateTime(
                IsoDateTime::parse("2026-09-03").expect("valid date"),
            )),
        ),
        (14, formula(CellValue::Number(2.0), "1+1")),
        (15, formula(CellValue::Text("ab".to_string()), r#""a"&"b""#)),
        (16, formula(CellValue::Text(String::new()), r#""""#)),
        (17, formula(CellValue::Bool(true), "1=1")),
        (18, formula(CellValue::Error(CellError::Div0), "1/0")),
        (19, formula(CellValue::Blank, "SUM(A1:B1)")),
        (20, formula(CellValue::Blank, "A1")),
        (21, None),
        (
            22,
            Some(CellData {
                value: CellValue::Blank,
                formula: None,
                style_id: Some(StyleId(1)),
            }),
        ),
        (23, None),
        (
            25,
            Some(CellData {
                value: CellValue::Number(1.0),
                formula: Some(FormulaSource {
                    text: "1".to_string(),
                    is_r1c1: false,
                }),
                style_id: Some(StyleId(1)),
            }),
        ),
        (26, constant(CellValue::Number(0.1))),
    ]
}

fn assert_matrix_model(worksheet: &WorksheetData, label: &str) {
    for (col, expected) in expected_matrix_cells() {
        assert_eq!(
            worksheet.cells.get(&(1, col)),
            expected.as_ref(),
            "{label}: column {col}"
        );
    }
    let rich_cell = worksheet
        .cells
        .get(&(1, RICH_TEXT_COLUMN))
        .unwrap_or_else(|| panic!("{label}: rich text cell must be materialized"));
    match &rich_cell.value {
        CellValue::RichText(value) => {
            assert_eq!(value.as_str(), "rich", "{label}: rich text display");
            assert!(
                std::str::from_utf8(value.raw_inner_xml())
                    .expect("rich text inner xml utf-8")
                    .contains("<b/>"),
                "{label}: rich text formatting run"
            );
        }
        other => panic!("{label}: expected rich text, got {other:?}"),
    }
    assert!(
        rich_cell.formula.is_none() && rich_cell.style_id.is_none(),
        "{label}: rich text cell channels"
    );
}

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
fn cell_value_matrix_loads_typed_channels_and_preserves_bytes_without_edits() {
    let codec = XlsxCodec;
    let fixture = matrix_workbook_bytes();
    let loaded = codec
        .load(&fixture, CommonLoadOptions::default())
        .expect("load cell value matrix");
    assert_matrix_model(
        loaded
            .state
            .worksheet_data_for_sheet(SheetId(1))
            .expect("loaded worksheet"),
        "load",
    );

    let saved = codec
        .save(&loaded, CommonSaveOptions::default())
        .expect("no-op save");
    assert_eq!(saved_part_text(&saved, SHEET_PART), matrix_sheet_xml());
    assert_eq!(
        saved_part_text(&saved, SHARED_STRINGS_PART),
        MATRIX_SHARED_STRINGS
    );
}

#[test]
fn cell_value_matrix_keeps_untouched_cell_bytes_through_unrelated_edit() {
    let codec = XlsxCodec;
    let mut loaded = codec
        .load(&matrix_workbook_bytes(), CommonLoadOptions::default())
        .expect("load cell value matrix");
    let worksheet = loaded
        .state
        .worksheet_data_for_sheet_mut(SheetId(1))
        .expect("worksheet data");
    worksheet.cells.insert(
        (1, 1),
        CellData {
            value: CellValue::Number(7.0),
            formula: None,
            style_id: None,
        },
    );
    worksheet.dirty = true;
    worksheet.dirty_cells.insert((1, 1));

    let saved = codec
        .save(&loaded, CommonSaveOptions::default())
        .expect("save unrelated edit");
    let saved_sheet = saved_part_text(&saved, SHEET_PART);
    assert!(saved_sheet.contains(r#"<c r="A1"><v>7</v></c>"#));
    for cell_xml in &MATRIX_CELLS[1..] {
        assert!(
            saved_sheet.contains(cell_xml),
            "unrelated edit must keep {cell_xml} verbatim in:\n{saved_sheet}"
        );
    }
    assert_eq!(
        saved_part_text(&saved, SHARED_STRINGS_PART),
        MATRIX_SHARED_STRINGS
    );

    let reopened = codec
        .load(&saved, CommonLoadOptions::default())
        .expect("reopen unrelated edit");
    let worksheet = reopened
        .state
        .worksheet_data_for_sheet(SheetId(1))
        .expect("reopened worksheet");
    assert_eq!(
        worksheet.cells.get(&(1, 1)),
        Some(&CellData {
            value: CellValue::Number(7.0),
            formula: None,
            style_id: None,
        })
    );
    for (col, expected) in expected_matrix_cells().into_iter().skip(1) {
        assert_eq!(
            worksheet.cells.get(&(1, col)),
            expected.as_ref(),
            "unrelated edit reopen: column {col}"
        );
    }
}

#[test]
fn cell_value_matrix_rewrites_touched_cells_canonically_and_reopens_identically() {
    let codec = XlsxCodec;
    let mut loaded = codec
        .load(&matrix_workbook_bytes(), CommonLoadOptions::default())
        .expect("load cell value matrix");
    let worksheet = loaded
        .state
        .worksheet_data_for_sheet_mut(SheetId(1))
        .expect("worksheet data");
    worksheet.dirty = true;
    worksheet
        .dirty_cells
        .extend((1..=MATRIX_CELLS.len() as u32).map(|col| (1, col)));

    let saved = codec
        .save(&loaded, CommonSaveOptions::default())
        .expect("save touched matrix");
    let saved_sheet = saved_part_text(&saved, SHEET_PART);
    for (source, touched) in MATRIX_CELLS.iter().zip(MATRIX_TOUCHED_CELLS) {
        let reference = source
            .split("r=\"")
            .nth(1)
            .and_then(|rest| rest.split('"').next())
            .expect("matrix cell reference");
        match touched {
            Some(expected) => assert!(
                saved_sheet.contains(expected),
                "touched rewrite of {source} must produce {expected} in:\n{saved_sheet}"
            ),
            None => assert!(
                !saved_sheet.contains(&format!("r=\"{reference}\"")),
                "unmaterialized {source} must disappear from:\n{saved_sheet}"
            ),
        }
    }
    assert!(saved_sheet.contains("<t>rich</t>"));
    assert!(saved_sheet.contains("<b/>"));

    let reopened = codec
        .load(&saved, CommonLoadOptions::default())
        .expect("reopen touched matrix");
    assert_matrix_model(
        reopened
            .state
            .worksheet_data_for_sheet(SheetId(1))
            .expect("reopened worksheet"),
        "touched reopen",
    );
}

#[test]
fn typed_cells_without_a_value_element_are_blank_in_both_element_forms() {
    let cells = parse_worksheet_cells(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="b"></c>
      <c r="B1" t="e"/>
      <c r="C1" t="str"></c>
      <c r="D1" t="s"/>
      <c r="E1" t="d"></c>
      <c r="F1" t="inlineStr"></c>
      <c r="G1" t="n"></c>
      <c r="H1" t="b" s="2"></c>
      <c r="I1" t="str"><f>A1</f></c>
      <c r="J1" t="d" s="2"/>
    </row>
  </sheetData>
</worksheet>"#,
        &[],
        SPREADSHEET_NAMESPACE,
        "/xl/worksheets/sheet1.xml",
    )
    .expect("typed cells without values")
    .cells;

    for col in 1..=7 {
        assert!(
            !cells.contains_key(&(1, col)),
            "unstyled typed cell without a value must not be materialized: column {col}"
        );
    }
    assert_eq!(
        cells.get(&(1, 8)),
        Some(&CellData {
            value: CellValue::Blank,
            formula: None,
            style_id: Some(StyleId(2)),
        })
    );
    assert_eq!(
        cells.get(&(1, 9)),
        Some(&CellData {
            value: CellValue::Blank,
            formula: Some(FormulaSource {
                text: "A1".to_string(),
                is_r1c1: false,
            }),
            style_id: None,
        })
    );
    assert_eq!(
        cells.get(&(1, 10)),
        Some(&CellData {
            value: CellValue::Blank,
            formula: None,
            style_id: Some(StyleId(2)),
        })
    );
}

#[test]
fn invalid_typed_cell_value_lexicals_fail_closed_with_part_and_cell_context() {
    let cases = [
        (
            r#"<c r="B2" t="b"><v>2</v></c>"#,
            "cell B2 has invalid boolean value lexical: 2",
        ),
        (
            r#"<c r="B2" t="b"><v>TRUE</v></c>"#,
            "cell B2 has invalid boolean value lexical: TRUE",
        ),
        (
            r#"<c r="B2" t="b"><v></v></c>"#,
            "cell B2 has an empty value lexical for cell type b",
        ),
        (
            r#"<c r="B2" t="e"><v></v></c>"#,
            "cell B2 has an empty value lexical for cell type e",
        ),
        (
            r#"<c r="B2"><v></v></c>"#,
            "cell B2 has an empty value lexical for cell type n",
        ),
        (
            r#"<c r="B2" t="n"><v/></c>"#,
            "cell B2 has an empty value lexical for cell type n",
        ),
        (
            r#"<c r="B2" t="s"><v></v></c>"#,
            "cell B2 has an empty value lexical for cell type s",
        ),
        (
            r#"<c r="B2" t="d"><v></v></c>"#,
            "cell B2 has an empty value lexical for cell type d",
        ),
        (
            r#"<c r="B2"><v>1,5</v></c>"#,
            "cell B2 invalid numeric value lexical: 1,5",
        ),
        (
            r#"<c r="B2" t="n"><v> 1</v></c>"#,
            "cell B2 invalid numeric value lexical:  1",
        ),
        (
            r#"<c r="B2"><v>0x10</v></c>"#,
            "cell B2 invalid numeric value lexical: 0x10",
        ),
        (
            r#"<c r="B2"><v>Infinity</v></c>"#,
            "cell B2 numeric value must be finite",
        ),
        (
            r#"<c r="B2" t="s"><v>abc</v></c>"#,
            "cell B2 has invalid shared string index: abc",
        ),
        (
            r#"<c r="B2" t="s"><v>-1</v></c>"#,
            "cell B2 has invalid shared string index: -1",
        ),
        (
            r#"<c r="B2" t="s"><v>1</v></c>"#,
            "cell B2 shared string index out of range: 1",
        ),
        (
            r#"<c r="B2" t="x"><v>1</v></c>"#,
            "cell B2 has unknown cell type: x",
        ),
        (r#"<c r="B2" t="x"/>"#, "cell B2 has unknown cell type: x"),
        (
            r#"<c r="B2" t="bool"></c>"#,
            "cell B2 has unknown cell type: bool",
        ),
        (
            r#"<c r="B2"><v>1</v><v>2</v></c>"#,
            "cell B2 declares more than one value element",
        ),
        (
            r#"<c r="B2"><v>1</v><v/></c>"#,
            "cell B2 declares more than one value element",
        ),
        (
            r#"<c r="B2" t="inlineStr"><v>1</v></c>"#,
            "cell B2 is an inline string cell but declares a value element",
        ),
    ];

    for (cell_xml, expected_message) in cases {
        let error = parse_worksheet_cells(
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="{SPREADSHEET_NAMESPACE}"><sheetData><row r="2">{cell_xml}</row></sheetData></worksheet>"#
            )
            .as_bytes(),
            &[CellValue::Text("only".to_string())],
            SPREADSHEET_NAMESPACE,
            "/xl/worksheets/sheet7.xml",
        )
        .err()
        .unwrap_or_else(|| panic!("{cell_xml} must fail closed"));

        assert_eq!(error.code, OmErrorCode::Parse, "{cell_xml}");
        assert_eq!(
            error.message,
            format!("/xl/worksheets/sheet7.xml: {expected_message}"),
            "{cell_xml}"
        );
    }
}

#[test]
fn value_elements_inside_inline_strings_do_not_reach_the_value_channel() {
    let cells = parse_worksheet_cells(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>keep</t><v>ignored</v></is></c>
    </row>
  </sheetData>
</worksheet>"#,
        &[],
        SPREADSHEET_NAMESPACE,
        "/xl/worksheets/sheet1.xml",
    )
    .expect("inline string with nested foreign value element")
    .cells;

    let value = &cells.get(&(1, 1)).expect("A1").value;
    assert_eq!(value.as_text(), Some("keep"));
}

#[test]
fn save_rejects_empty_unknown_error_lexical_before_serialization() {
    let codec = XlsxCodec;
    let mut loaded = codec
        .load(&synthetic_workbook_bytes(), CommonLoadOptions::default())
        .expect("load synthetic workbook");
    let worksheet = loaded
        .state
        .worksheet_data_for_sheet_mut(SheetId(1))
        .expect("worksheet data");
    worksheet.cells.insert(
        (1, 1),
        CellData {
            value: CellValue::Error(CellError::UnknownLexical(String::new())),
            formula: None,
            style_id: None,
        },
    );
    worksheet.dirty = true;
    worksheet.dirty_cells.insert((1, 1));

    let error = codec
        .save(&loaded, CommonSaveOptions::default())
        .expect_err("empty error lexical must not be serialized");
    assert_eq!(error.code, OmErrorCode::InvalidState);
    assert!(
        error.message.contains("error lexical must not be empty"),
        "{}",
        error.message
    );
    assert!(error.message.contains("R1C1") || error.message.contains("A1"));

    let error = rewrite_worksheet_xml(
        loaded
            .state
            .worksheet_data_for_sheet(SheetId(1))
            .expect("worksheet data"),
        None,
        SPREADSHEET_NAMESPACE,
    )
    .expect_err("worksheet rewrite must refuse the empty error lexical");
    assert_eq!(error.code, OmErrorCode::InvalidState);
    assert_eq!(
        error.message,
        "worksheet cell A1 error lexical must not be empty"
    );
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
