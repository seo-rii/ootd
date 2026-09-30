use super::*;
use excel_model::{FormulaGroup, FormulaGroupKind};
use office_common::{OmArray, RangeRef};

const SPREADSHEET_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const SHEET_PART: &str = "xl/worksheets/sheet1.xml";

/// Every formula group shape the codec models, keyed by coordinate so each saved cell can be
/// compared with its exact source XML.
///
/// - `B1:B3` is a shared group: master `B1`, children `B2` and `B3`.
/// - `C1:C2` is a legacy (CSE) array without dynamic-array cell metadata.
/// - `D1:D2` is a dynamic array whose anchor carries `cm`.
/// - `E1` is an explicit `t="normal"` formula.
/// - `F2:F3` is a one-variable data table anchored at `F2`.
const GROUP_CELLS: [((u32, u32), &str); 16] = [
    ((1, 1), r#"<c r="A1"><v>1</v></c>"#),
    (
        (1, 2),
        r#"<c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c>"#,
    ),
    (
        (1, 3),
        r#"<c r="C1"><f t="array" ref="C1:C2">A1:A2*10</f><v>10</v></c>"#,
    ),
    (
        (1, 4),
        r#"<c r="D1" cm="1"><f t="array" ref="D1:D2">A1:A2</f><v>1</v></c>"#,
    ),
    ((1, 5), r#"<c r="E1"><f t="normal">A1+1</f><v>2</v></c>"#),
    ((2, 1), r#"<c r="A2"><v>2</v></c>"#),
    ((2, 2), r#"<c r="B2"><f t="shared" si="0"/><v>4</v></c>"#),
    ((2, 3), r#"<c r="C2"><v>20</v></c>"#),
    ((2, 4), r#"<c r="D2"><v>2</v></c>"#),
    ((2, 5), r#"<c r="E2"><v>5</v></c>"#),
    (
        (2, 6),
        r#"<c r="F2"><f t="dataTable" ref="F2:F3" dt2D="0" dtr="0" r1="E2"/><v>6</v></c>"#,
    ),
    ((3, 1), r#"<c r="A3"><v>3</v></c>"#),
    ((3, 2), r#"<c r="B3"><f t="shared" si="0"/><v>6</v></c>"#),
    ((3, 5), r#"<c r="E3"><v>7</v></c>"#),
    ((3, 6), r#"<c r="F3"><v>8</v></c>"#),
    ((4, 1), r#"<c r="A4"><v>0</v></c>"#),
];

fn group_sheet_xml(cells: &[((u32, u32), &str)]) -> String {
    let mut rows = BTreeMap::<u32, String>::new();
    for ((row, _), xml) in cells {
        rows.entry(*row).or_default().push_str(xml);
    }
    let rows = rows
        .into_iter()
        .map(|(row, cells)| format!("<row r=\"{row}\">{cells}</row>"))
        .collect::<String>();
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<worksheet xmlns=\"{SPREADSHEET_NAMESPACE}\"><sheetData>{rows}</sheetData></worksheet>"
    )
}

fn workbook_with_sheet(sheet_xml: String) -> Vec<u8> {
    let mut package =
        OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("synthetic package");
    package
        .replace_part_bytes(SHEET_PART, sheet_xml.into_bytes())
        .expect("replace worksheet part");
    package.to_bytes().expect("formula group workbook bytes")
}

fn group_workbook_bytes() -> Vec<u8> {
    workbook_with_sheet(group_sheet_xml(&GROUP_CELLS))
}

fn saved_sheet_text(saved: &[u8]) -> String {
    let package = OpcPackage::from_bytes(saved).expect("saved package");
    String::from_utf8(
        package
            .part(SHEET_PART)
            .expect("saved worksheet")
            .bytes
            .clone(),
    )
    .expect("saved worksheet utf-8")
}

fn expected_formula_groups() -> BTreeMap<(u32, u32), FormulaGroup> {
    BTreeMap::from([
        (
            (1, 2),
            FormulaGroup {
                kind: FormulaGroupKind::Shared,
                range: Rect {
                    row_first: 1,
                    row_last: 3,
                    col_first: 2,
                    col_last: 2,
                },
                shared_index: Some(0),
                members: BTreeSet::from([(2, 2), (3, 2)]),
            },
        ),
        (
            (2, 6),
            FormulaGroup {
                kind: FormulaGroupKind::DataTable,
                range: Rect {
                    row_first: 2,
                    row_last: 3,
                    col_first: 6,
                    col_last: 6,
                },
                shared_index: None,
                members: BTreeSet::new(),
            },
        ),
    ])
}

fn assert_group_model(worksheet: &WorksheetData, context: &str) {
    assert_eq!(
        worksheet.formula_groups,
        expected_formula_groups(),
        "{context}: formula groups"
    );
    assert_eq!(
        worksheet.dynamic_array_formulas,
        BTreeSet::from([(1, 3), (1, 4)]),
        "{context}: array anchors"
    );
    assert_eq!(
        worksheet.spill_owners,
        BTreeMap::from([((2, 3), (1, 3)), ((2, 4), (1, 4))]),
        "{context}: array members"
    );
    assert_eq!(
        worksheet
            .cells
            .get(&(1, 2))
            .and_then(|cell| cell.formula.as_ref()),
        Some(&FormulaSource {
            text: "A1*2".to_string(),
            is_r1c1: false,
        }),
        "{context}: shared master formula"
    );
    assert_eq!(
        worksheet
            .cells
            .get(&(1, 5))
            .and_then(|cell| cell.formula.as_ref()),
        Some(&FormulaSource {
            text: "A1+1".to_string(),
            is_r1c1: false,
        }),
        "{context}: normal formula"
    );
}

fn load_groups() -> crate::LoadedXlsxWorkbook {
    XlsxCodec
        .load(&group_workbook_bytes(), CommonLoadOptions::default())
        .expect("load formula group workbook")
}

#[test]
fn formula_group_matrix_models_groups_and_preserves_bytes_without_edits() {
    let loaded = load_groups();
    assert_group_model(
        loaded
            .state
            .worksheet_data_for_sheet(SheetId(1))
            .expect("loaded worksheet"),
        "load",
    );

    let saved = XlsxCodec
        .save(&loaded, CommonSaveOptions::default())
        .expect("no-op save");
    assert_eq!(saved_sheet_text(&saved), group_sheet_xml(&GROUP_CELLS));
}

#[test]
fn formula_group_matrix_keeps_group_xml_through_unrelated_edit() {
    let mut loaded = load_groups();
    loaded
        .state
        .set_range_values(
            &RangeRef::single_cell(WorkbookId(0), SheetId(1), 4, 1),
            &OmArray::scalar(OmValue::Number(9.0)),
        )
        .expect("edit unrelated cell");

    let saved = XlsxCodec
        .save(&loaded, CommonSaveOptions::default())
        .expect("save unrelated edit");
    let saved_sheet = saved_sheet_text(&saved);
    assert!(
        saved_sheet.contains(r#"<c r="A4"><v>9</v></c>"#),
        "{saved_sheet}"
    );
    for (_, cell_xml) in &GROUP_CELLS[..GROUP_CELLS.len() - 1] {
        assert!(
            saved_sheet.contains(cell_xml),
            "unrelated edit must keep {cell_xml} verbatim in:\n{saved_sheet}"
        );
    }

    let reopened = XlsxCodec
        .load(&saved, CommonLoadOptions::default())
        .expect("reopen unrelated edit");
    assert_group_model(
        reopened
            .state
            .worksheet_data_for_sheet(SheetId(1))
            .expect("reopened worksheet"),
        "unrelated edit reopen",
    );
}

#[test]
fn formula_group_matrix_keeps_group_definitions_when_members_are_rewritten() {
    let mut loaded = load_groups();
    let worksheet = loaded
        .state
        .worksheet_data_for_sheet_mut(SheetId(1))
        .expect("worksheet data");
    worksheet.dirty = true;
    worksheet
        .dirty_cells
        .extend(GROUP_CELLS.iter().map(|(coordinates, _)| *coordinates));

    let saved = XlsxCodec
        .save(&loaded, CommonSaveOptions::default())
        .expect("save rewritten group members");
    let saved_sheet = saved_sheet_text(&saved);
    for expected in [
        r#"<c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c>"#,
        r#"<c r="B2"><f t="shared" si="0"/><v>4</v></c>"#,
        r#"<c r="B3"><f t="shared" si="0"/><v>6</v></c>"#,
        r#"<c r="C1"><f t="array" ref="C1:C2">A1:A2*10</f><v>10</v></c>"#,
        r#"<c r="D1" cm="1"><f t="array" ref="D1:D2">A1:A2</f><v>1</v></c>"#,
        r#"<c r="F2"><f t="dataTable" ref="F2:F3" dt2D="0" dtr="0" r1="E2"/><v>6</v></c>"#,
        r#"<c r="F3"><v>8</v></c>"#,
    ] {
        assert!(
            saved_sheet.contains(expected),
            "rewritten group member must keep {expected} in:\n{saved_sheet}"
        );
    }

    let reopened = XlsxCodec
        .load(&saved, CommonLoadOptions::default())
        .expect("reopen rewritten group members");
    assert_group_model(
        reopened
            .state
            .worksheet_data_for_sheet(SheetId(1))
            .expect("reopened worksheet"),
        "rewrite reopen",
    );
}

#[test]
fn formula_group_matrix_rejects_targeted_member_mutations_without_changing_state() {
    let mut loaded = load_groups();
    let before = loaded.state.clone();
    let cell = |row, col| RangeRef::single_cell(WorkbookId(0), SheetId(1), row, col);
    let value = OmArray::scalar(OmValue::Number(1.0));
    let formula = OmArray::scalar(OmValue::Text("=1+1".to_string()));

    for (row, col, anchor, label) in [
        (1, 2, "R1C2", "shared formula master"),
        (2, 2, "R1C2", "shared formula child"),
        (3, 2, "R1C2", "shared formula child"),
        (2, 6, "R2C6", "data table cell"),
        (3, 6, "R2C6", "data table cell"),
    ] {
        for (operation, result) in [
            (
                "value",
                loaded.state.set_range_values(&cell(row, col), &value),
            ),
            (
                "formula",
                loaded.state.set_range_formulas(&cell(row, col), &formula),
            ),
            ("clear", loaded.state.clear_range_contents(&cell(row, col))),
        ] {
            let error = result.expect_err("formula group member edit must be rejected");
            assert_eq!(
                error.code,
                OmErrorCode::InvalidState,
                "{operation} R{row}C{col}"
            );
            assert!(
                error.message.contains(label) && error.message.contains(anchor),
                "{operation} R{row}C{col}: {}",
                error.message
            );
        }
    }

    let rows = RangeRef::single_rect(
        WorkbookId(0),
        SheetId(1),
        Rect {
            row_first: 1,
            row_last: 3,
            col_first: 2,
            col_last: 2,
        },
    );
    let error = loaded
        .state
        .clear_range_with_change(&rows)
        .expect_err("clearing a shared group must be rejected");
    assert_eq!(error.code, OmErrorCode::InvalidState);
    assert_eq!(loaded.state, before);
}

type ParseCase = (
    &'static str,
    &'static [((u32, u32), &'static str)],
    &'static str,
);

#[test]
fn formula_group_parse_fails_closed_on_incoherent_groups() {
    let cases: [ParseCase; 7] = [
        (
            "unknown shared index",
            &[((1, 1), r#"<c r="A1"><f t="shared" si="3"/><v>1</v></c>"#)],
            "shared formula index 3",
        ),
        (
            "duplicate shared master",
            &[
                (
                    (1, 1),
                    r#"<c r="A1"><f t="shared" ref="A1:A2" si="0">1</f></c>"#,
                ),
                (
                    (1, 2),
                    r#"<c r="B1"><f t="shared" ref="B1:B2" si="0">1</f></c>"#,
                ),
            ],
            "duplicate shared formula index 0",
        ),
        (
            "child outside master range",
            &[
                (
                    (1, 1),
                    r#"<c r="A1"><f t="shared" ref="A1:A2" si="0">1</f></c>"#,
                ),
                ((3, 1), r#"<c r="A3"><f t="shared" si="0"/></c>"#),
            ],
            "outside shared formula range",
        ),
        (
            "invalid shared index",
            &[(
                (1, 1),
                r#"<c r="A1"><f t="shared" ref="A1" si="x">1</f></c>"#,
            )],
            "invalid shared formula index",
        ),
        (
            "data table anchor is not top-left",
            &[(
                (2, 2),
                r#"<c r="B2"><f t="dataTable" ref="A1:B2" r1="C1"/></c>"#,
            )],
            "is not the top-left",
        ),
        (
            "unknown formula type",
            &[((1, 1), r#"<c r="A1"><f t="vendor">1</f></c>"#)],
            "unsupported formula type vendor",
        ),
        (
            "shared group overlaps array range",
            &[
                ((1, 1), r#"<c r="A1"><f t="array" ref="A1:A2">1</f></c>"#),
                (
                    (2, 1),
                    r#"<c r="A2"><f t="shared" ref="A2" si="0">1</f></c>"#,
                ),
            ],
            "overlapping formula group",
        ),
    ];
    for (label, cells, expected) in cases {
        let error = XlsxCodec
            .load(
                &workbook_with_sheet(group_sheet_xml(cells)),
                CommonLoadOptions::default(),
            )
            .expect_err(label);
        assert_eq!(error.code, OmErrorCode::Parse, "{label}: {}", error.message);
        assert!(
            error.message.contains(expected),
            "{label}: expected {expected:?} in {}",
            error.message
        );
    }
}
