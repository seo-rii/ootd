use super::*;
use excel_model::{FormulaGroup, FormulaGroupKind};
use office_common::{OmArray, RangeRef};

const SHEET_PART: &str = "xl/worksheets/sheet1.xml";
const METADATA_PART: &str = "xl/metadata.xml";

fn part_text(saved: &[u8], part: &str) -> Option<String> {
    OpcPackage::from_bytes(saved)
        .expect("saved package")
        .part(part)
        .map(|part| String::from_utf8(part.bytes.clone()).expect("utf-8 part"))
}

fn with_sheet_rows(mut package: OpcPackage, rows: &str) -> Vec<u8> {
    let sheet = String::from_utf8(package.part(SHEET_PART).expect("sheet").bytes.clone())
        .expect("sheet utf-8")
        .replace("  </sheetData>", &format!("    {rows}\n  </sheetData>"));
    package
        .replace_part_bytes(SHEET_PART, sheet.into_bytes())
        .expect("replace sheet");
    package.to_bytes().expect("workbook bytes")
}

fn add_sequence(loaded: &mut crate::LoadedXlsxWorkbook, row: u32, col: u32) {
    loaded
        .state
        .set_range_dynamic_array_formulas(
            &RangeRef::single_cell(WorkbookId(0), SheetId(1), row, col),
            &OmArray::scalar(OmValue::Text("=SEQUENCE(3)".to_string())),
        )
        .expect("write dynamic array formula");
    let worksheet = loaded
        .state
        .worksheet_data_for_sheet_mut(SheetId(1))
        .expect("worksheet");
    worksheet.spill_ranges.insert(
        (row, col),
        Rect {
            row_first: row,
            row_last: row + 2,
            col_first: col,
            col_last: col,
        },
    );
}

#[test]
fn dynamic_array_save_creates_cell_metadata_so_the_formula_reopens_dynamic() {
    let bytes = synthetic_workbook_bytes();
    assert!(part_text(&bytes, METADATA_PART).is_none());
    let mut loaded = XlsxCodec
        .load(&bytes, CommonLoadOptions::default())
        .expect("load workbook without cell metadata");
    add_sequence(&mut loaded, 1, 13);

    let saved = XlsxCodec
        .save(&loaded, CommonSaveOptions::default())
        .expect("save new dynamic array");
    let metadata = part_text(&saved, METADATA_PART).expect("created cell metadata part");
    assert!(
        metadata.contains(r#"<metadataType name="XLDAPR""#),
        "{metadata}"
    );
    assert!(metadata.contains(r#"fDynamic="1""#), "{metadata}");
    let content_types = part_text(&saved, "[Content_Types].xml").expect("content types");
    assert!(
        content_types.contains(r#"<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>"#),
        "{content_types}"
    );
    let relationships = part_text(&saved, "xl/_rels/workbook.xml.rels").expect("workbook rels");
    assert!(
        relationships.contains(r#"Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml""#),
        "{relationships}"
    );
    let sheet = part_text(&saved, SHEET_PART).expect("sheet");
    assert!(
        sheet.contains(r#"<c r="M1" cm="1"><f t="array" ref="M1:M3">_xlfn.SEQUENCE(3)</f>"#),
        "{sheet}"
    );

    let reopened = XlsxCodec
        .load(&saved, CommonLoadOptions::default())
        .expect("reopen dynamic array");
    let worksheet = reopened
        .state
        .worksheet_data_for_sheet(SheetId(1))
        .expect("worksheet");
    assert!(worksheet.dynamic_array_formulas.contains(&(1, 13)));
    assert!(worksheet.formula_groups.is_empty());

    let saved_again = XlsxCodec
        .save(&reopened, CommonSaveOptions::default())
        .expect("no-op save of reopened workbook");
    assert_eq!(
        saved_again, saved,
        "a no-op save keeps the created metadata graph"
    );
}

#[test]
fn dynamic_array_save_reuses_existing_dynamic_cell_metadata() {
    let mut package = OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("package");
    add_dynamic_array_cell_metadata_part(&mut package);
    let bytes = with_sheet_rows(
        package,
        r#"<row r="10"><c r="J10" cm="1"><f t="array" ref="J10:J11">SEQUENCE(2)</f><v>1</v></c></row><row r="11"><c r="J11"><v>2</v></c></row>"#,
    );
    let mut loaded = XlsxCodec
        .load(&bytes, CommonLoadOptions::default())
        .expect("load workbook with dynamic-array metadata");
    assert_eq!(
        loaded.support_parts.cell_metadata.dynamic_array_indices,
        BTreeSet::from([1])
    );
    add_sequence(&mut loaded, 1, 13);

    let saved = XlsxCodec
        .save(&loaded, CommonSaveOptions::default())
        .expect("save second dynamic array");
    let sheet = part_text(&saved, SHEET_PART).expect("sheet");
    assert!(sheet.contains(r#"<c r="M1" cm="1">"#), "{sheet}");
    assert!(sheet.contains(r#"<c r="J10" cm="1">"#), "{sheet}");
    let package = OpcPackage::from_bytes(&saved).expect("saved package");
    assert!(package.part("xl/metadata1.xml").is_none());
    assert_eq!(
        part_text(&saved, METADATA_PART).expect("metadata"),
        DYNAMIC_ARRAY_CELL_METADATA_XML
    );
}

#[test]
fn array_formulas_without_dynamic_metadata_are_protected_legacy_arrays() {
    for (label, cell_metadata) in [("no cm", ""), ("unresolved cm", r#" cm="4""#)] {
        let bytes = with_sheet_rows(
            OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("package"),
            &format!(
                r#"<row r="10"><c r="J10"{cell_metadata}><f t="array" ref="J10:J11">A1:A2*10</f><v>10</v></c></row><row r="11"><c r="J11"><v>20</v></c></row>"#
            ),
        );
        let mut loaded = XlsxCodec
            .load(&bytes, CommonLoadOptions::default())
            .expect(label);
        let worksheet = loaded
            .state
            .worksheet_data_for_sheet(SheetId(1))
            .expect("worksheet");
        assert!(worksheet.dynamic_array_formulas.is_empty(), "{label}");
        assert!(worksheet.spill_ranges.is_empty(), "{label}");
        assert_eq!(
            worksheet.formula_groups.get(&(10, 10)),
            Some(&FormulaGroup {
                kind: FormulaGroupKind::LegacyArray,
                range: Rect {
                    row_first: 10,
                    row_last: 11,
                    col_first: 10,
                    col_last: 10,
                },
                shared_index: None,
                members: BTreeSet::new(),
            }),
            "{label}"
        );

        let error = loaded
            .state
            .set_range_values(
                &RangeRef::single_cell(WorkbookId(0), SheetId(1), 11, 10),
                &OmArray::scalar(OmValue::Number(1.0)),
            )
            .expect_err("legacy array member edit must fail closed");
        assert_eq!(error.code, OmErrorCode::InvalidState, "{label}");
        assert!(
            error
                .message
                .contains("legacy array cell R11C10; formula group anchor is R10C10"),
            "{label}: {}",
            error.message
        );

        let worksheet = loaded
            .state
            .worksheet_data_for_sheet_mut(SheetId(1))
            .expect("worksheet");
        worksheet.dirty = true;
        worksheet.dirty_cells.extend([(10, 10), (11, 10)]);
        let saved = XlsxCodec
            .save(&loaded, CommonSaveOptions::default())
            .expect("save touched legacy array");
        let sheet = part_text(&saved, SHEET_PART).expect("sheet");
        assert!(
            sheet.contains(&format!(
                r#"<c r="J10"{cell_metadata}><f t="array" ref="J10:J11">A1:A2*10</f><v>10</v></c>"#
            )),
            "{label}: {sheet}"
        );
        assert!(
            part_text(&saved, METADATA_PART).is_none(),
            "{label}: legacy arrays never create dynamic-array metadata"
        );
    }
}

#[test]
fn dynamic_array_save_refuses_to_rewrite_foreign_cell_metadata() {
    let mut package = OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("package");
    add_cell_metadata_part(
        &mut package,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="1"><bk/></futureMetadata><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>"#,
    );
    let mut loaded = XlsxCodec
        .load(
            &package.to_bytes().expect("bytes"),
            CommonLoadOptions::default(),
        )
        .expect("load workbook with foreign cell metadata");
    assert!(
        loaded
            .support_parts
            .cell_metadata
            .dynamic_array_indices
            .is_empty()
    );
    add_sequence(&mut loaded, 1, 13);

    let error = XlsxCodec
        .save(&loaded, CommonSaveOptions::default())
        .expect_err("appending to foreign cell metadata is not implemented");
    assert_eq!(error.code, OmErrorCode::Unsupported);
    assert!(
        error.message.contains("xl/metadata.xml"),
        "{}",
        error.message
    );
}

#[test]
fn malformed_cell_metadata_fails_closed_on_load() {
    for (label, cell_metadata, expected) in [
        (
            "unknown type",
            r#"<rc t="9" v="0"/>"#,
            "unknown metadata type 9",
        ),
        (
            "invalid type",
            r#"<rc t="x" v="0"/>"#,
            "invalid type index: x",
        ),
        (
            "missing value",
            r#"<rc t="1"/>"#,
            "missing its t or v index",
        ),
    ] {
        let mut package = OpcPackage::from_bytes(&synthetic_workbook_bytes()).expect("package");
        add_cell_metadata_part(
            &mut package,
            &DYNAMIC_ARRAY_CELL_METADATA_XML.replace(r#"<rc t="1" v="0"/>"#, cell_metadata),
        );
        let error = XlsxCodec
            .load(
                &package.to_bytes().expect("bytes"),
                CommonLoadOptions::default(),
            )
            .expect_err(label);
        assert_eq!(error.code, OmErrorCode::Parse, "{label}");
        assert!(
            error.message.contains("xl/metadata.xml") && error.message.contains(expected),
            "{label}: {}",
            error.message
        );
    }
}
