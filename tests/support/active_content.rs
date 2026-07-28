#[allow(dead_code)]
const WORKBOOK_XLSX_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
#[allow(dead_code)]
const WORKBOOK_XLSM_CONTENT_TYPE: &str = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";

pub const ACTIVE_PARTS: [(&str, &str); 9] = [
    ("xl/vbaProject.bin", "application/vnd.ms-office.vbaProject"),
    (
        "xl/vbaProjectSignatureAgile.bin",
        "application/vnd.ms-office.vbaProjectSignatureAgile",
    ),
    ("xl/vbaData.xml", "application/vnd.ms-office.vbaData+xml"),
    (
        "xl/macrosheets/sheet1.xml",
        "application/vnd.ms-excel.macrosheet+xml",
    ),
    (
        "xl/dialogsheets/sheet1.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
    ),
    (
        "xl/activeX/activeX1.xml",
        "application/vnd.ms-office.activeX+xml",
    ),
    (
        "xl/ctrlProps/ctrlProp1.xml",
        "application/vnd.ms-excel.controlproperties+xml",
    ),
    (
        "xl/embeddings/oleObject1.bin",
        "application/vnd.openxmlformats-officedocument.oleObject",
    ),
    ("customUI/customUI.xml", "application/xml"),
];

pub fn package_with_active_content(workbook_bytes: &[u8]) -> Vec<u8> {
    let mut package =
        office_opc::OpcPackage::from_bytes(workbook_bytes).expect("parse workbook fixture");

    let mut content_types = std::str::from_utf8(
        &package
            .part("[Content_Types].xml")
            .expect("content types")
            .bytes,
    )
    .expect("content types utf8")
    .to_string();
    let overrides = ACTIVE_PARTS
        .iter()
        .map(|(part_name, content_type)| {
            format!(r#"<Override PartName="/{part_name}" ContentType="{content_type}"/>"#)
        })
        .collect::<String>();
    content_types = content_types.replace("</Types>", &format!("{overrides}</Types>"));
    package
        .replace_part_bytes("[Content_Types].xml", content_types.into_bytes())
        .expect("replace content types");

    let package_relationships = std::str::from_utf8(
        &package
            .part("_rels/.rels")
            .expect("package relationships")
            .bytes,
    )
    .expect("package relationships utf8")
    .replace(
        "</Relationships>",
        r#"<Relationship Id="rIdCustomUi" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="/customUI/customUI.xml"/></Relationships>"#,
    );
    package
        .replace_part_bytes("_rels/.rels", package_relationships.into_bytes())
        .expect("replace package relationships");

    let workbook_relationships = std::str::from_utf8(
        &package
            .part("xl/_rels/workbook.xml.rels")
            .expect("workbook relationships")
            .bytes,
    )
    .expect("workbook relationships utf8")
    .replace(
        "</Relationships>",
        r#"<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
<Relationship Id="rIdVbaSignature" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignatureAgile" Target="vbaProjectSignatureAgile.bin"/>
<Relationship Id="rIdVbaData" Type="http://schemas.microsoft.com/office/2006/relationships/vbaData" Target="vbaData.xml"/>
<Relationship Id="rIdMacroSheet" Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="macrosheets/sheet1.xml"/>
<Relationship Id="rIdDialogSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet" Target="dialogsheets/sheet1.xml"/>
<Relationship Id="rIdActiveX" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" Target="activeX/activeX1.xml"/>
<Relationship Id="rIdOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>
</Relationships>"#,
    );
    package
        .replace_part_bytes(
            "xl/_rels/workbook.xml.rels",
            workbook_relationships.into_bytes(),
        )
        .expect("replace workbook relationships");

    for (part_name, content_type) in ACTIVE_PARTS {
        let bytes = if content_type.ends_with("+xml") || content_type == "application/xml" {
            b"<fixture/>".to_vec()
        } else {
            format!("fixture:{part_name}").into_bytes()
        };
        package
            .add_part(office_opc::OpcPart {
                name: part_name.to_string(),
                content_type: Some(content_type.to_string()),
                compression: office_opc::CompressionMethod::Stored,
                bytes,
            })
            .expect("add active-content fixture part");
    }

    package
        .to_bytes()
        .expect("serialize active-content workbook fixture")
}

#[allow(dead_code)]
pub fn macro_enabled_package_with_active_content(workbook_bytes: &[u8]) -> Vec<u8> {
    let active_content = package_with_active_content(workbook_bytes);
    let mut package =
        office_opc::OpcPackage::from_bytes(&active_content).expect("parse active-content fixture");
    let content_types = std::str::from_utf8(
        &package
            .part("[Content_Types].xml")
            .expect("content types")
            .bytes,
    )
    .expect("content types utf8")
    .replace(WORKBOOK_XLSX_CONTENT_TYPE, WORKBOOK_XLSM_CONTENT_TYPE);
    package
        .replace_part_bytes("[Content_Types].xml", content_types.into_bytes())
        .expect("retag macro-enabled workbook");
    package
        .to_bytes()
        .expect("serialize macro-enabled active-content fixture")
}

#[allow(dead_code)]
pub fn package_with_orphan_active_content_markers(workbook_bytes: &[u8]) -> Vec<u8> {
    let mut package =
        office_opc::OpcPackage::from_bytes(workbook_bytes).expect("parse workbook fixture");
    let content_types = std::str::from_utf8(
        &package
            .part("[Content_Types].xml")
            .expect("content types")
            .bytes,
    )
    .expect("content types utf8")
    .replace(
        "</Types>",
        r#"<Override PartName="/custom/missing-control.xml" ContentType="application/vnd.ms-office.activeX+xml"/></Types>"#,
    );
    package
        .replace_part_bytes("[Content_Types].xml", content_types.into_bytes())
        .expect("replace content types");
    let package_relationships = std::str::from_utf8(
        &package
            .part("_rels/.rels")
            .expect("package relationships")
            .bytes,
    )
    .expect("package relationships utf8")
    .replace(
        "</Relationships>",
        r#"<Relationship Id="rIdMissingPackage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="custom/missing-package.bin"/></Relationships>"#,
    );
    package
        .replace_part_bytes("_rels/.rels", package_relationships.into_bytes())
        .expect("replace package relationships");
    package
        .to_bytes()
        .expect("serialize orphan active-content marker fixture")
}
