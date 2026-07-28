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
pub fn package_with_active_content_closure(workbook_bytes: &[u8]) -> Vec<u8> {
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
    .replace(
        "</Types>",
        r#"<Default Extension="bin" ContentType="application/octet-stream"/></Types>"#,
    );
    package
        .replace_part_bytes("[Content_Types].xml", content_types.into_bytes())
        .expect("replace content types");

    let workbook_xml =
        std::str::from_utf8(&package.part("xl/workbook.xml").expect("workbook").bytes)
            .expect("workbook utf8")
            .replace(
                "  </sheets>",
                r#"    <sheet name="Macro1" sheetId="2" r:id="rIdMacroSheet"/>
    <sheet name="Dialog1" sheetId="3" r:id="rIdDialogSheet"/>
  </sheets>"#,
            );
    package
        .replace_part_bytes("xl/workbook.xml", workbook_xml.into_bytes())
        .expect("replace workbook sheets");

    let worksheet_xml = std::str::from_utf8(
        &package
            .part("xl/worksheets/sheet1.xml")
            .expect("worksheet")
            .bytes,
    )
    .expect("worksheet utf8")
    .replace(
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:linked=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">",
    )
    .replace(
        "</worksheet>",
        r#"  <controls><control linked:id="rIdControl"/></controls>
  <oleObjects><oleObject linked:id="rIdOle"/></oleObjects>
</worksheet>"#,
    );
    package
        .replace_part_bytes("xl/worksheets/sheet1.xml", worksheet_xml.into_bytes())
        .expect("replace worksheet");

    package
        .add_part(office_opc::OpcPart {
            name: "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
            content_type: Some(
                "application/vnd.openxmlformats-package.relationships+xml".to_string(),
            ),
            compression: office_opc::CompressionMethod::Stored,
            bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdControl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" Target="../activeX/activeX1.xml"/>
  <Relationship Id="rIdOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
  <Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/shared.bin"/>
</Relationships>"#
                .to_vec(),
        })
        .expect("add worksheet relationships");
    package
        .add_part(office_opc::OpcPart {
            name: "xl/activeX/_rels/activeX1.xml.rels".to_string(),
            content_type: Some(
                "application/vnd.openxmlformats-package.relationships+xml".to_string(),
            ),
            compression: office_opc::CompressionMethod::Stored,
            bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdExclusive" Type="http://example.com/relationships/payload" Target="/custom/active-exclusive.bin"/>
  <Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/shared.bin"/>
</Relationships>"#
                .to_vec(),
        })
        .expect("add ActiveX relationships");
    for (part_name, contents) in [
        ("custom/active-exclusive.bin", b"exclusive".as_slice()),
        ("xl/media/shared.bin", b"shared".as_slice()),
    ] {
        package
            .add_part(office_opc::OpcPart {
                name: part_name.to_string(),
                content_type: Some("application/octet-stream".to_string()),
                compression: office_opc::CompressionMethod::Stored,
                bytes: contents.to_vec(),
            })
            .expect("add active-content closure fixture part");
    }

    package
        .to_bytes()
        .expect("serialize active-content closure fixture")
}

#[allow(dead_code)]
pub fn macro_enabled_package_with_active_content(workbook_bytes: &[u8]) -> Vec<u8> {
    retag_macro_enabled(package_with_active_content(workbook_bytes))
}

#[allow(dead_code)]
pub fn macro_enabled_package_with_active_content_closure(workbook_bytes: &[u8]) -> Vec<u8> {
    retag_macro_enabled(package_with_active_content_closure(workbook_bytes))
}

fn retag_macro_enabled(active_content: Vec<u8>) -> Vec<u8> {
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
