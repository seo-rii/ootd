#[allow(dead_code)]
pub const EXTERNAL_DATA_PARTS: [&str; 6] = [
    "xl/externalLinks/externalLink1.xml",
    "xl/externalLinks/externalLink2.xml",
    "xl/externalLinks/externalLink3.xml",
    "xl/connections.xml",
    "xl/queryTables/queryTable1.xml",
    "xl/model/item.data",
];

pub fn package_with_external_data(workbook_bytes: &[u8]) -> Vec<u8> {
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
        r#"<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
<Override PartName="/xl/externalLinks/externalLink2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
<Override PartName="/xl/externalLinks/externalLink3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>
<Override PartName="/xl/queryTables/queryTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>
<Override PartName="/xl/model/item.data" ContentType="application/vnd.ms-excel.model"/>
</Types>"#,
    );
    package
        .replace_part_bytes("[Content_Types].xml", content_types.into_bytes())
        .expect("replace content types");

    let workbook_xml =
        std::str::from_utf8(&package.part("xl/workbook.xml").expect("workbook").bytes)
            .expect("workbook utf8")
            .replace(
                "</sheets>",
                r#"</sheets><externalReferences>
  <externalReference r:id="rIdExternalBook"/>
  <externalReference r:id="rIdDde"/>
  <externalReference r:id="rIdOle"/>
</externalReferences>"#,
            );
    package
        .replace_part_bytes("xl/workbook.xml", workbook_xml.into_bytes())
        .expect("replace workbook external references");

    let workbook_relationships = std::str::from_utf8(
        &package
            .part("xl/_rels/workbook.xml.rels")
            .expect("workbook relationships")
            .bytes,
    )
    .expect("workbook relationships utf8")
    .replace(
        "</Relationships>",
        r#"<Relationship Id="rIdExternalBook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
<Relationship Id="rIdDde" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink2.xml"/>
<Relationship Id="rIdOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink3.xml"/>
<Relationship Id="rIdConnections" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>
<Relationship Id="rIdModel" Type="http://schemas.microsoft.com/office/2007/relationships/model" Target="model/item.data"/>
</Relationships>"#,
    );
    package
        .replace_part_bytes(
            "xl/_rels/workbook.xml.rels",
            workbook_relationships.into_bytes(),
        )
        .expect("replace workbook relationships");

    let parts = [
        (
            "xl/externalLinks/externalLink1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
            br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rIdPath"><sheetNames><sheetName val="Cached"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1" t="n"><v>42</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>"#.as_slice(),
        ),
        (
            "xl/externalLinks/externalLink2.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
            br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><ddeLink ddeService="fixture-service" ddeTopic="fixture-topic"><ddeItems><ddeItem name="Item"><values rows="1" cols="1"><value t="str"><val>cached-dde</val></value></values></ddeItem></ddeItems></ddeLink></externalLink>"#.as_slice(),
        ),
        (
            "xl/externalLinks/externalLink3.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
            br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><oleLink progId="Fixture.Provider"><oleItems><oleItem name="Item"><values rows="1" cols="1"><value t="str"><val>cached-ole</val></value></values></oleItem></oleItems></oleLink></externalLink>"#.as_slice(),
        ),
        (
            "xl/connections.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml",
            br#"<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><connection id="1" name="OfflineFixture" type="4" refreshedVersion="8" saveData="1"><webPr url="https://example.invalid/data.csv"/></connection></connections>"#.as_slice(),
        ),
        (
            "xl/queryTables/queryTable1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
            br#"<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="OfflineFixture" connectionId="1" refreshOnLoad="1"><queryTableRefresh nextId="1"><queryTableFields count="0"/></queryTableRefresh></queryTable>"#.as_slice(),
        ),
        (
            "xl/model/item.data",
            "application/vnd.ms-excel.model",
            b"cached-data-model".as_slice(),
        ),
    ];
    for (name, content_type, bytes) in parts {
        package
            .add_part(office_opc::OpcPart {
                name: name.to_string(),
                content_type: Some(content_type.to_string()),
                compression: office_opc::CompressionMethod::Stored,
                bytes: bytes.to_vec(),
            })
            .expect("add external-data fixture part");
    }

    package
        .add_part(office_opc::OpcPart {
            name: "xl/externalLinks/_rels/externalLink1.xml.rels".to_string(),
            content_type: Some(
                "application/vnd.openxmlformats-package.relationships+xml".to_string(),
            ),
            compression: office_opc::CompressionMethod::Stored,
            bytes: br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPath" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///unreachable/external.xlsx" TargetMode="External"/></Relationships>"#.to_vec(),
        })
        .expect("add external workbook path relationship");
    package
        .add_part(office_opc::OpcPart {
            name: "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
            content_type: Some(
                "application/vnd.openxmlformats-package.relationships+xml".to_string(),
            ),
            compression: office_opc::CompressionMethod::Stored,
            bytes: br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdQueryTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" Target="../queryTables/queryTable1.xml"/></Relationships>"#.to_vec(),
        })
        .expect("add query table relationship");

    package
        .to_bytes()
        .expect("serialize external-data workbook fixture")
}
