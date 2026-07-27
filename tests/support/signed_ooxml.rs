const SIGNATURE_ORIGIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-package.digital-signature-origin";
const SIGNATURE_XML_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";

pub fn package_with_fake_digital_signature(workbook_bytes: &[u8]) -> Vec<u8> {
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
        &format!(
            r#"<Override PartName="/_xmlsignatures/origin.sigs" ContentType="{SIGNATURE_ORIGIN_CONTENT_TYPE}"/><Override PartName="/_xmlsignatures/sig1.xml" ContentType="{SIGNATURE_XML_CONTENT_TYPE}"/></Types>"#
        ),
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
        r#"<Relationship Id="rIdSignatureOrigin" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="_xmlsignatures/origin.sigs"/></Relationships>"#,
    );
    package
        .replace_part_bytes("_rels/.rels", package_relationships.into_bytes())
        .expect("replace package relationships");

    package
        .add_part(office_opc::OpcPart {
            name: "_xmlsignatures/origin.sigs".to_string(),
            content_type: Some(SIGNATURE_ORIGIN_CONTENT_TYPE.to_string()),
            compression: office_opc::CompressionMethod::Stored,
            bytes: Vec::new(),
        })
        .expect("add signature origin");
    package
        .add_part(office_opc::OpcPart {
            name: "_xmlsignatures/_rels/origin.sigs.rels".to_string(),
            content_type: Some(
                "application/vnd.openxmlformats-package.relationships+xml".to_string(),
            ),
            compression: office_opc::CompressionMethod::Stored,
            bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSignature" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="sig1.xml"/>
</Relationships>"#
                .to_vec(),
        })
        .expect("add signature relationships");
    package
        .add_part(office_opc::OpcPart {
            name: "_xmlsignatures/sig1.xml".to_string(),
            content_type: Some(SIGNATURE_XML_CONTENT_TYPE.to_string()),
            compression: office_opc::CompressionMethod::Stored,
            bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>"#
                .to_vec(),
        })
        .expect("add signature XML");

    package
        .to_bytes()
        .expect("serialize signed workbook fixture")
}

#[allow(dead_code)]
pub fn package_with_orphan_signature_part(workbook_bytes: &[u8]) -> Vec<u8> {
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
        &format!(
            r#"<Override PartName="/custom/orphan-signature.xml" ContentType="{SIGNATURE_XML_CONTENT_TYPE}"/></Types>"#
        ),
    );
    package
        .replace_part_bytes("[Content_Types].xml", content_types.into_bytes())
        .expect("replace content types");
    package
        .add_part(office_opc::OpcPart {
            name: "custom/orphan-signature.xml".to_string(),
            content_type: Some(SIGNATURE_XML_CONTENT_TYPE.to_string()),
            compression: office_opc::CompressionMethod::Stored,
            bytes: br#"<?xml version="1.0" encoding="UTF-8"?>
<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>"#
                .to_vec(),
        })
        .expect("add orphan signature XML");
    package
        .to_bytes()
        .expect("serialize orphan signature fixture")
}
