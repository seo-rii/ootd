use std::{fs, io::Cursor};

use office_idl::{
    CaptureOriginKind, IdlLoadError, MemberKind, OfficeIdlDocument, SupportState, TypeRefKind,
    load_json_slice, load_json_str,
};

#[test]
fn loads_a_minimal_idl_document() {
    let json = r#"
    {
      "library": "Excel",
      "version": "16.0",
      "metadata": {
        "namespace": "Microsoft.Office.Interop.Excel",
        "typeLibraryGuid": "{00020813-0000-0000-C000-000000000046}",
        "sourceInherits": ["IDispatch"]
      },
      "enums": [
        {
          "name": "ExampleEnum",
          "metadata": {
            "typeLibraryGuid": "{enum-guid}"
          },
          "values": [
            { "name": "A", "value": 1 }
          ]
        }
      ],
      "interfaces": [
        {
          "name": "Workbook",
          "kind": "dual",
          "metadata": {
            "iid": "{000208DA-0000-0000-C000-000000000046}",
            "sourceInherits": ["_Workbook"]
          },
          "members": [
            {
              "name": "Name",
              "memberKind": "property",
              "access": "read",
              "support": "implemented",
              "metadata": {
                "capture": {
                  "origins": [
                    {
                      "kind": "property_get",
                      "sourceInterface": "Workbook",
                      "sourceMember": "Name",
                      "dispId": 6
                    }
                  ]
                },
                "iid": "{member-guid}"
              }
            },
            {
              "name": "Range",
              "memberKind": "method",
              "support": "partial",
              "params": [
                {
                  "name": "address",
                  "type": {
                    "kind": "primitive",
                    "name": "string",
                    "aliasOf": "BSTR"
                  }
                }
              ],
              "returnType": {
                "kind": "interface",
                "name": "Range",
                "aliasOf": "Excel.Range",
                "metadata": {
                  "namespace": "Microsoft.Office.Interop.Excel",
                  "sourceInherits": ["INameable"],
                  "capture": {
                    "typeInfo": {
                      "kind": "interface",
                      "name": "Range",
                      "aliasOf": "Excel.Range"
                    }
                  }
                }
              },
              "metadata": {
                "capture": {
                  "typeInfo": {
                    "kind": "interface",
                    "name": "Range",
                    "aliasOf": "Excel.Range"
                  }
                }
              }
            }
          ]
        }
      ],
      "classes": [
        {
          "name": "WorkbookClass",
          "implements": ["Workbook"],
          "defaultInterface": "Workbook",
          "metadata": {
            "clsid": "{00020819-0000-0000-C000-000000000046}",
            "sourceDefaultInterface": "Workbook"
          }
        }
      ]
    }
    "#;

    let document = load_json_str(json).expect("document should parse");

    assert_eq!(document.library, "Excel");
    assert_eq!(document.version, "16.0");
    assert_eq!(
        document
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.type_library_guid.as_deref()),
        Some("{00020813-0000-0000-C000-000000000046}")
    );
    assert_eq!(
        document
            .metadata
            .as_ref()
            .map(|metadata| metadata.source_inherits.clone()),
        Some(vec!["IDispatch".to_string()])
    );
    assert_eq!(document.enums.len(), 1);
    assert_eq!(document.interfaces.len(), 1);
    assert_eq!(document.classes.len(), 1);
    assert_eq!(
        document.interfaces[0].members[0].member_kind,
        MemberKind::Property
    );
    assert_eq!(
        document.interfaces[0].members[0].support,
        SupportState::Implemented
    );
    assert_eq!(
        document.interfaces[0].members[1].params[0].type_ref.kind,
        TypeRefKind::Primitive
    );
    assert_eq!(
        document.interfaces[0]
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.iid.as_deref()),
        Some("{000208DA-0000-0000-C000-000000000046}")
    );
    assert_eq!(
        document.interfaces[0]
            .metadata
            .as_ref()
            .map(|metadata| metadata.source_inherits.clone()),
        Some(vec!["_Workbook".to_string()])
    );
    assert_eq!(
        document.classes[0]
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.clsid.as_deref()),
        Some("{00020819-0000-0000-C000-000000000046}")
    );
    assert_eq!(
        document.classes[0]
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.source_default_interface.as_deref()),
        Some("Workbook")
    );
    assert_eq!(
        document.interfaces[0].members[0]
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.capture.as_ref())
            .and_then(|capture| capture.origins.first())
            .map(|origin| origin.kind.clone()),
        Some(CaptureOriginKind::PropertyGet)
    );
    assert_eq!(
        document.interfaces[0].members[1]
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.metadata.as_ref())
            .and_then(|metadata| metadata.namespace.as_deref()),
        Some("Microsoft.Office.Interop.Excel")
    );
    assert_eq!(
        document.interfaces[0].members[1]
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.metadata.as_ref())
            .and_then(|metadata| metadata.source_inherits.first())
            .map(String::as_str),
        Some("INameable")
    );
    assert_eq!(
        document.interfaces[0].members[1]
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.metadata.as_ref())
            .and_then(|metadata| metadata.capture.as_ref())
            .and_then(|capture| capture.type_info.as_ref())
            .map(|type_ref| type_ref.name.as_str()),
        Some("Range")
    );
    assert_eq!(
        document.interfaces[0].members[1]
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.capture.as_ref())
            .and_then(|capture| capture.type_info.as_ref())
            .map(|type_ref| type_ref.name.as_str()),
        Some("Range")
    );
    assert_eq!(
        document.interfaces[0].members[1].params[0]
            .type_ref
            .alias_of
            .as_deref(),
        Some("BSTR")
    );
    assert_eq!(
        document.interfaces[0].members[1]
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
}

#[test]
fn round_trips_to_pretty_json() {
    let json = r#"
    {
      "library": "Excel",
      "version": "16.0",
      "interfaces": []
    }
    "#;

    let document = load_json_str(json).expect("document should parse");
    let pretty = document
        .to_json_pretty()
        .expect("document should serialize");

    assert!(pretty.contains("\"library\": \"Excel\""));
    assert!(pretty.contains("\"version\": \"16.0\""));
}

#[test]
fn round_trips_nested_type_metadata() {
    let json = r#"
    {
      "library": "Excel",
      "version": "16.0",
      "interfaces": [
        {
          "name": "Workbook",
          "kind": "dual",
          "members": [
            {
              "name": "Range",
              "memberKind": "method",
              "support": "implemented",
              "returnType": {
                "kind": "interface",
                "name": "Range",
                "aliasOf": "Excel.Range",
                "metadata": {
                  "namespace": "Microsoft.Office.Interop.Excel",
                  "sourceInherits": ["INameable"],
                  "capture": {
                    "typeInfo": {
                      "kind": "interface",
                      "name": "Range",
                      "aliasOf": "Excel.Range"
                    }
                  }
                }
              }
            }
          ]
        }
      ]
    }
    "#;

    let document = load_json_str(json).expect("document should parse");
    let pretty = document
        .to_json_pretty()
        .expect("document should serialize");
    let reparsed = load_json_str(&pretty).expect("pretty json should parse");

    assert_eq!(
        reparsed.interfaces[0].members[0]
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.metadata.as_ref())
            .and_then(|metadata| metadata.namespace.as_deref()),
        Some("Microsoft.Office.Interop.Excel")
    );
    assert_eq!(
        reparsed.interfaces[0].members[0]
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.metadata.as_ref())
            .and_then(|metadata| metadata.capture.as_ref())
            .and_then(|capture| capture.type_info.as_ref())
            .map(|type_ref| type_ref.alias_of.as_deref()),
        Some(Some("Excel.Range"))
    );
}

#[test]
fn loads_from_path() {
    let json = r#"
    {
      "library": "Excel",
      "version": "16.0",
      "interfaces": []
    }
    "#;

    let temp_dir = tempfile::tempdir().expect("temp dir");
    let path = temp_dir.path().join("office-idl.json");
    fs::write(&path, json).expect("write json");

    let document = OfficeIdlDocument::from_path(&path).expect("document should parse");
    assert_eq!(document.library, "Excel");
}

#[test]
fn loads_from_json_slice_and_reader_with_array_parameter_defaults() {
    let json = br#"
    {
      "library": "Excel",
      "version": "16.0",
      "interfaces": [
        {
          "name": "Worksheet",
          "kind": "dual",
          "members": [
            {
              "name": "SetValues",
              "memberKind": "method",
              "support": "partial",
              "params": [
                {
                  "name": "addresses",
                  "type": {
                    "kind": "array",
                    "name": "SafeArray",
                    "nullable": true,
                    "itemType": {
                      "kind": "primitive",
                      "name": "string",
                      "aliasOf": "BSTR"
                    }
                  },
                  "optional": true,
                  "byRef": true,
                  "defaultValue": ["A1", "B2"]
                }
              ]
            }
          ]
        }
      ]
    }
    "#;

    let slice_document = load_json_slice(json).expect("slice document should parse");
    let reader_document =
        OfficeIdlDocument::from_reader(Cursor::new(json)).expect("reader document should parse");
    let pretty = slice_document
        .to_json_pretty()
        .expect("document should serialize");
    let reparsed = load_json_str(&pretty).expect("pretty json should parse");

    assert_eq!(slice_document, reader_document);
    assert_eq!(slice_document, reparsed);
    assert_eq!(
        slice_document.interfaces[0].members[0].params[0]
            .type_ref
            .kind,
        TypeRefKind::Array
    );
    assert!(
        slice_document.interfaces[0].members[0].params[0]
            .type_ref
            .nullable
    );
    assert_eq!(
        slice_document.interfaces[0].members[0].params[0]
            .type_ref
            .item_type
            .as_ref()
            .map(|type_ref| type_ref.kind.clone()),
        Some(TypeRefKind::Primitive)
    );
    assert!(slice_document.interfaces[0].members[0].params[0].optional);
    assert!(slice_document.interfaces[0].members[0].params[0].by_ref);
    assert_eq!(
        slice_document.interfaces[0].members[0].params[0].default_value,
        Some(serde_json::json!(["A1", "B2"]))
    );
}

#[test]
fn pretty_json_omits_default_optional_fields() {
    let json = r#"
    {
      "library": "Excel",
      "version": "16.0",
      "interfaces": [
        {
          "name": "Application",
          "kind": "dual",
          "members": [
            {
              "name": "Visible",
              "memberKind": "property",
              "support": "stub"
            }
          ]
        }
      ]
    }
    "#;

    let pretty = load_json_str(json)
        .expect("document should parse")
        .to_json_pretty()
        .expect("document should serialize");

    assert!(!pretty.contains("\"access\""));
    assert!(!pretty.contains("\"params\""));
    assert!(!pretty.contains("\"inherits\""));
    assert!(!pretty.contains("\"metadata\""));
}

#[test]
fn from_path_reports_json_error_for_invalid_json() {
    let temp_dir = tempfile::tempdir().expect("temp dir");
    let path = temp_dir.path().join("office-idl.json");
    fs::write(&path, "{").expect("write invalid json");

    let error = OfficeIdlDocument::from_path(&path).expect_err("invalid json should fail");
    match error {
        IdlLoadError::Json(_) => {}
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn from_path_reports_io_error_for_missing_file() {
    let temp_dir = tempfile::tempdir().expect("temp dir");
    let path = temp_dir.path().join("missing-office-idl.json");

    let error = OfficeIdlDocument::from_path(&path).expect_err("missing file should fail");
    match error {
        IdlLoadError::Io(_) => {}
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn pretty_json_uses_camel_case_and_type_field_names_for_optional_metadata() {
    let json = r#"
    {
      "library": "Excel",
      "version": "16.0",
      "interfaces": [
        {
          "name": "Worksheet",
          "kind": "dual",
          "members": [
            {
              "name": "SetValue",
              "memberKind": "method",
              "support": "implemented",
              "params": [
                {
                  "name": "value",
                  "type": {
                    "kind": "primitive",
                    "name": "string",
                    "aliasOf": "BSTR"
                  },
                  "byRef": true,
                  "defaultValue": "fallback"
                }
              ],
              "metadata": {
                "capture": {
                  "origins": [
                    {
                      "kind": "method",
                      "sourceInterface": "Worksheet",
                      "sourceMember": "SetValue",
                      "dispId": 200
                    }
                  ]
                }
              }
            }
          ]
        }
      ],
      "classes": [
        {
          "name": "WorksheetClass",
          "defaultInterface": "Worksheet",
          "metadata": {
            "sourceDefaultInterface": "Worksheet"
          }
        }
      ]
    }
    "#;

    let pretty = load_json_str(json)
        .expect("document should parse")
        .to_json_pretty()
        .expect("document should serialize");

    assert!(pretty.contains("\"memberKind\": \"method\""));
    assert!(pretty.contains("\"type\": {"));
    assert!(pretty.contains("\"aliasOf\": \"BSTR\""));
    assert!(pretty.contains("\"byRef\": true"));
    assert!(pretty.contains("\"defaultValue\": \"fallback\""));
    assert!(pretty.contains("\"sourceInterface\": \"Worksheet\""));
    assert!(pretty.contains("\"sourceMember\": \"SetValue\""));
    assert!(pretty.contains("\"sourceDefaultInterface\": \"Worksheet\""));
}

#[test]
fn from_reader_reports_json_error_for_invalid_json() {
    let error =
        OfficeIdlDocument::from_reader(Cursor::new(b"{".as_slice())).expect_err("invalid json");

    match error {
        IdlLoadError::Json(_) => {}
        other => panic!("unexpected error: {other:?}"),
    }
}
