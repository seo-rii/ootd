use office_codegen::summarize_json;

#[test]
fn summarizes_supported_members() {
    let json = r#"
    {
      "library": "Excel",
      "version": "16.0",
      "enums": [],
      "interfaces": [
        {
          "name": "Application",
          "kind": "dual",
          "members": [
            {
              "name": "Name",
              "memberKind": "property",
              "support": "implemented"
            },
            {
              "name": "Open",
              "memberKind": "method",
              "support": "partial"
            },
            {
              "name": "Legacy",
              "memberKind": "method",
              "support": "unsupported"
            }
          ]
        }
      ],
      "classes": [
        { "name": "ApplicationClass" }
      ]
    }
    "#;

    let summary = summarize_json(json).expect("summary should parse");

    assert_eq!(summary.library, "Excel");
    assert_eq!(summary.version, "16.0");
    assert_eq!(summary.enum_count, 0);
    assert_eq!(summary.interface_count, 1);
    assert_eq!(summary.class_count, 1);
    assert_eq!(summary.member_count, 3);
    assert_eq!(summary.implemented_member_count, 1);
    assert_eq!(summary.partial_member_count, 1);
    assert_eq!(summary.unsupported_member_count, 1);
}
