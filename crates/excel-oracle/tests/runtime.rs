use excel_oracle::{
    CanonicalErrorKind, CaseInput, CaseOperation, CaseProbe, CaseProvenance, CaseSpec, CaseTier,
    EngineIdentity, EngineKind, ObservationResult, ObservedArray, ObservedObject, ObservedValue,
    run_runtime_case, sha256_hex,
};
use excel_runtime::ExcelRuntime;
use office_common::{ExcelProfile, FileFormat, SaveWorkbookSpec};

fn blank_workbook_bytes() -> Vec<u8> {
    let mut runtime = ExcelRuntime::new();
    let workbook = runtime.create_workbook().expect("create blank workbook");
    runtime
        .save_workbook(
            workbook,
            SaveWorkbookSpec {
                format: FileFormat::Xlsx,
                profile: ExcelProfile::Excel365,
                lossless: true,
            },
        )
        .expect("save blank workbook")
}

fn runtime_engine() -> EngineIdentity {
    EngineIdentity {
        kind: EngineKind::Ootd,
        version: env!("CARGO_PKG_VERSION").to_string(),
        build: "test".to_string(),
        channel: "source".to_string(),
        os: std::env::consts::OS.to_string(),
        architecture: std::env::consts::ARCH.to_string(),
        locale: "en-US".to_string(),
        timezone: "UTC".to_string(),
    }
}

fn range_case(input: &[u8]) -> CaseSpec {
    CaseSpec {
        schema_version: 1,
        id: "range.value2.array".to_string(),
        version: 1,
        tier: CaseTier::MustMatch,
        input: CaseInput {
            path: "ranges/value2-array.xlsx".to_string(),
            sha256: sha256_hex(input),
            provenance: CaseProvenance {
                source: "Microsoft Excel desktop".to_string(),
                producer: "ootd oracle corpus".to_string(),
            },
        },
        profile_id: "excel-win-en-us".to_string(),
        operations: vec![
            CaseOperation::Get {
                target: "workbook".to_string(),
                member: "Worksheets".to_string(),
                args: Vec::new(),
                bind: Some("worksheets".to_string()),
            },
            CaseOperation::Invoke {
                target: "worksheets".to_string(),
                member: "Item".to_string(),
                args: vec![ObservedValue::Number(1.0)],
                bind: Some("sheet".to_string()),
            },
            CaseOperation::Invoke {
                target: "sheet".to_string(),
                member: "Range".to_string(),
                args: vec![ObservedValue::Text("A1:B2".to_string())],
                bind: Some("range".to_string()),
            },
            CaseOperation::Set {
                target: "range".to_string(),
                member: "Value2".to_string(),
                value: ObservedValue::Array(ObservedArray {
                    rows: 2,
                    cols: 2,
                    values: vec![
                        ObservedValue::Number(1.0),
                        ObservedValue::Number(2.0),
                        ObservedValue::Number(3.0),
                        ObservedValue::Number(4.0),
                    ],
                }),
                args: Vec::new(),
            },
            CaseOperation::Get {
                target: "range".to_string(),
                member: "Value2".to_string(),
                args: Vec::new(),
                bind: None,
            },
        ],
        probes: vec![CaseProbe {
            id: "range-values".to_string(),
            target: "range".to_string(),
            member: "Value2".to_string(),
            args: Vec::new(),
        }],
    }
}

#[test]
fn executes_get_set_and_invoke_with_typed_arrays_and_symbolic_objects() {
    let input = blank_workbook_bytes();
    let case = range_case(&input);

    let output = run_runtime_case(&case, &input, runtime_engine()).expect("run case");

    assert_eq!(
        output.observation.operations[0].result,
        ObservationResult::Value(ObservedValue::Object(ObservedObject {
            type_name: "Object".to_string(),
            identity: "worksheets".to_string(),
        }))
    );
    assert_eq!(
        output.observation.operations[3].result,
        ObservationResult::Value(ObservedValue::Void)
    );
    assert_eq!(
        output.observation.probes[0].result,
        ObservationResult::Value(ObservedValue::Array(ObservedArray {
            rows: 2,
            cols: 2,
            values: vec![
                ObservedValue::Number(1.0),
                ObservedValue::Number(2.0),
                ObservedValue::Number(3.0),
                ObservedValue::Number(4.0),
            ],
        }))
    );
    let json = serde_json::to_string(&output.observation).expect("observation JSON");
    assert!(!json.contains("1000000"));
}

#[test]
fn records_dispatch_errors_and_continues_to_later_operations_and_probes() {
    let input = blank_workbook_bytes();
    let mut case = range_case(&input);
    case.operations.insert(
        0,
        CaseOperation::Invoke {
            target: "application".to_string(),
            member: "DefinitelyMissing".to_string(),
            args: Vec::new(),
            bind: None,
        },
    );

    let output = run_runtime_case(&case, &input, runtime_engine()).expect("run after OM error");

    let ObservationResult::Error(error) = &output.observation.operations[0].result else {
        panic!("missing member must be recorded as an error");
    };
    assert_eq!(error.kind, CanonicalErrorKind::NotFound);
    assert_eq!(output.observation.probes.len(), 1);
    assert_eq!(output.observation.operations.len(), case.operations.len());
}

#[test]
fn rejects_input_bytes_that_do_not_match_the_pinned_hash() {
    let input = blank_workbook_bytes();
    let case = range_case(&input);

    let error = run_runtime_case(&case, b"not the workbook", runtime_engine())
        .expect_err("hash mismatch must fail before opening");

    assert!(error.to_string().contains("sha256"));
}
