use excel_oracle::{
    CaseInput, CaseOperation, CaseProbe, CaseSpec, CaseTier, ComparisonPolicy, EngineIdentity,
    EngineKind, ObservationDocument, ObservationResult, ObservedArray, ObservedValue,
    OperationObservation, SaveReopenObservation, compare_observations,
};

const INPUT_SHA256: &str = "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef";

fn must_match_case() -> CaseSpec {
    CaseSpec {
        schema_version: 1,
        id: "formula2.sequence.basic".to_string(),
        version: 1,
        tier: CaseTier::MustMatch,
        input: CaseInput {
            path: "formula2/sequence-basic.xlsx".to_string(),
            sha256: INPUT_SHA256.to_string(),
        },
        operations: vec![
            CaseOperation::Get {
                target: "application".to_string(),
                member: "Workbooks".to_string(),
                args: Vec::new(),
                bind: Some("workbooks".to_string()),
            },
            CaseOperation::Calculate,
            CaseOperation::Save {
                workbook: "workbook".to_string(),
                output: "outputs/sequence-basic.xlsx".to_string(),
            },
        ],
        probes: vec![CaseProbe {
            id: "spill-values".to_string(),
            target: "sheet1-a1-a3".to_string(),
            member: "Value2".to_string(),
            args: Vec::new(),
        }],
    }
}

fn environment(kind: EngineKind) -> EngineIdentity {
    EngineIdentity {
        kind,
        version: "16.0".to_string(),
        build: "17928.20156".to_string(),
        channel: "Current".to_string(),
        os: "Windows 11".to_string(),
        architecture: "x64".to_string(),
        locale: "en-US".to_string(),
        timezone: "UTC".to_string(),
    }
}

fn observation(kind: EngineKind, rows: usize, cols: usize) -> ObservationDocument {
    ObservationDocument {
        schema_version: 1,
        case_id: "formula2.sequence.basic".to_string(),
        engine: environment(kind),
        operations: vec![OperationObservation {
            operation_index: 0,
            result: ObservationResult::Value(ObservedValue::Object("workbooks".to_string())),
        }],
        probes: vec![excel_oracle::ProbeObservation {
            id: "spill-values".to_string(),
            result: ObservationResult::Value(ObservedValue::Array(ObservedArray {
                rows,
                cols,
                values: vec![
                    ObservedValue::Number(1.0),
                    ObservedValue::Number(2.0),
                    ObservedValue::Number(3.0),
                ],
            })),
        }],
        save_reopen: Some(SaveReopenObservation {
            attempted: true,
            opened: true,
            repair_detected: Some(false),
        }),
    }
}

#[test]
fn round_trips_case_contract_and_typed_arrays() {
    let case = must_match_case();
    case.validate().expect("valid case");

    let json = case.to_json_pretty().expect("serialize case");
    let loaded = CaseSpec::from_json_str(&json).expect("reload case");

    assert_eq!(loaded, case);
    assert!(json.contains("\"mustMatch\""));
    assert!(json.contains("\"calculate\""));
}

#[test]
fn rejects_unsafe_input_paths_and_invalid_hashes() {
    for (path, sha256) in [
        ("../outside.xlsx", INPUT_SHA256),
        ("formula2\\case.xlsx", INPUT_SHA256),
        ("formula2/case.xlsx", "ABC"),
    ] {
        let mut case = must_match_case();
        case.input.path = path.to_string();
        case.input.sha256 = sha256.to_string();

        let error = case.validate().expect_err("invalid input must fail");
        assert!(error.to_string().contains("input"));
    }
}

#[test]
fn reports_array_shape_mismatches_with_a_stable_path() {
    let case = must_match_case();
    let oracle = observation(EngineKind::Excel, 3, 1);
    let runtime = observation(EngineKind::Ootd, 1, 3);

    let comparison = compare_observations(&case, &oracle, &runtime, ComparisonPolicy::default())
        .expect("valid comparison");

    assert!(!comparison.passed());
    assert_eq!(comparison.mismatches.len(), 2);
    assert_eq!(
        comparison.mismatches[0].path,
        "probes.spill-values.value.rows"
    );
    assert_eq!(
        comparison.mismatches[1].path,
        "probes.spill-values.value.cols"
    );
}

#[test]
fn must_match_save_cases_require_an_explicit_repair_result() {
    let case = must_match_case();
    let mut oracle = observation(EngineKind::Excel, 3, 1);
    oracle.save_reopen = Some(SaveReopenObservation {
        attempted: true,
        opened: true,
        repair_detected: None,
    });

    let error = oracle
        .validate_for_case(&case)
        .expect_err("unknown repair state must fail");

    assert!(error.to_string().contains("repairDetected"));
}
