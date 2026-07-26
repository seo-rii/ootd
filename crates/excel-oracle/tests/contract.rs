use excel_oracle::{
    CanonicalErrorKind, CaseArtifactRef, CaseInput, CaseOperation, CaseProbe, CaseProvenance,
    CaseSpec, CaseTier, ComparisonPolicy, EngineIdentity, EngineKind, NativeErrorDiagnostic,
    ObservationDocument, ObservationResult, ObservedArray, ObservedError, ObservedErrorKind,
    ObservedObject, ObservedValue, OperationObservation, OracleSuiteManifest, RunCaseRecord,
    RunCaseStatus, RunManifest, SaveReopenObservation, compare_observations, sha256_hex,
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
            provenance: CaseProvenance {
                source: "Microsoft Excel desktop".to_string(),
                producer: "ootd oracle corpus".to_string(),
            },
        },
        profile_id: "excel-win-en-us".to_string(),
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
        operations: vec![
            OperationObservation {
                operation_index: 0,
                result: ObservationResult::Value(ObservedValue::Object(ObservedObject {
                    type_name: "Workbooks".to_string(),
                    identity: "workbooks".to_string(),
                })),
            },
            OperationObservation {
                operation_index: 1,
                result: ObservationResult::Value(ObservedValue::Void),
            },
            OperationObservation {
                operation_index: 2,
                result: ObservationResult::Value(ObservedValue::Void),
            },
        ],
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
            normal_load_succeeded: true,
            repair_detected: Some(false),
            evidence: Some("Workbooks.Open CorruptLoad=xlNormalLoad".to_string()),
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
        normal_load_succeeded: true,
        repair_detected: None,
        evidence: Some("Workbooks.Open CorruptLoad=xlNormalLoad".to_string()),
    });

    let error = oracle
        .validate_for_case(&case)
        .expect_err("unknown repair state must fail");

    assert!(error.to_string().contains("repairDetected"));
}

#[test]
fn observations_must_cover_every_operation_and_represent_side_effects_as_void() {
    let case = must_match_case();
    let complete = observation(EngineKind::Excel, 3, 1);
    complete
        .validate_for_case(&case)
        .expect("complete operation observations");

    let mut incomplete = complete;
    incomplete.operations.pop();
    let error = incomplete
        .validate_for_case(&case)
        .expect_err("missing save observation must fail");

    assert!(error.to_string().contains("exactly cover"));
}

#[test]
fn structured_object_identity_round_trips_without_raw_handles() {
    let document = observation(EngineKind::Excel, 3, 1);
    let json = serde_json::to_string(&document).expect("serialize observation");

    assert!(json.contains(r#""typeName":"Workbooks""#));
    assert!(json.contains(r#""identity":"workbooks""#));
    assert!(!json.contains(r#""handle""#));
}

#[test]
fn requires_profile_provenance_and_normal_load_repair_evidence() {
    let mut case = must_match_case();
    case.profile_id.clear();
    assert!(
        case.validate()
            .expect_err("profile required")
            .to_string()
            .contains("profile")
    );

    let case = must_match_case();
    let mut document = observation(EngineKind::Excel, 3, 1);
    document.save_reopen.as_mut().expect("save reopen").evidence = None;
    let error = document
        .validate_for_case(&case)
        .expect_err("normal-load evidence required");
    assert!(error.to_string().contains("evidence"));
}

#[test]
fn rejects_unknown_fields_inside_tagged_operations() {
    let case = must_match_case();
    let json = case.to_json_pretty().expect("serialize case");
    let mutated = json.replacen(
        r#""operation": "get""#,
        r#""operation": "get", "unexpected": true"#,
        1,
    );

    let error = CaseSpec::from_json_str(&mutated).expect_err("unknown operation field must fail");
    assert!(error.to_string().contains("unknown field"));
}

#[test]
fn compares_canonical_errors_without_native_diagnostics() {
    let case = must_match_case();
    let mut oracle = observation(EngineKind::Excel, 3, 1);
    let mut runtime = observation(EngineKind::Ootd, 3, 1);
    oracle.operations[1].result = ObservationResult::Error(ObservedError {
        kind: CanonicalErrorKind::Calculation,
        code: "divisionByZero".to_string(),
        diagnostic: Some(NativeErrorDiagnostic {
            origin: ObservedErrorKind::ExcelCom,
            hresult: Some(-2_147_352_567),
            message: Some("localized Excel message".to_string()),
        }),
    });
    runtime.operations[1].result = ObservationResult::Error(ObservedError {
        kind: CanonicalErrorKind::Calculation,
        code: "divisionByZero".to_string(),
        diagnostic: Some(NativeErrorDiagnostic {
            origin: ObservedErrorKind::Ootd,
            hresult: None,
            message: Some("Calculation: division by zero".to_string()),
        }),
    });

    let comparison = compare_observations(&case, &oracle, &runtime, ComparisonPolicy::default())
        .expect("compare canonical errors");

    assert!(comparison.passed());
}

#[test]
fn default_number_comparison_is_exact_and_tolerance_is_opt_in() {
    let case = must_match_case();
    let oracle = observation(EngineKind::Excel, 3, 1);
    let mut runtime = observation(EngineKind::Ootd, 3, 1);
    let ObservationResult::Value(ObservedValue::Array(values)) = &mut runtime.probes[0].result
    else {
        panic!("array probe");
    };
    values.values[0] = ObservedValue::Number(1.0 + 5e-13);

    let exact = compare_observations(&case, &oracle, &runtime, ComparisonPolicy::default())
        .expect("exact comparison");
    let tolerant = compare_observations(
        &case,
        &oracle,
        &runtime,
        ComparisonPolicy {
            number_tolerance: 1e-12,
        },
    )
    .expect("tolerant comparison");

    assert!(!exact.passed());
    assert!(tolerant.passed());
}

fn manifest_fixture() -> (
    CaseSpec,
    Vec<u8>,
    Vec<u8>,
    OracleSuiteManifest,
    ObservationDocument,
    Vec<u8>,
    RunManifest,
) {
    let input_bytes = b"pinned xlsx fixture".to_vec();
    let mut case = must_match_case();
    case.input.sha256 = sha256_hex(&input_bytes);
    let case_bytes = case.to_json_pretty().expect("serialize case").into_bytes();
    let oracle = observation(EngineKind::Excel, 3, 1);
    let observation_bytes = serde_json::to_vec_pretty(&oracle).expect("serialize observation");
    let suite = OracleSuiteManifest {
        schema_version: 1,
        id: "excel-win-en-us-smoke".to_string(),
        profile_id: case.profile_id.clone(),
        expected_engine: environment(EngineKind::Excel),
        cases: vec![CaseArtifactRef {
            case_id: case.id.clone(),
            case_version: case.version,
            tier: case.tier,
            path: "cases/formula2.sequence.basic.json".to_string(),
            sha256: sha256_hex(&case_bytes),
            input_sha256: case.input.sha256.clone(),
        }],
    };
    let run = RunManifest {
        schema_version: 1,
        run_id: "excel-win-en-us-20260726-a".to_string(),
        profile_id: case.profile_id.clone(),
        engine: environment(EngineKind::Excel),
        cases: vec![RunCaseRecord {
            case_id: case.id.clone(),
            case_version: case.version,
            tier: case.tier,
            case_sha256: sha256_hex(&case_bytes),
            input_sha256: case.input.sha256.clone(),
            status: RunCaseStatus::Completed,
            observation_path: Some("observations/formula2.sequence.basic/oracle.json".to_string()),
            observation_sha256: Some(sha256_hex(&observation_bytes)),
            message: None,
        }],
    };

    (
        case,
        case_bytes,
        input_bytes,
        suite,
        oracle,
        observation_bytes,
        run,
    )
}

#[test]
fn validates_exact_byte_case_input_and_observation_artifacts() {
    let (case, case_bytes, input_bytes, suite, oracle, observation_bytes, run) = manifest_fixture();

    suite.validate().expect("valid suite");
    run.validate_for_suite(&suite).expect("complete run");
    let loaded_case = suite
        .load_case(&case.id, &case_bytes)
        .expect("case checksum and contract");
    suite.cases[0]
        .validate_input(&loaded_case, &input_bytes)
        .expect("input checksum");
    let loaded_observation = run.cases[0]
        .load_observation(&loaded_case, &run.engine, &observation_bytes)
        .expect("observation checksum and contract");

    assert_eq!(loaded_case, case);
    assert_eq!(loaded_observation, oracle);
    assert_eq!(
        sha256_hex(b"abc"),
        "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    );
}

#[test]
fn rejects_missing_extra_and_case_insensitive_duplicate_run_records() {
    let (_, _, _, suite, _, _, run) = manifest_fixture();

    let mut missing = run.clone();
    missing.cases.clear();
    assert!(
        missing
            .validate_for_suite(&suite)
            .expect_err("missing record")
            .to_string()
            .contains("exactly cover")
    );

    let mut duplicate = run;
    let mut second = duplicate.cases[0].clone();
    second.case_id = second.case_id.to_ascii_uppercase();
    duplicate.cases.push(second);
    assert!(
        duplicate
            .validate_for_suite(&suite)
            .expect_err("case-insensitive duplicate")
            .to_string()
            .contains("case-insensitive")
    );
}

#[test]
fn required_cases_cannot_be_failed_unsupported_or_skipped() {
    let (_, _, _, suite, _, _, mut run) = manifest_fixture();

    for status in [
        RunCaseStatus::Failed,
        RunCaseStatus::Unsupported,
        RunCaseStatus::Skipped,
    ] {
        run.cases[0].status = status;
        run.cases[0].observation_path = None;
        run.cases[0].observation_sha256 = None;
        run.cases[0].message = Some("runner did not complete".to_string());
        run.validate_for_suite(&suite)
            .expect("structurally valid incomplete run");
        assert!(
            run.validate_required_completeness(&suite)
                .expect_err("mustMatch status must block")
                .to_string()
                .contains("mustMatch")
        );
    }
}
