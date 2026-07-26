use std::collections::BTreeMap;

use excel_oracle::{
    CaseArtifactRef, CaseInput, CaseOperation, CaseProbe, CaseProvenance, CaseSpec, CaseTier,
    ComparisonPolicy, EngineIdentity, EngineKind, ObservationDocument, ObservationResult,
    ObservedValue, OperationObservation, OracleSuiteManifest, ProbeObservation, RunBundle,
    RunCaseRecord, RunCaseStatus, RunManifest, build_behavioral_differential_report, sha256_hex,
};
use office_codegen::DifferentialCaseStatus;

struct BehavioralFixture {
    suite: OracleSuiteManifest,
    cases: BTreeMap<String, Vec<u8>>,
    oracle: RunBundle,
    runtime: RunBundle,
}

fn engine(kind: EngineKind) -> EngineIdentity {
    EngineIdentity {
        kind,
        version: if kind == EngineKind::Excel {
            "16.0".to_string()
        } else {
            env!("CARGO_PKG_VERSION").to_string()
        },
        build: if kind == EngineKind::Excel {
            "17928.20156".to_string()
        } else {
            "test".to_string()
        },
        channel: "Current".to_string(),
        os: "Windows 11".to_string(),
        architecture: "x64".to_string(),
        locale: "en-US".to_string(),
        timezone: "UTC".to_string(),
    }
}

fn observation(kind: EngineKind, value: &str) -> ObservationDocument {
    ObservationDocument {
        schema_version: 1,
        case_id: "application.name".to_string(),
        engine: engine(kind),
        operations: vec![OperationObservation {
            operation_index: 0,
            result: ObservationResult::Value(ObservedValue::Void),
        }],
        probes: vec![ProbeObservation {
            id: "application-name".to_string(),
            result: ObservationResult::Value(ObservedValue::Text(value.to_string())),
        }],
        save_reopen: None,
    }
}

fn fixture(runtime_value: &str) -> BehavioralFixture {
    let input_sha256 = sha256_hex(b"pinned workbook");
    let case = CaseSpec {
        schema_version: 1,
        id: "application.name".to_string(),
        version: 1,
        tier: CaseTier::MustMatch,
        input: CaseInput {
            path: "ranges/application-name.xlsx".to_string(),
            sha256: input_sha256.clone(),
            provenance: CaseProvenance {
                source: "Microsoft Excel desktop".to_string(),
                producer: "ootd oracle corpus".to_string(),
            },
        },
        profile_id: "excel-win-en-us".to_string(),
        operations: vec![CaseOperation::Calculate],
        probes: vec![CaseProbe {
            id: "application-name".to_string(),
            target: "application".to_string(),
            member: "Name".to_string(),
            args: Vec::new(),
        }],
    };
    let case_bytes = case.to_json_pretty().expect("serialize case").into_bytes();
    let case_sha256 = sha256_hex(&case_bytes);
    let suite = OracleSuiteManifest {
        schema_version: 1,
        id: "excel-win-en-us-smoke".to_string(),
        profile_id: case.profile_id.clone(),
        expected_engine: engine(EngineKind::Excel),
        cases: vec![CaseArtifactRef {
            case_id: case.id.clone(),
            case_version: case.version,
            tier: case.tier,
            path: "cases/application.name.json".to_string(),
            sha256: case_sha256.clone(),
            input_sha256: input_sha256.clone(),
        }],
    };
    let oracle_observation = observation(EngineKind::Excel, "Microsoft Excel");
    let runtime_observation = observation(EngineKind::Ootd, runtime_value);
    let oracle_bytes = serde_json::to_vec_pretty(&oracle_observation).expect("oracle JSON");
    let runtime_bytes = serde_json::to_vec_pretty(&runtime_observation).expect("runtime JSON");
    let oracle_record = RunCaseRecord {
        case_id: case.id.clone(),
        case_version: case.version,
        tier: case.tier,
        case_sha256: case_sha256.clone(),
        input_sha256: input_sha256.clone(),
        status: RunCaseStatus::Completed,
        observation_path: Some("observations/application.name/oracle.json".to_string()),
        observation_sha256: Some(sha256_hex(&oracle_bytes)),
        message: None,
    };
    let runtime_record = RunCaseRecord {
        observation_path: Some("observations/application.name/runtime.json".to_string()),
        observation_sha256: Some(sha256_hex(&runtime_bytes)),
        ..oracle_record.clone()
    };
    let oracle = RunBundle {
        manifest: RunManifest {
            schema_version: 1,
            run_id: "excel-win-en-us-20260726-a".to_string(),
            profile_id: case.profile_id.clone(),
            engine: oracle_observation.engine.clone(),
            cases: vec![oracle_record],
        },
        observations: BTreeMap::from([(case.id.clone(), oracle_bytes)]),
    };
    let runtime = RunBundle {
        manifest: RunManifest {
            schema_version: 1,
            run_id: "ootd-win-en-us-20260726-a".to_string(),
            profile_id: case.profile_id.clone(),
            engine: runtime_observation.engine.clone(),
            cases: vec![runtime_record],
        },
        observations: BTreeMap::from([(case.id.clone(), runtime_bytes)]),
    };

    BehavioralFixture {
        suite,
        cases: BTreeMap::from([(case.id, case_bytes)]),
        oracle,
        runtime,
    }
}

#[test]
fn maps_equal_typed_observations_to_a_passing_differential_gate() {
    let fixture = fixture("Microsoft Excel");

    let report = build_behavioral_differential_report(
        &fixture.suite,
        &fixture.cases,
        &fixture.oracle,
        &fixture.runtime,
        ComparisonPolicy::default(),
    )
    .expect("build report");

    assert_eq!(report.cases[0].status, DifferentialCaseStatus::Passed);
    assert_eq!(report.cases[0].expected.as_deref(), Some("matched"));
    assert!(report.gate_summary().passed);
}

#[test]
fn maps_typed_mismatches_to_stable_artifact_links_and_paths() {
    let fixture = fixture("OOTD Spreadsheet");

    let report = build_behavioral_differential_report(
        &fixture.suite,
        &fixture.cases,
        &fixture.oracle,
        &fixture.runtime,
        ComparisonPolicy::default(),
    )
    .expect("build report");
    let result = &report.cases[0];

    assert_eq!(result.status, DifferentialCaseStatus::Failed);
    assert!(
        result
            .message
            .as_deref()
            .expect("mismatch message")
            .contains("probes.application-name.value")
    );
    assert_eq!(
        result
            .artifacts
            .get("oracleObservation")
            .map(String::as_str),
        Some("observations/application.name/oracle.json")
    );
    assert_eq!(
        result
            .artifacts
            .get("runtimeObservation")
            .map(String::as_str),
        Some("observations/application.name/runtime.json")
    );
    assert!(!report.gate_summary().passed);
}

#[test]
fn maps_required_unsupported_runtime_cases_to_blocking_missing_runtime() {
    let mut fixture = fixture("Microsoft Excel");
    let record = &mut fixture.runtime.manifest.cases[0];
    record.status = RunCaseStatus::Unsupported;
    record.observation_path = None;
    record.observation_sha256 = None;
    record.message = Some("member is not implemented".to_string());
    fixture.runtime.observations.clear();

    let report = build_behavioral_differential_report(
        &fixture.suite,
        &fixture.cases,
        &fixture.oracle,
        &fixture.runtime,
        ComparisonPolicy::default(),
    )
    .expect("build incomplete report");

    assert_eq!(
        report.cases[0].status,
        DifferentialCaseStatus::MissingRuntime
    );
    assert!(!report.gate_summary().passed);
}
