use std::fs;
use std::path::PathBuf;
use std::process::Command;

use excel_oracle::{
    CaseArtifactRef, CaseInput, CaseOperation, CaseProbe, CaseProvenance, CaseSpec, CaseTier,
    EngineIdentity, EngineKind, ObservationDocument, ObservationResult, ObservedValue,
    OperationObservation, OracleSuiteManifest, PinnedSuiteArtifacts, ProbeObservation,
    RunCaseRecord, RunCaseStatus, RunManifest, sha256_hex,
};

struct CorpusFixture {
    root: PathBuf,
    corpus_root: PathBuf,
    run_root: PathBuf,
    input_path: PathBuf,
    observation_path: PathBuf,
    input_bytes: Vec<u8>,
    observation_bytes: Vec<u8>,
}

impl CorpusFixture {
    fn create() -> Self {
        let root = std::env::temp_dir().join(format!(
            "ootd-oracle-corpus-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .expect("system clock")
                .as_nanos(),
        ));
        let corpus_root = root.join("corpus");
        let run_root = root.join("run-a");
        let input_path = corpus_root.join("inputs/application-name.xlsx");
        let case_path = corpus_root.join("cases/application.name.json");
        let suite_path = corpus_root.join("manifest/suite_manifest.json");
        let observation_path = run_root.join("observations/application.name/oracle.json");
        let run_path = run_root.join("manifest/run_manifest.json");
        for path in [
            &input_path,
            &case_path,
            &suite_path,
            &observation_path,
            &run_path,
        ] {
            fs::create_dir_all(path.parent().expect("fixture parent"))
                .expect("create fixture directory");
        }

        let input_bytes = b"pinned workbook bytes".to_vec();
        fs::write(&input_path, &input_bytes).expect("write input");
        let excel_engine = engine(EngineKind::Excel);
        let case = CaseSpec {
            schema_version: 1,
            id: "application.name".to_string(),
            version: 1,
            tier: CaseTier::MustMatch,
            input: CaseInput {
                path: "inputs/application-name.xlsx".to_string(),
                sha256: sha256_hex(&input_bytes),
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
        fs::write(&case_path, &case_bytes).expect("write case");
        let suite = OracleSuiteManifest {
            schema_version: 1,
            id: "excel-win-en-us-smoke".to_string(),
            profile_id: case.profile_id.clone(),
            expected_engine: excel_engine.clone(),
            cases: vec![CaseArtifactRef {
                case_id: case.id.clone(),
                case_version: case.version,
                tier: case.tier,
                path: "cases/application.name.json".to_string(),
                sha256: sha256_hex(&case_bytes),
                input_sha256: case.input.sha256.clone(),
            }],
        };
        fs::write(
            &suite_path,
            suite.to_json_pretty().expect("serialize suite"),
        )
        .expect("write suite");

        let observation = ObservationDocument {
            schema_version: 1,
            case_id: case.id.clone(),
            engine: excel_engine.clone(),
            operations: vec![OperationObservation {
                operation_index: 0,
                result: ObservationResult::Value(ObservedValue::Void),
            }],
            probes: vec![ProbeObservation {
                id: "application-name".to_string(),
                result: ObservationResult::Value(ObservedValue::Text(
                    "Microsoft Excel".to_string(),
                )),
            }],
            save_reopen: None,
        };
        let observation_bytes =
            serde_json::to_vec_pretty(&observation).expect("serialize observation");
        fs::write(&observation_path, &observation_bytes).expect("write observation");
        let run = RunManifest {
            schema_version: 1,
            run_id: "excel-win-en-us-20260810-a".to_string(),
            profile_id: case.profile_id.clone(),
            engine: excel_engine,
            cases: vec![RunCaseRecord {
                case_id: case.id,
                case_version: case.version,
                tier: case.tier,
                case_sha256: sha256_hex(&case_bytes),
                input_sha256: case.input.sha256,
                status: RunCaseStatus::Completed,
                observation_path: Some("observations/application.name/oracle.json".to_string()),
                observation_sha256: Some(sha256_hex(&observation_bytes)),
                message: None,
            }],
        };
        fs::write(
            &run_path,
            run.to_json_pretty(&suite).expect("serialize run"),
        )
        .expect("write run");

        Self {
            root,
            corpus_root,
            run_root,
            input_path,
            observation_path,
            input_bytes,
            observation_bytes,
        }
    }

    fn copy_run(&self, directory: &str, run_id: &str, probe_value: &str) -> PathBuf {
        let suite = PinnedSuiteArtifacts::load(&self.corpus_root).expect("load fixture suite");
        let source_manifest_path = self.run_root.join("manifest/run_manifest.json");
        let source_manifest = fs::read_to_string(source_manifest_path).expect("read source run");
        let mut manifest = RunManifest::from_json_str(&suite.manifest, &source_manifest)
            .expect("parse source run");
        manifest.run_id = run_id.to_string();

        let mut observation: ObservationDocument =
            serde_json::from_slice(&self.observation_bytes).expect("parse source observation");
        let ObservationResult::Value(ObservedValue::Text(value)) =
            &mut observation.probes[0].result
        else {
            panic!("expected text probe");
        };
        *value = probe_value.to_string();
        let observation_bytes =
            serde_json::to_vec_pretty(&observation).expect("serialize copied observation");
        manifest.cases[0].observation_sha256 = Some(sha256_hex(&observation_bytes));

        let run_root = self.root.join(directory);
        let observation_path = run_root.join("observations/application.name/oracle.json");
        let manifest_path = run_root.join("manifest/run_manifest.json");
        fs::create_dir_all(observation_path.parent().expect("observation parent"))
            .expect("create copied observation directory");
        fs::create_dir_all(manifest_path.parent().expect("manifest parent"))
            .expect("create copied manifest directory");
        fs::write(observation_path, observation_bytes).expect("write copied observation");
        fs::write(
            manifest_path,
            manifest
                .to_json_pretty(&suite.manifest)
                .expect("serialize copied run"),
        )
        .expect("write copied run");
        run_root
    }

    fn add_second_case_fragment(&self) -> PathBuf {
        let suite_path = self.corpus_root.join("manifest/suite_manifest.json");
        let suite_json = fs::read_to_string(&suite_path).expect("read source suite");
        let mut suite =
            OracleSuiteManifest::from_json_str(&suite_json).expect("parse source suite");
        let first_artifact = suite.cases[0].clone();
        let first_case_bytes = fs::read(self.corpus_root.join(&first_artifact.path))
            .expect("read first case artifact");
        let first_case = suite
            .load_case(&first_artifact.case_id, &first_case_bytes)
            .expect("load first case");
        let first_run_json = fs::read_to_string(self.run_root.join("manifest/run_manifest.json"))
            .expect("read first fragment manifest");
        let first_run =
            RunManifest::from_json_str(&suite, &first_run_json).expect("parse first fragment");

        let second_input_bytes = b"second pinned workbook bytes".to_vec();
        let second_input_relative = "inputs/application-version.xlsx";
        fs::write(
            self.corpus_root.join(second_input_relative),
            &second_input_bytes,
        )
        .expect("write second input");
        let mut second_case = first_case;
        second_case.id = "application.version".to_string();
        second_case.input.path = second_input_relative.to_string();
        second_case.input.sha256 = sha256_hex(&second_input_bytes);
        second_case.probes[0].id = "application-version".to_string();
        second_case.probes[0].member = "Version".to_string();
        let second_case_bytes = second_case
            .to_json_pretty()
            .expect("serialize second case")
            .into_bytes();
        let second_case_relative = "cases/application.version.json";
        fs::write(
            self.corpus_root.join(second_case_relative),
            &second_case_bytes,
        )
        .expect("write second case");
        suite.cases.push(CaseArtifactRef {
            case_id: second_case.id.clone(),
            case_version: second_case.version,
            tier: second_case.tier,
            path: second_case_relative.to_string(),
            sha256: sha256_hex(&second_case_bytes),
            input_sha256: second_case.input.sha256.clone(),
        });
        fs::write(
            &suite_path,
            suite.to_json_pretty().expect("serialize expanded suite"),
        )
        .expect("write expanded suite");

        let mut second_observation: ObservationDocument =
            serde_json::from_slice(&self.observation_bytes).expect("parse first observation");
        second_observation.case_id = second_case.id.clone();
        second_observation.probes[0].id = second_case.probes[0].id.clone();
        second_observation.probes[0].result =
            ObservationResult::Value(ObservedValue::Text("16.0".to_string()));
        let second_observation_bytes =
            serde_json::to_vec_pretty(&second_observation).expect("serialize second observation");
        let mut second_run = first_run;
        second_run.cases[0].case_id = second_case.id;
        second_run.cases[0].case_version = second_case.version;
        second_run.cases[0].tier = second_case.tier;
        second_run.cases[0].case_sha256 = sha256_hex(&second_case_bytes);
        second_run.cases[0].input_sha256 = second_case.input.sha256;
        second_run.cases[0].observation_path = Some("observations/oracle.json".to_string());
        second_run.cases[0].observation_sha256 = Some(sha256_hex(&second_observation_bytes));

        let fragment_root = self.root.join("run-fragment-b");
        let observation_path = fragment_root.join("observations/oracle.json");
        let manifest_path = fragment_root.join("manifest/run_manifest.json");
        fs::create_dir_all(
            observation_path
                .parent()
                .expect("fragment observation parent"),
        )
        .expect("create fragment observation directory");
        fs::create_dir_all(manifest_path.parent().expect("fragment manifest parent"))
            .expect("create fragment manifest directory");
        fs::write(observation_path, second_observation_bytes)
            .expect("write second fragment observation");
        fs::write(
            manifest_path,
            second_run
                .to_fragment_json_pretty(&suite)
                .expect("serialize second fragment manifest"),
        )
        .expect("write second fragment manifest");
        fragment_root
    }
}

impl Drop for CorpusFixture {
    fn drop(&mut self) {
        if self.root.exists() {
            fs::remove_dir_all(&self.root).expect("remove corpus fixture");
        }
    }
}

fn engine(kind: EngineKind) -> EngineIdentity {
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

fn run_oracle(args: &[&str], current_dir: &std::path::Path) -> std::process::Output {
    Command::new(env!("CARGO_BIN_EXE_excel-oracle"))
        .args(args)
        .current_dir(current_dir)
        .output()
        .expect("run excel-oracle")
}

#[test]
fn loads_exact_byte_suite_input_and_run_artifacts_from_disk() {
    let fixture = CorpusFixture::create();

    let suite = PinnedSuiteArtifacts::load(&fixture.corpus_root).expect("load pinned suite");
    assert_eq!(suite.manifest.cases.len(), 1);
    assert_eq!(
        suite.inputs.get("application.name"),
        Some(&fixture.input_bytes),
    );
    assert_eq!(
        suite.cases.keys().collect::<Vec<_>>(),
        vec![&"application.name".to_string()],
    );

    let run = suite.load_run(&fixture.run_root).expect("load pinned run");
    assert_eq!(run.manifest.cases.len(), 1);
    assert_eq!(
        run.observations.get("application.name"),
        Some(&fixture.observation_bytes),
    );
    run.manifest
        .validate_required_completeness(&suite.manifest)
        .expect("required corpus completeness");
}

#[test]
fn rejects_tampered_input_and_observation_bytes_before_replay() {
    let fixture = CorpusFixture::create();
    fs::write(&fixture.input_path, b"tampered input").expect("tamper input");
    let error = PinnedSuiteArtifacts::load(&fixture.corpus_root)
        .expect_err("tampered input must fail")
        .to_string();
    assert!(error.contains("input for case application.name exact-byte sha256"));

    fs::write(&fixture.input_path, &fixture.input_bytes).expect("restore input");
    let suite = PinnedSuiteArtifacts::load(&fixture.corpus_root).expect("load restored suite");
    fs::write(&fixture.observation_path, b"tampered observation").expect("tamper observation");
    let error = suite
        .load_run(&fixture.run_root)
        .expect_err("tampered observation must fail")
        .to_string();
    assert!(error.contains("observation for case application.name exact-byte sha256"));
}

#[test]
fn accepts_only_distinct_matching_repeated_excel_runs() {
    let fixture = CorpusFixture::create();
    let matching_root = fixture.copy_run("run-b", "excel-win-en-us-20260810-b", "Microsoft Excel");
    let suite = PinnedSuiteArtifacts::load(&fixture.corpus_root).expect("load pinned suite");
    let first = suite.load_run(&fixture.run_root).expect("load first run");
    let matching = suite.load_run(&matching_root).expect("load matching run");

    let evidence = suite
        .verify_repeated_excel_runs(&first, &matching)
        .expect("matching independent runs");
    assert_eq!(
        evidence.run_ids,
        [
            "excel-win-en-us-20260810-a".to_string(),
            "excel-win-en-us-20260810-b".to_string(),
        ],
    );
    assert_eq!(
        evidence.verified_case_ids,
        vec!["application.name".to_string()],
    );
    assert_eq!(evidence.engine, engine(EngineKind::Excel));

    let reused_id_root = fixture.copy_run(
        "run-reused",
        "excel-win-en-us-20260810-a",
        "Microsoft Excel",
    );
    let reused_id = suite.load_run(&reused_id_root).expect("load reused-id run");
    let error = suite
        .verify_repeated_excel_runs(&first, &reused_id)
        .expect_err("reused run id must not count as independent evidence")
        .to_string();
    assert!(error.contains("repeated Excel runs must use distinct run ids"));

    let drifted_root =
        fixture.copy_run("run-drifted", "excel-win-en-us-20260810-c", "Drifted Excel");
    let drifted = suite.load_run(&drifted_root).expect("load drifted run");
    let error = suite
        .verify_repeated_excel_runs(&first, &drifted)
        .expect_err("typed observation drift must fail")
        .to_string();
    assert!(error.contains(
        "repeated Excel runs diverged for case application.name at probes.application-name.value"
    ));

    let mut incomplete = matching.clone();
    incomplete.manifest.cases[0].status = RunCaseStatus::Unsupported;
    incomplete.manifest.cases[0].observation_path = None;
    incomplete.manifest.cases[0].observation_sha256 = None;
    incomplete.manifest.cases[0].message = Some("capture was unsupported".to_string());
    incomplete.observations.clear();
    let error = suite
        .verify_repeated_excel_runs(&first, &incomplete)
        .expect_err("incomplete must-match case must fail")
        .to_string();
    assert!(error.contains("mustMatch case application.name did not complete"));

    let mut non_excel = matching;
    non_excel.manifest.engine.kind = EngineKind::Ootd;
    let error = suite
        .verify_repeated_excel_runs(&first, &non_excel)
        .expect_err("non-Excel run must fail")
        .to_string();
    assert!(error.contains("repeated evidence requires two desktop Excel runs"));
}

#[test]
fn assembles_complete_suite_runs_from_case_fragments() {
    let fixture = CorpusFixture::create();
    let second_fragment_root = fixture.add_second_case_fragment();
    let suite = PinnedSuiteArtifacts::load(&fixture.corpus_root).expect("load expanded suite");
    let first = suite
        .load_run_fragment(&fixture.run_root)
        .expect("load first case fragment");
    let second = suite
        .load_run_fragment(&second_fragment_root)
        .expect("load second case fragment");

    let assembled = suite
        .assemble_run_fragments(&[first, second])
        .expect("assemble complete suite run");
    assert_eq!(
        assembled
            .manifest
            .cases
            .iter()
            .map(|record| record.case_id.as_str())
            .collect::<Vec<_>>(),
        vec!["application.name", "application.version"],
    );
    assert_eq!(
        assembled.manifest.cases[0].observation_path.as_deref(),
        Some("observations/application.name/oracle.json"),
    );
    assert_eq!(
        assembled.manifest.cases[1].observation_path.as_deref(),
        Some("observations/application.version/oracle.json"),
    );
    assembled
        .manifest
        .validate_required_completeness(&suite.manifest)
        .expect("assembled required completeness");
    assert_eq!(assembled.observations.len(), 2);
}

#[test]
fn rejects_missing_duplicate_and_cross_run_fragments() {
    let fixture = CorpusFixture::create();
    let second_fragment_root = fixture.add_second_case_fragment();
    let suite = PinnedSuiteArtifacts::load(&fixture.corpus_root).expect("load expanded suite");
    let first = suite
        .load_run_fragment(&fixture.run_root)
        .expect("load first case fragment");
    let second = suite
        .load_run_fragment(&second_fragment_root)
        .expect("load second case fragment");

    let error = suite
        .assemble_run_fragments(std::slice::from_ref(&first))
        .expect_err("missing fragment must fail")
        .to_string();
    assert!(error.contains("assembled run records must exactly cover the suite cases"));

    let error = suite
        .assemble_run_fragments(&[first.clone(), first.clone(), second.clone()])
        .expect_err("duplicate fragment must fail")
        .to_string();
    assert!(error.contains("duplicate fragment record for case application.name"));

    let mut tampered = second.clone();
    tampered
        .observations
        .get_mut("application.version")
        .expect("second observation")
        .push(b' ');
    let error = suite
        .assemble_run_fragments(&[first.clone(), tampered])
        .expect_err("tampered fragment observation must fail")
        .to_string();
    assert!(error.contains("observation for case application.version exact-byte sha256"));

    let mut cross_run = second;
    cross_run.manifest.run_id = "excel-win-en-us-20260810-other".to_string();
    let error = suite
        .assemble_run_fragments(&[first, cross_run])
        .expect_err("cross-run fragment must fail")
        .to_string();
    assert!(error.contains("fragment runId did not match the assembled run"));
}

#[test]
fn atomically_materializes_and_reloads_assembled_run_bundles() {
    let fixture = CorpusFixture::create();
    let second_fragment_root = fixture.add_second_case_fragment();
    let suite = PinnedSuiteArtifacts::load(&fixture.corpus_root).expect("load expanded suite");
    let first = suite
        .load_run_fragment(&fixture.run_root)
        .expect("load first fragment");
    let second = suite
        .load_run_fragment(&second_fragment_root)
        .expect("load second fragment");
    let assembled = suite
        .assemble_run_fragments(&[first, second])
        .expect("assemble run");
    let output_root = fixture.root.join("published-run");

    let receipt = suite
        .write_run_bundle(&assembled, &output_root)
        .expect("materialize assembled run");
    assert_eq!(receipt.root, output_root);
    assert_eq!(
        receipt.manifest_path,
        output_root.join("manifest/run_manifest.json"),
    );
    assert_eq!(receipt.observation_paths.len(), 2);
    assert_eq!(
        suite.load_run(&output_root).expect("reload published run"),
        assembled,
    );
}

#[test]
fn refuses_existing_destinations_and_tampered_bundles_without_partial_output() {
    let fixture = CorpusFixture::create();
    let suite = PinnedSuiteArtifacts::load(&fixture.corpus_root).expect("load suite");
    let assembled = suite
        .load_run(&fixture.run_root)
        .expect("load complete run");

    let existing_root = fixture.root.join("existing-run");
    fs::create_dir(&existing_root).expect("create existing output root");
    let sentinel = existing_root.join("sentinel.txt");
    fs::write(&sentinel, b"keep").expect("write sentinel");
    let error = suite
        .write_run_bundle(&assembled, &existing_root)
        .expect_err("existing output root must fail")
        .to_string();
    assert!(error.contains("run output root must not already exist"));
    assert_eq!(fs::read(&sentinel).expect("read sentinel"), b"keep");

    let mut tampered = assembled;
    tampered
        .observations
        .get_mut("application.name")
        .expect("observation")
        .push(b' ');
    let rejected_root = fixture.root.join("rejected-run");
    let error = suite
        .write_run_bundle(&tampered, &rejected_root)
        .expect_err("tampered bundle must fail before publication")
        .to_string();
    assert!(error.contains("observation for case application.name exact-byte sha256"));
    assert!(!rejected_root.exists());
    assert!(
        fs::read_dir(&fixture.root)
            .expect("list fixture root")
            .all(|entry| !entry
                .expect("fixture entry")
                .file_name()
                .to_string_lossy()
                .starts_with(".rejected-run.ootd-tmp-")),
    );
}

#[test]
fn cli_assembles_and_atomically_materializes_suite_fragments() {
    let fixture = CorpusFixture::create();
    let second_fragment_root = fixture.add_second_case_fragment();
    let output_root = fixture.root.join("cli-published-run");
    let output = run_oracle(
        &[
            "assemble-run",
            "--suite-root",
            fixture.corpus_root.to_str().expect("suite root UTF-8"),
            "--fragment-root",
            fixture.run_root.to_str().expect("first fragment UTF-8"),
            "--fragment-root",
            second_fragment_root
                .to_str()
                .expect("second fragment UTF-8"),
            "--output-root",
            output_root.to_str().expect("output root UTF-8"),
        ],
        &fixture.root,
    );

    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr),
    );
    let receipt: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("CLI receipt JSON");
    assert_eq!(receipt["runId"], "excel-win-en-us-20260810-a");
    assert_eq!(receipt["caseCount"], 2);
    assert_eq!(receipt["completedObservationCount"], 2);
    assert_eq!(
        receipt["outputRoot"],
        output_root.to_str().expect("output root UTF-8"),
    );
    let suite = PinnedSuiteArtifacts::load(&fixture.corpus_root).expect("reload suite");
    let run = suite.load_run(&output_root).expect("reload CLI output");
    assert_eq!(run.manifest.cases.len(), 2);
    assert_eq!(run.observations.len(), 2);
}

#[test]
fn cli_emits_capture_plan_only_for_exact_validated_suite_artifacts() {
    let fixture = CorpusFixture::create();
    let output = run_oracle(
        &[
            "capture-plan",
            "--suite-root",
            fixture.corpus_root.to_str().expect("suite root UTF-8"),
        ],
        &fixture.root,
    );

    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr),
    );
    let plan: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("capture plan JSON");
    assert_eq!(plan["schemaVersion"], 1);
    assert_eq!(plan["suiteId"], "excel-win-en-us-smoke");
    assert_eq!(plan["profileId"], "excel-win-en-us");
    assert_eq!(plan["expectedEngine"]["kind"], "excel");
    assert_eq!(plan["expectedEngine"]["channel"], "Current");
    assert_eq!(plan["caseCount"], 1);
    assert_eq!(plan["cases"][0]["caseId"], "application.name");
    assert_eq!(plan["cases"][0]["casePath"], "cases/application.name.json",);
    assert_eq!(
        plan["cases"][0]["inputPath"],
        "inputs/application-name.xlsx",
    );

    fs::write(&fixture.input_path, b"tampered input").expect("tamper input");
    let rejected = run_oracle(
        &[
            "capture-plan",
            "--suite-root",
            fixture.corpus_root.to_str().expect("suite root UTF-8"),
        ],
        &fixture.root,
    );
    assert_eq!(rejected.status.code(), Some(1));
    assert!(rejected.stdout.is_empty());
    assert!(
        String::from_utf8(rejected.stderr)
            .expect("stderr UTF-8")
            .contains("input for case application.name exact-byte sha256 did not match the suite"),
    );
}

#[test]
fn cli_rejects_incomplete_fragment_coverage_without_output() {
    let fixture = CorpusFixture::create();
    let _second_fragment_root = fixture.add_second_case_fragment();
    let output_root = fixture.root.join("cli-rejected-run");
    let output = run_oracle(
        &[
            "assemble-run",
            "--suite-root",
            fixture.corpus_root.to_str().expect("suite root UTF-8"),
            "--fragment-root",
            fixture.run_root.to_str().expect("first fragment UTF-8"),
            "--output-root",
            output_root.to_str().expect("output root UTF-8"),
        ],
        &fixture.root,
    );

    assert_eq!(output.status.code(), Some(1));
    assert!(
        String::from_utf8(output.stderr)
            .expect("stderr UTF-8")
            .contains("assembled run records must exactly cover the suite cases"),
    );
    assert!(!output_root.exists());
}

#[test]
fn cli_requires_each_assemble_run_argument() {
    let fixture = CorpusFixture::create();
    let output = run_oracle(&["assemble-run", "--suite-root"], &fixture.root);

    assert_eq!(output.status.code(), Some(2));
    assert!(
        String::from_utf8(output.stderr)
            .expect("stderr UTF-8")
            .contains("missing value for --suite-root"),
    );
}

#[cfg(unix)]
#[test]
fn rejects_symlinked_case_artifacts() {
    use std::os::unix::fs::symlink;

    let fixture = CorpusFixture::create();
    let case_path = fixture.corpus_root.join("cases/application.name.json");
    let outside = fixture.root.join("outside-case.json");
    fs::copy(&case_path, &outside).expect("copy outside case");
    fs::remove_file(&case_path).expect("remove original case");
    symlink(&outside, &case_path).expect("create case symlink");

    let error = PinnedSuiteArtifacts::load(&fixture.corpus_root)
        .expect_err("symlinked case must fail")
        .to_string();
    assert!(error.contains("case artifact must not be a symbolic link"));
}
