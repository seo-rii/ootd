use std::{
    collections::BTreeMap,
    fs,
    path::PathBuf,
    time::{SystemTime, UNIX_EPOCH},
};

use office_codegen::{
    CanonicalOmGenerationError, CodegenSummary, DifferentialCaseResult, DifferentialCaseStatus,
    DifferentialReportLoadError, OmCaptureBundleError, OmSourcesManifest, PiaCaptureClass,
    PiaCaptureInterface, PiaPublicSurfaceCapture, SourceRegistryManifest, TypelibIdentityCapture,
    build_coverage_report, build_coverage_report_from_json, build_coverage_report_from_path,
    build_differential_report, build_differential_report_with_source_context,
    build_focus_surface_registry, build_focus_surface_registry_from_json,
    build_focus_surface_registry_from_path, generate_canonical_office_idl_from_dir,
    load_capture_bundle, load_differential_gate_from_json, load_differential_gate_from_path,
    load_differential_gate_from_path_with_source_context, load_differential_report_from_json,
    load_differential_report_from_path, normalize_capture_bundle,
    normalize_capture_bundle_from_dir, normalize_pia_capture_json, summarize_capture_bundle,
    summarize_differential_gate, summarize_differential_gate_with_source_context,
    summarize_om_sources, summarize_om_sources_toml, summarize_source_registry,
    summarize_source_registry_toml, validate_differential_report_source_context,
    write_differential_gate_from_report_path_with_source_context, write_differential_gate_to_path,
    write_differential_report_to_path,
};
use office_idl::{AccessMode, CaptureOriginKind, InterfaceKind, OfficeIdlDocument};
use sha2::{Digest, Sha256};

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()
        .expect("repo root")
}

fn sha256_hex(contents: &[u8]) -> String {
    let digest = Sha256::digest(contents);
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        output.push_str(&format!("{byte:02x}"));
    }
    output
}

#[test]
fn pia_common_members_match_pinned_om_signatures() {
    let pia = PiaPublicSurfaceCapture::from_json_path(
        repo_root().join("specs/pinned/excel_pia_public_surface.template.json"),
    )
    .expect("pia template json")
    .to_office_idl_document();
    let om = OfficeIdlDocument::from_path(
        repo_root().join("specs/pinned/office_idl_excel_om.template.json"),
    )
    .expect("om template json");

    let mut pia_members = BTreeMap::new();
    for interface in &pia.interfaces {
        for member in &interface.members {
            pia_members.insert(
                (
                    interface.name.as_str(),
                    member.name.as_str(),
                    format!("{:?}", member.member_kind),
                ),
                member,
            );
        }
    }

    let type_shape = |type_ref: &office_idl::TypeRef| {
        (
            type_ref
                .alias_of
                .as_deref()
                .unwrap_or(type_ref.name.as_str())
                .to_string(),
            type_ref.nullable,
        )
    };
    let member_signature = |member: &office_idl::Member| {
        let params = member
            .params
            .iter()
            .map(|param| {
                let (type_name, nullable) = type_shape(&param.type_ref);
                (param.name.clone(), param.optional, type_name, nullable)
            })
            .collect::<Vec<_>>();
        let return_type = member.return_type.as_ref().map(type_shape);
        (params, return_type)
    };

    for interface in &om.interfaces {
        for member in &interface.members {
            let key = (
                interface.name.as_str(),
                member.name.as_str(),
                format!("{:?}", member.member_kind),
            );
            let Some(pia_member) = pia_members.get(&key) else {
                continue;
            };
            assert_eq!(
                member_signature(pia_member),
                member_signature(member),
                "{}.{} {:?}",
                interface.name,
                member.name,
                member.member_kind
            );
        }
    }
}

#[test]
fn loads_pinned_om_sources_manifest_and_reports_pending_capture() {
    let manifest =
        OmSourcesManifest::from_toml_path(repo_root().join("specs/pinned/om_sources.toml"))
            .expect("manifest");
    let summary = summarize_om_sources(&manifest);

    assert_eq!(manifest.manifest.status, "pending_capture");
    assert_eq!(summary.primary_artifact, "excel_type_library");
    assert_eq!(summary.secondary_artifact, "excel_primary_interop_assembly");
    assert!(summary.ready_for_windows_capture);
    assert_eq!(summary.machine_readable_artifact_count, 2);
    assert_eq!(summary.pending_output_count, 5);
    assert_eq!(
        summary.pending_outputs,
        vec![
            "raw_typelib_identity.json".to_string(),
            "excel_typelib_snapshot.idl".to_string(),
            "excel_typelib_snapshot.odl".to_string(),
            "excel_pia_identity.json".to_string(),
            "excel_pia_public_surface.json".to_string(),
        ]
    );
    assert_eq!(summary.behavior_doc_count, 3);
    assert_eq!(
        summary.unresolved_target_fields,
        vec![
            "product_family",
            "channel",
            "version",
            "build",
            "arch",
            "locale"
        ]
    );
}

#[test]
fn loads_source_registry_and_reports_enabled_test_corpus() {
    let registry_toml =
        fs::read_to_string(repo_root().join("specs/sources.toml")).expect("source registry");
    let manifest = SourceRegistryManifest::from_toml_str(&registry_toml).expect("registry");
    let summary = summarize_source_registry(&manifest);
    let summary_from_toml = summarize_source_registry_toml(&registry_toml).expect("summary");

    assert_eq!(summary, summary_from_toml);
    assert_eq!(summary.project_name, "excel-compat-core");
    assert_eq!(summary.default_profile, "excel_365");
    assert_eq!(summary.default_mode, "lossless");
    assert_eq!(summary.primary_om_artifact, "excel_type_library");
    assert_eq!(summary.secondary_om_artifact, "excel_pia");
    assert_eq!(summary.primary_docs_source, "excel_vba_reference");
    assert_eq!(summary.primary_ooxml_source, "ecma_376");
    assert_eq!(
        summary.enabled_corpus_groups,
        vec![
            "official_ms".to_string(),
            "open_source".to_string(),
            "synthetic".to_string(),
            "real_world".to_string(),
        ]
    );
    assert_eq!(
        summary.official_ms_corpus_sources,
        vec![
            "office_scripts_samples".to_string(),
            "data_validation_examples".to_string(),
            "power_bi_financial_sample".to_string(),
            "mos_excel_course_materials".to_string(),
            "mos_excel_expert_course_materials".to_string(),
        ]
    );
    assert_eq!(
        summary.open_source_corpus_sources,
        vec![
            "open_xml_sdk".to_string(),
            "apache_poi_test_data".to_string(),
            "libreoffice_sc_qa_unit_data".to_string(),
        ]
    );
    assert_eq!(summary.enabled_corpus_source_count, 10);
    assert_eq!(
        summary.validation_modes,
        vec![
            "openxml_validator".to_string(),
            "excel_oracle".to_string(),
            "render_snapshot".to_string(),
            "fuzz".to_string(),
        ]
    );
    assert_eq!(summary.profile_count, 3);
}

#[test]
fn builds_differential_report_with_source_registry_context() {
    let registry_toml =
        fs::read_to_string(repo_root().join("specs/sources.toml")).expect("source registry");
    let source_summary = summarize_source_registry_toml(&registry_toml).expect("source summary");
    let report = build_differential_report_with_source_context(
        "Excel",
        "16.0",
        &source_summary,
        vec![DifferentialCaseResult {
            name: "Range.Areas.Count".to_string(),
            surface: Some("Range".to_string()),
            member: Some("Areas".to_string()),
            status: DifferentialCaseStatus::Passed,
            expected: Some("2".to_string()),
            actual: Some("2".to_string()),
            message: None,
            artifacts: BTreeMap::new(),
        }],
    );

    let context = report.context.as_ref().expect("report context");
    assert_eq!(report.profile, "excel_365");
    assert_eq!(context.project_name, "excel-compat-core");
    assert_eq!(context.default_profile, "excel_365");
    assert_eq!(context.default_mode, "lossless");
    assert_eq!(context.primary_om_artifact, "excel_type_library");
    assert_eq!(context.primary_ooxml_source, "ecma_376");
    assert_eq!(
        context.enabled_corpus_groups,
        vec![
            "official_ms".to_string(),
            "open_source".to_string(),
            "synthetic".to_string(),
            "real_world".to_string(),
        ]
    );
    assert_eq!(context.enabled_corpus_source_count, 10);
    assert_eq!(
        context.validation_modes,
        vec![
            "openxml_validator".to_string(),
            "excel_oracle".to_string(),
            "render_snapshot".to_string(),
            "fuzz".to_string(),
        ]
    );

    let json = serde_json::to_string(&report).expect("report json");
    assert!(json.contains(r#""context""#));
    assert!(json.contains(r#""projectName":"excel-compat-core""#));
    let loaded = load_differential_report_from_json(&json).expect("load report");
    assert_eq!(loaded, report);
}

#[test]
fn validates_differential_report_source_registry_context() {
    let registry_toml =
        fs::read_to_string(repo_root().join("specs/sources.toml")).expect("source registry");
    let source_summary = summarize_source_registry_toml(&registry_toml).expect("source summary");
    let report = build_differential_report_with_source_context(
        "Excel",
        "16.0",
        &source_summary,
        vec![DifferentialCaseResult {
            name: "Application.Version".to_string(),
            surface: Some("Application".to_string()),
            member: Some("Version".to_string()),
            status: DifferentialCaseStatus::Passed,
            expected: Some("16.0".to_string()),
            actual: Some("16.0".to_string()),
            message: None,
            artifacts: BTreeMap::new(),
        }],
    );

    validate_differential_report_source_context(&report, &source_summary)
        .expect("matching source context");

    let report_without_context =
        build_differential_report("Excel", "16.0", "excel_365", report.cases.clone());
    let error =
        validate_differential_report_source_context(&report_without_context, &source_summary)
            .expect_err("missing context should fail");
    match error {
        DifferentialReportLoadError::Contract { message } => {
            assert!(message.contains("missing source registry context"));
        }
        other => panic!("unexpected error: {other:?}"),
    }

    let mut mismatched_summary = source_summary.clone();
    mismatched_summary.default_profile = "excel_2021".to_string();
    let error = validate_differential_report_source_context(&report, &mismatched_summary)
        .expect_err("mismatched context should fail");
    match error {
        DifferentialReportLoadError::Contract { message } => {
            assert!(message.contains("source context"));
            assert!(message.contains("excel_2021"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn summarizes_differential_gate_with_source_registry_context() {
    let registry_toml =
        fs::read_to_string(repo_root().join("specs/sources.toml")).expect("source registry");
    let source_summary = summarize_source_registry_toml(&registry_toml).expect("source summary");
    let report = build_differential_report_with_source_context(
        "Excel",
        "16.0",
        &source_summary,
        vec![
            DifferentialCaseResult {
                name: "Application.Name".to_string(),
                surface: Some("Application".to_string()),
                member: Some("Name".to_string()),
                status: DifferentialCaseStatus::Passed,
                expected: Some("Microsoft Excel".to_string()),
                actual: Some("Microsoft Excel".to_string()),
                message: None,
                artifacts: BTreeMap::new(),
            },
            DifferentialCaseResult {
                name: "Chart.SetSourceData".to_string(),
                surface: Some("Chart".to_string()),
                member: Some("SetSourceData".to_string()),
                status: DifferentialCaseStatus::MissingRuntime,
                expected: Some("series attached".to_string()),
                actual: None,
                message: Some("runtime chart setter is not implemented".to_string()),
                artifacts: BTreeMap::new(),
            },
        ],
    );

    let gate = summarize_differential_gate_with_source_context(&report, &source_summary)
        .expect("matching source context should gate");

    assert!(!gate.passed);
    assert_eq!(gate.blocking_case_count, 1);
    assert_eq!(gate.blocking_cases, vec!["Chart.SetSourceData".to_string()]);
    assert_eq!(gate.missing_runtime_count, 1);

    let report_without_context =
        build_differential_report("Excel", "16.0", "excel_365", report.cases.clone());
    let error =
        summarize_differential_gate_with_source_context(&report_without_context, &source_summary)
            .expect_err("missing context should fail before gate");
    match error {
        DifferentialReportLoadError::Contract { message } => {
            assert!(message.contains("missing source registry context"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn loads_differential_gate_from_path_with_source_registry_context() {
    let registry_toml =
        fs::read_to_string(repo_root().join("specs/sources.toml")).expect("source registry");
    let source_summary = summarize_source_registry_toml(&registry_toml).expect("source summary");
    let report = build_differential_report_with_source_context(
        "Excel",
        "16.0",
        &source_summary,
        vec![DifferentialCaseResult {
            name: "Application.Version".to_string(),
            surface: Some("Application".to_string()),
            member: Some("Version".to_string()),
            status: DifferentialCaseStatus::Passed,
            expected: Some("16.0".to_string()),
            actual: Some("16.0".to_string()),
            message: None,
            artifacts: BTreeMap::new(),
        }],
    );
    let unique_suffix = format!(
        "{}-{}",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("time")
            .as_nanos()
    );
    let path = std::env::temp_dir().join(format!(
        "ootd-differential-gate-context-{unique_suffix}.json"
    ));
    write_differential_report_to_path(&report, &path).expect("write report");

    let gate = load_differential_gate_from_path_with_source_context(&path, &source_summary)
        .expect("load context-aware gate");

    assert!(gate.passed);
    assert_eq!(gate.blocking_case_count, 0);

    let mut missing_context = report.clone();
    missing_context.context = None;
    fs::write(
        &path,
        serde_json::to_vec_pretty(&missing_context).expect("missing context json"),
    )
    .expect("rewrite missing context report");
    let error = load_differential_gate_from_path_with_source_context(&path, &source_summary)
        .expect_err("missing context should fail");
    match error {
        DifferentialReportLoadError::Contract { message } => {
            assert!(message.contains("missing source registry context"));
        }
        other => panic!("unexpected error: {other:?}"),
    }

    let mut stale_report = report;
    stale_report.case_count += 1;
    fs::write(
        &path,
        serde_json::to_vec_pretty(&stale_report).expect("stale report json"),
    )
    .expect("rewrite stale report");
    let error = load_differential_gate_from_path_with_source_context(&path, &source_summary)
        .expect_err("stale report should fail");
    match error {
        DifferentialReportLoadError::Contract { message } => {
            assert!(message.contains("caseCount"));
        }
        other => panic!("unexpected error: {other:?}"),
    }

    let _ = fs::remove_file(path);
}

#[test]
fn writes_differential_gate_from_report_path_with_source_registry_context() {
    let registry_toml =
        fs::read_to_string(repo_root().join("specs/sources.toml")).expect("source registry");
    let source_summary = summarize_source_registry_toml(&registry_toml).expect("source summary");
    let report = build_differential_report_with_source_context(
        "Excel",
        "16.0",
        &source_summary,
        vec![
            DifferentialCaseResult {
                name: "Application.Name".to_string(),
                surface: Some("Application".to_string()),
                member: Some("Name".to_string()),
                status: DifferentialCaseStatus::Passed,
                expected: Some("Microsoft Excel".to_string()),
                actual: Some("Microsoft Excel".to_string()),
                message: None,
                artifacts: BTreeMap::new(),
            },
            DifferentialCaseResult {
                name: "Range.Value2 multi-area".to_string(),
                surface: Some("Range".to_string()),
                member: Some("Value2".to_string()),
                status: DifferentialCaseStatus::Failed,
                expected: Some("[[1],[2]]".to_string()),
                actual: Some("1".to_string()),
                message: Some("runtime collapsed the reference to the first scalar".to_string()),
                artifacts: BTreeMap::new(),
            },
        ],
    );
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let report_path = std::env::temp_dir().join(format!(
        "ootd-differential-report-gate-{unique_suffix}.json"
    ));
    let gate_path =
        std::env::temp_dir().join(format!("ootd-differential-gate-{unique_suffix}.json"));
    write_differential_report_to_path(&report, &report_path).expect("write report");

    let gate = write_differential_gate_from_report_path_with_source_context(
        &report_path,
        &source_summary,
        &gate_path,
    )
    .expect("write context-aware gate");

    assert!(!gate.passed);
    assert_eq!(
        gate.blocking_cases,
        vec!["Range.Value2 multi-area".to_string()]
    );
    let gate_json = fs::read_to_string(&gate_path).expect("gate json");
    assert!(gate_json.contains(r#""passed": false"#));
    assert!(gate_json.contains(r#""blockingCaseCount": 1"#));
    assert!(gate_json.contains(r#""failedCaseCount": 1"#));

    let copied_gate_path =
        std::env::temp_dir().join(format!("ootd-differential-gate-copy-{unique_suffix}.json"));
    write_differential_gate_to_path(&gate, &copied_gate_path).expect("write copied gate");
    let copied_gate_json = fs::read_to_string(&copied_gate_path).expect("copied gate json");
    assert_eq!(copied_gate_json, gate_json);

    fs::remove_file(&report_path).expect("remove report");
    fs::remove_file(&gate_path).expect("remove gate");
    fs::remove_file(&copied_gate_path).expect("remove copied gate");
}

#[test]
fn loads_differential_gate_summary_and_rejects_stale_contract() {
    let report = build_differential_report(
        "Excel",
        "16.0",
        "excel_365",
        vec![DifferentialCaseResult {
            name: "Range.Value2 multi-area".to_string(),
            surface: Some("Range".to_string()),
            member: Some("Value2".to_string()),
            status: DifferentialCaseStatus::Failed,
            expected: Some("[[1],[2]]".to_string()),
            actual: Some("1".to_string()),
            message: Some("runtime collapsed the reference to the first scalar".to_string()),
            artifacts: BTreeMap::new(),
        }],
    );
    let gate = summarize_differential_gate(&report);
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let path =
        std::env::temp_dir().join(format!("ootd-differential-gate-load-{unique_suffix}.json"));
    write_differential_gate_to_path(&gate, &path).expect("write gate");

    let gate_json = fs::read_to_string(&path).expect("gate json");
    let loaded_from_json = load_differential_gate_from_json(&gate_json).expect("load gate json");
    let loaded_from_path = load_differential_gate_from_path(&path).expect("load gate path");

    assert_eq!(loaded_from_json, gate);
    assert_eq!(loaded_from_path, gate);

    let stale_count = gate_json.replace(r#""blockingCaseCount": 1"#, r#""blockingCaseCount": 2"#);
    let error = load_differential_gate_from_json(&stale_count)
        .expect_err("stale blocking count should fail");
    match error {
        DifferentialReportLoadError::Contract { message } => {
            assert!(message.contains("blockingCaseCount 2"));
            assert!(message.contains("blockingCases length 1"));
        }
        other => panic!("unexpected error: {other:?}"),
    }

    let stale_passed = gate_json.replace(r#""passed": false"#, r#""passed": true"#);
    let error =
        load_differential_gate_from_json(&stale_passed).expect_err("stale passed should fail");
    match error {
        DifferentialReportLoadError::Contract { message } => {
            assert!(message.contains("passed true"));
            assert!(message.contains("blockingCaseCount 1"));
        }
        other => panic!("unexpected error: {other:?}"),
    }

    let mut stale_gate = gate;
    stale_gate.blocking_case_count += 1;
    let write_error = write_differential_gate_to_path(&stale_gate, &path)
        .expect_err("stale gate should not write");
    match write_error {
        DifferentialReportLoadError::Contract { message } => {
            assert!(message.contains("blockingCaseCount 2"));
        }
        other => panic!("unexpected error: {other:?}"),
    }

    fs::remove_file(path).expect("remove gate");
}

#[test]
fn builds_differential_report_with_stable_status_counts() {
    let cases = vec![
        DifferentialCaseResult {
            name: "Application.Version".to_string(),
            surface: Some("Application".to_string()),
            member: Some("Version".to_string()),
            status: DifferentialCaseStatus::Passed,
            expected: Some("16.0".to_string()),
            actual: Some("16.0".to_string()),
            message: None,
            artifacts: BTreeMap::new(),
        },
        DifferentialCaseResult {
            name: "Range.Value2 multi-area".to_string(),
            surface: Some("Range".to_string()),
            member: Some("Value2".to_string()),
            status: DifferentialCaseStatus::Failed,
            expected: Some("[[1],[2]]".to_string()),
            actual: Some("1".to_string()),
            message: Some("runtime collapsed the reference to the first scalar".to_string()),
            artifacts: BTreeMap::from([(
                "runtimeTrace".to_string(),
                "reports/range_value2_runtime.json".to_string(),
            )]),
        },
        DifferentialCaseResult {
            name: "Chart.SetSourceData".to_string(),
            surface: Some("Chart".to_string()),
            member: Some("SetSourceData".to_string()),
            status: DifferentialCaseStatus::MissingOracle,
            expected: None,
            actual: Some("completed".to_string()),
            message: Some("Excel oracle result was not captured".to_string()),
            artifacts: BTreeMap::new(),
        },
        DifferentialCaseResult {
            name: "PivotChart refresh".to_string(),
            surface: Some("Chart".to_string()),
            member: Some("Refresh".to_string()),
            status: DifferentialCaseStatus::Unsupported,
            expected: None,
            actual: None,
            message: Some("pivot charts are preserve-only".to_string()),
            artifacts: BTreeMap::new(),
        },
        DifferentialCaseResult {
            name: "External link update".to_string(),
            surface: Some("Workbook".to_string()),
            member: Some("UpdateLink".to_string()),
            status: DifferentialCaseStatus::Skipped,
            expected: None,
            actual: None,
            message: Some("external workbook fixture is unavailable".to_string()),
            artifacts: BTreeMap::new(),
        },
        DifferentialCaseResult {
            name: "WorksheetFunction.XLookup".to_string(),
            surface: Some("WorksheetFunction".to_string()),
            member: Some("XLookup".to_string()),
            status: DifferentialCaseStatus::MissingRuntime,
            expected: Some("matched".to_string()),
            actual: None,
            message: Some("runtime function is not implemented".to_string()),
            artifacts: BTreeMap::new(),
        },
    ];

    let report = build_differential_report("Excel", "16.0", "excel_365", cases);

    assert_eq!(report.library, "Excel");
    assert_eq!(report.version, "16.0");
    assert_eq!(report.profile, "excel_365");
    assert_eq!(report.case_count, 6);
    assert_eq!(report.status_counts.passed, 1);
    assert_eq!(report.status_counts.failed, 1);
    assert_eq!(report.status_counts.missing_oracle, 1);
    assert_eq!(report.status_counts.missing_runtime, 1);
    assert_eq!(report.status_counts.unsupported, 1);
    assert_eq!(report.status_counts.skipped, 1);
    assert_eq!(report.cases[0].name, "Application.Version");

    let json = serde_json::to_string(&report).expect("report json");
    assert!(json.contains(r#""caseCount":6"#));
    assert!(json.contains(r#""missingOracle":1"#));
    assert!(json.contains(r#""missingRuntime":1"#));
    assert!(json.contains(r#""runtimeTrace":"reports/range_value2_runtime.json""#));
}

#[test]
fn summarizes_differential_report_gate_failures() {
    let report = build_differential_report(
        "Excel",
        "16.0",
        "excel_365",
        vec![
            DifferentialCaseResult {
                name: "Application.Version".to_string(),
                surface: Some("Application".to_string()),
                member: Some("Version".to_string()),
                status: DifferentialCaseStatus::Passed,
                expected: Some("16.0".to_string()),
                actual: Some("16.0".to_string()),
                message: None,
                artifacts: BTreeMap::new(),
            },
            DifferentialCaseResult {
                name: "Range.Value2 multi-area".to_string(),
                surface: Some("Range".to_string()),
                member: Some("Value2".to_string()),
                status: DifferentialCaseStatus::Failed,
                expected: Some("[[1],[2]]".to_string()),
                actual: Some("1".to_string()),
                message: Some("runtime collapsed the reference to the first scalar".to_string()),
                artifacts: BTreeMap::new(),
            },
            DifferentialCaseResult {
                name: "Chart.SetSourceData".to_string(),
                surface: Some("Chart".to_string()),
                member: Some("SetSourceData".to_string()),
                status: DifferentialCaseStatus::MissingOracle,
                expected: None,
                actual: Some("completed".to_string()),
                message: Some("Excel oracle result was not captured".to_string()),
                artifacts: BTreeMap::new(),
            },
            DifferentialCaseResult {
                name: "WorksheetFunction.XLookup".to_string(),
                surface: Some("WorksheetFunction".to_string()),
                member: Some("XLookup".to_string()),
                status: DifferentialCaseStatus::MissingRuntime,
                expected: Some("matched".to_string()),
                actual: None,
                message: Some("runtime function is not implemented".to_string()),
                artifacts: BTreeMap::new(),
            },
            DifferentialCaseResult {
                name: "PivotChart refresh".to_string(),
                surface: Some("Chart".to_string()),
                member: Some("Refresh".to_string()),
                status: DifferentialCaseStatus::Unsupported,
                expected: None,
                actual: None,
                message: Some("pivot charts are preserve-only".to_string()),
                artifacts: BTreeMap::new(),
            },
            DifferentialCaseResult {
                name: "External link update".to_string(),
                surface: Some("Workbook".to_string()),
                member: Some("UpdateLink".to_string()),
                status: DifferentialCaseStatus::Skipped,
                expected: None,
                actual: None,
                message: Some("external workbook fixture is unavailable".to_string()),
                artifacts: BTreeMap::new(),
            },
        ],
    );

    let gate = summarize_differential_gate(&report);

    assert!(!gate.passed);
    assert_eq!(gate.blocking_case_count, 3);
    assert_eq!(
        gate.blocking_cases,
        vec![
            "Range.Value2 multi-area".to_string(),
            "Chart.SetSourceData".to_string(),
            "WorksheetFunction.XLookup".to_string(),
        ]
    );
    assert_eq!(gate.failed_case_count, 1);
    assert_eq!(gate.incomplete_oracle_count, 1);
    assert_eq!(gate.missing_runtime_count, 1);
    assert_eq!(gate.unsupported_case_count, 1);
    assert_eq!(gate.skipped_case_count, 1);

    let json = serde_json::to_string(&gate).expect("gate json");
    assert!(json.contains(r#""passed":false"#));
    assert!(json.contains(r#""blockingCaseCount":3"#));
}

#[test]
fn differential_report_gate_passes_with_only_passed_skipped_and_unsupported_cases() {
    let report = build_differential_report(
        "Excel",
        "16.0",
        "excel_365",
        vec![
            DifferentialCaseResult {
                name: "Application.Name".to_string(),
                surface: Some("Application".to_string()),
                member: Some("Name".to_string()),
                status: DifferentialCaseStatus::Passed,
                expected: Some("Microsoft Excel".to_string()),
                actual: Some("Microsoft Excel".to_string()),
                message: None,
                artifacts: BTreeMap::new(),
            },
            DifferentialCaseResult {
                name: "PivotChart refresh".to_string(),
                surface: Some("Chart".to_string()),
                member: Some("Refresh".to_string()),
                status: DifferentialCaseStatus::Unsupported,
                expected: None,
                actual: None,
                message: Some("pivot charts are preserve-only".to_string()),
                artifacts: BTreeMap::new(),
            },
            DifferentialCaseResult {
                name: "External link update".to_string(),
                surface: Some("Workbook".to_string()),
                member: Some("UpdateLink".to_string()),
                status: DifferentialCaseStatus::Skipped,
                expected: None,
                actual: None,
                message: Some("external workbook fixture is unavailable".to_string()),
                artifacts: BTreeMap::new(),
            },
        ],
    );

    let gate = summarize_differential_gate(&report);

    assert!(gate.passed);
    assert_eq!(gate.blocking_case_count, 0);
    assert!(gate.blocking_cases.is_empty());
    assert_eq!(gate.unsupported_case_count, 1);
    assert_eq!(gate.skipped_case_count, 1);
}

#[test]
fn loads_differential_report_from_json_and_rejects_stale_counts() {
    let report = build_differential_report(
        "Excel",
        "16.0",
        "excel_365",
        vec![DifferentialCaseResult {
            name: "Application.Name".to_string(),
            surface: Some("Application".to_string()),
            member: Some("Name".to_string()),
            status: DifferentialCaseStatus::Passed,
            expected: Some("Microsoft Excel".to_string()),
            actual: Some("Microsoft Excel".to_string()),
            message: None,
            artifacts: BTreeMap::new(),
        }],
    );
    let json = serde_json::to_string(&report).expect("report json");

    let loaded = load_differential_report_from_json(&json).expect("load report");

    assert_eq!(loaded, report);

    let stale_case_count = json.replace(r#""caseCount":1"#, r#""caseCount":2"#);
    let error = load_differential_report_from_json(&stale_case_count)
        .expect_err("stale case count should fail");
    match error {
        DifferentialReportLoadError::Contract { message } => {
            assert!(message.contains("caseCount 2 did not match cases length 1"));
        }
        other => panic!("unexpected error: {other:?}"),
    }

    let stale_status_count = json.replace(r#""passed":1"#, r#""passed":0"#);
    let error = load_differential_report_from_json(&stale_status_count)
        .expect_err("stale status counts should fail");
    match error {
        DifferentialReportLoadError::Contract { message } => {
            assert!(message.contains("statusCounts"));
            assert!(message.contains("passed: 0"));
            assert!(message.contains("passed: 1"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn loads_differential_report_from_path_wrapper() {
    let report = build_differential_report(
        "Excel",
        "16.0",
        "excel_365",
        vec![DifferentialCaseResult {
            name: "Workbook.Worksheets.Count".to_string(),
            surface: Some("Workbook".to_string()),
            member: Some("Worksheets".to_string()),
            status: DifferentialCaseStatus::Failed,
            expected: Some("3".to_string()),
            actual: Some("2".to_string()),
            message: Some("chart sheets must not be counted as worksheets".to_string()),
            artifacts: BTreeMap::from([(
                "oracle".to_string(),
                "reports/workbook_worksheets_count_oracle.json".to_string(),
            )]),
        }],
    );
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let path = std::env::temp_dir().join(format!("ootd-differential-report-{unique_suffix}.json"));
    fs::write(&path, serde_json::to_string(&report).expect("report json")).expect("write report");

    let loaded = load_differential_report_from_path(&path).expect("load report path");

    assert_eq!(loaded, report);

    fs::remove_file(&path).expect("remove temp report");
}

#[test]
fn writes_differential_report_to_path_and_reloads_it() {
    let report = build_differential_report(
        "Excel",
        "16.0",
        "excel_365",
        vec![DifferentialCaseResult {
            name: "Range.Areas.Count".to_string(),
            surface: Some("Range".to_string()),
            member: Some("Areas".to_string()),
            status: DifferentialCaseStatus::Passed,
            expected: Some("2".to_string()),
            actual: Some("2".to_string()),
            message: None,
            artifacts: BTreeMap::from([(
                "fixture".to_string(),
                "corpus/range_multi_area.xlsx".to_string(),
            )]),
        }],
    );
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "ootd-differential-report-write-{unique_suffix}.json"
    ));

    write_differential_report_to_path(&report, &path).expect("write report");
    let loaded = load_differential_report_from_path(&path).expect("reload report");
    let json = fs::read_to_string(&path).expect("read written report");

    assert_eq!(loaded, report);
    assert!(json.contains('\n'));
    assert!(json.contains(r#""caseCount": 1"#));

    fs::remove_file(&path).expect("remove temp report");
}

#[test]
fn write_differential_report_rejects_stale_summary_counts() {
    let mut report = build_differential_report(
        "Excel",
        "16.0",
        "excel_365",
        vec![DifferentialCaseResult {
            name: "Application.Name".to_string(),
            surface: Some("Application".to_string()),
            member: Some("Name".to_string()),
            status: DifferentialCaseStatus::Passed,
            expected: Some("Microsoft Excel".to_string()),
            actual: Some("Microsoft Excel".to_string()),
            message: None,
            artifacts: BTreeMap::new(),
        }],
    );
    report.case_count = 2;
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "ootd-differential-report-stale-{unique_suffix}.json"
    ));

    let error =
        write_differential_report_to_path(&report, &path).expect_err("stale report should fail");

    match error {
        DifferentialReportLoadError::Contract { message } => {
            assert!(message.contains("caseCount 2 did not match cases length 1"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
    assert!(!path.exists());
}

#[test]
fn loads_office_idl_excel_om_template_and_summarizes_surface() {
    let document = OfficeIdlDocument::from_path(
        repo_root().join("specs/pinned/office_idl_excel_om.template.json"),
    )
    .expect("template json");
    let summary = CodegenSummary::from_document(&document);

    assert_eq!(document.library, "Excel");
    assert_eq!(document.version, "16.0");
    assert_eq!(
        document
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.namespace.as_deref()),
        Some("Microsoft.Office.Interop.Excel")
    );
    assert_eq!(
        document
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.namespace.as_deref()),
        Some("Microsoft.Office.Interop.Excel")
    );
    assert_eq!(summary.enum_count, 8);
    assert_eq!(summary.interface_count, 58);
    assert_eq!(summary.class_count, 3);
    assert_eq!(summary.member_count, 1470);
    assert_eq!(summary.stub_member_count, 1470);
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
        document.interfaces[0].members[0]
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Workbooks")
    );
    assert_eq!(
        document
            .interfaces
            .iter()
            .find(|interface| interface.name == "Worksheet")
            .expect("Worksheet")
            .metadata
            .as_ref()
            .map(|metadata| metadata.source_inherits.clone()),
        Some(vec![
            "IDispatch".to_string(),
            "Excel._Worksheet".to_string()
        ])
    );
    let workbooks = document
        .interfaces
        .iter()
        .find(|interface| interface.name == "Workbooks")
        .expect("Workbooks");
    let workbooks_open = workbooks
        .members
        .iter()
        .find(|member| member.name == "Open")
        .expect("Workbooks.Open");
    let expected_names = [
        "Filename",
        "UpdateLinks",
        "ReadOnly",
        "Format",
        "Password",
        "WriteResPassword",
        "IgnoreReadOnlyRecommended",
        "Origin",
        "Delimiter",
        "Editable",
        "Notify",
        "Converter",
        "AddToMru",
        "Local",
        "CorruptLoad",
    ];
    assert_eq!(workbooks_open.params.len(), expected_names.len());
    for (index, expected_name) in expected_names.iter().enumerate() {
        assert_eq!(workbooks_open.params[index].name, *expected_name);
        assert_eq!(workbooks_open.params[index].optional, index != 0);
    }
    let workbooks_item_member = workbooks
        .members
        .iter()
        .find(|member| member.name == "Item")
        .expect("Workbooks.Item");
    assert_eq!(workbooks_item_member.params.len(), 1);
    assert_eq!(workbooks_item_member.params[0].name, "Index");
    assert_eq!(
        workbooks_item_member
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Workbook")
    );
    let application = document
        .interfaces
        .iter()
        .find(|interface| interface.name == "Application")
        .expect("Application");
    let worksheet = document
        .interfaces
        .iter()
        .find(|interface| interface.name == "Worksheet")
        .expect("Worksheet");
    let worksheets = document
        .interfaces
        .iter()
        .find(|interface| interface.name == "Worksheets")
        .expect("Worksheets");
    assert_eq!(
        application
            .members
            .iter()
            .find(|member| member.name == "ActiveChart")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Chart")
    );
    let worksheet_chart_objects = worksheet
        .members
        .iter()
        .find(|member| member.name == "ChartObjects")
        .expect("Worksheet.ChartObjects");
    assert_eq!(worksheet_chart_objects.access, AccessMode::Read);
    assert_eq!(worksheet_chart_objects.params.len(), 1);
    assert!(worksheet_chart_objects.params[0].optional);
    assert_eq!(
        worksheet_chart_objects
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.ChartObjects")
    );
    let worksheets_creator_member = worksheets
        .members
        .iter()
        .find(|member| member.name == "Creator")
        .expect("Worksheets.Creator");
    assert_eq!(
        worksheets_creator_member
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Int32")
    );
    let worksheets_visible_member = worksheets
        .members
        .iter()
        .find(|member| member.name == "Visible")
        .expect("Worksheets.Visible");
    assert_eq!(worksheets_visible_member.access, AccessMode::Readwrite);
    assert_eq!(
        worksheets_visible_member
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("XlSheetVisibility")
    );
    let worksheets_delete_member = worksheets
        .members
        .iter()
        .find(|member| member.name == "Delete")
        .expect("Worksheets.Delete");
    assert_eq!(
        worksheets_delete_member
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT_BOOL")
    );
    for member_name in ["Copy", "Move"] {
        let member = worksheets
            .members
            .iter()
            .find(|member| member.name == member_name)
            .expect(member_name);
        assert_eq!(member.params.len(), 2, "{member_name}");
        assert!(member.params.iter().all(|param| param.optional));
    }
    let worksheets_print_out_member = worksheets
        .members
        .iter()
        .find(|member| member.name == "PrintOut")
        .expect("Worksheets.PrintOut");
    assert_eq!(worksheets_print_out_member.params.len(), 8);
    assert!(
        worksheets_print_out_member
            .params
            .iter()
            .all(|param| param.optional)
    );
    let worksheets_select_member = worksheets
        .members
        .iter()
        .find(|member| member.name == "Select")
        .expect("Worksheets.Select");
    assert_eq!(worksheets_select_member.params.len(), 1);
    assert!(worksheets_select_member.params[0].optional);
}

#[test]
fn normalizes_pia_capture_template_into_office_idl_surface() {
    let capture = PiaPublicSurfaceCapture::from_json_path(
        repo_root().join("specs/pinned/excel_pia_public_surface.template.json"),
    )
    .expect("capture json");
    let document = capture.to_office_idl_document();
    let summary = CodegenSummary::from_document(&document);

    assert_eq!(document.library, "Excel");
    assert_eq!(document.version, "16.0");
    assert_eq!(summary.enum_count, 1);
    assert_eq!(summary.interface_count, 6);
    assert_eq!(summary.class_count, 3);
    assert_eq!(summary.member_count, 64);

    let application = document
        .interfaces
        .iter()
        .find(|interface| interface.name == "Application")
        .expect("Application");
    let worksheet = document
        .interfaces
        .iter()
        .find(|interface| interface.name == "Worksheet")
        .expect("Worksheet");
    let workbook = document
        .interfaces
        .iter()
        .find(|interface| interface.name == "Workbook")
        .expect("Workbook");
    let range = document
        .interfaces
        .iter()
        .find(|interface| interface.name == "Range")
        .expect("Range");
    let workbooks = document
        .interfaces
        .iter()
        .find(|interface| interface.name == "Workbooks")
        .expect("Workbooks");
    let worksheets = document
        .interfaces
        .iter()
        .find(|interface| interface.name == "Worksheets")
        .expect("Worksheets");
    let name_member = worksheet
        .members
        .iter()
        .find(|member| member.name == "Name")
        .expect("Name");
    assert_eq!(name_member.access, AccessMode::Readwrite);
    assert_eq!(name_member.params.len(), 0);
    assert_eq!(
        name_member
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.capture.as_ref())
            .map(|capture| capture.origins.len()),
        Some(2)
    );
    assert_eq!(
        name_member
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.capture.as_ref())
            .and_then(|capture| capture.type_info.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("BSTR")
    );
    assert_eq!(
        workbooks
            .members
            .iter()
            .find(|member| member.name == "Parent")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Application")
    );
    let add_member = workbooks
        .members
        .iter()
        .find(|member| member.name == "Add")
        .expect("Add");
    assert_eq!(add_member.params.len(), 1);
    assert_eq!(add_member.params[0].name, "Template");
    assert!(add_member.params[0].optional);
    assert_eq!(
        add_member
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Workbook")
    );
    let worksheets_add_member = worksheets
        .members
        .iter()
        .find(|member| member.name == "Add")
        .expect("Worksheets.Add");
    assert_eq!(worksheets_add_member.params.len(), 4);
    assert!(
        worksheets_add_member
            .params
            .iter()
            .all(|param| param.optional)
    );
    assert_eq!(
        worksheets_add_member
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Worksheet")
    );
    assert_eq!(
        application
            .members
            .iter()
            .find(|member| member.name == "ActiveCell")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    assert_eq!(
        application
            .members
            .iter()
            .find(|member| member.name == "Worksheets")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Worksheets")
    );
    let calculate_full_rebuild = application
        .members
        .iter()
        .find(|member| member.name == "CalculateFullRebuild")
        .expect("Application.CalculateFullRebuild");
    assert_eq!(calculate_full_rebuild.params.len(), 0);
    let application_goto = application
        .members
        .iter()
        .find(|member| member.name == "Goto")
        .expect("Application.Goto");
    assert_eq!(application_goto.access, AccessMode::Read);
    assert_eq!(application_goto.params.len(), 2);
    assert!(application_goto.params[0].optional);
    assert!(application_goto.params[1].optional);
    assert!(application_goto.return_type.is_none());
    let application_range = application
        .members
        .iter()
        .find(|member| member.name == "Range")
        .expect("Application.Range");
    assert_eq!(application_range.access, AccessMode::Read);
    assert_eq!(application_range.params.len(), 2);
    assert!(application_range.params[1].optional);
    assert_eq!(
        application_range
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let application_cells = application
        .members
        .iter()
        .find(|member| member.name == "Cells")
        .expect("Application.Cells");
    assert_eq!(application_cells.access, AccessMode::Read);
    assert_eq!(application_cells.params.len(), 0);
    assert_eq!(
        application_cells
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let application_rows = application
        .members
        .iter()
        .find(|member| member.name == "Rows")
        .expect("Application.Rows");
    assert_eq!(application_rows.access, AccessMode::Read);
    assert_eq!(
        application_rows
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let application_columns = application
        .members
        .iter()
        .find(|member| member.name == "Columns")
        .expect("Application.Columns");
    assert_eq!(application_columns.access, AccessMode::Read);
    assert_eq!(
        application_columns
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let application_intersect = application
        .members
        .iter()
        .find(|member| member.name == "Intersect")
        .expect("Application.Intersect");
    assert_eq!(application_intersect.access, AccessMode::Read);
    assert_eq!(application_intersect.params.len(), 30);
    for (index, param) in application_intersect.params.iter().enumerate() {
        assert_eq!(param.name, format!("Arg{}", index + 1));
        assert_eq!(param.optional, index >= 2);
    }
    assert_eq!(
        application_intersect
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let application_union = application
        .members
        .iter()
        .find(|member| member.name == "Union")
        .expect("Application.Union");
    assert_eq!(application_union.access, AccessMode::Read);
    assert_eq!(application_union.params.len(), 30);
    for (index, param) in application_union.params.iter().enumerate() {
        assert_eq!(param.name, format!("Arg{}", index + 1));
        assert_eq!(param.optional, index >= 2);
    }
    assert_eq!(
        application_union
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let worksheet_activate = worksheet
        .members
        .iter()
        .find(|member| member.name == "Activate")
        .expect("Worksheet.Activate");
    assert_eq!(worksheet_activate.access, AccessMode::Read);
    assert_eq!(worksheet_activate.params.len(), 0);
    assert!(worksheet_activate.return_type.is_none());
    let worksheet_delete = worksheet
        .members
        .iter()
        .find(|member| member.name == "Delete")
        .expect("Worksheet.Delete");
    assert_eq!(worksheet_delete.access, AccessMode::Read);
    assert_eq!(worksheet_delete.params.len(), 0);
    assert_eq!(
        worksheet_delete
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT_BOOL")
    );
    let worksheet_move = worksheet
        .members
        .iter()
        .find(|member| member.name == "Move")
        .expect("Worksheet.Move");
    assert_eq!(worksheet_move.access, AccessMode::Read);
    assert_eq!(worksheet_move.params.len(), 2);
    assert!(worksheet_move.params[0].optional);
    assert!(worksheet_move.params[1].optional);
    assert!(worksheet_move.return_type.is_none());
    assert_eq!(
        worksheet
            .members
            .iter()
            .find(|member| member.name == "Rows")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    assert_eq!(
        worksheet
            .members
            .iter()
            .find(|member| member.name == "Columns")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    assert_eq!(
        range
            .members
            .iter()
            .find(|member| member.name == "Text")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("BSTR")
    );
    assert_eq!(
        range
            .members
            .iter()
            .find(|member| member.name == "Rows")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let range_item = range
        .members
        .iter()
        .find(|member| member.name == "Item")
        .expect("Range.Item");
    assert_eq!(range_item.params.len(), 2);
    assert!(range_item.params[1].optional);
    assert_eq!(
        range_item
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let range_address = range
        .members
        .iter()
        .find(|member| member.name == "Address")
        .expect("Range.Address");
    assert_eq!(range_address.access, AccessMode::Read);
    let expected_address_names = [
        "RowAbsolute",
        "ColumnAbsolute",
        "ReferenceStyle",
        "External",
        "RelativeTo",
    ];
    assert_eq!(range_address.params.len(), expected_address_names.len());
    for (index, expected_name) in expected_address_names.iter().enumerate() {
        assert_eq!(range_address.params[index].name, *expected_name);
        assert!(range_address.params[index].optional);
    }
    assert_eq!(
        range_address
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("BSTR")
    );
    let range_offset = range
        .members
        .iter()
        .find(|member| member.name == "Offset")
        .expect("Range.Offset");
    assert_eq!(range_offset.params.len(), 2);
    assert!(range_offset.params[0].optional);
    assert!(range_offset.params[1].optional);
    assert_eq!(
        range_offset
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let range_resize = range
        .members
        .iter()
        .find(|member| member.name == "Resize")
        .expect("Range.Resize");
    assert_eq!(range_resize.params.len(), 2);
    assert!(range_resize.params[0].optional);
    assert!(range_resize.params[1].optional);
    assert_eq!(
        range_resize
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let range_select = range
        .members
        .iter()
        .find(|member| member.name == "Select")
        .expect("Range.Select");
    assert_eq!(range_select.access, AccessMode::Read);
    assert_eq!(range_select.params.len(), 0);
    assert!(range_select.return_type.is_none());
    let range_clear_contents = range
        .members
        .iter()
        .find(|member| member.name == "ClearContents")
        .expect("Range.ClearContents");
    assert_eq!(range_clear_contents.access, AccessMode::Read);
    assert_eq!(range_clear_contents.params.len(), 0);
    assert!(range_clear_contents.return_type.is_none());
    assert_eq!(
        range
            .members
            .iter()
            .find(|member| member.name == "HasFormula")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT")
    );
    assert_eq!(
        range
            .members
            .iter()
            .find(|member| member.name == "CurrentRegion")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    assert_eq!(
        range
            .members
            .iter()
            .find(|member| member.name == "EntireRow")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    assert_eq!(
        range
            .members
            .iter()
            .find(|member| member.name == "EntireColumn")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    assert_eq!(
        range
            .members
            .iter()
            .find(|member| member.name == "Cells")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    assert_eq!(
        range
            .members
            .iter()
            .find(|member| member.name == "Columns")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    assert_eq!(
        application
            .members
            .iter()
            .find(|member| member.name == "Selection")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    assert_eq!(
        worksheets
            .members
            .iter()
            .find(|member| member.name == "Parent")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Workbook")
    );
    assert_eq!(
        workbook
            .members
            .iter()
            .find(|member| member.name == "Path")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("BSTR")
    );
    assert_eq!(
        workbook
            .members
            .iter()
            .find(|member| member.name == "FullName")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("BSTR")
    );
    assert_eq!(
        workbook
            .members
            .iter()
            .find(|member| member.name == "ReadOnly")
            .and_then(|member| member.return_type.as_ref())
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT_BOOL")
    );
    let saved_member = workbook
        .members
        .iter()
        .find(|member| member.name == "Saved")
        .expect("Saved");
    assert_eq!(saved_member.access, AccessMode::Readwrite);
    assert_eq!(
        saved_member
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT_BOOL")
    );
    let value_member = range
        .members
        .iter()
        .find(|member| member.name == "Value")
        .expect("Value");
    assert_eq!(value_member.access, AccessMode::Readwrite);
    assert_eq!(
        value_member
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT")
    );
    let formula_member = range
        .members
        .iter()
        .find(|member| member.name == "Formula")
        .expect("Formula");
    assert_eq!(formula_member.access, AccessMode::Readwrite);
    assert_eq!(
        formula_member
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT")
    );
}

#[test]
fn validates_capture_bundle_against_raw_typelib_identity_template() {
    let typelib = TypelibIdentityCapture::from_json_path(
        repo_root().join("specs/pinned/raw_typelib_identity.template.json"),
    )
    .expect("typelib json");
    let capture = PiaPublicSurfaceCapture::from_json_path(
        repo_root().join("specs/pinned/excel_pia_public_surface.template.json"),
    )
    .expect("capture json");
    let (document, summary) = normalize_capture_bundle(&typelib, &capture).expect("bundle");

    assert_eq!(summary.library, "Excel");
    assert_eq!(summary.version, "16.0");
    assert_eq!(
        summary.type_library_guid,
        "{00020813-0000-0000-C000-000000000046}"
    );
    assert_eq!(summary.interface_iid_count, 6);
    assert_eq!(summary.coclass_clsid_count, 3);
    assert!(summary.missing_pia_interfaces.is_empty());
    assert!(summary.missing_pia_classes.is_empty());
    assert_eq!(document.interfaces.len(), 6);
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
            .and_then(|metadata| metadata.namespace.as_deref()),
        Some("Microsoft.Office.Interop.Excel")
    );
    assert_eq!(
        document.interfaces[0]
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.iid.as_deref()),
        Some("{000208D5-0000-0000-C000-000000000046}")
    );
    assert_eq!(
        document.interfaces[0]
            .metadata
            .as_ref()
            .map(|metadata| metadata.source_inherits.clone()),
        Some(vec!["IDispatch".to_string()])
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
        document.classes[0]
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.clsid.as_deref()),
        Some("{00024500-0000-0000-C000-000000000046}")
    );
    assert_eq!(
        document.classes[0]
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.source_default_interface.as_deref()),
        Some("Application")
    );
}

#[test]
fn rejects_capture_bundle_when_library_identity_mismatches() {
    let typelib = TypelibIdentityCapture::from_json_path(
        repo_root().join("specs/pinned/raw_typelib_identity.template.json"),
    )
    .expect("typelib json");
    let mut capture = PiaPublicSurfaceCapture::from_json_path(
        repo_root().join("specs/pinned/excel_pia_public_surface.template.json"),
    )
    .expect("capture json");
    capture.library = "Word".to_string();

    let error = normalize_capture_bundle(&typelib, &capture).expect_err("mismatch");
    match error {
        OmCaptureBundleError::LibraryMismatch { typelib, pia } => {
            assert_eq!(typelib, "Excel");
            assert_eq!(pia, "Word");
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn rejects_capture_bundle_when_namespace_identity_mismatches() {
    let typelib = TypelibIdentityCapture::from_json_path(
        repo_root().join("specs/pinned/raw_typelib_identity.template.json"),
    )
    .expect("typelib json");
    let mut capture = PiaPublicSurfaceCapture::from_json_path(
        repo_root().join("specs/pinned/excel_pia_public_surface.template.json"),
    )
    .expect("capture json");
    capture.namespace = "Other.Interop.Excel".to_string();

    let error = normalize_capture_bundle(&typelib, &capture).expect_err("mismatch");
    match error {
        OmCaptureBundleError::NamespaceMismatch { typelib, pia } => {
            assert_eq!(typelib, "Microsoft.Office.Interop.Excel");
            assert_eq!(pia, "Other.Interop.Excel");
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn rejects_capture_bundle_when_version_identity_mismatches() {
    let typelib = TypelibIdentityCapture::from_json_path(
        repo_root().join("specs/pinned/raw_typelib_identity.template.json"),
    )
    .expect("typelib json");
    let mut capture = PiaPublicSurfaceCapture::from_json_path(
        repo_root().join("specs/pinned/excel_pia_public_surface.template.json"),
    )
    .expect("capture json");
    capture.version = "15.0".to_string();

    let error = normalize_capture_bundle(&typelib, &capture).expect_err("mismatch");
    match error {
        OmCaptureBundleError::VersionMismatch { typelib, pia } => {
            assert_eq!(typelib, "16.0");
            assert_eq!(pia, "15.0");
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn summarizes_capture_bundle_with_missing_pia_interfaces_and_classes() {
    let typelib = TypelibIdentityCapture::from_json_path(
        repo_root().join("specs/pinned/raw_typelib_identity.template.json"),
    )
    .expect("typelib json");
    let mut capture = PiaPublicSurfaceCapture::from_json_path(
        repo_root().join("specs/pinned/excel_pia_public_surface.template.json"),
    )
    .expect("capture json");
    capture.interfaces.push(PiaCaptureInterface {
        name: "GhostWorksheet".to_string(),
        kind: InterfaceKind::Dual,
        inherits: vec!["IDispatch".to_string()],
        members: Vec::new(),
        metadata: None,
    });
    capture.classes.push(PiaCaptureClass {
        name: "GhostWorksheetClass".to_string(),
        implements: vec!["GhostWorksheet".to_string()],
        default_interface: Some("GhostWorksheet".to_string()),
        metadata: None,
    });

    let summary = summarize_capture_bundle(&typelib, &capture).expect("bundle summary");

    assert_eq!(summary.library, "Excel");
    assert_eq!(summary.version, "16.0");
    assert_eq!(
        summary.missing_pia_interfaces,
        vec!["GhostWorksheet".to_string()]
    );
    assert_eq!(
        summary.missing_pia_classes,
        vec!["GhostWorksheetClass".to_string()]
    );
}

#[test]
fn normalizes_property_set_only_capture_member_into_write_property() {
    let document = normalize_pia_capture_json(
        r#"
        {
          "library": "Excel",
          "version": "16.0",
          "namespace": "Microsoft.Office.Interop.Excel",
          "enums": [],
          "interfaces": [
            {
              "name": "Worksheet",
              "kind": "dual",
              "inherits": ["IDispatch"],
              "members": [
                {
                  "name": "Name",
                  "memberKind": "property_set",
                  "params": [
                    {
                      "name": "locale",
                      "type": {
                        "kind": "variant",
                        "name": "VARIANT",
                        "aliasOf": "VARIANT"
                      },
                      "optional": true
                    },
                    {
                      "name": "value",
                      "type": {
                        "kind": "primitive",
                        "name": "String",
                        "aliasOf": "BSTR"
                      }
                    }
                  ],
                  "dispId": 110
                }
              ]
            }
          ],
          "classes": []
        }
        "#,
    )
    .expect("normalized document");

    let worksheet = document
        .interfaces
        .iter()
        .find(|interface| interface.name == "Worksheet")
        .expect("Worksheet");
    let name_member = worksheet
        .members
        .iter()
        .find(|member| member.name == "Name")
        .expect("Worksheet.Name");

    assert_eq!(name_member.access, AccessMode::Write);
    assert_eq!(name_member.disp_id, Some(110));
    assert_eq!(name_member.params.len(), 1);
    assert_eq!(name_member.params[0].name, "locale");
    assert!(name_member.params[0].optional);
    assert_eq!(
        name_member
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("BSTR")
    );
    assert_eq!(
        name_member
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.capture.as_ref())
            .map(|capture| {
                capture
                    .origins
                    .iter()
                    .map(|origin| origin.kind.clone())
                    .collect::<Vec<_>>()
            }),
        Some(vec![CaptureOriginKind::PropertySet])
    );
}

#[test]
fn builds_coverage_report_for_each_support_state_bucket() {
    let document = OfficeIdlDocument::from_json_str(
        r#"
        {
          "library": "Excel",
          "version": "16.0",
          "enums": [],
          "interfaces": [
            {
              "name": "Application",
              "kind": "dual",
              "members": [
                { "name": "Generated", "memberKind": "method", "support": "generated_only" },
                { "name": "Stubbed", "memberKind": "method", "support": "stub" },
                { "name": "Partial", "memberKind": "method", "support": "partial" },
                { "name": "Implemented", "memberKind": "method", "support": "implemented" },
                { "name": "Oracle", "memberKind": "method", "support": "oracle_verified" },
                { "name": "Unsupported", "memberKind": "method", "support": "unsupported" }
              ]
            }
          ],
          "classes": []
        }
        "#,
    )
    .expect("document");

    let coverage = build_coverage_report(&document);
    let application = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Application")
        .expect("Application coverage");

    assert_eq!(coverage.member_count, 6);
    assert_eq!(coverage.support_counts.generated_only, 1);
    assert_eq!(coverage.support_counts.stub, 1);
    assert_eq!(coverage.support_counts.partial, 1);
    assert_eq!(coverage.support_counts.implemented, 1);
    assert_eq!(coverage.support_counts.oracle_verified, 1);
    assert_eq!(coverage.support_counts.unsupported, 1);
    assert_eq!(
        coverage.missing_focus_surfaces,
        vec![
            "WorksheetFunction".to_string(),
            "Workbooks".to_string(),
            "Workbook".to_string(),
            "Worksheets".to_string(),
            "Sheets".to_string(),
            "Charts".to_string(),
            "Worksheet".to_string(),
            "Range".to_string(),
            "Areas".to_string(),
            "Names".to_string(),
            "Name".to_string(),
            "ChartObjects".to_string(),
            "ChartObject".to_string(),
            "ShapeRange".to_string(),
            "Chart".to_string(),
            "ChartArea".to_string(),
            "PlotArea".to_string(),
            "ChartTitle".to_string(),
            "Legend".to_string(),
            "LegendEntries".to_string(),
            "LegendEntry".to_string(),
            "LegendKey".to_string(),
            "DataTable".to_string(),
            "ChartFormat".to_string(),
            "Adjustments".to_string(),
            "FillFormat".to_string(),
            "GlowFormat".to_string(),
            "LineFormat".to_string(),
            "PictureFormat".to_string(),
            "Crop".to_string(),
            "ShadowFormat".to_string(),
            "SoftEdgeFormat".to_string(),
            "TextFrame2".to_string(),
            "ThreeDFormat".to_string(),
            "ChartGroups".to_string(),
            "ChartGroup".to_string(),
            "CategoryCollection".to_string(),
            "ChartCategory".to_string(),
            "SeriesLines".to_string(),
            "DropLines".to_string(),
            "HiLoLines".to_string(),
            "UpBars".to_string(),
            "DownBars".to_string(),
            "Axes".to_string(),
            "Axis".to_string(),
            "TickLabels".to_string(),
            "Gridlines".to_string(),
            "DisplayUnitLabel".to_string(),
            "AxisTitle".to_string(),
            "SeriesCollection".to_string(),
            "Series".to_string(),
            "LeaderLines".to_string(),
            "Border".to_string(),
            "DataLabels".to_string(),
            "DataLabel".to_string(),
            "Points".to_string(),
            "Point".to_string()
        ]
    );
    assert_eq!(
        application.generated_only_members,
        vec!["Generated".to_string()]
    );
    assert_eq!(application.stub_members, vec!["Stubbed".to_string()]);
    assert_eq!(application.partial_members, vec!["Partial".to_string()]);
    assert_eq!(
        application.implemented_members,
        vec!["Implemented".to_string()]
    );
    assert_eq!(
        application.oracle_verified_members,
        vec!["Oracle".to_string()]
    );
    assert_eq!(
        application.unsupported_members,
        vec!["Unsupported".to_string()]
    );
}

#[test]
fn reports_manifest_not_ready_when_windows_capture_requirements_are_missing() {
    let manifest_toml = fs::read_to_string(repo_root().join("specs/pinned/om_sources.toml"))
        .expect("manifest template")
        .replace(r#"host_os = "windows""#, r#"host_os = "linux""#)
        .replace(
            "requires_installed_excel = true",
            "requires_installed_excel = false",
        )
        .replace(
            "requires_windows_sdk = true",
            "requires_windows_sdk = false",
        )
        .replace(
            "requires_dotnet_framework_tooling = true",
            "requires_dotnet_framework_tooling = false",
        );

    let summary = summarize_om_sources_toml(&manifest_toml).expect("manifest summary");

    assert!(!summary.ready_for_windows_capture);
    assert_eq!(
        summary.pending_outputs,
        vec![
            "raw_typelib_identity.json".to_string(),
            "excel_typelib_snapshot.idl".to_string(),
            "excel_typelib_snapshot.odl".to_string(),
            "excel_pia_identity.json".to_string(),
            "excel_pia_public_surface.json".to_string(),
        ]
    );
    assert_eq!(
        summary.unresolved_target_fields,
        vec![
            "product_family",
            "channel",
            "version",
            "build",
            "arch",
            "locale"
        ]
    );
}

#[test]
fn reports_json_error_when_typelib_capture_is_invalid() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root = std::env::temp_dir().join(format!("ootd-step3-invalid-{unique_suffix}"));
    let raw_dir = bundle_root.join("raw");
    let snapshots_dir = bundle_root.join("snapshots");

    fs::create_dir_all(&raw_dir).expect("raw dir");
    fs::create_dir_all(&snapshots_dir).expect("snapshots dir");
    fs::write(raw_dir.join("raw_typelib_identity.json"), "{").expect("invalid typelib");
    fs::write(
        snapshots_dir.join("excel_pia_public_surface.json"),
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template"),
    )
    .expect("write pia surface");

    let output_path = bundle_root.join("manifest/office_idl_excel_om.json");
    let error = generate_canonical_office_idl_from_dir(&bundle_root, &output_path)
        .expect_err("invalid typelib json should fail");
    match error {
        CanonicalOmGenerationError::Json { path, .. } => {
            assert_eq!(path, raw_dir.join("raw_typelib_identity.json"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn builds_focus_surface_registry_from_json_wrapper() {
    let registry = build_focus_surface_registry_from_json(
        r#"
        {
          "library": "Excel",
          "version": "16.0",
          "interfaces": [
            {
              "name": "Application",
              "kind": "dual",
              "members": [
                { "name": "Visible", "memberKind": "property", "support": "implemented" }
              ]
            }
          ],
          "classes": []
        }
        "#,
    )
    .expect("registry");

    assert_eq!(registry.library, "Excel");
    assert_eq!(registry.focus_surfaces.len(), 1);
    assert_eq!(
        registry.missing_focus_surfaces,
        vec![
            "WorksheetFunction".to_string(),
            "Workbooks".to_string(),
            "Workbook".to_string(),
            "Worksheets".to_string(),
            "Sheets".to_string(),
            "Charts".to_string(),
            "Worksheet".to_string(),
            "Range".to_string(),
            "Areas".to_string(),
            "Names".to_string(),
            "Name".to_string(),
            "ChartObjects".to_string(),
            "ChartObject".to_string(),
            "ShapeRange".to_string(),
            "Chart".to_string(),
            "ChartArea".to_string(),
            "PlotArea".to_string(),
            "ChartTitle".to_string(),
            "Legend".to_string(),
            "LegendEntries".to_string(),
            "LegendEntry".to_string(),
            "LegendKey".to_string(),
            "DataTable".to_string(),
            "ChartFormat".to_string(),
            "Adjustments".to_string(),
            "FillFormat".to_string(),
            "GlowFormat".to_string(),
            "LineFormat".to_string(),
            "PictureFormat".to_string(),
            "Crop".to_string(),
            "ShadowFormat".to_string(),
            "SoftEdgeFormat".to_string(),
            "TextFrame2".to_string(),
            "ThreeDFormat".to_string(),
            "ChartGroups".to_string(),
            "ChartGroup".to_string(),
            "CategoryCollection".to_string(),
            "ChartCategory".to_string(),
            "SeriesLines".to_string(),
            "DropLines".to_string(),
            "HiLoLines".to_string(),
            "UpBars".to_string(),
            "DownBars".to_string(),
            "Axes".to_string(),
            "Axis".to_string(),
            "TickLabels".to_string(),
            "Gridlines".to_string(),
            "DisplayUnitLabel".to_string(),
            "AxisTitle".to_string(),
            "SeriesCollection".to_string(),
            "Series".to_string(),
            "LeaderLines".to_string(),
            "Border".to_string(),
            "DataLabels".to_string(),
            "DataLabel".to_string(),
            "Points".to_string(),
            "Point".to_string()
        ]
    );
    assert_eq!(registry.focus_surfaces[0].name, "Application");
    assert_eq!(registry.focus_surfaces[0].member_count, 1);
}

#[test]
fn builds_coverage_report_from_path_wrapper() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let path = std::env::temp_dir().join(format!("ootd-coverage-wrapper-{unique_suffix}.json"));
    let json = r#"
    {
      "library": "Excel",
      "version": "16.0",
      "interfaces": [
        {
          "name": "Application",
          "kind": "dual",
          "members": [
            { "name": "Open", "memberKind": "method", "support": "partial" }
          ]
        }
      ],
      "classes": []
    }
    "#;
    fs::write(&path, json).expect("write document");

    let report_from_path = build_coverage_report_from_path(&path).expect("coverage report");
    let report_from_json = build_coverage_report_from_json(json).expect("coverage report");

    assert_eq!(report_from_path, report_from_json);
    assert_eq!(report_from_path.support_counts.partial, 1);
    assert_eq!(report_from_path.missing_focus_surfaces.len(), 57);

    fs::remove_file(&path).expect("remove temp document");
}

#[test]
fn builds_focus_surface_registry_from_path_wrapper() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let path = std::env::temp_dir().join(format!("ootd-registry-wrapper-{unique_suffix}.json"));
    let json = r#"
    {
      "library": "Excel",
      "version": "16.0",
      "interfaces": [
        {
          "name": "Workbook",
          "kind": "dual",
          "members": [
            { "name": "Name", "memberKind": "property", "support": "stub" }
          ]
        }
      ],
      "classes": []
    }
    "#;
    fs::write(&path, json).expect("write document");

    let registry_from_path =
        build_focus_surface_registry_from_path(&path).expect("focus surface registry");
    let registry_from_json =
        build_focus_surface_registry_from_json(json).expect("focus surface registry");

    assert_eq!(registry_from_path, registry_from_json);
    assert_eq!(registry_from_path.focus_surfaces[0].name, "Workbook");

    fs::remove_file(&path).expect("remove temp document");
}

#[test]
fn normalize_capture_bundle_from_dir_requires_materialized_typelib_identity() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root = std::env::temp_dir().join(format!("ootd-step3-missing-{unique_suffix}"));
    fs::create_dir_all(bundle_root.join("snapshots")).expect("snapshots dir");
    fs::write(
        bundle_root.join("snapshots/excel_pia_public_surface.json"),
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template"),
    )
    .expect("write pia surface");

    let error = normalize_capture_bundle_from_dir(&bundle_root)
        .expect_err("missing typelib identity should fail");
    match error {
        CanonicalOmGenerationError::Io { action, path, .. } => {
            assert_eq!(action, "read typelib identity capture");
            assert_eq!(path, bundle_root.join("raw/raw_typelib_identity.json"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn normalize_capture_bundle_from_dir_validates_manifest_checksum_contract() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root = std::env::temp_dir().join(format!("ootd-step3-contract-{unique_suffix}"));
    let raw_dir = bundle_root.join("raw");
    let snapshots_dir = bundle_root.join("snapshots");
    let manifest_dir = bundle_root.join("manifest");

    fs::create_dir_all(&raw_dir).expect("raw dir");
    fs::create_dir_all(&snapshots_dir).expect("snapshots dir");
    fs::create_dir_all(&manifest_dir).expect("manifest dir");
    fs::write(
        raw_dir.join("raw_typelib_identity.json"),
        fs::read_to_string(repo_root().join("specs/pinned/raw_typelib_identity.template.json"))
            .expect("typelib template"),
    )
    .expect("write typelib");
    fs::write(
        raw_dir.join("excel_pia_identity.json"),
        r#"{"assembly":"Excel"}"#,
    )
    .expect("write pia identity");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.idl"),
        "library Excel {}",
    )
    .expect("write idl");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.odl"),
        "odl Excel {}",
    )
    .expect("write odl");
    fs::write(
        snapshots_dir.join("excel_pia_public_surface.json"),
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template"),
    )
    .expect("write pia surface");
    fs::write(
        manifest_dir.join("capture_manifest.json"),
        r#"{
  "expectedCaptureOutputs": [
    "raw_typelib_identity.json",
    "excel_typelib_snapshot.idl",
    "excel_typelib_snapshot.odl",
    "excel_pia_identity.json",
    "excel_pia_public_surface.json"
  ],
  "writableOutputs": {
    "raw_typelib_identity": "C:\\capture\\raw\\raw_typelib_identity.json",
    "excel_typelib_snapshot_idl": "C:\\capture\\snapshots\\excel_typelib_snapshot.idl",
    "excel_typelib_snapshot_odl": "C:\\capture\\snapshots\\excel_typelib_snapshot.odl",
    "excel_pia_identity": "C:\\capture\\raw\\excel_pia_identity.json",
    "excel_pia_public_surface": "C:\\capture\\snapshots\\excel_pia_public_surface.json"
  }
}"#,
    )
    .expect("write manifest");
    fs::write(
        manifest_dir.join("output_checksums.json"),
        r#"{
  "raw/raw_typelib_identity.json": "sha"
}"#,
    )
    .expect("write incomplete checksums");

    let error = normalize_capture_bundle_from_dir(&bundle_root)
        .expect_err("incomplete checksum contract should fail");
    match error {
        CanonicalOmGenerationError::CaptureBundleContract { message } => {
            assert!(message.contains("output_checksums.json payload names"));
            assert!(message.contains("excel_pia_public_surface.json"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn normalize_capture_bundle_from_dir_rejects_duplicate_manifest_expected_outputs() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root =
        std::env::temp_dir().join(format!("ootd-step3-duplicate-expected-{unique_suffix}"));
    let manifest_dir = bundle_root.join("manifest");

    fs::create_dir_all(&manifest_dir).expect("manifest dir");
    fs::write(
        manifest_dir.join("capture_manifest.json"),
        r#"{
  "expectedCaptureOutputs": [
    "raw_typelib_identity.json",
    "raw_typelib_identity.json",
    "excel_typelib_snapshot.idl",
    "excel_typelib_snapshot.odl",
    "excel_pia_identity.json",
    "excel_pia_public_surface.json"
  ],
  "writableOutputs": {
    "raw_typelib_identity": "C:\\capture\\raw\\raw_typelib_identity.json",
    "excel_typelib_snapshot_idl": "C:\\capture\\snapshots\\excel_typelib_snapshot.idl",
    "excel_typelib_snapshot_odl": "C:\\capture\\snapshots\\excel_typelib_snapshot.odl",
    "excel_pia_identity": "C:\\capture\\raw\\excel_pia_identity.json",
    "excel_pia_public_surface": "C:\\capture\\snapshots\\excel_pia_public_surface.json"
  }
}"#,
    )
    .expect("write manifest");

    let error = normalize_capture_bundle_from_dir(&bundle_root)
        .expect_err("duplicate manifest expected output should fail");
    match error {
        CanonicalOmGenerationError::CaptureBundleContract { message } => {
            assert!(message.contains("expectedCaptureOutputs"));
            assert!(message.contains("duplicate entry raw_typelib_identity.json"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn normalize_capture_bundle_from_dir_rejects_unexpected_writable_output_keys() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root =
        std::env::temp_dir().join(format!("ootd-step3-extra-writable-output-{unique_suffix}"));
    let raw_dir = bundle_root.join("raw");
    let snapshots_dir = bundle_root.join("snapshots");
    let manifest_dir = bundle_root.join("manifest");

    fs::create_dir_all(&raw_dir).expect("raw dir");
    fs::create_dir_all(&snapshots_dir).expect("snapshots dir");
    fs::create_dir_all(&manifest_dir).expect("manifest dir");
    fs::write(
        raw_dir.join("raw_typelib_identity.json"),
        fs::read_to_string(repo_root().join("specs/pinned/raw_typelib_identity.template.json"))
            .expect("typelib template"),
    )
    .expect("write typelib");
    fs::write(
        raw_dir.join("excel_pia_identity.json"),
        r#"{"assembly":"Excel"}"#,
    )
    .expect("write pia identity");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.idl"),
        "library Excel {}",
    )
    .expect("write idl");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.odl"),
        "odl Excel {}",
    )
    .expect("write odl");
    fs::write(
        snapshots_dir.join("excel_pia_public_surface.json"),
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template"),
    )
    .expect("write pia surface");
    fs::write(
        manifest_dir.join("capture_manifest.json"),
        r#"{
  "expectedCaptureOutputs": [
    "raw_typelib_identity.json",
    "excel_typelib_snapshot.idl",
    "excel_typelib_snapshot.odl",
    "excel_pia_identity.json",
    "excel_pia_public_surface.json"
  ],
  "writableOutputs": {
    "raw_typelib_identity": "C:\\capture\\raw\\raw_typelib_identity.json",
    "excel_typelib_snapshot_idl": "C:\\capture\\snapshots\\excel_typelib_snapshot.idl",
    "excel_typelib_snapshot_odl": "C:\\capture\\snapshots\\excel_typelib_snapshot.odl",
    "excel_pia_identity": "C:\\capture\\raw\\excel_pia_identity.json",
    "excel_pia_public_surface": "C:\\capture\\snapshots\\excel_pia_public_surface.json",
    "unplanned_payload": "C:\\capture\\extra\\unplanned_payload.json"
  }
}"#,
    )
    .expect("write manifest");
    fs::write(
        manifest_dir.join("output_checksums.json"),
        format!(
            r#"{{
  "raw/raw_typelib_identity.json": "{}",
  "snapshots/excel_typelib_snapshot.idl": "{}",
  "snapshots/excel_typelib_snapshot.odl": "{}",
  "raw/excel_pia_identity.json": "{}",
  "snapshots/excel_pia_public_surface.json": "{}",
  "logs/capture.log": "{}"
}}"#,
            sha256_hex(
                fs::read(raw_dir.join("raw_typelib_identity.json"))
                    .expect("typelib")
                    .as_slice()
            ),
            sha256_hex(b"library Excel {}"),
            sha256_hex(b"odl Excel {}"),
            sha256_hex(br#"{"assembly":"Excel"}"#),
            sha256_hex(
                fs::read(snapshots_dir.join("excel_pia_public_surface.json"))
                    .expect("pia")
                    .as_slice()
            ),
            sha256_hex(b"capture log")
        ),
    )
    .expect("write checksums");

    let error = normalize_capture_bundle_from_dir(&bundle_root)
        .expect_err("unexpected writable output key should fail");
    match error {
        CanonicalOmGenerationError::CaptureBundleContract { message } => {
            assert!(message.contains("writableOutputs keys"));
            assert!(message.contains("unplanned_payload"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn load_capture_bundle_allows_known_auxiliary_writable_output_keys() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root =
        std::env::temp_dir().join(format!("ootd-step3-aux-writable-output-{unique_suffix}"));
    let raw_dir = bundle_root.join("raw");
    let snapshots_dir = bundle_root.join("snapshots");
    let manifest_dir = bundle_root.join("manifest");
    let logs_dir = bundle_root.join("logs");

    fs::create_dir_all(&raw_dir).expect("raw dir");
    fs::create_dir_all(&snapshots_dir).expect("snapshots dir");
    fs::create_dir_all(&manifest_dir).expect("manifest dir");
    fs::create_dir_all(&logs_dir).expect("logs dir");
    fs::write(
        raw_dir.join("raw_typelib_identity.json"),
        fs::read_to_string(repo_root().join("specs/pinned/raw_typelib_identity.template.json"))
            .expect("typelib template"),
    )
    .expect("write typelib");
    fs::write(
        raw_dir.join("excel_pia_identity.json"),
        r#"{"assembly":"Excel"}"#,
    )
    .expect("write pia identity");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.idl"),
        "library Excel {}",
    )
    .expect("write idl");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.odl"),
        "odl Excel {}",
    )
    .expect("write odl");
    fs::write(
        snapshots_dir.join("excel_pia_public_surface.json"),
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template"),
    )
    .expect("write pia surface");
    fs::write(logs_dir.join("capture.log"), "capture log").expect("write log");
    fs::write(
        manifest_dir.join("capture_manifest.json"),
        r#"{
  "expectedCaptureOutputs": [
    "raw_typelib_identity.json",
    "excel_typelib_snapshot.idl",
    "excel_typelib_snapshot.odl",
    "excel_pia_identity.json",
    "excel_pia_public_surface.json"
  ],
  "writableOutputs": {
    "raw_typelib_identity": "C:\\capture\\raw\\raw_typelib_identity.json",
    "excel_typelib_snapshot_idl": "C:\\capture\\snapshots\\excel_typelib_snapshot.idl",
    "excel_typelib_snapshot_odl": "C:\\capture\\snapshots\\excel_typelib_snapshot.odl",
    "excel_pia_identity": "C:\\capture\\raw\\excel_pia_identity.json",
    "excel_pia_public_surface": "C:\\capture\\snapshots\\excel_pia_public_surface.json",
    "capture_log": "C:\\capture\\logs\\capture.log",
    "capture_manifest": "C:\\capture\\manifest\\capture_manifest.json",
    "output_checksums": "C:\\capture\\manifest\\output_checksums.json"
  }
}"#,
    )
    .expect("write manifest");
    fs::write(
        manifest_dir.join("output_checksums.json"),
        format!(
            r#"{{
  "raw/raw_typelib_identity.json": "{}",
  "snapshots/excel_typelib_snapshot.idl": "{}",
  "snapshots/excel_typelib_snapshot.odl": "{}",
  "raw/excel_pia_identity.json": "{}",
  "snapshots/excel_pia_public_surface.json": "{}"
}}"#,
            sha256_hex(
                fs::read(raw_dir.join("raw_typelib_identity.json"))
                    .expect("typelib")
                    .as_slice()
            ),
            sha256_hex(b"library Excel {}"),
            sha256_hex(b"odl Excel {}"),
            sha256_hex(br#"{"assembly":"Excel"}"#),
            sha256_hex(
                fs::read(snapshots_dir.join("excel_pia_public_surface.json"))
                    .expect("pia")
                    .as_slice()
            )
        ),
    )
    .expect("write checksums");

    load_capture_bundle(&bundle_root).expect("known auxiliary writable outputs should pass");
}

#[test]
fn normalize_capture_bundle_from_dir_rejects_wrong_writable_output_paths() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root =
        std::env::temp_dir().join(format!("ootd-step3-wrong-writable-path-{unique_suffix}"));
    let raw_dir = bundle_root.join("raw");
    let snapshots_dir = bundle_root.join("snapshots");
    let manifest_dir = bundle_root.join("manifest");

    fs::create_dir_all(&raw_dir).expect("raw dir");
    fs::create_dir_all(&snapshots_dir).expect("snapshots dir");
    fs::create_dir_all(&manifest_dir).expect("manifest dir");
    fs::write(
        raw_dir.join("raw_typelib_identity.json"),
        fs::read_to_string(repo_root().join("specs/pinned/raw_typelib_identity.template.json"))
            .expect("typelib template"),
    )
    .expect("write typelib");
    fs::write(
        raw_dir.join("excel_pia_identity.json"),
        r#"{"assembly":"Excel"}"#,
    )
    .expect("write pia identity");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.idl"),
        "library Excel {}",
    )
    .expect("write idl");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.odl"),
        "odl Excel {}",
    )
    .expect("write odl");
    fs::write(
        snapshots_dir.join("excel_pia_public_surface.json"),
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template"),
    )
    .expect("write pia surface");
    fs::write(
        manifest_dir.join("capture_manifest.json"),
        r#"{
  "expectedCaptureOutputs": [
    "raw_typelib_identity.json",
    "excel_typelib_snapshot.idl",
    "excel_typelib_snapshot.odl",
    "excel_pia_identity.json",
    "excel_pia_public_surface.json"
  ],
  "writableOutputs": {
    "raw_typelib_identity": "C:\\capture\\raw\\raw_typelib_identity.json",
    "excel_typelib_snapshot_idl": "C:\\capture\\snapshots\\excel_typelib_snapshot.idl",
    "excel_typelib_snapshot_odl": "C:\\capture\\snapshots\\excel_typelib_snapshot.odl",
    "excel_pia_identity": "C:\\capture\\extra\\excel_pia_identity.json",
    "excel_pia_public_surface": "C:\\capture\\snapshots\\excel_pia_public_surface.json"
  }
}"#,
    )
    .expect("write manifest");
    fs::write(
        manifest_dir.join("output_checksums.json"),
        format!(
            r#"{{
  "raw/raw_typelib_identity.json": "{}",
  "snapshots/excel_typelib_snapshot.idl": "{}",
  "snapshots/excel_typelib_snapshot.odl": "{}",
  "raw/excel_pia_identity.json": "{}",
  "snapshots/excel_pia_public_surface.json": "{}"
}}"#,
            sha256_hex(
                fs::read(raw_dir.join("raw_typelib_identity.json"))
                    .expect("typelib")
                    .as_slice()
            ),
            sha256_hex(b"library Excel {}"),
            sha256_hex(b"odl Excel {}"),
            sha256_hex(br#"{"assembly":"Excel"}"#),
            sha256_hex(
                fs::read(snapshots_dir.join("excel_pia_public_surface.json"))
                    .expect("pia")
                    .as_slice()
            )
        ),
    )
    .expect("write checksums");

    let error = normalize_capture_bundle_from_dir(&bundle_root)
        .expect_err("wrong writable output path should fail");
    match error {
        CanonicalOmGenerationError::CaptureBundleContract { message } => {
            assert!(message.contains("writableOutputs.excel_pia_identity"));
            assert!(message.contains("extra\\excel_pia_identity.json"));
            assert!(message.contains("raw/excel_pia_identity.json"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn normalize_capture_bundle_from_dir_validates_embedded_receipt_contract() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root = std::env::temp_dir().join(format!("ootd-step3-receipt-{unique_suffix}"));
    let raw_dir = bundle_root.join("raw");
    let snapshots_dir = bundle_root.join("snapshots");
    let manifest_dir = bundle_root.join("manifest");

    fs::create_dir_all(&raw_dir).expect("raw dir");
    fs::create_dir_all(&snapshots_dir).expect("snapshots dir");
    fs::create_dir_all(&manifest_dir).expect("manifest dir");
    fs::write(
        raw_dir.join("raw_typelib_identity.json"),
        fs::read_to_string(repo_root().join("specs/pinned/raw_typelib_identity.template.json"))
            .expect("typelib template"),
    )
    .expect("write typelib");
    fs::write(
        raw_dir.join("excel_pia_identity.json"),
        r#"{"assembly":"Excel"}"#,
    )
    .expect("write pia identity");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.idl"),
        "library Excel {}",
    )
    .expect("write idl");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.odl"),
        "odl Excel {}",
    )
    .expect("write odl");
    fs::write(
        snapshots_dir.join("excel_pia_public_surface.json"),
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template"),
    )
    .expect("write pia surface");
    fs::write(
        manifest_dir.join("capture_manifest.json"),
        r#"{
  "expectedCaptureOutputs": [
    "raw_typelib_identity.json",
    "excel_typelib_snapshot.idl",
    "excel_typelib_snapshot.odl",
    "excel_pia_identity.json",
    "excel_pia_public_surface.json"
  ],
  "writableOutputs": {
    "raw_typelib_identity": "C:\\capture\\raw\\raw_typelib_identity.json",
    "excel_typelib_snapshot_idl": "C:\\capture\\snapshots\\excel_typelib_snapshot.idl",
    "excel_typelib_snapshot_odl": "C:\\capture\\snapshots\\excel_typelib_snapshot.odl",
    "excel_pia_identity": "C:\\capture\\raw\\excel_pia_identity.json",
    "excel_pia_public_surface": "C:\\capture\\snapshots\\excel_pia_public_surface.json"
  },
  "executionReceipt": {
    "expectedCaptureOutputs": [
      "raw_typelib_identity.json",
      "excel_typelib_snapshot.idl",
      "excel_typelib_snapshot.odl",
      "excel_pia_identity.json",
      "excel_pia_public_surface.json"
    ],
    "commandResults": [
      { "name": "powershell_capture_reflection", "status": "completed" }
    ],
    "manualStepResults": [
      { "name": "oleview_snapshot_export", "status": "pending" }
    ]
  }
}"#,
    )
    .expect("write manifest");
    fs::write(
        manifest_dir.join("output_checksums.json"),
        format!(
            r#"{{
  "raw/raw_typelib_identity.json": "{}",
  "snapshots/excel_typelib_snapshot.idl": "{}",
  "snapshots/excel_typelib_snapshot.odl": "{}",
  "raw/excel_pia_identity.json": "{}",
  "snapshots/excel_pia_public_surface.json": "{}"
}}"#,
            sha256_hex(
                fs::read(raw_dir.join("raw_typelib_identity.json"))
                    .expect("typelib")
                    .as_slice()
            ),
            sha256_hex(b"library Excel {}"),
            sha256_hex(b"odl Excel {}"),
            sha256_hex(br#"{"assembly":"Excel"}"#),
            sha256_hex(
                fs::read(snapshots_dir.join("excel_pia_public_surface.json"))
                    .expect("pia")
                    .as_slice()
            )
        ),
    )
    .expect("write checksums");

    let error = normalize_capture_bundle_from_dir(&bundle_root)
        .expect_err("pending embedded receipt should fail");
    match error {
        CanonicalOmGenerationError::CaptureBundleContract { message } => {
            assert!(message.contains("manualStepResults.oleview_snapshot_export"));
            assert!(message.contains("not completed"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn normalize_capture_bundle_from_dir_rejects_duplicate_embedded_receipt_results() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root =
        std::env::temp_dir().join(format!("ootd-step3-duplicate-receipt-{unique_suffix}"));
    let raw_dir = bundle_root.join("raw");
    let snapshots_dir = bundle_root.join("snapshots");
    let manifest_dir = bundle_root.join("manifest");

    fs::create_dir_all(&raw_dir).expect("raw dir");
    fs::create_dir_all(&snapshots_dir).expect("snapshots dir");
    fs::create_dir_all(&manifest_dir).expect("manifest dir");
    fs::write(
        raw_dir.join("raw_typelib_identity.json"),
        fs::read_to_string(repo_root().join("specs/pinned/raw_typelib_identity.template.json"))
            .expect("typelib template"),
    )
    .expect("write typelib");
    fs::write(
        raw_dir.join("excel_pia_identity.json"),
        r#"{"assembly":"Excel"}"#,
    )
    .expect("write pia identity");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.idl"),
        "library Excel {}",
    )
    .expect("write idl");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.odl"),
        "odl Excel {}",
    )
    .expect("write odl");
    fs::write(
        snapshots_dir.join("excel_pia_public_surface.json"),
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template"),
    )
    .expect("write pia surface");
    fs::write(
        manifest_dir.join("capture_manifest.json"),
        r#"{
  "expectedCaptureOutputs": [
    "raw_typelib_identity.json",
    "excel_typelib_snapshot.idl",
    "excel_typelib_snapshot.odl",
    "excel_pia_identity.json",
    "excel_pia_public_surface.json"
  ],
  "writableOutputs": {
    "raw_typelib_identity": "C:\\capture\\raw\\raw_typelib_identity.json",
    "excel_typelib_snapshot_idl": "C:\\capture\\snapshots\\excel_typelib_snapshot.idl",
    "excel_typelib_snapshot_odl": "C:\\capture\\snapshots\\excel_typelib_snapshot.odl",
    "excel_pia_identity": "C:\\capture\\raw\\excel_pia_identity.json",
    "excel_pia_public_surface": "C:\\capture\\snapshots\\excel_pia_public_surface.json"
  },
  "executionReceipt": {
    "expectedCaptureOutputs": [
      "raw_typelib_identity.json",
      "excel_typelib_snapshot.idl",
      "excel_typelib_snapshot.odl",
      "excel_pia_identity.json",
      "excel_pia_public_surface.json"
    ],
    "commandResults": [
      { "name": "powershell_capture_reflection", "status": "completed" }
    ],
    "manualStepResults": [
      { "name": "oleview_snapshot_export", "status": "completed" },
      { "name": "oleview_snapshot_export", "status": "completed" }
    ]
  }
}"#,
    )
    .expect("write manifest");
    fs::write(
        manifest_dir.join("output_checksums.json"),
        format!(
            r#"{{
  "raw/raw_typelib_identity.json": "{}",
  "snapshots/excel_typelib_snapshot.idl": "{}",
  "snapshots/excel_typelib_snapshot.odl": "{}",
  "raw/excel_pia_identity.json": "{}",
  "snapshots/excel_pia_public_surface.json": "{}"
}}"#,
            sha256_hex(
                fs::read(raw_dir.join("raw_typelib_identity.json"))
                    .expect("typelib")
                    .as_slice()
            ),
            sha256_hex(b"library Excel {}"),
            sha256_hex(b"odl Excel {}"),
            sha256_hex(br#"{"assembly":"Excel"}"#),
            sha256_hex(
                fs::read(snapshots_dir.join("excel_pia_public_surface.json"))
                    .expect("pia")
                    .as_slice()
            )
        ),
    )
    .expect("write checksums");

    let error = normalize_capture_bundle_from_dir(&bundle_root)
        .expect_err("duplicate embedded receipt result should fail");
    match error {
        CanonicalOmGenerationError::CaptureBundleContract { message } => {
            assert!(message.contains("manualStepResults"));
            assert!(message.contains("duplicate result oleview_snapshot_export"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn normalize_capture_bundle_from_dir_rejects_duplicate_embedded_receipt_expected_outputs() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root = std::env::temp_dir().join(format!(
        "ootd-step3-duplicate-receipt-expected-{unique_suffix}"
    ));
    let raw_dir = bundle_root.join("raw");
    let snapshots_dir = bundle_root.join("snapshots");
    let manifest_dir = bundle_root.join("manifest");

    fs::create_dir_all(&raw_dir).expect("raw dir");
    fs::create_dir_all(&snapshots_dir).expect("snapshots dir");
    fs::create_dir_all(&manifest_dir).expect("manifest dir");
    fs::write(
        raw_dir.join("raw_typelib_identity.json"),
        fs::read_to_string(repo_root().join("specs/pinned/raw_typelib_identity.template.json"))
            .expect("typelib template"),
    )
    .expect("write typelib");
    fs::write(
        raw_dir.join("excel_pia_identity.json"),
        r#"{"assembly":"Excel"}"#,
    )
    .expect("write pia identity");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.idl"),
        "library Excel {}",
    )
    .expect("write idl");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.odl"),
        "odl Excel {}",
    )
    .expect("write odl");
    fs::write(
        snapshots_dir.join("excel_pia_public_surface.json"),
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template"),
    )
    .expect("write pia surface");
    fs::write(
        manifest_dir.join("capture_manifest.json"),
        r#"{
  "expectedCaptureOutputs": [
    "raw_typelib_identity.json",
    "excel_typelib_snapshot.idl",
    "excel_typelib_snapshot.odl",
    "excel_pia_identity.json",
    "excel_pia_public_surface.json"
  ],
  "writableOutputs": {
    "raw_typelib_identity": "C:\\capture\\raw\\raw_typelib_identity.json",
    "excel_typelib_snapshot_idl": "C:\\capture\\snapshots\\excel_typelib_snapshot.idl",
    "excel_typelib_snapshot_odl": "C:\\capture\\snapshots\\excel_typelib_snapshot.odl",
    "excel_pia_identity": "C:\\capture\\raw\\excel_pia_identity.json",
    "excel_pia_public_surface": "C:\\capture\\snapshots\\excel_pia_public_surface.json"
  },
  "executionReceipt": {
    "expectedCaptureOutputs": [
      "raw_typelib_identity.json",
      "raw_typelib_identity.json",
      "excel_typelib_snapshot.idl",
      "excel_typelib_snapshot.odl",
      "excel_pia_identity.json",
      "excel_pia_public_surface.json"
    ],
    "commandResults": [
      { "name": "powershell_capture_reflection", "status": "completed" }
    ],
    "manualStepResults": [
      { "name": "oleview_snapshot_export", "status": "completed" }
    ]
  }
}"#,
    )
    .expect("write manifest");
    fs::write(
        manifest_dir.join("output_checksums.json"),
        format!(
            r#"{{
  "raw/raw_typelib_identity.json": "{}",
  "snapshots/excel_typelib_snapshot.idl": "{}",
  "snapshots/excel_typelib_snapshot.odl": "{}",
  "raw/excel_pia_identity.json": "{}",
  "snapshots/excel_pia_public_surface.json": "{}"
}}"#,
            sha256_hex(
                fs::read(raw_dir.join("raw_typelib_identity.json"))
                    .expect("typelib")
                    .as_slice()
            ),
            sha256_hex(b"library Excel {}"),
            sha256_hex(b"odl Excel {}"),
            sha256_hex(br#"{"assembly":"Excel"}"#),
            sha256_hex(
                fs::read(snapshots_dir.join("excel_pia_public_surface.json"))
                    .expect("pia")
                    .as_slice()
            )
        ),
    )
    .expect("write checksums");

    let error = normalize_capture_bundle_from_dir(&bundle_root)
        .expect_err("duplicate embedded receipt expected output should fail");
    match error {
        CanonicalOmGenerationError::CaptureBundleContract { message } => {
            assert!(message.contains("executionReceipt.expectedCaptureOutputs"));
            assert!(message.contains("duplicate entry raw_typelib_identity.json"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn normalize_capture_bundle_from_dir_requires_checksum_listed_payload_files() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root =
        std::env::temp_dir().join(format!("ootd-step3-missing-payload-{unique_suffix}"));
    let raw_dir = bundle_root.join("raw");
    let snapshots_dir = bundle_root.join("snapshots");
    let manifest_dir = bundle_root.join("manifest");

    fs::create_dir_all(&raw_dir).expect("raw dir");
    fs::create_dir_all(&snapshots_dir).expect("snapshots dir");
    fs::create_dir_all(&manifest_dir).expect("manifest dir");
    fs::write(
        raw_dir.join("raw_typelib_identity.json"),
        fs::read_to_string(repo_root().join("specs/pinned/raw_typelib_identity.template.json"))
            .expect("typelib template"),
    )
    .expect("write typelib");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.idl"),
        "library Excel {}",
    )
    .expect("write idl");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.odl"),
        "odl Excel {}",
    )
    .expect("write odl");
    fs::write(
        snapshots_dir.join("excel_pia_public_surface.json"),
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template"),
    )
    .expect("write pia surface");
    fs::write(
        manifest_dir.join("capture_manifest.json"),
        r#"{
  "expectedCaptureOutputs": [
    "raw_typelib_identity.json",
    "excel_typelib_snapshot.idl",
    "excel_typelib_snapshot.odl",
    "excel_pia_identity.json",
    "excel_pia_public_surface.json"
  ],
  "writableOutputs": {
    "raw_typelib_identity": "C:\\capture\\raw\\raw_typelib_identity.json",
    "excel_typelib_snapshot_idl": "C:\\capture\\snapshots\\excel_typelib_snapshot.idl",
    "excel_typelib_snapshot_odl": "C:\\capture\\snapshots\\excel_typelib_snapshot.odl",
    "excel_pia_identity": "C:\\capture\\raw\\excel_pia_identity.json",
    "excel_pia_public_surface": "C:\\capture\\snapshots\\excel_pia_public_surface.json"
  }
}"#,
    )
    .expect("write manifest");
    fs::write(
        manifest_dir.join("output_checksums.json"),
        r#"{
  "raw/raw_typelib_identity.json": "sha",
  "snapshots/excel_typelib_snapshot.idl": "sha",
  "snapshots/excel_typelib_snapshot.odl": "sha",
  "raw/excel_pia_identity.json": "sha",
  "snapshots/excel_pia_public_surface.json": "sha"
}"#,
    )
    .expect("write checksums");

    let error = normalize_capture_bundle_from_dir(&bundle_root)
        .expect_err("missing checksum-listed payload should fail");
    match error {
        CanonicalOmGenerationError::CaptureBundleContract { message } => {
            assert!(message.contains("raw/excel_pia_identity.json"));
            assert!(message.contains("did not exist"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn normalize_capture_bundle_from_dir_validates_checksum_digests() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root = std::env::temp_dir().join(format!("ootd-step3-bad-checksum-{unique_suffix}"));
    let raw_dir = bundle_root.join("raw");
    let snapshots_dir = bundle_root.join("snapshots");
    let manifest_dir = bundle_root.join("manifest");

    fs::create_dir_all(&raw_dir).expect("raw dir");
    fs::create_dir_all(&snapshots_dir).expect("snapshots dir");
    fs::create_dir_all(&manifest_dir).expect("manifest dir");
    fs::write(
        raw_dir.join("raw_typelib_identity.json"),
        fs::read_to_string(repo_root().join("specs/pinned/raw_typelib_identity.template.json"))
            .expect("typelib template"),
    )
    .expect("write typelib");
    fs::write(
        raw_dir.join("excel_pia_identity.json"),
        r#"{"assembly":"Excel"}"#,
    )
    .expect("write pia identity");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.idl"),
        "library Excel {}",
    )
    .expect("write idl");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.odl"),
        "odl Excel {}",
    )
    .expect("write odl");
    fs::write(
        snapshots_dir.join("excel_pia_public_surface.json"),
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template"),
    )
    .expect("write pia surface");
    fs::write(
        manifest_dir.join("capture_manifest.json"),
        r#"{
  "expectedCaptureOutputs": [
    "raw_typelib_identity.json",
    "excel_typelib_snapshot.idl",
    "excel_typelib_snapshot.odl",
    "excel_pia_identity.json",
    "excel_pia_public_surface.json"
  ],
  "writableOutputs": {
    "raw_typelib_identity": "C:\\capture\\raw\\raw_typelib_identity.json",
    "excel_typelib_snapshot_idl": "C:\\capture\\snapshots\\excel_typelib_snapshot.idl",
    "excel_typelib_snapshot_odl": "C:\\capture\\snapshots\\excel_typelib_snapshot.odl",
    "excel_pia_identity": "C:\\capture\\raw\\excel_pia_identity.json",
    "excel_pia_public_surface": "C:\\capture\\snapshots\\excel_pia_public_surface.json"
  }
}"#,
    )
    .expect("write manifest");
    fs::write(
        manifest_dir.join("output_checksums.json"),
        format!(
            r#"{{
  "raw/raw_typelib_identity.json": "not-a-real-sha",
  "snapshots/excel_typelib_snapshot.idl": "{}",
  "snapshots/excel_typelib_snapshot.odl": "{}",
  "raw/excel_pia_identity.json": "{}",
  "snapshots/excel_pia_public_surface.json": "{}"
}}"#,
            sha256_hex(b"library Excel {}"),
            sha256_hex(b"odl Excel {}"),
            sha256_hex(br#"{"assembly":"Excel"}"#),
            sha256_hex(
                fs::read(snapshots_dir.join("excel_pia_public_surface.json"))
                    .expect("pia")
                    .as_slice()
            )
        ),
    )
    .expect("write checksums");

    let error =
        normalize_capture_bundle_from_dir(&bundle_root).expect_err("checksum mismatch should fail");
    match error {
        CanonicalOmGenerationError::CaptureBundleContract { message } => {
            assert!(message.contains("raw/raw_typelib_identity.json"));
            assert!(message.contains("actual"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn normalize_capture_bundle_from_dir_rejects_non_relative_checksum_paths() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root =
        std::env::temp_dir().join(format!("ootd-step3-absolute-checksum-{unique_suffix}"));
    let raw_dir = bundle_root.join("raw");
    let snapshots_dir = bundle_root.join("snapshots");
    let manifest_dir = bundle_root.join("manifest");

    fs::create_dir_all(&raw_dir).expect("raw dir");
    fs::create_dir_all(&snapshots_dir).expect("snapshots dir");
    fs::create_dir_all(&manifest_dir).expect("manifest dir");
    fs::write(
        raw_dir.join("raw_typelib_identity.json"),
        fs::read_to_string(repo_root().join("specs/pinned/raw_typelib_identity.template.json"))
            .expect("typelib template"),
    )
    .expect("write typelib");
    fs::write(
        raw_dir.join("excel_pia_identity.json"),
        r#"{"assembly":"Excel"}"#,
    )
    .expect("write pia identity");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.idl"),
        "library Excel {}",
    )
    .expect("write idl");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.odl"),
        "odl Excel {}",
    )
    .expect("write odl");
    fs::write(
        snapshots_dir.join("excel_pia_public_surface.json"),
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template"),
    )
    .expect("write pia surface");
    fs::write(
        manifest_dir.join("capture_manifest.json"),
        r#"{
  "expectedCaptureOutputs": [
    "raw_typelib_identity.json",
    "excel_typelib_snapshot.idl",
    "excel_typelib_snapshot.odl",
    "excel_pia_identity.json",
    "excel_pia_public_surface.json"
  ],
  "writableOutputs": {
    "raw_typelib_identity": "C:\\capture\\raw\\raw_typelib_identity.json",
    "excel_typelib_snapshot_idl": "C:\\capture\\snapshots\\excel_typelib_snapshot.idl",
    "excel_typelib_snapshot_odl": "C:\\capture\\snapshots\\excel_typelib_snapshot.odl",
    "excel_pia_identity": "C:\\capture\\raw\\excel_pia_identity.json",
    "excel_pia_public_surface": "C:\\capture\\snapshots\\excel_pia_public_surface.json"
  }
}"#,
    )
    .expect("write manifest");
    fs::write(
        manifest_dir.join("output_checksums.json"),
        format!(
            r#"{{
  "C:\\capture\\raw\\raw_typelib_identity.json": "{}",
  "C:\\capture\\snapshots\\excel_typelib_snapshot.idl": "{}",
  "C:\\capture\\snapshots\\excel_typelib_snapshot.odl": "{}",
  "C:\\capture\\raw\\excel_pia_identity.json": "{}",
  "C:\\capture\\snapshots\\excel_pia_public_surface.json": "{}"
}}"#,
            sha256_hex(
                fs::read(raw_dir.join("raw_typelib_identity.json"))
                    .expect("typelib")
                    .as_slice()
            ),
            sha256_hex(b"library Excel {}"),
            sha256_hex(b"odl Excel {}"),
            sha256_hex(br#"{"assembly":"Excel"}"#),
            sha256_hex(
                fs::read(snapshots_dir.join("excel_pia_public_surface.json"))
                    .expect("pia")
                    .as_slice()
            )
        ),
    )
    .expect("write checksums");

    let error = normalize_capture_bundle_from_dir(&bundle_root)
        .expect_err("absolute checksum path should fail");
    match error {
        CanonicalOmGenerationError::CaptureBundleContract { message } => {
            assert!(message.contains("was not bundle-relative"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn normalize_capture_bundle_from_dir_rejects_duplicate_checksum_payload_names() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root =
        std::env::temp_dir().join(format!("ootd-step3-duplicate-checksum-{unique_suffix}"));
    let raw_dir = bundle_root.join("raw");
    let extra_dir = bundle_root.join("extra");
    let snapshots_dir = bundle_root.join("snapshots");
    let manifest_dir = bundle_root.join("manifest");

    fs::create_dir_all(&raw_dir).expect("raw dir");
    fs::create_dir_all(&extra_dir).expect("extra dir");
    fs::create_dir_all(&snapshots_dir).expect("snapshots dir");
    fs::create_dir_all(&manifest_dir).expect("manifest dir");
    fs::write(
        raw_dir.join("raw_typelib_identity.json"),
        fs::read_to_string(repo_root().join("specs/pinned/raw_typelib_identity.template.json"))
            .expect("typelib template"),
    )
    .expect("write typelib");
    fs::write(
        extra_dir.join("raw_typelib_identity.json"),
        fs::read(raw_dir.join("raw_typelib_identity.json")).expect("typelib"),
    )
    .expect("write duplicate typelib");
    fs::write(
        raw_dir.join("excel_pia_identity.json"),
        r#"{"assembly":"Excel"}"#,
    )
    .expect("write pia identity");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.idl"),
        "library Excel {}",
    )
    .expect("write idl");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.odl"),
        "odl Excel {}",
    )
    .expect("write odl");
    fs::write(
        snapshots_dir.join("excel_pia_public_surface.json"),
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template"),
    )
    .expect("write pia surface");
    fs::write(
        manifest_dir.join("capture_manifest.json"),
        r#"{
  "expectedCaptureOutputs": [
    "raw_typelib_identity.json",
    "excel_typelib_snapshot.idl",
    "excel_typelib_snapshot.odl",
    "excel_pia_identity.json",
    "excel_pia_public_surface.json"
  ],
  "writableOutputs": {
    "raw_typelib_identity": "C:\\capture\\raw\\raw_typelib_identity.json",
    "excel_typelib_snapshot_idl": "C:\\capture\\snapshots\\excel_typelib_snapshot.idl",
    "excel_typelib_snapshot_odl": "C:\\capture\\snapshots\\excel_typelib_snapshot.odl",
    "excel_pia_identity": "C:\\capture\\raw\\excel_pia_identity.json",
    "excel_pia_public_surface": "C:\\capture\\snapshots\\excel_pia_public_surface.json"
  }
}"#,
    )
    .expect("write manifest");
    fs::write(
        manifest_dir.join("output_checksums.json"),
        format!(
            r#"{{
  "raw/raw_typelib_identity.json": "{}",
  "extra/raw_typelib_identity.json": "{}",
  "snapshots/excel_typelib_snapshot.idl": "{}",
  "snapshots/excel_typelib_snapshot.odl": "{}",
  "raw/excel_pia_identity.json": "{}",
  "snapshots/excel_pia_public_surface.json": "{}"
}}"#,
            sha256_hex(
                fs::read(raw_dir.join("raw_typelib_identity.json"))
                    .expect("typelib")
                    .as_slice()
            ),
            sha256_hex(
                fs::read(extra_dir.join("raw_typelib_identity.json"))
                    .expect("duplicate typelib")
                    .as_slice()
            ),
            sha256_hex(b"library Excel {}"),
            sha256_hex(b"odl Excel {}"),
            sha256_hex(br#"{"assembly":"Excel"}"#),
            sha256_hex(
                fs::read(snapshots_dir.join("excel_pia_public_surface.json"))
                    .expect("pia")
                    .as_slice()
            )
        ),
    )
    .expect("write checksums");

    let error = normalize_capture_bundle_from_dir(&bundle_root)
        .expect_err("duplicate checksum payload name should fail");
    match error {
        CanonicalOmGenerationError::CaptureBundleContract { message } => {
            assert!(message.contains("raw_typelib_identity.json"));
            assert!(message.contains("expected exactly one"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn normalize_capture_bundle_from_dir_rejects_wrong_checksum_payload_paths() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root =
        std::env::temp_dir().join(format!("ootd-step3-wrong-checksum-path-{unique_suffix}"));
    let raw_dir = bundle_root.join("raw");
    let extra_dir = bundle_root.join("extra");
    let snapshots_dir = bundle_root.join("snapshots");
    let manifest_dir = bundle_root.join("manifest");

    fs::create_dir_all(&raw_dir).expect("raw dir");
    fs::create_dir_all(&extra_dir).expect("extra dir");
    fs::create_dir_all(&snapshots_dir).expect("snapshots dir");
    fs::create_dir_all(&manifest_dir).expect("manifest dir");
    fs::write(
        raw_dir.join("raw_typelib_identity.json"),
        fs::read_to_string(repo_root().join("specs/pinned/raw_typelib_identity.template.json"))
            .expect("typelib template"),
    )
    .expect("write typelib");
    fs::write(
        extra_dir.join("excel_pia_identity.json"),
        r#"{"assembly":"Excel"}"#,
    )
    .expect("write misplaced pia identity");
    fs::write(
        raw_dir.join("excel_pia_identity.json"),
        r#"{"assembly":"Excel"}"#,
    )
    .expect("write canonical pia identity");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.idl"),
        "library Excel {}",
    )
    .expect("write idl");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.odl"),
        "odl Excel {}",
    )
    .expect("write odl");
    fs::write(
        snapshots_dir.join("excel_pia_public_surface.json"),
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template"),
    )
    .expect("write pia surface");
    fs::write(
        manifest_dir.join("capture_manifest.json"),
        r#"{
  "expectedCaptureOutputs": [
    "raw_typelib_identity.json",
    "excel_typelib_snapshot.idl",
    "excel_typelib_snapshot.odl",
    "excel_pia_identity.json",
    "excel_pia_public_surface.json"
  ],
  "writableOutputs": {
    "raw_typelib_identity": "C:\\capture\\raw\\raw_typelib_identity.json",
    "excel_typelib_snapshot_idl": "C:\\capture\\snapshots\\excel_typelib_snapshot.idl",
    "excel_typelib_snapshot_odl": "C:\\capture\\snapshots\\excel_typelib_snapshot.odl",
    "excel_pia_identity": "C:\\capture\\raw\\excel_pia_identity.json",
    "excel_pia_public_surface": "C:\\capture\\snapshots\\excel_pia_public_surface.json"
  }
}"#,
    )
    .expect("write manifest");
    fs::write(
        manifest_dir.join("output_checksums.json"),
        format!(
            r#"{{
  "raw/raw_typelib_identity.json": "{}",
  "snapshots/excel_typelib_snapshot.idl": "{}",
  "snapshots/excel_typelib_snapshot.odl": "{}",
  "extra/excel_pia_identity.json": "{}",
  "snapshots/excel_pia_public_surface.json": "{}"
}}"#,
            sha256_hex(
                fs::read(raw_dir.join("raw_typelib_identity.json"))
                    .expect("typelib")
                    .as_slice()
            ),
            sha256_hex(b"library Excel {}"),
            sha256_hex(b"odl Excel {}"),
            sha256_hex(br#"{"assembly":"Excel"}"#),
            sha256_hex(
                fs::read(snapshots_dir.join("excel_pia_public_surface.json"))
                    .expect("pia")
                    .as_slice()
            )
        ),
    )
    .expect("write checksums");

    let error = normalize_capture_bundle_from_dir(&bundle_root)
        .expect_err("wrong checksum payload path should fail");
    match error {
        CanonicalOmGenerationError::CaptureBundleContract { message } => {
            assert!(message.contains("extra/excel_pia_identity.json"));
            assert!(message.contains("raw/excel_pia_identity.json"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn normalize_capture_bundle_from_dir_rejects_unexpected_checksum_payload_names() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root =
        std::env::temp_dir().join(format!("ootd-step3-extra-checksum-payload-{unique_suffix}"));
    let raw_dir = bundle_root.join("raw");
    let extra_dir = bundle_root.join("extra");
    let snapshots_dir = bundle_root.join("snapshots");
    let manifest_dir = bundle_root.join("manifest");

    fs::create_dir_all(&raw_dir).expect("raw dir");
    fs::create_dir_all(&extra_dir).expect("extra dir");
    fs::create_dir_all(&snapshots_dir).expect("snapshots dir");
    fs::create_dir_all(&manifest_dir).expect("manifest dir");
    fs::write(
        raw_dir.join("raw_typelib_identity.json"),
        fs::read_to_string(repo_root().join("specs/pinned/raw_typelib_identity.template.json"))
            .expect("typelib template"),
    )
    .expect("write typelib");
    fs::write(
        raw_dir.join("excel_pia_identity.json"),
        r#"{"assembly":"Excel"}"#,
    )
    .expect("write pia identity");
    fs::write(
        extra_dir.join("unplanned_payload.json"),
        r#"{"extra":true}"#,
    )
    .expect("write extra payload");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.idl"),
        "library Excel {}",
    )
    .expect("write idl");
    fs::write(
        snapshots_dir.join("excel_typelib_snapshot.odl"),
        "odl Excel {}",
    )
    .expect("write odl");
    fs::write(
        snapshots_dir.join("excel_pia_public_surface.json"),
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template"),
    )
    .expect("write pia surface");
    fs::write(
        manifest_dir.join("capture_manifest.json"),
        r#"{
  "expectedCaptureOutputs": [
    "raw_typelib_identity.json",
    "excel_typelib_snapshot.idl",
    "excel_typelib_snapshot.odl",
    "excel_pia_identity.json",
    "excel_pia_public_surface.json"
  ],
  "writableOutputs": {
    "raw_typelib_identity": "C:\\capture\\raw\\raw_typelib_identity.json",
    "excel_typelib_snapshot_idl": "C:\\capture\\snapshots\\excel_typelib_snapshot.idl",
    "excel_typelib_snapshot_odl": "C:\\capture\\snapshots\\excel_typelib_snapshot.odl",
    "excel_pia_identity": "C:\\capture\\raw\\excel_pia_identity.json",
    "excel_pia_public_surface": "C:\\capture\\snapshots\\excel_pia_public_surface.json"
  }
}"#,
    )
    .expect("write manifest");
    fs::write(
        manifest_dir.join("output_checksums.json"),
        format!(
            r#"{{
  "raw/raw_typelib_identity.json": "{}",
  "snapshots/excel_typelib_snapshot.idl": "{}",
  "snapshots/excel_typelib_snapshot.odl": "{}",
  "raw/excel_pia_identity.json": "{}",
  "snapshots/excel_pia_public_surface.json": "{}",
  "extra/unplanned_payload.json": "{}"
}}"#,
            sha256_hex(
                fs::read(raw_dir.join("raw_typelib_identity.json"))
                    .expect("typelib")
                    .as_slice()
            ),
            sha256_hex(b"library Excel {}"),
            sha256_hex(b"odl Excel {}"),
            sha256_hex(br#"{"assembly":"Excel"}"#),
            sha256_hex(
                fs::read(snapshots_dir.join("excel_pia_public_surface.json"))
                    .expect("pia")
                    .as_slice()
            ),
            sha256_hex(br#"{"extra":true}"#)
        ),
    )
    .expect("write checksums");

    let error = normalize_capture_bundle_from_dir(&bundle_root)
        .expect_err("unexpected checksum payload should fail");
    match error {
        CanonicalOmGenerationError::CaptureBundleContract { message } => {
            assert!(message.contains("unplanned_payload.json"));
            assert!(message.contains("outside allowed"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn writes_canonical_office_idl_json_from_bundle_inputs() {
    let unique_suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let bundle_root = std::env::temp_dir().join(format!("ootd-step3-{unique_suffix}"));
    let raw_dir = bundle_root.join("raw");
    let snapshots_dir = bundle_root.join("snapshots");
    let manifest_dir = bundle_root.join("manifest");

    fs::create_dir_all(&raw_dir).expect("raw dir");
    fs::create_dir_all(&snapshots_dir).expect("snapshots dir");
    fs::create_dir_all(&manifest_dir).expect("manifest dir");

    let typelib_template =
        fs::read_to_string(repo_root().join("specs/pinned/raw_typelib_identity.template.json"))
            .expect("typelib template");
    let pia_template =
        fs::read_to_string(repo_root().join("specs/pinned/excel_pia_public_surface.template.json"))
            .expect("pia template");

    fs::write(raw_dir.join("raw_typelib_identity.json"), typelib_template).expect("write typelib");
    fs::write(
        snapshots_dir.join("excel_pia_public_surface.json"),
        pia_template,
    )
    .expect("write pia surface");

    let output_path = manifest_dir.join("office_idl_excel_om.json");
    let generation = generate_canonical_office_idl_from_dir(&bundle_root, &output_path)
        .expect("generate canonical document");

    let round_trip = OfficeIdlDocument::from_path(&output_path).expect("round trip");
    let round_trip_summary = CodegenSummary::from_document(&round_trip);

    let template = OfficeIdlDocument::from_path(
        repo_root().join("specs/pinned/office_idl_excel_om.template.json"),
    )
    .expect("template json");

    assert_eq!(generation.output_path, output_path);
    assert_eq!(generation.bundle_paths.bundle_root_path, bundle_root);
    assert_eq!(
        generation.bundle_paths.raw_typelib_identity_path,
        raw_dir.join("raw_typelib_identity.json")
    );
    assert_eq!(
        generation.bundle_paths.excel_pia_public_surface_path,
        snapshots_dir.join("excel_pia_public_surface.json")
    );
    assert_eq!(
        generation.bundle_paths.capture_manifest_path,
        manifest_dir.join("capture_manifest.json")
    );
    assert_eq!(
        generation.bundle_paths.output_checksums_path,
        manifest_dir.join("output_checksums.json")
    );
    assert_eq!(round_trip_summary.enum_count, 1);
    assert_eq!(round_trip_summary.interface_count, 6);
    assert_eq!(round_trip_summary.class_count, 3);
    assert_eq!(round_trip_summary.member_count, 64);
    assert_eq!(generation.summary.library, "Excel");
    assert_eq!(generation.summary.version, "16.0");
    assert_eq!(
        generation.summary.type_library_guid,
        "{00020813-0000-0000-C000-000000000046}"
    );
    assert_eq!(generation.summary.interface_iid_count, 6);
    assert_eq!(generation.summary.coclass_clsid_count, 3);
    assert!(generation.summary.missing_pia_interfaces.is_empty());
    assert!(generation.summary.missing_pia_classes.is_empty());
    assert_eq!(
        round_trip
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.namespace.as_deref()),
        Some("Microsoft.Office.Interop.Excel")
    );
    assert_eq!(
        round_trip
            .metadata
            .as_ref()
            .and_then(|metadata| metadata.type_library_guid.as_deref()),
        Some("{00020813-0000-0000-C000-000000000046}")
    );
    assert_eq!(round_trip.library, template.library);
    assert_eq!(round_trip.version, template.version);
    for interface in &round_trip.interfaces {
        assert!(
            template
                .interfaces
                .iter()
                .any(|template_interface| template_interface.name == interface.name),
            "{} should be present in the pinned template",
            interface.name
        );
    }
    assert_eq!(round_trip.classes.len(), template.classes.len());
}

#[test]
fn summarizes_focus_surface_registry_and_coverage_from_template_document() {
    let document = OfficeIdlDocument::from_path(
        repo_root().join("specs/pinned/office_idl_excel_om.template.json"),
    )
    .expect("template json");
    let registry = build_focus_surface_registry(&document);
    let coverage = build_coverage_report(&document);

    assert_eq!(registry.library, "Excel");
    assert_eq!(registry.version, "16.0");
    assert_eq!(registry.focus_surfaces.len(), 58);
    assert!(registry.missing_focus_surfaces.is_empty());

    let application = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Application")
        .expect("Application");
    let worksheet_function = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "WorksheetFunction")
        .expect("WorksheetFunction");
    let workbooks = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Workbooks")
        .expect("Workbooks");
    let workbook = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Workbook")
        .expect("Workbook");
    let worksheets = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Worksheets")
        .expect("Worksheets");
    let sheets = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Sheets")
        .expect("Sheets");
    let charts = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Charts")
        .expect("Charts");
    let worksheet = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Worksheet")
        .expect("Worksheet");
    let range = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Range")
        .expect("Range");
    let areas = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Areas")
        .expect("Areas");
    let names = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Names")
        .expect("Names");
    let name = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Name")
        .expect("Name");
    let chart_objects = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartObjects")
        .expect("ChartObjects");
    let chart_object = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartObject")
        .expect("ChartObject");
    let shape_range = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ShapeRange")
        .expect("ShapeRange");
    let chart = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Chart")
        .expect("Chart");
    let chart_area = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartArea")
        .expect("ChartArea");
    let plot_area = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "PlotArea")
        .expect("PlotArea");
    let chart_title = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartTitle")
        .expect("ChartTitle");
    let legend = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Legend")
        .expect("Legend");
    let legend_entries = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "LegendEntries")
        .expect("LegendEntries");
    let legend_entry = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "LegendEntry")
        .expect("LegendEntry");
    let legend_key = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "LegendKey")
        .expect("LegendKey");
    let data_table = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "DataTable")
        .expect("DataTable");
    let chart_format = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartFormat")
        .expect("ChartFormat");
    let adjustments = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Adjustments")
        .expect("Adjustments");
    let fill_format = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "FillFormat")
        .expect("FillFormat");
    let glow_format = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "GlowFormat")
        .expect("GlowFormat");
    let line_format = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "LineFormat")
        .expect("LineFormat");
    let picture_format = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "PictureFormat")
        .expect("PictureFormat");
    let crop = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Crop")
        .expect("Crop");
    let shadow_format = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ShadowFormat")
        .expect("ShadowFormat");
    let soft_edge_format = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "SoftEdgeFormat")
        .expect("SoftEdgeFormat");
    let text_frame2 = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "TextFrame2")
        .expect("TextFrame2");
    let three_d_format = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ThreeDFormat")
        .expect("ThreeDFormat");
    let chart_groups = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartGroups")
        .expect("ChartGroups");
    let chart_group = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartGroup")
        .expect("ChartGroup");
    let category_collection = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "CategoryCollection")
        .expect("CategoryCollection");
    let chart_category = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartCategory")
        .expect("ChartCategory");
    let series_lines = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "SeriesLines")
        .expect("SeriesLines");
    let drop_lines = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "DropLines")
        .expect("DropLines");
    let hi_lo_lines = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "HiLoLines")
        .expect("HiLoLines");
    let up_bars = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "UpBars")
        .expect("UpBars");
    let down_bars = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "DownBars")
        .expect("DownBars");
    let axes = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Axes")
        .expect("Axes");
    let axis = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Axis")
        .expect("Axis");
    let tick_labels = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "TickLabels")
        .expect("TickLabels");
    let gridlines = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Gridlines")
        .expect("Gridlines");
    let display_unit_label = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "DisplayUnitLabel")
        .expect("DisplayUnitLabel");
    let axis_title = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "AxisTitle")
        .expect("AxisTitle");
    let series_collection = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "SeriesCollection")
        .expect("SeriesCollection");
    let series = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Series")
        .expect("Series");
    let leader_lines = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "LeaderLines")
        .expect("LeaderLines");
    let border = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Border")
        .expect("Border");
    let data_labels = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "DataLabels")
        .expect("DataLabels");
    let data_label = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "DataLabel")
        .expect("DataLabel");
    let points = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Points")
        .expect("Points");
    let point = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Point")
        .expect("Point");

    assert_eq!(application.member_count, 44);
    assert_eq!(worksheet_function.member_count, 526);
    assert_eq!(workbooks.member_count, 7);
    assert_eq!(workbook.member_count, 26);
    assert_eq!(worksheets.member_count, 13);
    assert_eq!(sheets.member_count, 13);
    assert_eq!(charts.member_count, 13);
    assert_eq!(worksheet.member_count, 28);
    assert_eq!(range.member_count, 51);
    assert_eq!(areas.member_count, 5);
    assert_eq!(names.member_count, 5);
    assert_eq!(name.member_count, 6);
    assert_eq!(chart_objects.member_count, 25);
    assert_eq!(chart_object.member_count, 35);
    assert_eq!(shape_range.member_count, 46);
    assert_eq!(chart.member_count, 80);
    assert_eq!(chart_area.member_count, 16);
    assert_eq!(plot_area.member_count, 19);
    assert_eq!(chart_title.member_count, 16);
    assert_eq!(legend.member_count, 16);
    assert_eq!(legend_entries.member_count, 5);
    assert_eq!(legend_entry.member_count, 11);
    assert_eq!(legend_key.member_count, 9);
    assert_eq!(data_table.member_count, 11);
    assert_eq!(chart_format.member_count, 13);
    assert_eq!(adjustments.member_count, 5);
    assert_eq!(fill_format.member_count, 5);
    assert_eq!(glow_format.member_count, 6);
    assert_eq!(line_format.member_count, 15);
    assert_eq!(picture_format.member_count, 15);
    assert_eq!(crop.member_count, 11);
    assert_eq!(shadow_format.member_count, 10);
    assert_eq!(soft_edge_format.member_count, 5);
    assert_eq!(text_frame2.member_count, 14);
    assert_eq!(three_d_format.member_count, 18);
    assert_eq!(chart_groups.member_count, 5);
    assert_eq!(chart_group.member_count, 33);
    assert_eq!(category_collection.member_count, 5);
    assert_eq!(chart_category.member_count, 6);
    assert_eq!(series_lines.member_count, 10);
    assert_eq!(drop_lines.member_count, 10);
    assert_eq!(hi_lo_lines.member_count, 10);
    assert_eq!(up_bars.member_count, 10);
    assert_eq!(down_bars.member_count, 10);
    assert_eq!(axes.member_count, 5);
    assert_eq!(axis.member_count, 45);
    assert_eq!(tick_labels.member_count, 16);
    assert_eq!(gridlines.member_count, 8);
    assert_eq!(display_unit_label.member_count, 16);
    assert_eq!(axis_title.member_count, 16);
    assert_eq!(series_collection.member_count, 7);
    assert_eq!(series.member_count, 28);
    assert_eq!(leader_lines.member_count, 7);
    assert_eq!(border.member_count, 9);
    assert_eq!(data_labels.member_count, 27);
    assert_eq!(data_label.member_count, 25);
    assert_eq!(points.member_count, 5);
    assert_eq!(point.member_count, 14);
    assert_eq!(
        application.default_coclasses,
        vec!["Application".to_string()]
    );
    assert_eq!(workbook.default_coclasses, vec!["Workbook".to_string()]);
    assert_eq!(worksheet.default_coclasses, vec!["Worksheet".to_string()]);

    let application_workbooks = application
        .members
        .iter()
        .find(|member| member.name == "Workbooks")
        .expect("Application.Workbooks");
    assert_eq!(application_workbooks.disp_id, Some(572));
    assert_eq!(application_workbooks.access, AccessMode::Read);
    assert_eq!(
        application_workbooks.capture_origin_kinds,
        vec![CaptureOriginKind::PropertyGet]
    );
    let application_active_cell = application
        .members
        .iter()
        .find(|member| member.name == "ActiveCell")
        .expect("Application.ActiveCell");
    assert_eq!(application_active_cell.access, AccessMode::Read);
    assert_eq!(
        application_active_cell
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let application_selection = application
        .members
        .iter()
        .find(|member| member.name == "Selection")
        .expect("Application.Selection");
    assert_eq!(application_selection.access, AccessMode::Read);
    assert_eq!(
        application_selection
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let application_user_name = application
        .members
        .iter()
        .find(|member| member.name == "UserName")
        .expect("Application.UserName");
    assert_eq!(application_user_name.access, AccessMode::Readwrite);
    assert_eq!(
        application_user_name.capture_origin_kinds,
        vec![
            CaptureOriginKind::PropertyGet,
            CaptureOriginKind::PropertySet
        ]
    );
    assert_eq!(
        application_user_name
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("BSTR")
    );
    let application_default_file_path = application
        .members
        .iter()
        .find(|member| member.name == "DefaultFilePath")
        .expect("Application.DefaultFilePath");
    assert_eq!(application_default_file_path.access, AccessMode::Readwrite);
    assert_eq!(
        application_default_file_path.capture_origin_kinds,
        vec![
            CaptureOriginKind::PropertyGet,
            CaptureOriginKind::PropertySet
        ]
    );
    assert_eq!(
        application_default_file_path
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("BSTR")
    );
    let application_caption = application
        .members
        .iter()
        .find(|member| member.name == "Caption")
        .expect("Application.Caption");
    assert_eq!(application_caption.access, AccessMode::Readwrite);
    assert_eq!(
        application_caption
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("BSTR")
    );
    let application_display_formula_bar = application
        .members
        .iter()
        .find(|member| member.name == "DisplayFormulaBar")
        .expect("Application.DisplayFormulaBar");
    assert_eq!(
        application_display_formula_bar.access,
        AccessMode::Readwrite
    );
    assert_eq!(
        application_display_formula_bar
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT_BOOL")
    );
    let application_display_scroll_bars = application
        .members
        .iter()
        .find(|member| member.name == "DisplayScrollBars")
        .expect("Application.DisplayScrollBars");
    assert_eq!(
        application_display_scroll_bars.access,
        AccessMode::Readwrite
    );
    assert_eq!(
        application_display_scroll_bars
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT_BOOL")
    );
    let application_display_full_screen = application
        .members
        .iter()
        .find(|member| member.name == "DisplayFullScreen")
        .expect("Application.DisplayFullScreen");
    assert_eq!(
        application_display_full_screen.access,
        AccessMode::Readwrite
    );
    assert_eq!(
        application_display_full_screen
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT_BOOL")
    );
    let application_use_system_separators = application
        .members
        .iter()
        .find(|member| member.name == "UseSystemSeparators")
        .expect("Application.UseSystemSeparators");
    assert_eq!(
        application_use_system_separators.access,
        AccessMode::Readwrite
    );
    assert_eq!(
        application_use_system_separators
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT_BOOL")
    );
    let application_decimal_separator = application
        .members
        .iter()
        .find(|member| member.name == "DecimalSeparator")
        .expect("Application.DecimalSeparator");
    assert_eq!(application_decimal_separator.access, AccessMode::Readwrite);
    assert_eq!(
        application_decimal_separator
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("BSTR")
    );
    let application_thousands_separator = application
        .members
        .iter()
        .find(|member| member.name == "ThousandsSeparator")
        .expect("Application.ThousandsSeparator");
    assert_eq!(
        application_thousands_separator.access,
        AccessMode::Readwrite
    );
    assert_eq!(
        application_thousands_separator
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("BSTR")
    );
    let application_international = application
        .members
        .iter()
        .find(|member| member.name == "International")
        .expect("Application.International");
    assert_eq!(application_international.access, AccessMode::Read);
    assert_eq!(application_international.params.len(), 1);
    assert_eq!(
        application_international
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT")
    );
    let application_show_windows_in_taskbar = application
        .members
        .iter()
        .find(|member| member.name == "ShowWindowsInTaskbar")
        .expect("Application.ShowWindowsInTaskbar");
    assert_eq!(
        application_show_windows_in_taskbar.access,
        AccessMode::Readwrite
    );
    assert_eq!(
        application_show_windows_in_taskbar
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT_BOOL")
    );
    let application_interactive = application
        .members
        .iter()
        .find(|member| member.name == "Interactive")
        .expect("Application.Interactive");
    assert_eq!(application_interactive.access, AccessMode::Readwrite);
    assert_eq!(
        application_interactive
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT_BOOL")
    );

    let worksheet_name = worksheet
        .members
        .iter()
        .find(|member| member.name == "Name")
        .expect("Worksheet.Name");
    assert_eq!(worksheet_name.disp_id, Some(110));
    assert_eq!(worksheet_name.access, AccessMode::Readwrite);
    assert_eq!(
        worksheet_name.capture_origin_kinds,
        vec![
            CaptureOriginKind::PropertyGet,
            CaptureOriginKind::PropertySet
        ]
    );

    let worksheet_range = worksheet
        .members
        .iter()
        .find(|member| member.name == "Range")
        .expect("Worksheet.Range");
    assert_eq!(worksheet_range.disp_id, Some(197));
    assert_eq!(worksheet_range.access, AccessMode::Read);
    assert_eq!(
        worksheet_range
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let workbook_name = workbook
        .members
        .iter()
        .find(|member| member.name == "Name")
        .expect("Workbook.Name");
    assert_eq!(workbook_name.access, AccessMode::Read);
    let workbook_parent = workbook
        .members
        .iter()
        .find(|member| member.name == "Parent")
        .expect("Workbook.Parent");
    assert_eq!(workbook_parent.access, AccessMode::Read);
    assert_eq!(
        workbook_parent
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Application")
    );
    let workbook_path = workbook
        .members
        .iter()
        .find(|member| member.name == "Path")
        .expect("Workbook.Path");
    assert_eq!(workbook_path.access, AccessMode::Read);
    let workbook_full_name = workbook
        .members
        .iter()
        .find(|member| member.name == "FullName")
        .expect("Workbook.FullName");
    assert_eq!(workbook_full_name.access, AccessMode::Read);
    let workbook_date1904 = workbook
        .members
        .iter()
        .find(|member| member.name == "Date1904")
        .expect("Workbook.Date1904");
    assert_eq!(workbook_date1904.access, AccessMode::Readwrite);
    assert_eq!(
        workbook_date1904.capture_origin_kinds,
        vec![
            CaptureOriginKind::PropertyGet,
            CaptureOriginKind::PropertySet
        ]
    );
    assert_eq!(
        workbook_date1904
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT_BOOL")
    );
    let workbook_is_addin = workbook
        .members
        .iter()
        .find(|member| member.name == "IsAddin")
        .expect("Workbook.IsAddin");
    assert_eq!(workbook_is_addin.access, AccessMode::Readwrite);
    assert_eq!(
        workbook_is_addin.capture_origin_kinds,
        vec![
            CaptureOriginKind::PropertyGet,
            CaptureOriginKind::PropertySet
        ]
    );
    assert_eq!(
        workbook_is_addin
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT_BOOL")
    );
    let workbook_read_only = workbook
        .members
        .iter()
        .find(|member| member.name == "ReadOnly")
        .expect("Workbook.ReadOnly");
    assert_eq!(workbook_read_only.access, AccessMode::Read);
    let workbook_saved = workbook
        .members
        .iter()
        .find(|member| member.name == "Saved")
        .expect("Workbook.Saved");
    assert_eq!(workbook_saved.access, AccessMode::Readwrite);
    let workbook_save_as = workbook
        .members
        .iter()
        .find(|member| member.name == "SaveAs")
        .expect("Workbook.SaveAs");
    assert_eq!(workbook_save_as.access, AccessMode::Read);
    let expected_save_as_names = [
        "Filename",
        "FileFormat",
        "Password",
        "WriteResPassword",
        "ReadOnlyRecommended",
        "CreateBackup",
        "AccessMode",
        "ConflictResolution",
        "AddToMru",
        "TextCodepage",
        "TextVisualLayout",
        "Local",
    ];
    assert_eq!(workbook_save_as.params.len(), expected_save_as_names.len());
    for (index, expected_name) in expected_save_as_names.iter().enumerate() {
        assert_eq!(workbook_save_as.params[index].name, *expected_name);
        assert_eq!(workbook_save_as.params[index].optional, index != 0);
    }
    let workbook_refresh_all = workbook
        .members
        .iter()
        .find(|member| member.name == "RefreshAll")
        .expect("Workbook.RefreshAll");
    assert_eq!(workbook_refresh_all.access, AccessMode::Read);
    assert_eq!(workbook_refresh_all.params.len(), 0);
    assert!(workbook_refresh_all.return_type.is_none());
    let workbook_close = workbook
        .members
        .iter()
        .find(|member| member.name == "Close")
        .expect("Workbook.Close");
    assert_eq!(workbook_close.access, AccessMode::Read);
    assert_eq!(workbook_close.params.len(), 3);
    assert_eq!(workbook_close.params[0].name, "SaveChanges");
    assert!(workbook_close.params[0].optional);
    assert_eq!(workbook_close.params[1].name, "Filename");
    assert!(workbook_close.params[1].optional);
    assert_eq!(workbook_close.params[2].name, "RouteWorkbook");
    assert!(workbook_close.params[2].optional);
    let application_worksheets = application
        .members
        .iter()
        .find(|member| member.name == "Worksheets")
        .expect("Application.Worksheets");
    assert_eq!(application_worksheets.access, AccessMode::Read);
    assert_eq!(
        application_worksheets
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Worksheets")
    );
    let worksheet_parent = worksheet
        .members
        .iter()
        .find(|member| member.name == "Parent")
        .expect("Worksheet.Parent");
    assert_eq!(worksheet_parent.access, AccessMode::Read);
    assert_eq!(
        worksheet_parent
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Workbook")
    );
    let worksheet_index = worksheet
        .members
        .iter()
        .find(|member| member.name == "Index")
        .expect("Worksheet.Index");
    assert_eq!(worksheet_index.access, AccessMode::Read);
    let application_calculate = application
        .members
        .iter()
        .find(|member| member.name == "Calculate")
        .expect("Application.Calculate");
    assert_eq!(application_calculate.access, AccessMode::Read);
    assert_eq!(application_calculate.params.len(), 0);
    assert!(application_calculate.return_type.is_none());
    let application_calculate_full = application
        .members
        .iter()
        .find(|member| member.name == "CalculateFull")
        .expect("Application.CalculateFull");
    assert_eq!(application_calculate_full.access, AccessMode::Read);
    assert_eq!(application_calculate_full.params.len(), 0);
    assert!(application_calculate_full.return_type.is_none());
    let application_calculate_full_rebuild = application
        .members
        .iter()
        .find(|member| member.name == "CalculateFullRebuild")
        .expect("Application.CalculateFullRebuild");
    assert_eq!(application_calculate_full_rebuild.access, AccessMode::Read);
    assert_eq!(application_calculate_full_rebuild.params.len(), 0);
    let application_goto = application
        .members
        .iter()
        .find(|member| member.name == "Goto")
        .expect("Application.Goto");
    assert_eq!(application_goto.access, AccessMode::Read);
    assert_eq!(application_goto.params.len(), 2);
    assert!(application_goto.params[0].optional);
    assert!(application_goto.params[1].optional);
    assert!(application_goto.return_type.is_none());
    let application_range = application
        .members
        .iter()
        .find(|member| member.name == "Range")
        .expect("Application.Range");
    assert_eq!(application_range.access, AccessMode::Read);
    assert_eq!(application_range.params.len(), 2);
    assert!(application_range.params[1].optional);
    assert_eq!(
        application_range
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let application_cells = application
        .members
        .iter()
        .find(|member| member.name == "Cells")
        .expect("Application.Cells");
    assert_eq!(application_cells.access, AccessMode::Read);
    assert_eq!(application_cells.params.len(), 0);
    assert_eq!(
        application_cells
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let application_rows = application
        .members
        .iter()
        .find(|member| member.name == "Rows")
        .expect("Application.Rows");
    assert_eq!(application_rows.access, AccessMode::Read);
    assert_eq!(
        application_rows
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let application_columns = application
        .members
        .iter()
        .find(|member| member.name == "Columns")
        .expect("Application.Columns");
    assert_eq!(application_columns.access, AccessMode::Read);
    assert_eq!(
        application_columns
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let application_intersect = application
        .members
        .iter()
        .find(|member| member.name == "Intersect")
        .expect("Application.Intersect");
    assert_eq!(application_intersect.access, AccessMode::Read);
    assert_eq!(application_intersect.params.len(), 30);
    for (index, param) in application_intersect.params.iter().enumerate() {
        assert_eq!(param.name, format!("Arg{}", index + 1));
        assert_eq!(param.optional, index >= 2);
    }
    assert_eq!(
        application_intersect
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let application_union = application
        .members
        .iter()
        .find(|member| member.name == "Union")
        .expect("Application.Union");
    assert_eq!(application_union.access, AccessMode::Read);
    assert_eq!(application_union.params.len(), 30);
    for (index, param) in application_union.params.iter().enumerate() {
        assert_eq!(param.name, format!("Arg{}", index + 1));
        assert_eq!(param.optional, index >= 2);
    }
    assert_eq!(
        application_union
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let worksheet_activate = worksheet
        .members
        .iter()
        .find(|member| member.name == "Activate")
        .expect("Worksheet.Activate");
    assert_eq!(worksheet_activate.access, AccessMode::Read);
    assert_eq!(worksheet_activate.params.len(), 0);
    assert!(worksheet_activate.return_type.is_none());
    let worksheet_calculate = worksheet
        .members
        .iter()
        .find(|member| member.name == "Calculate")
        .expect("Worksheet.Calculate");
    assert_eq!(worksheet_calculate.access, AccessMode::Read);
    assert_eq!(worksheet_calculate.params.len(), 0);
    assert!(worksheet_calculate.return_type.is_none());
    let worksheet_rows = worksheet
        .members
        .iter()
        .find(|member| member.name == "Rows")
        .expect("Worksheet.Rows");
    assert_eq!(worksheet_rows.access, AccessMode::Read);
    assert_eq!(
        worksheet_rows
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let worksheet_columns = worksheet
        .members
        .iter()
        .find(|member| member.name == "Columns")
        .expect("Worksheet.Columns");
    assert_eq!(worksheet_columns.access, AccessMode::Read);
    assert_eq!(
        worksheet_columns
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let range_value2 = range
        .members
        .iter()
        .find(|member| member.name == "Value2")
        .expect("Range.Value2");
    assert_eq!(range_value2.access, AccessMode::Readwrite);
    let range_value = range
        .members
        .iter()
        .find(|member| member.name == "Value")
        .expect("Range.Value");
    assert_eq!(range_value.access, AccessMode::Readwrite);
    let range_formula = range
        .members
        .iter()
        .find(|member| member.name == "Formula")
        .expect("Range.Formula");
    assert_eq!(range_formula.access, AccessMode::Readwrite);
    let range_formula_r1c1 = range
        .members
        .iter()
        .find(|member| member.name == "FormulaR1C1")
        .expect("Range.FormulaR1C1");
    assert_eq!(range_formula_r1c1.access, AccessMode::Readwrite);
    let range_formula_local = range
        .members
        .iter()
        .find(|member| member.name == "FormulaLocal")
        .expect("Range.FormulaLocal");
    assert_eq!(range_formula_local.access, AccessMode::Readwrite);
    let range_formula_r1c1_local = range
        .members
        .iter()
        .find(|member| member.name == "FormulaR1C1Local")
        .expect("Range.FormulaR1C1Local");
    assert_eq!(range_formula_r1c1_local.access, AccessMode::Readwrite);
    let range_formula2 = range
        .members
        .iter()
        .find(|member| member.name == "Formula2")
        .expect("Range.Formula2");
    assert_eq!(range_formula2.access, AccessMode::Readwrite);
    let range_formula2_r1c1 = range
        .members
        .iter()
        .find(|member| member.name == "Formula2R1C1")
        .expect("Range.Formula2R1C1");
    assert_eq!(range_formula2_r1c1.access, AccessMode::Readwrite);
    let range_formula2_local = range
        .members
        .iter()
        .find(|member| member.name == "Formula2Local")
        .expect("Range.Formula2Local");
    assert_eq!(range_formula2_local.access, AccessMode::Readwrite);
    let range_formula2_r1c1_local = range
        .members
        .iter()
        .find(|member| member.name == "Formula2R1C1Local")
        .expect("Range.Formula2R1C1Local");
    assert_eq!(range_formula2_r1c1_local.access, AccessMode::Readwrite);
    let range_text = range
        .members
        .iter()
        .find(|member| member.name == "Text")
        .expect("Range.Text");
    assert_eq!(range_text.access, AccessMode::Read);
    assert_eq!(
        range_text
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("BSTR")
    );
    let range_has_formula = range
        .members
        .iter()
        .find(|member| member.name == "HasFormula")
        .expect("Range.HasFormula");
    assert_eq!(range_has_formula.access, AccessMode::Read);
    assert_eq!(
        range_has_formula
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT")
    );
    let range_count = range
        .members
        .iter()
        .find(|member| member.name == "Count")
        .expect("Range.Count");
    assert_eq!(range_count.access, AccessMode::Read);
    let range_current_region = range
        .members
        .iter()
        .find(|member| member.name == "CurrentRegion")
        .expect("Range.CurrentRegion");
    assert_eq!(range_current_region.access, AccessMode::Read);
    assert_eq!(
        range_current_region
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let range_entire_row = range
        .members
        .iter()
        .find(|member| member.name == "EntireRow")
        .expect("Range.EntireRow");
    assert_eq!(range_entire_row.access, AccessMode::Read);
    assert_eq!(
        range_entire_row
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let range_entire_column = range
        .members
        .iter()
        .find(|member| member.name == "EntireColumn")
        .expect("Range.EntireColumn");
    assert_eq!(range_entire_column.access, AccessMode::Read);
    assert_eq!(
        range_entire_column
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let range_cells = range
        .members
        .iter()
        .find(|member| member.name == "Cells")
        .expect("Range.Cells");
    assert_eq!(range_cells.access, AccessMode::Read);
    assert_eq!(
        range_cells
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let range_rows = range
        .members
        .iter()
        .find(|member| member.name == "Rows")
        .expect("Range.Rows");
    assert_eq!(range_rows.access, AccessMode::Read);
    assert_eq!(
        range_rows
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let range_columns = range
        .members
        .iter()
        .find(|member| member.name == "Columns")
        .expect("Range.Columns");
    assert_eq!(range_columns.access, AccessMode::Read);
    assert_eq!(
        range_columns
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let range_item = range
        .members
        .iter()
        .find(|member| member.name == "Item")
        .expect("Range.Item");
    assert_eq!(range_item.access, AccessMode::Read);
    assert_eq!(range_item.params.len(), 2);
    assert!(range_item.params[1].optional);
    assert_eq!(
        range_item
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let range_address = range
        .members
        .iter()
        .find(|member| member.name == "Address")
        .expect("Range.Address");
    assert_eq!(range_address.access, AccessMode::Read);
    let expected_address_names = [
        "RowAbsolute",
        "ColumnAbsolute",
        "ReferenceStyle",
        "External",
        "RelativeTo",
    ];
    assert_eq!(range_address.params.len(), expected_address_names.len());
    for (index, expected_name) in expected_address_names.iter().enumerate() {
        assert_eq!(range_address.params[index].name, *expected_name);
        assert!(range_address.params[index].optional);
    }
    assert_eq!(
        range_address
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("BSTR")
    );
    let range_offset = range
        .members
        .iter()
        .find(|member| member.name == "Offset")
        .expect("Range.Offset");
    assert_eq!(range_offset.access, AccessMode::Read);
    assert_eq!(range_offset.params.len(), 2);
    assert!(range_offset.params[0].optional);
    assert!(range_offset.params[1].optional);
    assert_eq!(
        range_offset
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let range_resize = range
        .members
        .iter()
        .find(|member| member.name == "Resize")
        .expect("Range.Resize");
    assert_eq!(range_resize.access, AccessMode::Read);
    assert_eq!(range_resize.params.len(), 2);
    assert!(range_resize.params[0].optional);
    assert!(range_resize.params[1].optional);
    assert_eq!(
        range_resize
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let range_end = range
        .members
        .iter()
        .find(|member| member.name == "End")
        .expect("Range.End");
    assert_eq!(range_end.access, AccessMode::Read);
    assert_eq!(range_end.params.len(), 1);
    assert_eq!(range_end.params[0].name, "Direction");
    assert_eq!(
        range_end
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Range")
    );
    let range_find = range
        .members
        .iter()
        .find(|member| member.name == "Find")
        .expect("Range.Find");
    assert_eq!(range_find.access, AccessMode::Read);
    assert_eq!(range_find.params.len(), 9);
    assert_eq!(range_find.params[0].name, "What");
    assert!(!range_find.params[0].optional);
    assert_eq!(range_find.params[8].name, "SearchFormat");
    assert!(range_find.params[8].optional);
    let range_find_return = range_find.return_type.as_ref().expect("Range.Find return");
    assert_eq!(range_find_return.alias_of.as_deref(), Some("Excel.Range"));
    assert!(range_find_return.nullable);
    for member_name in ["FindNext", "FindPrevious"] {
        let find_member = range
            .members
            .iter()
            .find(|member| member.name == member_name)
            .expect(member_name);
        assert_eq!(find_member.access, AccessMode::Read);
        assert_eq!(find_member.params.len(), 1);
        assert_eq!(find_member.params[0].name, "After");
        assert!(find_member.params[0].optional);
        let return_type = find_member.return_type.as_ref().expect("Find return");
        assert_eq!(return_type.alias_of.as_deref(), Some("Excel.Range"));
        assert!(return_type.nullable);
    }
    let range_replace = range
        .members
        .iter()
        .find(|member| member.name == "Replace")
        .expect("Range.Replace");
    assert_eq!(range_replace.access, AccessMode::Read);
    assert_eq!(range_replace.params.len(), 8);
    assert_eq!(range_replace.params[0].name, "What");
    assert!(!range_replace.params[0].optional);
    assert_eq!(range_replace.params[1].name, "Replacement");
    assert!(!range_replace.params[1].optional);
    assert_eq!(range_replace.params[2].name, "LookAt");
    assert!(range_replace.params[2].optional);
    assert_eq!(range_replace.params[3].name, "SearchOrder");
    assert!(range_replace.params[3].optional);
    assert_eq!(range_replace.params[4].name, "MatchCase");
    assert!(range_replace.params[4].optional);
    assert_eq!(range_replace.params[5].name, "MatchByte");
    assert!(range_replace.params[5].optional);
    assert_eq!(range_replace.params[6].name, "SearchFormat");
    assert!(range_replace.params[6].optional);
    assert_eq!(range_replace.params[7].name, "ReplaceFormat");
    assert!(range_replace.params[7].optional);
    assert_eq!(
        range_replace
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT_BOOL")
    );
    let range_sort = range
        .members
        .iter()
        .find(|member| member.name == "Sort")
        .expect("Range.Sort");
    assert_eq!(range_sort.access, AccessMode::Read);
    assert_eq!(range_sort.params.len(), 15);
    assert!(range_sort.params.iter().all(|param| param.optional));
    assert_eq!(range_sort.params[0].name, "Key1");
    assert_eq!(range_sort.params[1].name, "Order1");
    assert_eq!(range_sort.params[2].name, "Key2");
    assert_eq!(range_sort.params[7].name, "Header");
    assert_eq!(range_sort.params[10].name, "Orientation");
    assert_eq!(range_sort.params[14].name, "DataOption3");
    assert_eq!(
        range_sort
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("VARIANT")
    );
    let range_delete = range
        .members
        .iter()
        .find(|member| member.name == "Delete")
        .expect("Range.Delete");
    assert_eq!(range_delete.access, AccessMode::Read);
    assert_eq!(range_delete.params.len(), 1);
    assert_eq!(range_delete.params[0].name, "Shift");
    assert!(range_delete.params[0].optional);
    assert!(range_delete.return_type.is_none());
    let range_insert = range
        .members
        .iter()
        .find(|member| member.name == "Insert")
        .expect("Range.Insert");
    assert_eq!(range_insert.access, AccessMode::Read);
    assert_eq!(range_insert.params.len(), 1);
    assert_eq!(range_insert.params[0].name, "Shift");
    assert!(range_insert.params[0].optional);
    assert!(range_insert.return_type.is_none());
    let range_copy = range
        .members
        .iter()
        .find(|member| member.name == "Copy")
        .expect("Range.Copy");
    assert_eq!(range_copy.access, AccessMode::Read);
    assert_eq!(range_copy.params.len(), 1);
    assert!(range_copy.params[0].optional);
    assert!(range_copy.return_type.is_none());
    let range_cut = range
        .members
        .iter()
        .find(|member| member.name == "Cut")
        .expect("Range.Cut");
    assert_eq!(range_cut.access, AccessMode::Read);
    assert_eq!(range_cut.params.len(), 1);
    assert!(range_cut.params[0].optional);
    assert!(range_cut.return_type.is_none());
    let range_paste_special = range
        .members
        .iter()
        .find(|member| member.name == "PasteSpecial")
        .expect("Range.PasteSpecial");
    assert_eq!(range_paste_special.access, AccessMode::Read);
    assert_eq!(range_paste_special.params.len(), 4);
    assert_eq!(range_paste_special.params[0].name, "Paste");
    assert_eq!(range_paste_special.params[1].name, "Operation");
    assert_eq!(range_paste_special.params[2].name, "SkipBlanks");
    assert_eq!(range_paste_special.params[3].name, "Transpose");
    assert!(
        range_paste_special
            .params
            .iter()
            .all(|param| param.optional)
    );
    assert!(range_paste_special.return_type.is_none());
    let range_fill_down = range
        .members
        .iter()
        .find(|member| member.name == "FillDown")
        .expect("Range.FillDown");
    assert_eq!(range_fill_down.access, AccessMode::Read);
    assert_eq!(range_fill_down.params.len(), 0);
    assert!(range_fill_down.return_type.is_none());
    let range_fill_right = range
        .members
        .iter()
        .find(|member| member.name == "FillRight")
        .expect("Range.FillRight");
    assert_eq!(range_fill_right.access, AccessMode::Read);
    assert_eq!(range_fill_right.params.len(), 0);
    assert!(range_fill_right.return_type.is_none());
    let range_fill_up = range
        .members
        .iter()
        .find(|member| member.name == "FillUp")
        .expect("Range.FillUp");
    assert_eq!(range_fill_up.access, AccessMode::Read);
    assert_eq!(range_fill_up.params.len(), 0);
    assert!(range_fill_up.return_type.is_none());
    let range_fill_left = range
        .members
        .iter()
        .find(|member| member.name == "FillLeft")
        .expect("Range.FillLeft");
    assert_eq!(range_fill_left.access, AccessMode::Read);
    assert_eq!(range_fill_left.params.len(), 0);
    assert!(range_fill_left.return_type.is_none());
    let range_select = range
        .members
        .iter()
        .find(|member| member.name == "Select")
        .expect("Range.Select");
    assert_eq!(range_select.access, AccessMode::Read);
    assert_eq!(range_select.params.len(), 0);
    assert!(range_select.return_type.is_none());
    let range_calculate = range
        .members
        .iter()
        .find(|member| member.name == "Calculate")
        .expect("Range.Calculate");
    assert_eq!(range_calculate.access, AccessMode::Read);
    assert_eq!(range_calculate.params.len(), 0);
    assert!(range_calculate.return_type.is_none());
    let range_clear_contents = range
        .members
        .iter()
        .find(|member| member.name == "ClearContents")
        .expect("Range.ClearContents");
    assert_eq!(range_clear_contents.access, AccessMode::Read);
    assert_eq!(range_clear_contents.params.len(), 0);
    assert!(range_clear_contents.return_type.is_none());
    let range_clear = range
        .members
        .iter()
        .find(|member| member.name == "Clear")
        .expect("Range.Clear");
    assert_eq!(range_clear.access, AccessMode::Read);
    assert_eq!(range_clear.params.len(), 0);
    assert!(range_clear.return_type.is_none());
    let range_clear_formats = range
        .members
        .iter()
        .find(|member| member.name == "ClearFormats")
        .expect("Range.ClearFormats");
    assert_eq!(range_clear_formats.access, AccessMode::Read);
    assert_eq!(range_clear_formats.params.len(), 0);
    assert!(range_clear_formats.return_type.is_none());
    let range_row = range
        .members
        .iter()
        .find(|member| member.name == "Row")
        .expect("Range.Row");
    assert_eq!(range_row.access, AccessMode::Read);
    let range_column = range
        .members
        .iter()
        .find(|member| member.name == "Column")
        .expect("Range.Column");
    assert_eq!(range_column.access, AccessMode::Read);
    let range_parent = range
        .members
        .iter()
        .find(|member| member.name == "Parent")
        .expect("Range.Parent");
    assert_eq!(range_parent.access, AccessMode::Read);
    assert_eq!(
        range_parent
            .return_type
            .as_ref()
            .and_then(|type_ref| type_ref.alias_of.as_deref()),
        Some("Excel.Worksheet")
    );

    assert_eq!(coverage.library, "Excel");
    assert_eq!(coverage.version, "16.0");
    assert_eq!(coverage.member_count, 1470);
    assert_eq!(coverage.support_counts.stub, 1470);
    assert!(coverage.missing_focus_surfaces.is_empty());

    let application_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Application")
        .expect("Application coverage");
    let worksheet_function_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "WorksheetFunction")
        .expect("WorksheetFunction coverage");
    let workbooks_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Workbooks")
        .expect("Workbooks coverage");
    let workbook_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Workbook")
        .expect("Workbook coverage");
    let worksheets_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Worksheets")
        .expect("Worksheets coverage");
    let sheets_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Sheets")
        .expect("Sheets coverage");
    let charts_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Charts")
        .expect("Charts coverage");
    let worksheet_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Worksheet")
        .expect("Worksheet coverage");
    let range_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Range")
        .expect("Range coverage");
    let areas_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Areas")
        .expect("Areas coverage");
    let names_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Names")
        .expect("Names coverage");
    let name_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Name")
        .expect("Name coverage");
    let chart_objects_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartObjects")
        .expect("ChartObjects coverage");
    let chart_object_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartObject")
        .expect("ChartObject coverage");
    let shape_range_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ShapeRange")
        .expect("ShapeRange coverage");
    let chart_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Chart")
        .expect("Chart coverage");
    let chart_area_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartArea")
        .expect("ChartArea coverage");
    let plot_area_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "PlotArea")
        .expect("PlotArea coverage");
    let chart_title_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartTitle")
        .expect("ChartTitle coverage");
    let legend_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Legend")
        .expect("Legend coverage");
    let legend_entries_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "LegendEntries")
        .expect("LegendEntries coverage");
    let legend_entry_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "LegendEntry")
        .expect("LegendEntry coverage");
    let legend_key_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "LegendKey")
        .expect("LegendKey coverage");
    let data_table_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "DataTable")
        .expect("DataTable coverage");
    let chart_format_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartFormat")
        .expect("ChartFormat coverage");
    let adjustments_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Adjustments")
        .expect("Adjustments coverage");
    let fill_format_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "FillFormat")
        .expect("FillFormat coverage");
    let glow_format_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "GlowFormat")
        .expect("GlowFormat coverage");
    let line_format_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "LineFormat")
        .expect("LineFormat coverage");
    let picture_format_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "PictureFormat")
        .expect("PictureFormat coverage");
    let crop_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Crop")
        .expect("Crop coverage");
    let shadow_format_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ShadowFormat")
        .expect("ShadowFormat coverage");
    let soft_edge_format_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "SoftEdgeFormat")
        .expect("SoftEdgeFormat coverage");
    let text_frame2_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "TextFrame2")
        .expect("TextFrame2 coverage");
    let three_d_format_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ThreeDFormat")
        .expect("ThreeDFormat coverage");
    let chart_groups_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartGroups")
        .expect("ChartGroups coverage");
    let chart_group_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartGroup")
        .expect("ChartGroup coverage");
    let category_collection_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "CategoryCollection")
        .expect("CategoryCollection coverage");
    let chart_category_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "ChartCategory")
        .expect("ChartCategory coverage");
    let series_lines_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "SeriesLines")
        .expect("SeriesLines coverage");
    let drop_lines_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "DropLines")
        .expect("DropLines coverage");
    let hi_lo_lines_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "HiLoLines")
        .expect("HiLoLines coverage");
    let up_bars_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "UpBars")
        .expect("UpBars coverage");
    let down_bars_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "DownBars")
        .expect("DownBars coverage");
    let axes_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Axes")
        .expect("Axes coverage");
    let axis_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Axis")
        .expect("Axis coverage");
    let tick_labels_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "TickLabels")
        .expect("TickLabels coverage");
    let gridlines_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Gridlines")
        .expect("Gridlines coverage");
    let display_unit_label_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "DisplayUnitLabel")
        .expect("DisplayUnitLabel coverage");
    let axis_title_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "AxisTitle")
        .expect("AxisTitle coverage");
    let series_collection_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "SeriesCollection")
        .expect("SeriesCollection coverage");
    let series_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Series")
        .expect("Series coverage");
    let leader_lines_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "LeaderLines")
        .expect("LeaderLines coverage");
    let border_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Border")
        .expect("Border coverage");
    let data_labels_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "DataLabels")
        .expect("DataLabels coverage");
    let data_label_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "DataLabel")
        .expect("DataLabel coverage");
    let points_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Points")
        .expect("Points coverage");
    let point_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Point")
        .expect("Point coverage");

    assert_eq!(application_coverage.member_count, 44);
    assert_eq!(application_coverage.support_counts.stub, 44);
    assert_eq!(
        application_coverage.stub_members,
        vec![
            "Workbooks".to_string(),
            "Worksheets".to_string(),
            "Sheets".to_string(),
            "Charts".to_string(),
            "Names".to_string(),
            "WorksheetFunction".to_string(),
            "ActiveWorkbook".to_string(),
            "ActiveSheet".to_string(),
            "ActiveCell".to_string(),
            "Selection".to_string(),
            "ActiveChart".to_string(),
            "Name".to_string(),
            "Version".to_string(),
            "UserName".to_string(),
            "DefaultFilePath".to_string(),
            "Caption".to_string(),
            "DisplayAlerts".to_string(),
            "Calculation".to_string(),
            "ScreenUpdating".to_string(),
            "EnableEvents".to_string(),
            "StatusBar".to_string(),
            "CutCopyMode".to_string(),
            "DisplayStatusBar".to_string(),
            "DisplayFormulaBar".to_string(),
            "DisplayScrollBars".to_string(),
            "DisplayFullScreen".to_string(),
            "UseSystemSeparators".to_string(),
            "DecimalSeparator".to_string(),
            "ThousandsSeparator".to_string(),
            "International".to_string(),
            "ShowWindowsInTaskbar".to_string(),
            "Interactive".to_string(),
            "Cells".to_string(),
            "Rows".to_string(),
            "Columns".to_string(),
            "Quit".to_string(),
            "Calculate".to_string(),
            "CalculateFull".to_string(),
            "CalculateFullRebuild".to_string(),
            "Evaluate".to_string(),
            "Goto".to_string(),
            "Range".to_string(),
            "Intersect".to_string(),
            "Union".to_string()
        ]
    );

    assert_eq!(worksheet_function_coverage.member_count, 526);
    assert_eq!(worksheet_function_coverage.support_counts.stub, 526);
    assert_eq!(
        worksheet_function_coverage.stub_members,
        vec![
            "Application".to_string(),
            "Creator".to_string(),
            "Parent".to_string(),
            "Sum".to_string(),
            "Average".to_string(),
            "Count".to_string(),
            "CountA".to_string(),
            "CountBlank".to_string(),
            "AverageA".to_string(),
            "MaxA".to_string(),
            "MinA".to_string(),
            "Min".to_string(),
            "Max".to_string(),
            "Product".to_string(),
            "SumIf".to_string(),
            "CountIf".to_string(),
            "AverageIf".to_string(),
            "CountIfs".to_string(),
            "SumIfs".to_string(),
            "AverageIfs".to_string(),
            "MinIfs".to_string(),
            "MaxIfs".to_string(),
            "Abs".to_string(),
            "Int".to_string(),
            "Round".to_string(),
            "Power".to_string(),
            "Sqrt".to_string(),
            "Mod".to_string(),
            "Sign".to_string(),
            "RoundUp".to_string(),
            "RoundDown".to_string(),
            "Even".to_string(),
            "Sin".to_string(),
            "Cos".to_string(),
            "Tan".to_string(),
            "Asin".to_string(),
            "Acos".to_string(),
            "Atan".to_string(),
            "Atan2".to_string(),
            "Degrees".to_string(),
            "Radians".to_string(),
            "Pi".to_string(),
            "Sinh".to_string(),
            "Cosh".to_string(),
            "Tanh".to_string(),
            "Asinh".to_string(),
            "Acosh".to_string(),
            "Atanh".to_string(),
            "Sec".to_string(),
            "Csc".to_string(),
            "Cot".to_string(),
            "Sech".to_string(),
            "SqrtPi".to_string(),
            "BesselI".to_string(),
            "BesselJ".to_string(),
            "BesselK".to_string(),
            "BesselY".to_string(),
            "Fisher".to_string(),
            "FisherInv".to_string(),
            "Erf".to_string(),
            "Erf_Precise".to_string(),
            "Erfc".to_string(),
            "Erfc_Precise".to_string(),
            "GammaLn".to_string(),
            "GammaLn_Precise".to_string(),
            "Odd".to_string(),
            "Gcd".to_string(),
            "Lcm".to_string(),
            "Fact".to_string(),
            "FactDouble".to_string(),
            "Combin".to_string(),
            "Combina".to_string(),
            "Permut".to_string(),
            "PermutationA".to_string(),
            "Multinomial".to_string(),
            "MRound".to_string(),
            "Quotient".to_string(),
            "Ceiling".to_string(),
            "Floor".to_string(),
            "Ceiling_Math".to_string(),
            "Floor_Math".to_string(),
            "Ceiling_Precise".to_string(),
            "Floor_Precise".to_string(),
            "Iso_Ceiling".to_string(),
            "IsEven".to_string(),
            "IsOdd".to_string(),
            "Exp".to_string(),
            "Ln".to_string(),
            "Log".to_string(),
            "Log10".to_string(),
            "LogNormDist".to_string(),
            "LogInv".to_string(),
            "Trunc".to_string(),
            "BitAnd".to_string(),
            "BitOr".to_string(),
            "BitXor".to_string(),
            "BitLShift".to_string(),
            "BitRShift".to_string(),
            "Base".to_string(),
            "Dec2Bin".to_string(),
            "Dec2Oct".to_string(),
            "Dec2Hex".to_string(),
            "Bin2Hex".to_string(),
            "Bin2Oct".to_string(),
            "Oct2Bin".to_string(),
            "Oct2Hex".to_string(),
            "Hex2Bin".to_string(),
            "Hex2Oct".to_string(),
            "Complex".to_string(),
            "ImConjugate".to_string(),
            "ImCos".to_string(),
            "ImCosh".to_string(),
            "ImCot".to_string(),
            "ImCsc".to_string(),
            "ImCsch".to_string(),
            "ImDiv".to_string(),
            "ImExp".to_string(),
            "ImLn".to_string(),
            "ImLog10".to_string(),
            "ImLog2".to_string(),
            "ImPower".to_string(),
            "ImProduct".to_string(),
            "ImSec".to_string(),
            "ImSech".to_string(),
            "ImSin".to_string(),
            "ImSinh".to_string(),
            "ImSqrt".to_string(),
            "ImSub".to_string(),
            "ImSum".to_string(),
            "ImTan".to_string(),
            "Len".to_string(),
            "LeftB".to_string(),
            "RightB".to_string(),
            "MidB".to_string(),
            "Asc".to_string(),
            "DBCS".to_string(),
            "Jis".to_string(),
            "Char".to_string(),
            "UniChar".to_string(),
            "Left".to_string(),
            "Right".to_string(),
            "Mid".to_string(),
            "Upper".to_string(),
            "Lower".to_string(),
            "Trim".to_string(),
            "Clean".to_string(),
            "Exact".to_string(),
            "Value".to_string(),
            "Find".to_string(),
            "Search".to_string(),
            "Rept".to_string(),
            "Replace".to_string(),
            "ReplaceB".to_string(),
            "Substitute".to_string(),
            "Proper".to_string(),
            "Concatenate".to_string(),
            "Concat".to_string(),
            "TextJoin".to_string(),
            "TextBefore".to_string(),
            "TextAfter".to_string(),
            "TextSplit".to_string(),
            "T".to_string(),
            "EncodeUrl".to_string(),
            "ValueToText".to_string(),
            "ArrayToText".to_string(),
            "NumberValue".to_string(),
            "Dollar".to_string(),
            "Fixed".to_string(),
            "Roman".to_string(),
            "BahtText".to_string(),
            "Address".to_string(),
            "FormulaText".to_string(),
            "Date".to_string(),
            "Year".to_string(),
            "Month".to_string(),
            "Day".to_string(),
            "Days".to_string(),
            "Time".to_string(),
            "Today".to_string(),
            "Now".to_string(),
            "Rand".to_string(),
            "RandBetween".to_string(),
            "Hour".to_string(),
            "Minute".to_string(),
            "Second".to_string(),
            "Weekday".to_string(),
            "EDate".to_string(),
            "EoMonth".to_string(),
            "Days360".to_string(),
            "YearFrac".to_string(),
            "DatedIf".to_string(),
            "WeekNum".to_string(),
            "IsoWeekNum".to_string(),
            "DateValue".to_string(),
            "TimeValue".to_string(),
            "NetworkDays".to_string(),
            "WorkDay".to_string(),
            "WorkDay_Intl".to_string(),
            "NetworkDays_Intl".to_string(),
            "IfError".to_string(),
            "IfNa".to_string(),
            "IsError".to_string(),
            "IsNA".to_string(),
            "IsBlank".to_string(),
            "IsNumber".to_string(),
            "IsText".to_string(),
            "IsLogical".to_string(),
            "IsNonText".to_string(),
            "IsRef".to_string(),
            "N".to_string(),
            "Type".to_string(),
            "Error_Type".to_string(),
            "FV".to_string(),
            "PV".to_string(),
            "PMT".to_string(),
            "NPER".to_string(),
            "RATE".to_string(),
            "IPMT".to_string(),
            "PPMT".to_string(),
            "NPV".to_string(),
            "IRR".to_string(),
            "MIRR".to_string(),
            "CumIPmt".to_string(),
            "CumPrinc".to_string(),
            "Effect".to_string(),
            "Nominal".to_string(),
            "RRI".to_string(),
            "PDuration".to_string(),
            "FVSchedule".to_string(),
            "XNPV".to_string(),
            "XIRR".to_string(),
            "SLN".to_string(),
            "Syd".to_string(),
            "Db".to_string(),
            "Ddb".to_string(),
            "Vdb".to_string(),
            "Disc".to_string(),
            "PriceDisc".to_string(),
            "Intrate".to_string(),
            "Received".to_string(),
            "YieldDisc".to_string(),
            "TBillEq".to_string(),
            "TBillPrice".to_string(),
            "TBillYield".to_string(),
            "CoupDayBs".to_string(),
            "CoupDays".to_string(),
            "CoupDaysNc".to_string(),
            "CoupNcd".to_string(),
            "CoupNum".to_string(),
            "CoupPcd".to_string(),
            "Price".to_string(),
            "Yield".to_string(),
            "PriceMat".to_string(),
            "YieldMat".to_string(),
            "Duration".to_string(),
            "MDuration".to_string(),
            "AccrInt".to_string(),
            "AccrIntM".to_string(),
            "OddFPrice".to_string(),
            "OddFYield".to_string(),
            "OddLPrice".to_string(),
            "OddLYield".to_string(),
            "AmorDegrc".to_string(),
            "AmorLinc".to_string(),
            "Acot".to_string(),
            "Acoth".to_string(),
            "Coth".to_string(),
            "Csch".to_string(),
            "Delta".to_string(),
            "GeStep".to_string(),
            "Gamma".to_string(),
            "Gauss".to_string(),
            "Phi".to_string(),
            "Standardize".to_string(),
            "Confidence".to_string(),
            "Confidence_Norm".to_string(),
            "Confidence_T".to_string(),
            "Binom_Dist".to_string(),
            "BinomDist".to_string(),
            "Binom_Dist_Range".to_string(),
            "Binom_Inv".to_string(),
            "CritBinom".to_string(),
            "NegBinom_Dist".to_string(),
            "HypGeom_Dist".to_string(),
            "HypGeomDist".to_string(),
            "NegBinomDist".to_string(),
            "If".to_string(),
            "Not".to_string(),
            "And".to_string(),
            "Or".to_string(),
            "Xor".to_string(),
            "AveDev".to_string(),
            "Kurt".to_string(),
            "Skew".to_string(),
            "Skew_P".to_string(),
            "TrimMean".to_string(),
            "Prob".to_string(),
            "Median".to_string(),
            "Mode".to_string(),
            "SumSq".to_string(),
            "DevSq".to_string(),
            "Var_P".to_string(),
            "Var_S".to_string(),
            "StDev_P".to_string(),
            "StDev_S".to_string(),
            "VarP".to_string(),
            "StDevP".to_string(),
            "VarA".to_string(),
            "VarPA".to_string(),
            "StDevA".to_string(),
            "StDevPA".to_string(),
            "Var".to_string(),
            "StDev".to_string(),
            "GeoMean".to_string(),
            "HarMean".to_string(),
            "Percentile".to_string(),
            "Quartile".to_string(),
            "PercentRank".to_string(),
            "Rank".to_string(),
            "Percentile_Inc".to_string(),
            "Percentile_Exc".to_string(),
            "Quartile_Inc".to_string(),
            "Quartile_Exc".to_string(),
            "PercentRank_Inc".to_string(),
            "PercentRank_Exc".to_string(),
            "Rank_Avg".to_string(),
            "Rank_Eq".to_string(),
            "Large".to_string(),
            "Small".to_string(),
            "Correl".to_string(),
            "Pearson".to_string(),
            "Covar".to_string(),
            "Covariance_P".to_string(),
            "Covariance_S".to_string(),
            "Slope".to_string(),
            "Intercept".to_string(),
            "RSq".to_string(),
            "Forecast".to_string(),
            "Steyx".to_string(),
            "Expon_Dist".to_string(),
            "ExponDist".to_string(),
            "Poisson".to_string(),
            "Weibull".to_string(),
            "Poisson_Dist".to_string(),
            "Weibull_Dist".to_string(),
            "Norm_Dist".to_string(),
            "NormDist".to_string(),
            "Norm_S_Dist".to_string(),
            "Norm_Inv".to_string(),
            "NormInv".to_string(),
            "Norm_S_Inv".to_string(),
            "NormSDist".to_string(),
            "NormSInv".to_string(),
            "LogNorm_Dist".to_string(),
            "LogNorm_Inv".to_string(),
            "Beta_Dist".to_string(),
            "BetaDist".to_string(),
            "Beta_Inv".to_string(),
            "BetaInv".to_string(),
            "Gamma_Dist".to_string(),
            "GammaDist".to_string(),
            "Gamma_Inv".to_string(),
            "GammaInv".to_string(),
            "ChiSq_Dist".to_string(),
            "ChiSq_Dist_RT".to_string(),
            "ChiDist".to_string(),
            "ChiSq_Inv".to_string(),
            "ChiSq_Inv_RT".to_string(),
            "ChiInv".to_string(),
            "F_Dist".to_string(),
            "F_Dist_RT".to_string(),
            "FDist".to_string(),
            "F_Inv".to_string(),
            "F_Inv_RT".to_string(),
            "FInv".to_string(),
            "T_Dist".to_string(),
            "T_Dist_RT".to_string(),
            "T_Dist_2T".to_string(),
            "TDist".to_string(),
            "T_Inv".to_string(),
            "T_Inv_2T".to_string(),
            "TInv".to_string(),
            "ChiSq_Test".to_string(),
            "F_Test".to_string(),
            "T_Test".to_string(),
            "Z_Test".to_string(),
            "Cell".to_string(),
            "Choose".to_string(),
            "CubeKpiMember".to_string(),
            "CubeMember".to_string(),
            "CubeMemberProperty".to_string(),
            "CubeRankedMember".to_string(),
            "CubeSet".to_string(),
            "DetectLanguage".to_string(),
            "DGet".to_string(),
            "FilterXML".to_string(),
            "GetPivotData".to_string(),
            "HLookup".to_string(),
            "Hyperlink".to_string(),
            "Ifs".to_string(),
            "Image".to_string(),
            "Index".to_string(),
            "Indirect".to_string(),
            "Info".to_string(),
            "Lambda".to_string(),
            "Let".to_string(),
            "Lookup".to_string(),
            "MakeArray".to_string(),
            "Offset".to_string(),
            "Phonetic".to_string(),
            "Reduce".to_string(),
            "RegexExtract".to_string(),
            "RegexReplace".to_string(),
            "Scan".to_string(),
            "Switch".to_string(),
            "Text".to_string(),
            "Translate".to_string(),
            "TrimRange".to_string(),
            "VLookup".to_string(),
            "WebService".to_string(),
            "XLookup".to_string(),
            "Arabic".to_string(),
            "Bin2Dec".to_string(),
            "Decimal".to_string(),
            "Hex2Dec".to_string(),
            "Oct2Dec".to_string(),
            "Convert".to_string(),
            "EuroConvert".to_string(),
            "ImAbs".to_string(),
            "Imaginary".to_string(),
            "ImArgument".to_string(),
            "ImReal".to_string(),
            "Code".to_string(),
            "Unicode".to_string(),
            "LenB".to_string(),
            "FindB".to_string(),
            "SearchB".to_string(),
            "RegexTest".to_string(),
            "Areas".to_string(),
            "Column".to_string(),
            "Columns".to_string(),
            "Cols".to_string(),
            "Row".to_string(),
            "Rows".to_string(),
            "Sheet".to_string(),
            "Sheets".to_string(),
            "IsFormula".to_string(),
            "IsErr".to_string(),
            "Na".to_string(),
            "IsOmitted".to_string(),
            "Match".to_string(),
            "XMatch".to_string(),
            "ChooseCols".to_string(),
            "ChooseRows".to_string(),
            "Drop".to_string(),
            "Expand".to_string(),
            "Filter".to_string(),
            "HStack".to_string(),
            "Sort".to_string(),
            "SortBy".to_string(),
            "Take".to_string(),
            "ToCol".to_string(),
            "ToRow".to_string(),
            "Transpose".to_string(),
            "Unique".to_string(),
            "VStack".to_string(),
            "WrapCols".to_string(),
            "WrapRows".to_string(),
            "ByCol".to_string(),
            "ByRow".to_string(),
            "Map".to_string(),
            "PercentOf".to_string(),
            "Sequence".to_string(),
            "RandArray".to_string(),
            "Aggregate".to_string(),
            "Subtotal".to_string(),
            "SumProduct".to_string(),
            "SumX2MY2".to_string(),
            "SumX2PY2".to_string(),
            "SumXMY2".to_string(),
            "Seriessum".to_string(),
            "Frequency".to_string(),
            "MDeterm".to_string(),
            "MInverse".to_string(),
            "MMult".to_string(),
            "MUnit".to_string(),
            "Growth".to_string(),
            "LinEst".to_string(),
            "LogEst".to_string(),
            "Trend".to_string(),
            "Mode_Mult".to_string(),
            "Mode_Sngl".to_string(),
            "Forecast_Linear".to_string(),
            "Forecast_Ets".to_string(),
            "Forecast_Ets_ConfInt".to_string(),
            "Forecast_Ets_Seasonality".to_string(),
            "Forecast_Ets_Stat".to_string(),
            "ChiTest".to_string(),
            "FTest".to_string(),
            "TTest".to_string(),
            "ZTest".to_string(),
            "DAverage".to_string(),
            "DCount".to_string(),
            "DCountA".to_string(),
            "DMax".to_string(),
            "DMin".to_string(),
            "DProduct".to_string(),
            "DStDev".to_string(),
            "DStDevP".to_string(),
            "DSum".to_string(),
            "DVar".to_string(),
            "DVarP".to_string(),
            "DollarDe".to_string(),
            "DollarFr".to_string(),
            "IsPmt".to_string(),
            "CubeSetCount".to_string(),
            "CubeValue".to_string(),
            "Call".to_string(),
            "Register_ID".to_string(),
            "RTD".to_string(),
            "FieldValue".to_string(),
            "GroupBy".to_string(),
            "PivotBy".to_string(),
            "Py".to_string(),
            "Copilot".to_string(),
            "StockHistory".to_string(),
        ]
    );

    assert_eq!(workbooks_coverage.member_count, 7);
    assert_eq!(workbooks_coverage.support_counts.stub, 7);
    assert_eq!(
        workbooks_coverage.stub_members,
        vec![
            "Count".to_string(),
            "Parent".to_string(),
            "Application".to_string(),
            "Item".to_string(),
            "Add".to_string(),
            "Open".to_string(),
            "Close".to_string()
        ]
    );

    assert_eq!(workbook_coverage.member_count, 26);
    assert_eq!(workbook_coverage.support_counts.stub, 26);
    assert_eq!(
        workbook_coverage.stub_members,
        vec![
            "Worksheets".to_string(),
            "Sheets".to_string(),
            "Charts".to_string(),
            "Names".to_string(),
            "ActiveSheet".to_string(),
            "Activate".to_string(),
            "Name".to_string(),
            "Parent".to_string(),
            "Application".to_string(),
            "Path".to_string(),
            "FullName".to_string(),
            "FileFormat".to_string(),
            "Date1904".to_string(),
            "IsAddin".to_string(),
            "HasVBProject".to_string(),
            "ReadOnly".to_string(),
            "Saved".to_string(),
            "Save".to_string(),
            "SaveAs".to_string(),
            "SaveCopyAs".to_string(),
            "RefreshAll".to_string(),
            "CheckSpelling".to_string(),
            "ExportAsFixedFormat".to_string(),
            "PrintPreview".to_string(),
            "PrintOut".to_string(),
            "Close".to_string()
        ]
    );

    assert_eq!(worksheets_coverage.member_count, 13);
    assert_eq!(worksheets_coverage.support_counts.stub, 13);
    assert_eq!(
        worksheets_coverage.stub_members,
        vec![
            "Count".to_string(),
            "Parent".to_string(),
            "Application".to_string(),
            "Creator".to_string(),
            "Visible".to_string(),
            "Add".to_string(),
            "Item".to_string(),
            "Delete".to_string(),
            "Copy".to_string(),
            "Move".to_string(),
            "PrintPreview".to_string(),
            "PrintOut".to_string(),
            "Select".to_string()
        ]
    );

    assert_eq!(sheets_coverage.member_count, 13);
    assert_eq!(sheets_coverage.support_counts.stub, 13);
    assert_eq!(
        sheets_coverage.stub_members,
        vec![
            "Count".to_string(),
            "Parent".to_string(),
            "Application".to_string(),
            "Creator".to_string(),
            "Visible".to_string(),
            "Add".to_string(),
            "Item".to_string(),
            "Delete".to_string(),
            "Copy".to_string(),
            "Move".to_string(),
            "PrintPreview".to_string(),
            "PrintOut".to_string(),
            "Select".to_string()
        ]
    );

    assert_eq!(charts_coverage.member_count, 13);
    assert_eq!(charts_coverage.support_counts.stub, 13);
    assert_eq!(
        charts_coverage.stub_members,
        vec![
            "Count".to_string(),
            "Parent".to_string(),
            "Application".to_string(),
            "Creator".to_string(),
            "Visible".to_string(),
            "Add".to_string(),
            "Item".to_string(),
            "Delete".to_string(),
            "Copy".to_string(),
            "Move".to_string(),
            "PrintPreview".to_string(),
            "PrintOut".to_string(),
            "Select".to_string()
        ]
    );

    assert_eq!(worksheet_coverage.member_count, 28);
    assert_eq!(worksheet_coverage.support_counts.stub, 28);
    assert_eq!(
        worksheet_coverage.stub_members,
        vec![
            "Name".to_string(),
            "Parent".to_string(),
            "Application".to_string(),
            "Names".to_string(),
            "Index".to_string(),
            "Next".to_string(),
            "Previous".to_string(),
            "Visible".to_string(),
            "Type".to_string(),
            "Range".to_string(),
            "UsedRange".to_string(),
            "Cells".to_string(),
            "Rows".to_string(),
            "Columns".to_string(),
            "ChartObjects".to_string(),
            "Activate".to_string(),
            "Select".to_string(),
            "Calculate".to_string(),
            "Evaluate".to_string(),
            "Paste".to_string(),
            "PasteSpecial".to_string(),
            "CheckSpelling".to_string(),
            "ExportAsFixedFormat".to_string(),
            "PrintPreview".to_string(),
            "PrintOut".to_string(),
            "Delete".to_string(),
            "Move".to_string(),
            "Copy".to_string()
        ]
    );

    assert_eq!(range_coverage.member_count, 51);
    assert_eq!(range_coverage.support_counts.stub, 51);
    assert_eq!(
        range_coverage.stub_members,
        vec![
            "Value".to_string(),
            "Value2".to_string(),
            "Formula".to_string(),
            "FormulaR1C1".to_string(),
            "FormulaLocal".to_string(),
            "FormulaR1C1Local".to_string(),
            "Formula2".to_string(),
            "Formula2R1C1".to_string(),
            "Formula2Local".to_string(),
            "Formula2R1C1Local".to_string(),
            "Text".to_string(),
            "HasFormula".to_string(),
            "Address".to_string(),
            "Parent".to_string(),
            "Areas".to_string(),
            "Application".to_string(),
            "Worksheet".to_string(),
            "Row".to_string(),
            "Column".to_string(),
            "Count".to_string(),
            "CountLarge".to_string(),
            "CurrentRegion".to_string(),
            "EntireRow".to_string(),
            "EntireColumn".to_string(),
            "Cells".to_string(),
            "Rows".to_string(),
            "Columns".to_string(),
            "Item".to_string(),
            "Offset".to_string(),
            "Resize".to_string(),
            "End".to_string(),
            "Find".to_string(),
            "FindNext".to_string(),
            "FindPrevious".to_string(),
            "Replace".to_string(),
            "Sort".to_string(),
            "Delete".to_string(),
            "Insert".to_string(),
            "Copy".to_string(),
            "Cut".to_string(),
            "CopyPicture".to_string(),
            "PasteSpecial".to_string(),
            "FillDown".to_string(),
            "FillRight".to_string(),
            "FillUp".to_string(),
            "FillLeft".to_string(),
            "Select".to_string(),
            "Calculate".to_string(),
            "ClearContents".to_string(),
            "Clear".to_string(),
            "ClearFormats".to_string()
        ]
    );

    assert_eq!(areas_coverage.member_count, 5);
    assert_eq!(areas_coverage.support_counts.stub, 5);
    assert_eq!(
        areas_coverage.stub_members,
        vec![
            "Count".to_string(),
            "Item".to_string(),
            "Application".to_string(),
            "Creator".to_string(),
            "Parent".to_string()
        ]
    );

    assert_eq!(names_coverage.member_count, 5);
    assert_eq!(names_coverage.support_counts.stub, 5);
    assert_eq!(
        names_coverage.stub_members,
        vec![
            "Count".to_string(),
            "Item".to_string(),
            "Add".to_string(),
            "Application".to_string(),
            "Parent".to_string()
        ]
    );

    assert_eq!(name_coverage.member_count, 6);
    assert_eq!(name_coverage.support_counts.stub, 6);
    assert_eq!(
        name_coverage.stub_members,
        vec![
            "Name".to_string(),
            "RefersTo".to_string(),
            "RefersToRange".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Delete".to_string()
        ]
    );

    assert_eq!(chart_objects_coverage.member_count, 25);
    assert_eq!(chart_objects_coverage.support_counts.stub, 25);
    assert_eq!(
        chart_objects_coverage.stub_members,
        vec![
            "Count".to_string(),
            "Item".to_string(),
            "Add".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Left".to_string(),
            "Top".to_string(),
            "Width".to_string(),
            "Height".to_string(),
            "Placement".to_string(),
            "Visible".to_string(),
            "ProtectChartObject".to_string(),
            "PrintObject".to_string(),
            "Locked".to_string(),
            "RoundedCorners".to_string(),
            "ShapeRange".to_string(),
            "Select".to_string(),
            "Copy".to_string(),
            "BringToFront".to_string(),
            "Cut".to_string(),
            "Duplicate".to_string(),
            "CopyPicture".to_string(),
            "Delete".to_string(),
            "SendToBack".to_string()
        ]
    );

    assert_eq!(chart_object_coverage.member_count, 35);
    assert_eq!(chart_object_coverage.support_counts.stub, 35);
    assert_eq!(
        chart_object_coverage.stub_members,
        vec![
            "Name".to_string(),
            "Chart".to_string(),
            "Index".to_string(),
            "ZOrder".to_string(),
            "Placement".to_string(),
            "Left".to_string(),
            "Top".to_string(),
            "Width".to_string(),
            "Height".to_string(),
            "Visible".to_string(),
            "OnAction".to_string(),
            "PrintObject".to_string(),
            "Locked".to_string(),
            "ProtectChartObject".to_string(),
            "RoundedCorners".to_string(),
            "ShapeRange".to_string(),
            "TopLeftCell".to_string(),
            "BottomRightCell".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Activate".to_string(),
            "Select".to_string(),
            "BringToFront".to_string(),
            "Copy".to_string(),
            "Cut".to_string(),
            "Duplicate".to_string(),
            "CopyPicture".to_string(),
            "Delete".to_string(),
            "SendToBack".to_string(),
            "IncrementLeft".to_string(),
            "IncrementTop".to_string(),
            "IncrementRotation".to_string(),
            "ScaleWidth".to_string(),
            "ScaleHeight".to_string()
        ]
    );

    assert_eq!(shape_range_coverage.member_count, 46);
    assert_eq!(shape_range_coverage.support_counts.stub, 46);
    assert_eq!(
        shape_range_coverage.stub_members,
        vec![
            "Count".to_string(),
            "Item".to_string(),
            "Name".to_string(),
            "Chart".to_string(),
            "Index".to_string(),
            "ID".to_string(),
            "Type".to_string(),
            "AutoShapeType".to_string(),
            "AlternativeText".to_string(),
            "Title".to_string(),
            "HasChart".to_string(),
            "HasSmartArt".to_string(),
            "LockAspectRatio".to_string(),
            "Rotation".to_string(),
            "HorizontalFlip".to_string(),
            "VerticalFlip".to_string(),
            "ZOrder".to_string(),
            "ZOrderPosition".to_string(),
            "Placement".to_string(),
            "Left".to_string(),
            "Top".to_string(),
            "Width".to_string(),
            "Height".to_string(),
            "Visible".to_string(),
            "OnAction".to_string(),
            "PrintObject".to_string(),
            "Locked".to_string(),
            "ProtectChartObject".to_string(),
            "RoundedCorners".to_string(),
            "TopLeftCell".to_string(),
            "BottomRightCell".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Copy".to_string(),
            "Cut".to_string(),
            "Duplicate".to_string(),
            "CopyPicture".to_string(),
            "Delete".to_string(),
            "Flip".to_string(),
            "Select".to_string(),
            "IncrementLeft".to_string(),
            "IncrementTop".to_string(),
            "IncrementRotation".to_string(),
            "ScaleWidth".to_string(),
            "ScaleHeight".to_string()
        ]
    );

    assert_eq!(chart_coverage.member_count, 80);
    assert_eq!(chart_coverage.support_counts.stub, 80);
    assert_eq!(
        chart_coverage.stub_members,
        vec![
            "Name".to_string(),
            "ChartType".to_string(),
            "ChartStyle".to_string(),
            "Index".to_string(),
            "BarShape".to_string(),
            "Elevation".to_string(),
            "HeightPercent".to_string(),
            "Rotation".to_string(),
            "DepthPercent".to_string(),
            "GapDepth".to_string(),
            "Perspective".to_string(),
            "RightAngleAxes".to_string(),
            "DisplayBlanksAs".to_string(),
            "PlotVisibleOnly".to_string(),
            "ShowDataLabelsOverMaximum".to_string(),
            "ProtectContents".to_string(),
            "ProtectDrawingObjects".to_string(),
            "ProtectData".to_string(),
            "ProtectFormatting".to_string(),
            "ProtectSelection".to_string(),
            "ProtectionMode".to_string(),
            "ChartArea".to_string(),
            "PlotArea".to_string(),
            "HasTitle".to_string(),
            "ChartTitle".to_string(),
            "HasDataTable".to_string(),
            "DataTable".to_string(),
            "HasAxis".to_string(),
            "HasLegend".to_string(),
            "Legend".to_string(),
            "ChartGroups".to_string(),
            "AreaGroups".to_string(),
            "BarGroups".to_string(),
            "ColumnGroups".to_string(),
            "DoughnutGroups".to_string(),
            "LineGroups".to_string(),
            "PieGroups".to_string(),
            "RadarGroups".to_string(),
            "SurfaceGroups".to_string(),
            "XYGroups".to_string(),
            "Axes".to_string(),
            "SeriesCollection".to_string(),
            "FullSeriesCollection".to_string(),
            "ChartObjects".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Next".to_string(),
            "Previous".to_string(),
            "Visible".to_string(),
            "Activate".to_string(),
            "Select".to_string(),
            "Copy".to_string(),
            "Move".to_string(),
            "Protect".to_string(),
            "Unprotect".to_string(),
            "Refresh".to_string(),
            "CheckSpelling".to_string(),
            "Deselect".to_string(),
            "ClearToMatchColorStyle".to_string(),
            "ClearToMatchStyle".to_string(),
            "ApplyLayout".to_string(),
            "ApplyCustomType".to_string(),
            "SetSourceData".to_string(),
            "Location".to_string(),
            "ChartWizard".to_string(),
            "ApplyDataLabels".to_string(),
            "ApplyChartTemplate".to_string(),
            "SaveChartTemplate".to_string(),
            "SetDefaultChart".to_string(),
            "SetBackgroundPicture".to_string(),
            "Paste".to_string(),
            "Evaluate".to_string(),
            "CopyPicture".to_string(),
            "SetElement".to_string(),
            "Export".to_string(),
            "ExportAsFixedFormat".to_string(),
            "PrintPreview".to_string(),
            "PrintOut".to_string(),
            "Delete".to_string()
        ]
    );

    assert_eq!(chart_area_coverage.member_count, 16);
    assert_eq!(chart_area_coverage.support_counts.stub, 16);
    assert_eq!(
        chart_area_coverage.stub_members,
        vec![
            "Name".to_string(),
            "Format".to_string(),
            "Border".to_string(),
            "Left".to_string(),
            "Top".to_string(),
            "Width".to_string(),
            "Height".to_string(),
            "RoundedCorners".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Select".to_string(),
            "Copy".to_string(),
            "Clear".to_string(),
            "ClearFormats".to_string(),
            "ClearContents".to_string()
        ]
    );

    assert_eq!(plot_area_coverage.member_count, 19);
    assert_eq!(plot_area_coverage.support_counts.stub, 19);
    assert_eq!(
        plot_area_coverage.stub_members,
        vec![
            "Name".to_string(),
            "Format".to_string(),
            "Border".to_string(),
            "Left".to_string(),
            "Top".to_string(),
            "Width".to_string(),
            "Height".to_string(),
            "InsideLeft".to_string(),
            "InsideTop".to_string(),
            "InsideWidth".to_string(),
            "InsideHeight".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Select".to_string(),
            "Copy".to_string(),
            "Clear".to_string(),
            "ClearFormats".to_string(),
            "ClearContents".to_string()
        ]
    );

    assert_eq!(chart_title_coverage.member_count, 16);
    assert_eq!(chart_title_coverage.support_counts.stub, 16);
    assert_eq!(
        chart_title_coverage.stub_members,
        vec![
            "Name".to_string(),
            "Format".to_string(),
            "Border".to_string(),
            "Text".to_string(),
            "Caption".to_string(),
            "Left".to_string(),
            "Top".to_string(),
            "Width".to_string(),
            "Height".to_string(),
            "Orientation".to_string(),
            "ReadingOrder".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Select".to_string(),
            "Delete".to_string()
        ]
    );

    assert_eq!(legend_coverage.member_count, 16);
    assert_eq!(legend_coverage.support_counts.stub, 16);
    assert_eq!(
        legend_coverage.stub_members,
        vec![
            "Name".to_string(),
            "Format".to_string(),
            "Border".to_string(),
            "Position".to_string(),
            "IncludeInLayout".to_string(),
            "Left".to_string(),
            "Top".to_string(),
            "Width".to_string(),
            "Height".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Select".to_string(),
            "LegendEntries".to_string(),
            "Clear".to_string(),
            "Delete".to_string()
        ]
    );

    assert_eq!(legend_entries_coverage.member_count, 5);
    assert_eq!(legend_entries_coverage.support_counts.stub, 5);
    assert_eq!(
        legend_entries_coverage.stub_members,
        vec![
            "Count".to_string(),
            "Item".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string()
        ]
    );

    assert_eq!(legend_entry_coverage.member_count, 11);
    assert_eq!(legend_entry_coverage.support_counts.stub, 11);
    assert_eq!(
        legend_entry_coverage.stub_members,
        vec![
            "Index".to_string(),
            "Format".to_string(),
            "LegendKey".to_string(),
            "Left".to_string(),
            "Top".to_string(),
            "Width".to_string(),
            "Height".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Select".to_string()
        ]
    );

    assert_eq!(legend_key_coverage.member_count, 9);
    assert_eq!(legend_key_coverage.support_counts.stub, 9);
    assert_eq!(
        legend_key_coverage.stub_members,
        vec![
            "Format".to_string(),
            "Border".to_string(),
            "Left".to_string(),
            "Top".to_string(),
            "Width".to_string(),
            "Height".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string()
        ]
    );

    assert_eq!(data_table_coverage.member_count, 11);
    assert_eq!(data_table_coverage.support_counts.stub, 11);
    assert_eq!(
        data_table_coverage.stub_members,
        vec![
            "HasBorderHorizontal".to_string(),
            "HasBorderVertical".to_string(),
            "HasBorderOutline".to_string(),
            "ShowLegendKey".to_string(),
            "Format".to_string(),
            "Border".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Select".to_string(),
            "Delete".to_string()
        ]
    );

    assert_eq!(chart_format_coverage.member_count, 13);
    assert_eq!(chart_format_coverage.support_counts.stub, 13);
    assert_eq!(
        chart_format_coverage.stub_members,
        vec![
            "Creator".to_string(),
            "Application".to_string(),
            "Adjustments".to_string(),
            "AutoShapeType".to_string(),
            "Parent".to_string(),
            "Fill".to_string(),
            "Glow".to_string(),
            "Line".to_string(),
            "PictureFormat".to_string(),
            "Shadow".to_string(),
            "SoftEdge".to_string(),
            "TextFrame2".to_string(),
            "ThreeD".to_string()
        ]
    );

    assert_eq!(adjustments_coverage.member_count, 5);
    assert_eq!(adjustments_coverage.support_counts.stub, 5);
    assert_eq!(
        adjustments_coverage.stub_members,
        vec![
            "Application".to_string(),
            "Count".to_string(),
            "Creator".to_string(),
            "Item".to_string(),
            "Parent".to_string()
        ]
    );

    assert_eq!(picture_format_coverage.member_count, 15);
    assert_eq!(picture_format_coverage.support_counts.stub, 15);
    assert_eq!(
        picture_format_coverage.stub_members,
        vec![
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Brightness".to_string(),
            "Contrast".to_string(),
            "ColorType".to_string(),
            "Crop".to_string(),
            "IncrementBrightness".to_string(),
            "IncrementContrast".to_string(),
            "CropLeft".to_string(),
            "CropTop".to_string(),
            "CropRight".to_string(),
            "CropBottom".to_string(),
            "TransparentBackground".to_string(),
            "TransparencyColor".to_string()
        ]
    );
    assert_eq!(crop_coverage.member_count, 11);
    assert_eq!(crop_coverage.support_counts.stub, 11);
    assert_eq!(
        crop_coverage.stub_members,
        vec![
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "PictureHeight".to_string(),
            "PictureOffsetX".to_string(),
            "PictureOffsetY".to_string(),
            "PictureWidth".to_string(),
            "ShapeHeight".to_string(),
            "ShapeLeft".to_string(),
            "ShapeTop".to_string(),
            "ShapeWidth".to_string()
        ]
    );
    assert_eq!(shadow_format_coverage.member_count, 10);
    assert_eq!(shadow_format_coverage.support_counts.stub, 10);
    assert_eq!(
        shadow_format_coverage.stub_members,
        vec![
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Blur".to_string(),
            "Obscured".to_string(),
            "OffsetX".to_string(),
            "OffsetY".to_string(),
            "Size".to_string(),
            "Visible".to_string(),
            "Transparency".to_string()
        ]
    );
    assert_eq!(three_d_format_coverage.member_count, 18);
    assert_eq!(three_d_format_coverage.support_counts.stub, 18);
    assert_eq!(
        three_d_format_coverage.stub_members,
        vec![
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Visible".to_string(),
            "Depth".to_string(),
            "BevelBottomDepth".to_string(),
            "BevelBottomInset".to_string(),
            "BevelTopDepth".to_string(),
            "BevelTopInset".to_string(),
            "ContourWidth".to_string(),
            "FieldOfView".to_string(),
            "LightAngle".to_string(),
            "Perspective".to_string(),
            "ProjectText".to_string(),
            "RotationX".to_string(),
            "RotationY".to_string(),
            "RotationZ".to_string(),
            "Z".to_string()
        ]
    );
    assert_eq!(glow_format_coverage.member_count, 6);
    assert_eq!(glow_format_coverage.support_counts.stub, 6);
    assert_eq!(
        glow_format_coverage.stub_members,
        vec![
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Radius".to_string(),
            "Transparency".to_string(),
            "Visible".to_string()
        ]
    );
    assert_eq!(soft_edge_format_coverage.member_count, 5);
    assert_eq!(soft_edge_format_coverage.support_counts.stub, 5);
    assert_eq!(
        soft_edge_format_coverage.stub_members,
        vec![
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Radius".to_string(),
            "Type".to_string()
        ]
    );
    assert_eq!(text_frame2_coverage.member_count, 14);
    assert_eq!(text_frame2_coverage.support_counts.stub, 14);
    assert_eq!(
        text_frame2_coverage.stub_members,
        vec![
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "AutoSize".to_string(),
            "HasText".to_string(),
            "HorizontalAnchor".to_string(),
            "MarginBottom".to_string(),
            "MarginLeft".to_string(),
            "MarginRight".to_string(),
            "MarginTop".to_string(),
            "NoTextRotation".to_string(),
            "Orientation".to_string(),
            "VerticalAnchor".to_string(),
            "WordWrap".to_string()
        ]
    );
    assert_eq!(fill_format_coverage.member_count, 5);
    assert_eq!(fill_format_coverage.support_counts.stub, 5);
    assert_eq!(
        fill_format_coverage.stub_members,
        vec![
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Visible".to_string(),
            "Transparency".to_string()
        ]
    );
    assert_eq!(line_format_coverage.member_count, 15);
    assert_eq!(line_format_coverage.support_counts.stub, 15);
    assert_eq!(
        line_format_coverage.stub_members,
        vec![
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "BeginArrowheadLength".to_string(),
            "BeginArrowheadStyle".to_string(),
            "BeginArrowheadWidth".to_string(),
            "DashStyle".to_string(),
            "EndArrowheadLength".to_string(),
            "EndArrowheadStyle".to_string(),
            "EndArrowheadWidth".to_string(),
            "InsetPen".to_string(),
            "Style".to_string(),
            "Weight".to_string(),
            "Visible".to_string(),
            "Transparency".to_string()
        ]
    );

    assert_eq!(chart_groups_coverage.member_count, 5);
    assert_eq!(chart_groups_coverage.support_counts.stub, 5);
    assert_eq!(
        chart_groups_coverage.stub_members,
        vec![
            "Count".to_string(),
            "Item".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string()
        ]
    );

    assert_eq!(chart_group_coverage.member_count, 33);
    assert_eq!(chart_group_coverage.support_counts.stub, 33);
    assert_eq!(
        chart_group_coverage.stub_members,
        vec![
            "ChartType".to_string(),
            "Index".to_string(),
            "AxisGroup".to_string(),
            "SeriesLines".to_string(),
            "DropLines".to_string(),
            "HiLoLines".to_string(),
            "UpBars".to_string(),
            "DownBars".to_string(),
            "RadarAxisLabels".to_string(),
            "VaryByCategories".to_string(),
            "GapWidth".to_string(),
            "Overlap".to_string(),
            "HasRadarAxisLabels".to_string(),
            "HasSeriesLines".to_string(),
            "HasDropLines".to_string(),
            "HasHiLoLines".to_string(),
            "HasUpDownBars".to_string(),
            "FirstSliceAngle".to_string(),
            "Explosion".to_string(),
            "BubbleScale".to_string(),
            "ShowNegativeBubbles".to_string(),
            "Has3DShading".to_string(),
            "DoughnutHoleSize".to_string(),
            "SecondPlotSize".to_string(),
            "SizeRepresents".to_string(),
            "SplitType".to_string(),
            "SplitValue".to_string(),
            "SeriesCollection".to_string(),
            "CategoryCollection".to_string(),
            "FullCategoryCollection".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string()
        ]
    );

    assert_eq!(category_collection_coverage.member_count, 5);
    assert_eq!(category_collection_coverage.support_counts.stub, 5);
    assert_eq!(
        category_collection_coverage.stub_members,
        vec![
            "Count".to_string(),
            "Item".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string()
        ]
    );

    assert_eq!(chart_category_coverage.member_count, 6);
    assert_eq!(chart_category_coverage.support_counts.stub, 6);
    assert_eq!(
        chart_category_coverage.stub_members,
        vec![
            "Name".to_string(),
            "Index".to_string(),
            "IsFiltered".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string()
        ]
    );

    for coverage_entry in [
        series_lines_coverage,
        drop_lines_coverage,
        hi_lo_lines_coverage,
        up_bars_coverage,
        down_bars_coverage,
    ] {
        assert_eq!(coverage_entry.member_count, 10);
        assert_eq!(coverage_entry.support_counts.stub, 10);
        assert_eq!(
            coverage_entry.stub_members,
            vec![
                "Name".to_string(),
                "Format".to_string(),
                "Border".to_string(),
                "Creator".to_string(),
                "Application".to_string(),
                "Parent".to_string(),
                "Select".to_string(),
                "Delete".to_string(),
                "ClearFormats".to_string(),
                "Copy".to_string()
            ]
        );
    }

    assert_eq!(axes_coverage.member_count, 5);
    assert_eq!(axes_coverage.support_counts.stub, 5);
    assert_eq!(
        axes_coverage.stub_members,
        vec![
            "Count".to_string(),
            "Item".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string()
        ]
    );

    assert_eq!(axis_coverage.member_count, 45);
    assert_eq!(axis_coverage.support_counts.stub, 45);
    assert_eq!(
        axis_coverage.stub_members,
        vec![
            "Type".to_string(),
            "Format".to_string(),
            "Border".to_string(),
            "AxisGroup".to_string(),
            "AxisBetweenCategories".to_string(),
            "CategoryType".to_string(),
            "DisplayUnit".to_string(),
            "DisplayUnitCustom".to_string(),
            "HasDisplayUnitLabel".to_string(),
            "DisplayUnitLabel".to_string(),
            "BaseUnit".to_string(),
            "BaseUnitIsAuto".to_string(),
            "HasTitle".to_string(),
            "HasMajorGridlines".to_string(),
            "HasMinorGridlines".to_string(),
            "MajorGridlines".to_string(),
            "MinorGridlines".to_string(),
            "AxisTitle".to_string(),
            "ReversePlotOrder".to_string(),
            "ScaleType".to_string(),
            "LogBase".to_string(),
            "Crosses".to_string(),
            "CrossesAt".to_string(),
            "MinimumScale".to_string(),
            "MaximumScale".to_string(),
            "MajorUnit".to_string(),
            "MinorUnit".to_string(),
            "MinimumScaleIsAuto".to_string(),
            "MaximumScaleIsAuto".to_string(),
            "MajorUnitIsAuto".to_string(),
            "MinorUnitIsAuto".to_string(),
            "MajorUnitScale".to_string(),
            "MinorUnitScale".to_string(),
            "MajorTickMark".to_string(),
            "MinorTickMark".to_string(),
            "TickLabelPosition".to_string(),
            "TickLabels".to_string(),
            "TickLabelSpacing".to_string(),
            "TickLabelSpacingIsAuto".to_string(),
            "TickMarkSpacing".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Select".to_string(),
            "Delete".to_string()
        ]
    );

    assert_eq!(tick_labels_coverage.member_count, 16);
    assert_eq!(tick_labels_coverage.support_counts.stub, 16);
    assert_eq!(
        tick_labels_coverage.stub_members,
        vec![
            "Name".to_string(),
            "AutoScaleFont".to_string(),
            "Depth".to_string(),
            "Format".to_string(),
            "NumberFormat".to_string(),
            "NumberFormatLocal".to_string(),
            "MultiLevel".to_string(),
            "NumberFormatLinked".to_string(),
            "Offset".to_string(),
            "Orientation".to_string(),
            "ReadingOrder".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Select".to_string(),
            "Delete".to_string()
        ]
    );

    assert_eq!(gridlines_coverage.member_count, 8);
    assert_eq!(gridlines_coverage.support_counts.stub, 8);
    assert_eq!(
        gridlines_coverage.stub_members,
        vec![
            "Name".to_string(),
            "Format".to_string(),
            "Border".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Select".to_string(),
            "Delete".to_string()
        ]
    );

    for coverage_entry in [display_unit_label_coverage, axis_title_coverage] {
        assert_eq!(coverage_entry.member_count, 16);
        assert_eq!(coverage_entry.support_counts.stub, 16);
        assert_eq!(
            coverage_entry.stub_members,
            vec![
                "Name".to_string(),
                "Format".to_string(),
                "Border".to_string(),
                "Text".to_string(),
                "Caption".to_string(),
                "Left".to_string(),
                "Top".to_string(),
                "Width".to_string(),
                "Height".to_string(),
                "Orientation".to_string(),
                "ReadingOrder".to_string(),
                "Creator".to_string(),
                "Application".to_string(),
                "Parent".to_string(),
                "Select".to_string(),
                "Delete".to_string()
            ]
        );
    }

    assert_eq!(series_collection_coverage.member_count, 7);
    assert_eq!(series_collection_coverage.support_counts.stub, 7);
    assert_eq!(
        series_collection_coverage.stub_members,
        vec![
            "Count".to_string(),
            "Item".to_string(),
            "NewSeries".to_string(),
            "Add".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string()
        ]
    );

    assert_eq!(series_coverage.member_count, 28);
    assert_eq!(series_coverage.support_counts.stub, 28);
    assert_eq!(
        series_coverage.stub_members,
        vec![
            "Name".to_string(),
            "Format".to_string(),
            "ChartType".to_string(),
            "Values".to_string(),
            "XValues".to_string(),
            "BubbleSizes".to_string(),
            "BarShape".to_string(),
            "Smooth".to_string(),
            "MarkerStyle".to_string(),
            "MarkerSize".to_string(),
            "Formula".to_string(),
            "AxisGroup".to_string(),
            "HasDataLabels".to_string(),
            "HasLeaderLines".to_string(),
            "LeaderLines".to_string(),
            "DataLabels".to_string(),
            "Points".to_string(),
            "PlotOrder".to_string(),
            "InvertIfNegative".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Select".to_string(),
            "ApplyDataLabels".to_string(),
            "ClearFormats".to_string(),
            "Copy".to_string(),
            "Paste".to_string(),
            "Delete".to_string()
        ]
    );

    assert_eq!(leader_lines_coverage.member_count, 7);
    assert_eq!(leader_lines_coverage.support_counts.stub, 7);
    assert_eq!(
        leader_lines_coverage.stub_members,
        vec![
            "Format".to_string(),
            "Border".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Select".to_string(),
            "Delete".to_string()
        ]
    );
    assert_eq!(border_coverage.member_count, 9);
    assert_eq!(border_coverage.support_counts.stub, 9);
    assert_eq!(
        border_coverage.stub_members,
        vec![
            "Application".to_string(),
            "Color".to_string(),
            "ColorIndex".to_string(),
            "Creator".to_string(),
            "LineStyle".to_string(),
            "Parent".to_string(),
            "ThemeColor".to_string(),
            "TintAndShade".to_string(),
            "Weight".to_string()
        ]
    );

    assert_eq!(data_labels_coverage.member_count, 27);
    assert_eq!(data_labels_coverage.support_counts.stub, 27);
    assert_eq!(
        data_labels_coverage.stub_members,
        vec![
            "Name".to_string(),
            "Format".to_string(),
            "Border".to_string(),
            "Count".to_string(),
            "Item".to_string(),
            "Type".to_string(),
            "ShowLegendKey".to_string(),
            "HasLeaderLines".to_string(),
            "ShowSeriesName".to_string(),
            "ShowCategoryName".to_string(),
            "ShowValue".to_string(),
            "ShowPercentage".to_string(),
            "ShowBubbleSize".to_string(),
            "NumberFormat".to_string(),
            "NumberFormatLocal".to_string(),
            "NumberFormatLinked".to_string(),
            "Position".to_string(),
            "Separator".to_string(),
            "Orientation".to_string(),
            "ReadingOrder".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Delete".to_string(),
            "Propagate".to_string(),
            "Select".to_string(),
            "ClearFormats".to_string()
        ]
    );

    assert_eq!(data_label_coverage.member_count, 25);
    assert_eq!(data_label_coverage.support_counts.stub, 25);
    assert_eq!(
        data_label_coverage.stub_members,
        vec![
            "Name".to_string(),
            "Format".to_string(),
            "Border".to_string(),
            "Index".to_string(),
            "Type".to_string(),
            "ShowLegendKey".to_string(),
            "HasLeaderLines".to_string(),
            "ShowSeriesName".to_string(),
            "ShowCategoryName".to_string(),
            "ShowValue".to_string(),
            "ShowPercentage".to_string(),
            "ShowBubbleSize".to_string(),
            "NumberFormat".to_string(),
            "NumberFormatLocal".to_string(),
            "NumberFormatLinked".to_string(),
            "Position".to_string(),
            "Separator".to_string(),
            "Orientation".to_string(),
            "ReadingOrder".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "Delete".to_string(),
            "Select".to_string(),
            "ClearFormats".to_string()
        ]
    );

    assert_eq!(points_coverage.member_count, 5);
    assert_eq!(points_coverage.support_counts.stub, 5);
    assert_eq!(
        points_coverage.stub_members,
        vec![
            "Count".to_string(),
            "Item".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string()
        ]
    );

    assert_eq!(point_coverage.member_count, 14);
    assert_eq!(point_coverage.support_counts.stub, 14);
    assert_eq!(
        point_coverage.stub_members,
        vec![
            "Name".to_string(),
            "Format".to_string(),
            "Border".to_string(),
            "Index".to_string(),
            "Explosion".to_string(),
            "HasDataLabel".to_string(),
            "DataLabel".to_string(),
            "Creator".to_string(),
            "Application".to_string(),
            "Parent".to_string(),
            "ApplyDataLabels".to_string(),
            "Copy".to_string(),
            "Select".to_string(),
            "ClearFormats".to_string()
        ]
    );
}
