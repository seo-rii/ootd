use std::collections::{BTreeMap, BTreeSet};

use office_codegen::{DifferentialCaseResult, DifferentialCaseStatus, DifferentialReport};

use crate::{
    CaseSpec, CaseTier, ComparisonPolicy, EngineKind, ObservationDocument, OracleContractError,
    OracleSuiteManifest, RunCaseRecord, RunCaseStatus, RunManifest, compare_observations,
};

#[derive(Debug, Clone, PartialEq)]
pub struct RunBundle {
    pub manifest: RunManifest,
    pub observations: BTreeMap<String, Vec<u8>>,
}

impl RunBundle {
    pub(crate) fn validate(&self, suite: &OracleSuiteManifest) -> Result<(), OracleContractError> {
        self.manifest.validate_for_suite(suite)?;
        self.validate_observation_coverage()
    }

    pub(crate) fn validate_fragment(
        &self,
        suite: &OracleSuiteManifest,
    ) -> Result<(), OracleContractError> {
        self.manifest.validate_fragment_for_suite(suite)?;
        self.validate_observation_coverage()
    }

    fn validate_observation_coverage(&self) -> Result<(), OracleContractError> {
        let expected = self
            .manifest
            .cases
            .iter()
            .filter(|record| record.status == RunCaseStatus::Completed)
            .map(|record| record.case_id.as_str())
            .collect::<BTreeSet<_>>();
        let actual = self
            .observations
            .keys()
            .map(String::as_str)
            .collect::<BTreeSet<_>>();
        if expected != actual {
            return Err(OracleContractError::new(
                "run observation artifacts must exactly cover completed cases",
            ));
        }
        Ok(())
    }

    fn record(&self, case_id: &str) -> &RunCaseRecord {
        self.manifest
            .cases
            .iter()
            .find(|record| record.case_id == case_id)
            .expect("validated exact run coverage")
    }

    fn observation(&self, case: &CaseSpec) -> Result<ObservationDocument, OracleContractError> {
        let record = self.record(&case.id);
        let bytes = self
            .observations
            .get(&case.id)
            .expect("validated completed observation coverage");
        record.load_observation(case, &self.manifest.engine, bytes)
    }
}

pub fn build_behavioral_differential_report(
    suite: &OracleSuiteManifest,
    case_artifacts: &BTreeMap<String, Vec<u8>>,
    oracle: &RunBundle,
    runtime: &RunBundle,
    policy: ComparisonPolicy,
) -> Result<DifferentialReport, OracleContractError> {
    suite.validate()?;
    oracle.validate(suite)?;
    runtime.validate(suite)?;
    if oracle.manifest.engine.kind != EngineKind::Excel {
        return Err(OracleContractError::new(
            "behavioral Oracle bundle must use the Excel engine kind",
        ));
    }
    if runtime.manifest.engine.kind != EngineKind::Ootd {
        return Err(OracleContractError::new(
            "behavioral runtime bundle must use the OOTD engine kind",
        ));
    }
    let expected_case_ids = suite
        .cases
        .iter()
        .map(|case| case.case_id.as_str())
        .collect::<BTreeSet<_>>();
    let actual_case_ids = case_artifacts
        .keys()
        .map(String::as_str)
        .collect::<BTreeSet<_>>();
    if expected_case_ids != actual_case_ids {
        return Err(OracleContractError::new(
            "case artifacts must exactly cover the suite cases",
        ));
    }

    let mut results = Vec::with_capacity(suite.cases.len());
    for expected in &suite.cases {
        let case = suite.load_case(
            &expected.case_id,
            &case_artifacts[expected.case_id.as_str()],
        )?;
        let oracle_record = oracle.record(&case.id);
        let runtime_record = runtime.record(&case.id);
        let result = match (oracle_record.status, runtime_record.status) {
            (RunCaseStatus::Completed, RunCaseStatus::Completed) => {
                let oracle_observation = oracle.observation(&case)?;
                let runtime_observation = runtime.observation(&case)?;
                let comparison =
                    compare_observations(&case, &oracle_observation, &runtime_observation, policy)?;
                comparison_result(&case, oracle_record, runtime_record, comparison)
            }
            (RunCaseStatus::Completed, runtime_status) => incomplete_result(
                &case,
                IncompleteSide::Runtime,
                runtime_status,
                runtime_record,
                Some(oracle_record),
            )?,
            (oracle_status, RunCaseStatus::Completed) => incomplete_result(
                &case,
                IncompleteSide::Oracle,
                oracle_status,
                oracle_record,
                Some(runtime_record),
            )?,
            _ => {
                return Err(OracleContractError::new(format!(
                    "case {} was incomplete in both Oracle and runtime bundles",
                    case.id
                )));
            }
        };
        results.push(result);
    }

    let report_profile = suite
        .profile_id
        .chars()
        .map(|character| {
            if character.is_ascii_alphanumeric() || character == '_' {
                character
            } else {
                '_'
            }
        })
        .collect::<String>();
    let report = DifferentialReport::from_cases(
        "Excel",
        suite.expected_engine.version.clone(),
        report_profile,
        results,
    );
    report.validate().map_err(|error| {
        OracleContractError::new(format!("invalid behavioral differential report: {error:?}"))
    })?;
    report.gate_summary().validate().map_err(|error| {
        OracleContractError::new(format!("invalid behavioral differential gate: {error:?}"))
    })?;
    Ok(report)
}

fn comparison_result(
    case: &CaseSpec,
    oracle: &RunCaseRecord,
    runtime: &RunCaseRecord,
    comparison: crate::ObservationComparison,
) -> DifferentialCaseResult {
    let passed = comparison.passed();
    DifferentialCaseResult {
        name: case.id.clone(),
        surface: None,
        member: None,
        status: if passed {
            DifferentialCaseStatus::Passed
        } else {
            DifferentialCaseStatus::Failed
        },
        expected: Some(if passed {
            "matched".to_string()
        } else {
            "see oracleObservation".to_string()
        }),
        actual: Some(if passed {
            "matched".to_string()
        } else {
            "see runtimeObservation".to_string()
        }),
        message: comparison
            .mismatches
            .first()
            .map(|mismatch| format!("first mismatch at {}", mismatch.path)),
        artifacts: observation_artifacts(Some(oracle), Some(runtime)),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum IncompleteSide {
    Oracle,
    Runtime,
}

fn incomplete_result(
    case: &CaseSpec,
    side: IncompleteSide,
    status: RunCaseStatus,
    incomplete: &RunCaseRecord,
    completed: Option<&RunCaseRecord>,
) -> Result<DifferentialCaseResult, OracleContractError> {
    let report_status = if case.tier == CaseTier::MustMatch || status == RunCaseStatus::Failed {
        match side {
            IncompleteSide::Oracle => DifferentialCaseStatus::MissingOracle,
            IncompleteSide::Runtime => DifferentialCaseStatus::MissingRuntime,
        }
    } else {
        match status {
            RunCaseStatus::Unsupported => DifferentialCaseStatus::Unsupported,
            RunCaseStatus::Skipped => DifferentialCaseStatus::Skipped,
            RunCaseStatus::Failed => unreachable!(),
            RunCaseStatus::Completed => {
                return Err(OracleContractError::new(
                    "completed case was routed through incomplete reporting",
                ));
            }
        }
    };
    let completed_summary = completed.map(|_| match side {
        IncompleteSide::Oracle => "see runtimeObservation".to_string(),
        IncompleteSide::Runtime => "see oracleObservation".to_string(),
    });
    let (expected, actual) = match report_status {
        DifferentialCaseStatus::MissingOracle => (None, completed_summary),
        DifferentialCaseStatus::MissingRuntime => (completed_summary, None),
        DifferentialCaseStatus::Unsupported | DifferentialCaseStatus::Skipped => (None, None),
        DifferentialCaseStatus::Passed | DifferentialCaseStatus::Failed => unreachable!(),
    };
    let (oracle, runtime) = match side {
        IncompleteSide::Oracle => (None, completed),
        IncompleteSide::Runtime => (completed, None),
    };
    Ok(DifferentialCaseResult {
        name: case.id.clone(),
        surface: None,
        member: None,
        status: report_status,
        expected,
        actual,
        message: incomplete.message.clone(),
        artifacts: observation_artifacts(oracle, runtime),
    })
}

fn observation_artifacts(
    oracle: Option<&RunCaseRecord>,
    runtime: Option<&RunCaseRecord>,
) -> BTreeMap<String, String> {
    let mut artifacts = BTreeMap::new();
    if let Some(path) = oracle.and_then(|record| record.observation_path.as_ref()) {
        artifacts.insert("oracleObservation".to_string(), path.clone());
    }
    if let Some(path) = runtime.and_then(|record| record.observation_path.as_ref()) {
        artifacts.insert("runtimeObservation".to_string(), path.clone());
    }
    artifacts
}
