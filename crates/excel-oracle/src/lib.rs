//! Cross-platform contracts for behavioral comparison between desktop Excel and OOTD.

use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, BTreeSet};
use std::fmt::{Display, Formatter};

pub const ORACLE_SCHEMA_VERSION: u32 = 1;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OracleContractError {
    message: String,
}

impl OracleContractError {
    fn new(message: impl Into<String>) -> Self {
        Self {
            message: message.into(),
        }
    }
}

impl Display for OracleContractError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(&self.message)
    }
}

impl std::error::Error for OracleContractError {}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum CaseTier {
    MustMatch,
    Informational,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CaseProvenance {
    pub source: String,
    pub producer: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CaseInput {
    pub path: String,
    pub sha256: String,
    pub provenance: CaseProvenance,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "operation", rename_all = "camelCase", deny_unknown_fields)]
pub enum CaseOperation {
    Get {
        target: String,
        member: String,
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        args: Vec<ObservedValue>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        bind: Option<String>,
    },
    Set {
        target: String,
        member: String,
        value: ObservedValue,
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        args: Vec<ObservedValue>,
    },
    Invoke {
        target: String,
        member: String,
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        args: Vec<ObservedValue>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        bind: Option<String>,
    },
    Calculate,
    Save {
        workbook: String,
        output: String,
    },
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CaseProbe {
    pub id: String,
    pub target: String,
    pub member: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub args: Vec<ObservedValue>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CaseSpec {
    pub schema_version: u32,
    pub id: String,
    pub version: u32,
    pub tier: CaseTier,
    pub input: CaseInput,
    pub profile_id: String,
    pub operations: Vec<CaseOperation>,
    pub probes: Vec<CaseProbe>,
}

impl CaseSpec {
    pub fn from_json_str(input: &str) -> Result<Self, OracleContractError> {
        let value: Self = serde_json::from_str(input)
            .map_err(|error| OracleContractError::new(format!("invalid case JSON: {error}")))?;
        value.validate()?;
        Ok(value)
    }

    pub fn to_json_pretty(&self) -> Result<String, OracleContractError> {
        self.validate()?;
        serde_json::to_string_pretty(self)
            .map_err(|error| OracleContractError::new(format!("failed to serialize case: {error}")))
    }

    pub fn validate(&self) -> Result<(), OracleContractError> {
        if self.schema_version != ORACLE_SCHEMA_VERSION {
            return Err(OracleContractError::new(format!(
                "unsupported case schemaVersion {}",
                self.schema_version
            )));
        }
        if self.id.is_empty()
            || self.id.trim() != self.id
            || !self
                .id
                .bytes()
                .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'))
        {
            return Err(OracleContractError::new(
                "case id must be a trimmed ASCII identifier",
            ));
        }
        if self.version == 0 {
            return Err(OracleContractError::new("case version must be positive"));
        }
        if self.profile_id.is_empty()
            || self.profile_id.trim() != self.profile_id
            || !self
                .profile_id
                .bytes()
                .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'))
        {
            return Err(OracleContractError::new(
                "case profileId must be a trimmed ASCII identifier",
            ));
        }
        if self.input.path.is_empty()
            || self.input.path.starts_with('/')
            || self.input.path.contains('\\')
            || self
                .input
                .path
                .split('/')
                .any(|segment| segment.is_empty() || matches!(segment, "." | ".."))
        {
            return Err(OracleContractError::new(
                "input path must be a safe forward-slash relative path",
            ));
        }
        if self.input.sha256.len() != 64
            || !self
                .input
                .sha256
                .bytes()
                .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
        {
            return Err(OracleContractError::new(
                "input sha256 must contain 64 lowercase hexadecimal characters",
            ));
        }
        for (label, value) in [
            ("source", self.input.provenance.source.as_str()),
            ("producer", self.input.provenance.producer.as_str()),
        ] {
            if value.is_empty() || value.trim() != value {
                return Err(OracleContractError::new(format!(
                    "input provenance {label} must be non-empty and trimmed"
                )));
            }
        }
        if self.operations.is_empty() {
            return Err(OracleContractError::new(
                "case operations must not be empty",
            ));
        }
        let mut bindings = BTreeSet::new();
        for operation in &self.operations {
            match operation {
                CaseOperation::Get {
                    target,
                    member,
                    args,
                    bind,
                }
                | CaseOperation::Invoke {
                    target,
                    member,
                    args,
                    bind,
                } => {
                    if target.is_empty() || member.is_empty() {
                        return Err(OracleContractError::new(
                            "operation target and member must not be empty",
                        ));
                    }
                    for value in args {
                        value.validate()?;
                    }
                    if let Some(binding) = bind
                        && (binding.is_empty() || !bindings.insert(binding.clone()))
                    {
                        return Err(OracleContractError::new(
                            "operation bindings must be non-empty and unique",
                        ));
                    }
                }
                CaseOperation::Set {
                    target,
                    member,
                    value,
                    args,
                } => {
                    if target.is_empty() || member.is_empty() {
                        return Err(OracleContractError::new(
                            "operation target and member must not be empty",
                        ));
                    }
                    value.validate()?;
                    for value in args {
                        value.validate()?;
                    }
                }
                CaseOperation::Calculate => {}
                CaseOperation::Save { workbook, output } => {
                    if workbook.is_empty()
                        || output.is_empty()
                        || output.starts_with('/')
                        || output.contains('\\')
                        || output
                            .split('/')
                            .any(|segment| segment.is_empty() || matches!(segment, "." | ".."))
                    {
                        return Err(OracleContractError::new(
                            "save workbook and output must use safe relative identifiers",
                        ));
                    }
                }
            }
        }
        if self.probes.is_empty() {
            return Err(OracleContractError::new("case probes must not be empty"));
        }
        let mut probe_ids = BTreeSet::new();
        for probe in &self.probes {
            if probe.id.is_empty()
                || probe.target.is_empty()
                || probe.member.is_empty()
                || !probe_ids.insert(probe.id.clone())
            {
                return Err(OracleContractError::new(
                    "probe ids must be unique and probe fields must not be empty",
                ));
            }
            for value in &probe.args {
                value.validate()?;
            }
        }
        Ok(())
    }

    fn has_save_operation(&self) -> bool {
        self.operations
            .iter()
            .any(|operation| matches!(operation, CaseOperation::Save { .. }))
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum EngineKind {
    Excel,
    Ootd,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct EngineIdentity {
    pub kind: EngineKind,
    pub version: String,
    pub build: String,
    pub channel: String,
    pub os: String,
    pub architecture: String,
    pub locale: String,
    pub timezone: String,
}

impl EngineIdentity {
    fn validate(&self) -> Result<(), OracleContractError> {
        for (label, value) in [
            ("version", self.version.as_str()),
            ("build", self.build.as_str()),
            ("channel", self.channel.as_str()),
            ("os", self.os.as_str()),
            ("architecture", self.architecture.as_str()),
            ("locale", self.locale.as_str()),
            ("timezone", self.timezone.as_str()),
        ] {
            if value.is_empty() || value.trim() != value {
                return Err(OracleContractError::new(format!(
                    "engine {label} must be non-empty and trimmed"
                )));
            }
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ObservedErrorKind {
    ExcelCom,
    Ootd,
    Runner,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum CanonicalErrorKind {
    InvalidArgument,
    NotFound,
    TypeMismatch,
    Unsupported,
    InvalidState,
    Io,
    Parse,
    Calculation,
    External,
    ApplicationDefined,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeErrorDiagnostic {
    pub origin: ObservedErrorKind,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hresult: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub message: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ObservedError {
    pub kind: CanonicalErrorKind,
    pub code: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub diagnostic: Option<NativeErrorDiagnostic>,
}

impl ObservedError {
    fn validate(&self) -> Result<(), OracleContractError> {
        if self.code.is_empty() || self.code.trim() != self.code {
            return Err(OracleContractError::new(
                "observed error codes must be non-empty and trimmed",
            ));
        }
        if self
            .diagnostic
            .as_ref()
            .and_then(|diagnostic| diagnostic.message.as_deref())
            .is_some_and(|message| message.is_empty() || message.trim() != message)
        {
            return Err(OracleContractError::new(
                "native error diagnostic messages must be non-empty and trimmed",
            ));
        }
        Ok(())
    }

    fn semantically_eq(&self, other: &Self) -> bool {
        self.kind == other.kind && self.code == other.code
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ObservedCellError {
    pub code: String,
    pub cv_err: u16,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ObservedObject {
    pub type_name: String,
    pub identity: String,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ObservedArray {
    pub rows: usize,
    pub cols: usize,
    pub values: Vec<ObservedValue>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "type",
    content = "value",
    rename_all = "camelCase",
    deny_unknown_fields
)]
pub enum ObservedValue {
    Void,
    Missing,
    Empty,
    Null,
    Bool(bool),
    Number(f64),
    Text(String),
    CellError(ObservedCellError),
    Object(ObservedObject),
    Array(ObservedArray),
}

impl ObservedValue {
    fn validate(&self) -> Result<(), OracleContractError> {
        match self {
            Self::Number(value) if !value.is_finite() => {
                Err(OracleContractError::new("observed numbers must be finite"))
            }
            Self::Object(value)
                if value.type_name.is_empty()
                    || value.type_name.trim() != value.type_name
                    || value.identity.is_empty()
                    || value.identity.trim() != value.identity =>
            {
                Err(OracleContractError::new(
                    "observed object type names and identities must be non-empty and trimmed",
                ))
            }
            Self::CellError(error)
                if error.code.is_empty()
                    || error.code.trim() != error.code
                    || error.cv_err == 0 =>
            {
                Err(OracleContractError::new(
                    "observed cell errors require a trimmed code and positive cvErr",
                ))
            }
            Self::Array(array) => {
                if array.rows == 0
                    || array.cols == 0
                    || array.rows.checked_mul(array.cols) != Some(array.values.len())
                {
                    return Err(OracleContractError::new(
                        "observed array dimensions must match non-empty values",
                    ));
                }
                for value in &array.values {
                    value.validate()?;
                }
                Ok(())
            }
            _ => Ok(()),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "status",
    content = "result",
    rename_all = "camelCase",
    deny_unknown_fields
)]
pub enum ObservationResult {
    Value(ObservedValue),
    Error(ObservedError),
}

impl ObservationResult {
    fn validate(&self) -> Result<(), OracleContractError> {
        match self {
            Self::Value(value) => value.validate(),
            Self::Error(error) => error.validate(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct OperationObservation {
    pub operation_index: usize,
    pub result: ObservationResult,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ProbeObservation {
    pub id: String,
    pub result: ObservationResult,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SaveReopenObservation {
    pub attempted: bool,
    pub normal_load_succeeded: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub repair_detected: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub evidence: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ObservationDocument {
    pub schema_version: u32,
    pub case_id: String,
    pub engine: EngineIdentity,
    pub operations: Vec<OperationObservation>,
    pub probes: Vec<ProbeObservation>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub save_reopen: Option<SaveReopenObservation>,
}

impl ObservationDocument {
    pub fn validate_for_case(&self, case: &CaseSpec) -> Result<(), OracleContractError> {
        case.validate()?;
        if self.schema_version != ORACLE_SCHEMA_VERSION {
            return Err(OracleContractError::new(format!(
                "unsupported observation schemaVersion {}",
                self.schema_version
            )));
        }
        if self.case_id != case.id {
            return Err(OracleContractError::new(
                "observation caseId did not match the case",
            ));
        }
        self.engine.validate()?;
        if self.operations.len() != case.operations.len() {
            return Err(OracleContractError::new(
                "operation observations must exactly cover the case operations",
            ));
        }
        for (expected_index, operation) in self.operations.iter().enumerate() {
            if operation.operation_index != expected_index {
                return Err(OracleContractError::new(
                    "operation observations must exactly cover ordered indexes from zero",
                ));
            }
            operation.result.validate()?;
        }
        let expected_probe_ids = case
            .probes
            .iter()
            .map(|probe| probe.id.as_str())
            .collect::<BTreeSet<_>>();
        let mut actual_probe_ids = BTreeSet::new();
        for probe in &self.probes {
            if probe.id.is_empty() || !actual_probe_ids.insert(probe.id.as_str()) {
                return Err(OracleContractError::new(
                    "probe observations must have unique non-empty ids",
                ));
            }
            probe.result.validate()?;
        }
        if expected_probe_ids != actual_probe_ids {
            return Err(OracleContractError::new(
                "probe observations must exactly cover the case probes",
            ));
        }
        if case.has_save_operation() {
            let save_reopen = self.save_reopen.as_ref().ok_or_else(|| {
                OracleContractError::new("save cases require a saveReopen observation")
            })?;
            if !save_reopen.attempted {
                return Err(OracleContractError::new(
                    "saveReopen must record an attempted reopen",
                ));
            }
            if !save_reopen.normal_load_succeeded && save_reopen.repair_detected == Some(false) {
                return Err(OracleContractError::new(
                    "repairDetected=false requires a successful normal load",
                ));
            }
            if save_reopen
                .evidence
                .as_deref()
                .is_some_and(|evidence| evidence.is_empty() || evidence.trim() != evidence)
            {
                return Err(OracleContractError::new(
                    "saveReopen evidence must be non-empty and trimmed",
                ));
            }
            if case.tier == CaseTier::MustMatch {
                if !save_reopen.normal_load_succeeded {
                    return Err(OracleContractError::new(
                        "mustMatch save cases require a successful normal load",
                    ));
                }
                if save_reopen.repair_detected != Some(false) {
                    return Err(OracleContractError::new(
                        "mustMatch save cases require repairDetected=false",
                    ));
                }
                if save_reopen.evidence.is_none() {
                    return Err(OracleContractError::new(
                        "mustMatch save cases require normal-load evidence",
                    ));
                }
            }
        } else if self.save_reopen.is_some() {
            return Err(OracleContractError::new(
                "saveReopen is only valid for cases with a save operation",
            ));
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ComparisonPolicy {
    pub number_tolerance: f64,
}

impl Default for ComparisonPolicy {
    fn default() -> Self {
        Self {
            number_tolerance: 0.0,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ObservationMismatch {
    pub path: String,
    pub expected: String,
    pub actual: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ObservationComparison {
    pub case_id: String,
    pub mismatches: Vec<ObservationMismatch>,
}

impl ObservationComparison {
    pub fn passed(&self) -> bool {
        self.mismatches.is_empty()
    }
}

pub fn compare_observations(
    case: &CaseSpec,
    oracle: &ObservationDocument,
    runtime: &ObservationDocument,
    policy: ComparisonPolicy,
) -> Result<ObservationComparison, OracleContractError> {
    if !policy.number_tolerance.is_finite() || policy.number_tolerance < 0.0 {
        return Err(OracleContractError::new(
            "comparison number tolerance must be finite and non-negative",
        ));
    }
    oracle.validate_for_case(case)?;
    runtime.validate_for_case(case)?;
    if oracle.engine.kind != EngineKind::Excel {
        return Err(OracleContractError::new(
            "oracle observation must use the Excel engine kind",
        ));
    }
    if runtime.engine.kind != EngineKind::Ootd {
        return Err(OracleContractError::new(
            "runtime observation must use the OOTD engine kind",
        ));
    }

    let mut mismatches = Vec::new();
    compare_operation_results(oracle, runtime, policy, &mut mismatches);
    compare_probe_results(oracle, runtime, policy, &mut mismatches);
    if oracle.save_reopen != runtime.save_reopen {
        mismatches.push(ObservationMismatch {
            path: "saveReopen".to_string(),
            expected: format!("{:?}", oracle.save_reopen),
            actual: format!("{:?}", runtime.save_reopen),
        });
    }

    Ok(ObservationComparison {
        case_id: case.id.clone(),
        mismatches,
    })
}

fn compare_operation_results(
    oracle: &ObservationDocument,
    runtime: &ObservationDocument,
    policy: ComparisonPolicy,
    mismatches: &mut Vec<ObservationMismatch>,
) {
    let expected = oracle
        .operations
        .iter()
        .map(|observation| (observation.operation_index, &observation.result))
        .collect::<BTreeMap<_, _>>();
    let actual = runtime
        .operations
        .iter()
        .map(|observation| (observation.operation_index, &observation.result))
        .collect::<BTreeMap<_, _>>();
    for index in expected
        .keys()
        .chain(actual.keys())
        .copied()
        .collect::<BTreeSet<_>>()
    {
        match (expected.get(&index), actual.get(&index)) {
            (Some(expected), Some(actual)) => compare_result(
                &format!("operations.{index}"),
                expected,
                actual,
                policy,
                mismatches,
            ),
            (Some(expected), None) => mismatches.push(ObservationMismatch {
                path: format!("operations.{index}"),
                expected: format!("{expected:?}"),
                actual: "missing".to_string(),
            }),
            (None, Some(actual)) => mismatches.push(ObservationMismatch {
                path: format!("operations.{index}"),
                expected: "missing".to_string(),
                actual: format!("{actual:?}"),
            }),
            (None, None) => unreachable!(),
        }
    }
}

fn compare_probe_results(
    oracle: &ObservationDocument,
    runtime: &ObservationDocument,
    policy: ComparisonPolicy,
    mismatches: &mut Vec<ObservationMismatch>,
) {
    let expected = oracle
        .probes
        .iter()
        .map(|observation| (observation.id.as_str(), &observation.result))
        .collect::<BTreeMap<_, _>>();
    let actual = runtime
        .probes
        .iter()
        .map(|observation| (observation.id.as_str(), &observation.result))
        .collect::<BTreeMap<_, _>>();
    for (id, expected) in expected {
        compare_result(
            &format!("probes.{id}"),
            expected,
            actual[id],
            policy,
            mismatches,
        );
    }
}

fn compare_result(
    path: &str,
    expected: &ObservationResult,
    actual: &ObservationResult,
    policy: ComparisonPolicy,
    mismatches: &mut Vec<ObservationMismatch>,
) {
    match (expected, actual) {
        (ObservationResult::Value(expected), ObservationResult::Value(actual)) => {
            compare_value(
                &format!("{path}.value"),
                expected,
                actual,
                policy,
                mismatches,
            );
        }
        (ObservationResult::Error(expected), ObservationResult::Error(actual))
            if expected.semantically_eq(actual) => {}
        _ if expected == actual => {}
        _ => mismatches.push(ObservationMismatch {
            path: path.to_string(),
            expected: format!("{expected:?}"),
            actual: format!("{actual:?}"),
        }),
    }
}

fn compare_value(
    path: &str,
    expected: &ObservedValue,
    actual: &ObservedValue,
    policy: ComparisonPolicy,
    mismatches: &mut Vec<ObservationMismatch>,
) {
    match (expected, actual) {
        (ObservedValue::Number(expected), ObservedValue::Number(actual))
            if (expected - actual).abs() <= policy.number_tolerance => {}
        (ObservedValue::Array(expected), ObservedValue::Array(actual)) => {
            if expected.rows != actual.rows {
                mismatches.push(ObservationMismatch {
                    path: format!("{path}.rows"),
                    expected: expected.rows.to_string(),
                    actual: actual.rows.to_string(),
                });
            }
            if expected.cols != actual.cols {
                mismatches.push(ObservationMismatch {
                    path: format!("{path}.cols"),
                    expected: expected.cols.to_string(),
                    actual: actual.cols.to_string(),
                });
            }
            for (index, (expected, actual)) in
                expected.values.iter().zip(&actual.values).enumerate()
            {
                compare_value(
                    &format!("{path}.values.{index}"),
                    expected,
                    actual,
                    policy,
                    mismatches,
                );
            }
            if expected.values.len() != actual.values.len() {
                mismatches.push(ObservationMismatch {
                    path: format!("{path}.values.length"),
                    expected: expected.values.len().to_string(),
                    actual: actual.values.len().to_string(),
                });
            }
        }
        _ if expected == actual => {}
        _ => mismatches.push(ObservationMismatch {
            path: path.to_string(),
            expected: format!("{expected:?}"),
            actual: format!("{actual:?}"),
        }),
    }
}
