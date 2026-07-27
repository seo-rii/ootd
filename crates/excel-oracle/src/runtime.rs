use std::collections::BTreeMap;

use excel_runtime::ExcelRuntime;
use office_common::{
    CellError, ExcelProfile, FileFormat, ObjectHandle, OmArray, OmError, OmErrorCode, OmValue,
    OpenWorkbookSpec,
};

use crate::{
    CanonicalErrorKind, CaseOperation, CaseSpec, EngineIdentity, EngineKind, NativeErrorDiagnostic,
    ObservationDocument, ObservationResult, ObservedArray, ObservedCellError, ObservedError,
    ObservedErrorKind, ObservedObject, ObservedValue, OperationObservation, OracleContractError,
    ProbeObservation, sha256_hex,
};

#[derive(Debug, Clone, PartialEq)]
pub struct RuntimeCaseOutput {
    pub observation: ObservationDocument,
}

pub fn run_runtime_case(
    case: &CaseSpec,
    input_bytes: &[u8],
    engine: EngineIdentity,
) -> Result<RuntimeCaseOutput, OracleContractError> {
    case.validate()?;
    if engine.kind != EngineKind::Ootd {
        return Err(OracleContractError::new(
            "runtime case engine must use the OOTD kind",
        ));
    }
    if sha256_hex(input_bytes) != case.input.sha256 {
        return Err(OracleContractError::new(format!(
            "input sha256 for case {} did not match",
            case.id
        )));
    }
    if case
        .operations
        .iter()
        .any(|operation| matches!(operation, CaseOperation::Save { .. }))
    {
        return Err(OracleContractError::new(
            "runtime save cases require an external Excel reopen verifier",
        ));
    }

    let mut runtime = ExcelRuntime::new();
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: input_bytes.to_vec(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .map_err(|error| runtime_harness_error("open workbook", error))?;
    let application = runtime.root_application();
    let mut objects = BTreeMap::from([
        ("application".to_string(), application),
        ("workbook".to_string(), workbook.0),
    ]);
    let mut identities = BTreeMap::from([
        (application, "application".to_string()),
        (workbook.0, "workbook".to_string()),
    ]);
    let mut next_object_identity = 1_u64;
    let mut operations = Vec::with_capacity(case.operations.len());

    for (operation_index, operation) in case.operations.iter().enumerate() {
        let result = execute_operation(
            &mut runtime,
            operation,
            &mut objects,
            &mut identities,
            &mut next_object_identity,
        )?;
        operations.push(OperationObservation {
            operation_index,
            result,
        });
    }

    let mut probes = Vec::with_capacity(case.probes.len());
    for probe in &case.probes {
        let target = resolve_target(&objects, &probe.target)?;
        let args = observed_args_to_om(&probe.args, &objects)?;
        let result = runtime
            .dispatch_get(target, &probe.member, &args)
            .map(|value| observe_om_value(value, &mut identities, &mut next_object_identity))
            .map(ObservationResult::Value)
            .unwrap_or_else(|error| ObservationResult::Error(observe_om_error(error)));
        probes.push(ProbeObservation {
            id: probe.id.clone(),
            result,
        });
    }

    let observation = ObservationDocument {
        schema_version: crate::ORACLE_SCHEMA_VERSION,
        case_id: case.id.clone(),
        engine,
        operations,
        probes,
        save_reopen: None,
    };
    observation.validate_for_case(case)?;
    Ok(RuntimeCaseOutput { observation })
}

fn execute_operation(
    runtime: &mut ExcelRuntime,
    operation: &CaseOperation,
    objects: &mut BTreeMap<String, ObjectHandle>,
    identities: &mut BTreeMap<ObjectHandle, String>,
    next_object_identity: &mut u64,
) -> Result<ObservationResult, OracleContractError> {
    match operation {
        CaseOperation::Get {
            target,
            member,
            args,
            bind,
        } => {
            let target = resolve_target(objects, target)?;
            let args = observed_args_to_om(args, objects)?;
            observe_dispatch_result(
                runtime.dispatch_get(target, member, &args),
                bind.as_deref(),
                objects,
                identities,
                next_object_identity,
            )
        }
        CaseOperation::Set {
            target,
            member,
            value,
            args,
        } => {
            let target = resolve_target(objects, target)?;
            let value = observed_value_to_om(value, objects)?;
            let args = observed_args_to_om(args, objects)?;
            Ok(match runtime.dispatch_set(target, member, value, &args) {
                Ok(()) => ObservationResult::Value(ObservedValue::Void),
                Err(error) => ObservationResult::Error(observe_om_error(error)),
            })
        }
        CaseOperation::Invoke {
            target,
            member,
            args,
            bind,
        } => {
            let target = resolve_target(objects, target)?;
            let args = observed_args_to_om(args, objects)?;
            observe_dispatch_result(
                runtime.dispatch_invoke(target, member, &args),
                bind.as_deref(),
                objects,
                identities,
                next_object_identity,
            )
        }
        CaseOperation::Calculate => Ok(
            match runtime.dispatch_invoke(runtime.root_application(), "Calculate", &[]) {
                Ok(_) => ObservationResult::Value(ObservedValue::Void),
                Err(error) => ObservationResult::Error(observe_om_error(error)),
            },
        ),
        CaseOperation::Save { .. } => Err(OracleContractError::new(
            "runtime save cases require an external Excel reopen verifier",
        )),
    }
}

fn observe_dispatch_result(
    result: Result<OmValue, OmError>,
    bind: Option<&str>,
    objects: &mut BTreeMap<String, ObjectHandle>,
    identities: &mut BTreeMap<ObjectHandle, String>,
    next_object_identity: &mut u64,
) -> Result<ObservationResult, OracleContractError> {
    match result {
        Ok(OmValue::Object(handle)) if bind.is_some() => {
            let bind = bind.expect("checked binding");
            if objects.contains_key(bind) {
                return Err(OracleContractError::new(format!(
                    "object binding {bind} already exists"
                )));
            }
            objects.insert(bind.to_string(), handle);
            let identity = identities
                .entry(handle)
                .or_insert_with(|| bind.to_string())
                .clone();
            Ok(ObservationResult::Value(ObservedValue::Object(
                ObservedObject {
                    type_name: "Object".to_string(),
                    identity,
                },
            )))
        }
        Ok(_value) if bind.is_some() => Err(OracleContractError::new(format!(
            "object binding {} received a non-object result",
            bind.expect("checked binding")
        ))),
        Ok(value) => Ok(ObservationResult::Value(observe_om_value(
            value,
            identities,
            next_object_identity,
        ))),
        Err(error) => Ok(ObservationResult::Error(observe_om_error(error))),
    }
}

fn resolve_target(
    objects: &BTreeMap<String, ObjectHandle>,
    target: &str,
) -> Result<ObjectHandle, OracleContractError> {
    objects
        .get(target)
        .copied()
        .ok_or_else(|| OracleContractError::new(format!("object binding {target} was not defined")))
}

fn observed_args_to_om(
    args: &[ObservedValue],
    objects: &BTreeMap<String, ObjectHandle>,
) -> Result<Vec<OmValue>, OracleContractError> {
    args.iter()
        .map(|value| observed_value_to_om(value, objects))
        .collect()
}

fn observed_value_to_om(
    value: &ObservedValue,
    objects: &BTreeMap<String, ObjectHandle>,
) -> Result<OmValue, OracleContractError> {
    match value {
        ObservedValue::Void => Err(OracleContractError::new(
            "void cannot be used as an operation argument",
        )),
        ObservedValue::Missing => Ok(OmValue::Missing),
        ObservedValue::Empty => Ok(OmValue::Empty),
        ObservedValue::Null => Ok(OmValue::Null),
        ObservedValue::Bool(value) => Ok(OmValue::Bool(*value)),
        ObservedValue::Number(value) => Ok(OmValue::Number(*value)),
        ObservedValue::Text(value) => Ok(OmValue::Text(value.clone())),
        ObservedValue::CellError(error) => Ok(OmValue::Error(cell_error_from_cv_err(error.cv_err))),
        ObservedValue::Object(object) => {
            resolve_target(objects, &object.identity).map(OmValue::Object)
        }
        ObservedValue::Array(array) => {
            let values = array
                .values
                .iter()
                .map(|value| observed_value_to_om(value, objects))
                .collect::<Result<Vec<_>, _>>()?;
            OmArray::new(array.rows, array.cols, values)
                .map(OmValue::Array)
                .map_err(|error| runtime_harness_error("build array argument", error))
        }
    }
}

fn observe_om_value(
    value: OmValue,
    identities: &mut BTreeMap<ObjectHandle, String>,
    next_object_identity: &mut u64,
) -> ObservedValue {
    match value {
        OmValue::Missing => ObservedValue::Missing,
        OmValue::Empty => ObservedValue::Empty,
        OmValue::Null => ObservedValue::Null,
        OmValue::Bool(value) => ObservedValue::Bool(value),
        OmValue::Number(value) => ObservedValue::Number(value),
        OmValue::Text(value) => ObservedValue::Text(value),
        OmValue::Error(error) => ObservedValue::CellError(observe_cell_error(error)),
        OmValue::Object(handle) => {
            let identity = identities
                .entry(handle)
                .or_insert_with(|| {
                    let identity = format!("object_{:04}", *next_object_identity);
                    *next_object_identity += 1;
                    identity
                })
                .clone();
            ObservedValue::Object(ObservedObject {
                type_name: "Object".to_string(),
                identity,
            })
        }
        OmValue::Array(array) => ObservedValue::Array(ObservedArray {
            rows: array.rows,
            cols: array.cols,
            values: array
                .values
                .into_iter()
                .map(|value| observe_om_value(value, identities, next_object_identity))
                .collect(),
        }),
    }
}

fn observe_om_error(error: OmError) -> ObservedError {
    let kind = match error.code {
        OmErrorCode::InvalidArgument => CanonicalErrorKind::InvalidArgument,
        OmErrorCode::NotFound => CanonicalErrorKind::NotFound,
        OmErrorCode::TypeMismatch => CanonicalErrorKind::TypeMismatch,
        OmErrorCode::Unsupported => CanonicalErrorKind::Unsupported,
        OmErrorCode::InvalidState => CanonicalErrorKind::InvalidState,
        OmErrorCode::Io => CanonicalErrorKind::Io,
        OmErrorCode::Parse => CanonicalErrorKind::Parse,
        OmErrorCode::ResourceLimit => CanonicalErrorKind::ResourceLimit,
        OmErrorCode::EncryptedWorkbookUnsupported => CanonicalErrorKind::Unsupported,
        OmErrorCode::SignedPackageMutationUnsupported => CanonicalErrorKind::Unsupported,
        OmErrorCode::Calculation => CanonicalErrorKind::Calculation,
        OmErrorCode::External => CanonicalErrorKind::External,
    };
    let code = match error.code {
        OmErrorCode::InvalidArgument => "invalidArgument",
        OmErrorCode::NotFound => "notFound",
        OmErrorCode::TypeMismatch => "typeMismatch",
        OmErrorCode::Unsupported => "unsupported",
        OmErrorCode::InvalidState => "invalidState",
        OmErrorCode::Io => "io",
        OmErrorCode::Parse => "parse",
        OmErrorCode::ResourceLimit => "resourceLimit",
        OmErrorCode::EncryptedWorkbookUnsupported => "encryptedWorkbookUnsupported",
        OmErrorCode::SignedPackageMutationUnsupported => "signedPackageMutationUnsupported",
        OmErrorCode::Calculation => "calculation",
        OmErrorCode::External => "external",
    };
    ObservedError {
        kind,
        code: code.to_string(),
        diagnostic: Some(NativeErrorDiagnostic {
            origin: ObservedErrorKind::Ootd,
            hresult: None,
            message: Some(error.message),
        }),
    }
}

fn observe_cell_error(error: CellError) -> ObservedCellError {
    let (code, cv_err) = match error {
        CellError::Null => ("#NULL!", 2000),
        CellError::Div0 => ("#DIV/0!", 2007),
        CellError::Value => ("#VALUE!", 2015),
        CellError::Ref => ("#REF!", 2023),
        CellError::Name => ("#NAME?", 2029),
        CellError::Num => ("#NUM!", 2036),
        CellError::NA => ("#N/A", 2042),
        CellError::GettingData => ("#GETTING_DATA", 2043),
        CellError::Spill => ("#SPILL!", 2045),
        CellError::Connect => ("#CONNECT!", 2046),
        CellError::Blocked => ("#BLOCKED!", 2047),
        CellError::Unknown => ("#UNKNOWN!", 2048),
        CellError::Field => ("#FIELD!", 2049),
        CellError::Calc => ("#CALC!", 2050),
        CellError::Busy => ("#BUSY!", 2051),
        CellError::Python => ("#PYTHON!", 2052),
        CellError::Timeout => ("#TIMEOUT!", 2053),
    };
    ObservedCellError {
        code: code.to_string(),
        cv_err,
    }
}

fn cell_error_from_cv_err(cv_err: u16) -> CellError {
    match cv_err {
        2000 => CellError::Null,
        2007 => CellError::Div0,
        2015 => CellError::Value,
        2023 => CellError::Ref,
        2029 => CellError::Name,
        2036 => CellError::Num,
        2042 => CellError::NA,
        2043 => CellError::GettingData,
        2045 => CellError::Spill,
        2046 => CellError::Connect,
        2047 => CellError::Blocked,
        2048 => CellError::Unknown,
        2049 => CellError::Field,
        2050 => CellError::Calc,
        2051 => CellError::Busy,
        2052 => CellError::Python,
        2053 => CellError::Timeout,
        _ => CellError::Unknown,
    }
}

fn runtime_harness_error(context: &str, error: OmError) -> OracleContractError {
    OracleContractError::new(format!("runtime harness failed to {context}: {error}"))
}
