use std::{
    collections::BTreeMap,
    fmt, fs,
    path::{Path, PathBuf},
    process::Command,
};

use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

#[derive(Debug, Clone, PartialEq, Eq, Deserialize, Serialize)]
pub struct CaptureTemplate {
    pub capture: CaptureSection,
    pub target: CaptureTarget,
    pub excel: CaptureExcel,
    pub paths: CapturePaths,
    pub tooling: CaptureTooling,
    pub validation: CaptureValidation,
    pub outputs: CaptureOutputs,
    pub downstream: CaptureDownstream,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize, Serialize)]
pub struct CaptureSection {
    pub name: String,
    pub status: String,
    pub workspace: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize, Serialize)]
pub struct CaptureTarget {
    pub host_os: String,
    pub product_family: String,
    pub channel: String,
    pub version: String,
    pub build: String,
    pub arch: String,
    pub locale: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize, Serialize)]
pub struct CaptureExcel {
    pub type_library_major: u32,
    pub type_library_minor: u32,
    pub pia_assembly: String,
    pub pia_file: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize, Serialize)]
pub struct CapturePaths {
    pub excel_install_root: String,
    pub typelib_container: String,
    pub pia_path: String,
    pub output_dir: String,
    pub capture_root: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize, Serialize)]
pub struct CaptureTooling {
    pub oleview: String,
    pub tlbimp: String,
    pub regasm: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize, Serialize)]
pub struct CaptureValidation {
    pub fail_if_output_dir_exists: bool,
    pub fail_if_tooling_missing: bool,
    pub fail_if_paths_missing: bool,
    pub fail_if_typelib_version_mismatch: bool,
    pub write_capture_manifest: bool,
    pub write_output_checksums: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize, Serialize)]
pub struct CaptureOutputs {
    pub capture_manifest: String,
    pub capture_log: String,
    pub raw_typelib_identity: String,
    pub excel_typelib_snapshot_idl: String,
    pub excel_typelib_snapshot_odl: String,
    pub excel_pia_identity: String,
    pub excel_pia_public_surface: String,
    pub output_checksums: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize, Serialize)]
pub struct CaptureDownstream {
    pub office_idl_excel_om: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct NormalizedCapturePaths {
    pub excel_install_root: String,
    pub typelib_container: String,
    pub pia_path: String,
    pub output_dir: String,
    pub capture_root: String,
    pub oleview: String,
    pub tlbimp: String,
    pub regasm: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CaptureOutputLayout {
    pub capture_manifest: String,
    pub capture_log: String,
    pub raw_typelib_identity: String,
    pub excel_typelib_snapshot_idl: String,
    pub excel_typelib_snapshot_odl: String,
    pub excel_pia_identity: String,
    pub excel_pia_public_surface: String,
    pub output_checksums: String,
    pub office_idl_excel_om: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CapturePlan {
    pub template: CaptureTemplate,
    pub normalized_paths: NormalizedCapturePaths,
    pub output_layout: CaptureOutputLayout,
    pub unresolved_fields: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CapturePlanSummary {
    pub capture_name: String,
    pub capture_workspace: String,
    pub target_host_os: String,
    pub target_product_family: String,
    pub target_channel: String,
    pub target_version: String,
    pub target_build: String,
    pub target_arch: String,
    pub target_locale: String,
    pub output_dir: String,
    pub capture_root: String,
    pub output_paths: Vec<String>,
    pub pending_capture_outputs: Vec<String>,
    pub downstream_path: String,
    pub unresolved_fields: Vec<String>,
    pub ready_to_run: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CaptureArtifacts {
    pub capture_log: String,
    pub raw_typelib_identity: String,
    pub excel_typelib_snapshot_idl: String,
    pub excel_typelib_snapshot_odl: String,
    pub excel_pia_identity: String,
    pub excel_pia_public_surface: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CaptureWriteResult {
    pub written_paths: Vec<PathBuf>,
    pub manifest_path: Option<PathBuf>,
    pub output_checksums_path: Option<PathBuf>,
    pub downstream_path: PathBuf,
    pub checksums: BTreeMap<String, String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct CaptureExecutionPlan {
    pub script_path: String,
    pub direct_exec_launcher_path: String,
    pub direct_exec_status_path: String,
    pub generated_interop_assembly: String,
    pub pending_capture_outputs: Vec<String>,
    pub manual_steps: Vec<CaptureManualStep>,
    pub commands: Vec<CaptureCommand>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct CaptureManualStep {
    pub name: String,
    pub instructions: String,
    pub outputs: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct CaptureCommand {
    pub name: String,
    pub program: String,
    pub args: Vec<String>,
    pub outputs: Vec<String>,
    pub condition: Option<String>,
    pub notes: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct CaptureExecutionReceipt {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub started_at_utc: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub completed_at_utc: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub expected_capture_outputs: Vec<String>,
    #[serde(default)]
    pub host: CaptureHostIdentity,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub command_results: Vec<CaptureCommandResult>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub manual_step_results: Vec<CaptureManualStepResult>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct CaptureHostIdentity {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub computer_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub user_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub host_os: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub host_arch: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub powershell_version: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CaptureCommandResult {
    pub name: String,
    pub status: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub exit_code: Option<i32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub detail: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CaptureManualStepResult {
    pub name: String,
    pub status: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub detail: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CaptureExecutionBundleWriteResult {
    pub script_path: PathBuf,
    pub direct_exec_launcher_path: PathBuf,
    pub execution_plan_path: PathBuf,
    pub direct_exec_status_template_path: PathBuf,
    pub execution_receipt_template_path: PathBuf,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CaptureBundleCompletionResult {
    pub execution_receipt_path: PathBuf,
    pub write_result: CaptureWriteResult,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CaptureDirectExecStatus {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub started_at_utc: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub completed_at_utc: Option<String>,
    pub status: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub exit_code: Option<i32>,
    pub launcher_path: String,
    pub capture_script_path: String,
    pub execution_receipt_path: String,
    pub capture_manifest_path: String,
    pub output_checksums_path: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CaptureDirectExecRunResult {
    pub launcher_path: PathBuf,
    pub direct_exec_status_path: PathBuf,
    pub direct_exec_status: CaptureDirectExecStatus,
    pub launcher_exit_code: i32,
    pub launcher_stdout: String,
    pub launcher_stderr: String,
    pub completion_result: CaptureBundleCompletionResult,
}

#[derive(Debug)]
pub enum CapturePlanError {
    Io {
        path: PathBuf,
        source: std::io::Error,
    },
    Toml {
        path: PathBuf,
        source: toml::de::Error,
    },
    InvalidField {
        field: &'static str,
        value: String,
        message: &'static str,
    },
}

#[derive(Debug)]
pub enum CaptureWriteError {
    CreateDir {
        path: PathBuf,
        source: std::io::Error,
    },
    Write {
        path: PathBuf,
        source: std::io::Error,
    },
    Json {
        path: PathBuf,
        source: serde_json::Error,
    },
}

#[derive(Debug)]
pub enum CaptureBundleCompletionError {
    Io {
        path: PathBuf,
        source: std::io::Error,
    },
    Json {
        path: PathBuf,
        source: serde_json::Error,
    },
    MissingArtifact {
        logical_name: &'static str,
        path: PathBuf,
    },
    Write(CaptureWriteError),
}

#[derive(Debug)]
pub enum CaptureDirectExecError {
    Io {
        action: &'static str,
        path: PathBuf,
        source: std::io::Error,
    },
    Json {
        path: PathBuf,
        source: serde_json::Error,
    },
    MissingArtifact {
        logical_name: &'static str,
        path: PathBuf,
    },
    UnsupportedHostOs {
        actual: String,
    },
    LauncherFailed {
        exit_code: Option<i32>,
        status_path: PathBuf,
        direct_exec_status: CaptureDirectExecStatus,
        launcher_stdout: String,
        launcher_stderr: String,
    },
    Completion(CaptureBundleCompletionError),
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
struct CaptureManifestRecord {
    capture: CaptureSection,
    target: CaptureTarget,
    excel: CaptureExcel,
    normalized_paths: NormalizedCapturePaths,
    validation: CaptureValidation,
    writable_outputs: BTreeMap<String, String>,
    expected_capture_outputs: Vec<String>,
    downstream_output: String,
    unresolved_fields: Vec<String>,
    ready_to_run: bool,
    checksums: BTreeMap<String, String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    execution_receipt: Option<CaptureExecutionReceipt>,
}

impl CaptureTemplate {
    pub fn from_toml_str(input: &str) -> Result<Self, CapturePlanError> {
        toml::from_str(input).map_err(|source| CapturePlanError::Toml {
            path: PathBuf::from("<inline>"),
            source,
        })
    }

    pub fn from_toml_path(path: impl AsRef<Path>) -> Result<Self, CapturePlanError> {
        let path = path.as_ref();
        let input = fs::read_to_string(path).map_err(|source| CapturePlanError::Io {
            path: path.to_path_buf(),
            source,
        })?;
        Self::from_toml_str(&input)
    }
}

impl CaptureExecutionReceipt {
    pub fn from_json_str(input: &str) -> Result<Self, serde_json::Error> {
        serde_json::from_str(input)
    }

    pub fn from_json_path(path: impl AsRef<Path>) -> Result<Self, CaptureBundleCompletionError> {
        let path = path.as_ref();
        let input = fs::read_to_string(path).map_err(|source| {
            if source.kind() == std::io::ErrorKind::NotFound {
                CaptureBundleCompletionError::MissingArtifact {
                    logical_name: "execution_receipt",
                    path: path.to_path_buf(),
                }
            } else {
                CaptureBundleCompletionError::Io {
                    path: path.to_path_buf(),
                    source,
                }
            }
        })?;
        Self::from_json_str(&input).map_err(|source| CaptureBundleCompletionError::Json {
            path: path.to_path_buf(),
            source,
        })
    }
}

impl CaptureDirectExecStatus {
    pub fn from_json_str(input: &str) -> Result<Self, serde_json::Error> {
        serde_json::from_str(input)
    }

    pub fn from_json_path(path: impl AsRef<Path>) -> Result<Self, CaptureDirectExecError> {
        let path = path.as_ref();
        let input = fs::read_to_string(path).map_err(|source| {
            if source.kind() == std::io::ErrorKind::NotFound {
                CaptureDirectExecError::MissingArtifact {
                    logical_name: "direct_exec_status",
                    path: path.to_path_buf(),
                }
            } else {
                CaptureDirectExecError::Io {
                    action: "read",
                    path: path.to_path_buf(),
                    source,
                }
            }
        })?;
        Self::from_json_str(&input).map_err(|source| CaptureDirectExecError::Json {
            path: path.to_path_buf(),
            source,
        })
    }
}

impl CapturePlan {
    pub fn from_toml_str(input: &str) -> Result<Self, CapturePlanError> {
        let template = CaptureTemplate::from_toml_str(input)?;
        Self::from_template(template)
    }

    pub fn from_toml_path(path: impl AsRef<Path>) -> Result<Self, CapturePlanError> {
        let template = CaptureTemplate::from_toml_path(path)?;
        Self::from_template(template)
    }

    pub fn from_template(template: CaptureTemplate) -> Result<Self, CapturePlanError> {
        validate_capture_section(&template.capture)?;
        validate_target_section(&template.target)?;
        validate_excel_section(&template.excel)?;

        let unresolved_fields = collect_unresolved_fields(&template);
        let normalized_paths = NormalizedCapturePaths {
            excel_install_root: normalize_windows_path(
                "paths.excel_install_root",
                &template.paths.excel_install_root,
            )?,
            typelib_container: normalize_windows_path(
                "paths.typelib_container",
                &template.paths.typelib_container,
            )?,
            pia_path: normalize_windows_path("paths.pia_path", &template.paths.pia_path)?,
            output_dir: normalize_windows_path("paths.output_dir", &template.paths.output_dir)?,
            capture_root: normalize_windows_path(
                "paths.capture_root",
                &template.paths.capture_root,
            )?,
            oleview: normalize_windows_path("tooling.oleview", &template.tooling.oleview)?,
            tlbimp: normalize_windows_path("tooling.tlbimp", &template.tooling.tlbimp)?,
            regasm: normalize_windows_path("tooling.regasm", &template.tooling.regasm)?,
        };
        validate_capture_root(&normalized_paths.output_dir, &normalized_paths.capture_root)?;

        let output_layout = CaptureOutputLayout {
            capture_manifest: normalize_output_path(
                "outputs.capture_manifest",
                &template.outputs.capture_manifest,
                Some("manifest"),
            )?,
            capture_log: normalize_output_path(
                "outputs.capture_log",
                &template.outputs.capture_log,
                Some("logs"),
            )?,
            raw_typelib_identity: normalize_output_path(
                "outputs.raw_typelib_identity",
                &template.outputs.raw_typelib_identity,
                Some("raw"),
            )?,
            excel_typelib_snapshot_idl: normalize_output_path(
                "outputs.excel_typelib_snapshot_idl",
                &template.outputs.excel_typelib_snapshot_idl,
                Some("snapshots"),
            )?,
            excel_typelib_snapshot_odl: normalize_output_path(
                "outputs.excel_typelib_snapshot_odl",
                &template.outputs.excel_typelib_snapshot_odl,
                Some("snapshots"),
            )?,
            excel_pia_identity: normalize_output_path(
                "outputs.excel_pia_identity",
                &template.outputs.excel_pia_identity,
                Some("raw"),
            )?,
            excel_pia_public_surface: normalize_output_path(
                "outputs.excel_pia_public_surface",
                &template.outputs.excel_pia_public_surface,
                Some("snapshots"),
            )?,
            output_checksums: normalize_output_path(
                "outputs.output_checksums",
                &template.outputs.output_checksums,
                Some("manifest"),
            )?,
            office_idl_excel_om: normalize_output_path(
                "downstream.office_idl_excel_om",
                &template.downstream.office_idl_excel_om,
                None,
            )?,
        };

        Ok(Self {
            template,
            normalized_paths,
            output_layout,
            unresolved_fields,
        })
    }

    pub fn artifact_path(&self, artifact: &str) -> String {
        join_windows_path(&self.normalized_paths.capture_root, artifact)
    }

    pub fn execution_script_path(&self) -> String {
        self.artifact_path(execution_script_relative_path())
    }

    pub fn direct_exec_launcher_path(&self) -> String {
        self.artifact_path(direct_exec_launcher_relative_path())
    }

    pub fn write_artifacts(
        &self,
        host_root: impl AsRef<Path>,
        artifacts: &CaptureArtifacts,
    ) -> Result<CaptureWriteResult, CaptureWriteError> {
        self.write_artifacts_with_receipt(host_root, artifacts, None)
    }

    pub fn write_artifacts_with_receipt(
        &self,
        host_root: impl AsRef<Path>,
        artifacts: &CaptureArtifacts,
        execution_receipt: Option<&CaptureExecutionReceipt>,
    ) -> Result<CaptureWriteResult, CaptureWriteError> {
        let host_root = host_root.as_ref();
        let writable_outputs = [
            (
                &self.output_layout.capture_log,
                artifacts.capture_log.as_bytes(),
            ),
            (
                &self.output_layout.raw_typelib_identity,
                artifacts.raw_typelib_identity.as_bytes(),
            ),
            (
                &self.output_layout.excel_typelib_snapshot_idl,
                artifacts.excel_typelib_snapshot_idl.as_bytes(),
            ),
            (
                &self.output_layout.excel_typelib_snapshot_odl,
                artifacts.excel_typelib_snapshot_odl.as_bytes(),
            ),
            (
                &self.output_layout.excel_pia_identity,
                artifacts.excel_pia_identity.as_bytes(),
            ),
            (
                &self.output_layout.excel_pia_public_surface,
                artifacts.excel_pia_public_surface.as_bytes(),
            ),
        ];

        let mut written_paths = Vec::new();
        let mut checksums = BTreeMap::new();

        for (relative_path, contents) in writable_outputs {
            let path = host_root.join(relative_path_for_host(relative_path));
            write_file(&path, contents)?;
            written_paths.push(path);
            checksums.insert(relative_path.to_string(), sha256_hex(contents));
        }

        let output_checksums_path = if self.template.validation.write_output_checksums {
            let path = host_root.join(relative_path_for_host(&self.output_layout.output_checksums));
            let payload = serde_json::to_vec_pretty(&checksums).map_err(|source| {
                CaptureWriteError::Json {
                    path: path.clone(),
                    source,
                }
            })?;
            write_file(&path, &payload)?;
            written_paths.push(path.clone());
            Some(path)
        } else {
            None
        };

        let manifest_path = if self.template.validation.write_capture_manifest {
            let writable_outputs = self
                .writable_output_map()
                .into_iter()
                .map(|(name, relative_path)| (name.to_string(), self.artifact_path(&relative_path)))
                .collect();
            let path = host_root.join(relative_path_for_host(&self.output_layout.capture_manifest));
            let manifest = CaptureManifestRecord {
                capture: self.template.capture.clone(),
                target: self.template.target.clone(),
                excel: self.template.excel.clone(),
                normalized_paths: self.normalized_paths.clone(),
                validation: self.template.validation.clone(),
                writable_outputs,
                expected_capture_outputs: self.summary().pending_capture_outputs,
                downstream_output: self.artifact_path(&self.output_layout.office_idl_excel_om),
                unresolved_fields: self.unresolved_fields.clone(),
                ready_to_run: self.unresolved_fields.is_empty(),
                checksums: checksums.clone(),
                execution_receipt: execution_receipt.cloned(),
            };
            let payload =
                serde_json::to_vec_pretty(&manifest).map_err(|source| CaptureWriteError::Json {
                    path: path.clone(),
                    source,
                })?;
            write_file(&path, &payload)?;
            written_paths.push(path.clone());
            Some(path)
        } else {
            None
        };

        Ok(CaptureWriteResult {
            written_paths,
            manifest_path,
            output_checksums_path,
            downstream_path: host_root.join(relative_path_for_host(
                &self.output_layout.office_idl_excel_om,
            )),
            checksums,
        })
    }

    pub fn materialize_execution_bundle(
        &self,
        host_root: impl AsRef<Path>,
    ) -> Result<CaptureExecutionBundleWriteResult, CaptureWriteError> {
        let host_root = host_root.as_ref();
        let execution_plan = self.execution_plan();
        let script_path = host_root.join(relative_path_for_host(execution_script_relative_path()));
        let direct_exec_launcher_path =
            host_root.join(relative_path_for_host(direct_exec_launcher_relative_path()));
        let execution_plan_path =
            host_root.join(relative_path_for_host(execution_plan_relative_path()));
        let direct_exec_status_template_path = host_root.join(relative_path_for_host(
            direct_exec_status_template_relative_path(),
        ));
        let execution_receipt_template_path = host_root.join(relative_path_for_host(
            execution_receipt_template_relative_path(),
        ));
        let materialized_dirs = [
            script_path.parent().map(PathBuf::from),
            direct_exec_launcher_path.parent().map(PathBuf::from),
            execution_plan_path.parent().map(PathBuf::from),
            direct_exec_status_template_path.parent().map(PathBuf::from),
            execution_receipt_template_path.parent().map(PathBuf::from),
            host_root
                .join(relative_path_for_host(&self.output_layout.capture_log))
                .parent()
                .map(PathBuf::from),
            host_root
                .join(relative_path_for_host(
                    &self.output_layout.raw_typelib_identity,
                ))
                .parent()
                .map(PathBuf::from),
            host_root
                .join(relative_path_for_host(
                    &self.output_layout.excel_typelib_snapshot_idl,
                ))
                .parent()
                .map(PathBuf::from),
            host_root
                .join(relative_path_for_host(
                    &self.output_layout.excel_pia_public_surface,
                ))
                .parent()
                .map(PathBuf::from),
            host_root
                .join(relative_path_for_host(
                    &execution_plan.generated_interop_assembly,
                ))
                .parent()
                .map(PathBuf::from),
        ];
        for directory in materialized_dirs.into_iter().flatten() {
            fs::create_dir_all(&directory).map_err(|source| CaptureWriteError::CreateDir {
                path: directory.clone(),
                source,
            })?;
        }

        write_file(&script_path, self.render_powershell_script().as_bytes())?;
        write_file(
            &direct_exec_launcher_path,
            self.render_direct_exec_launcher().as_bytes(),
        )?;

        let execution_plan_payload =
            serde_json::to_vec_pretty(&execution_plan).map_err(|source| {
                CaptureWriteError::Json {
                    path: execution_plan_path.clone(),
                    source,
                }
            })?;
        write_file(&execution_plan_path, &execution_plan_payload)?;

        let direct_exec_status_payload = serde_json::to_vec_pretty(&CaptureDirectExecStatus {
            started_at_utc: None,
            completed_at_utc: None,
            status: "pending".to_string(),
            exit_code: None,
            launcher_path: self.direct_exec_launcher_path(),
            capture_script_path: self.execution_script_path(),
            execution_receipt_path: self.artifact_path(execution_receipt_relative_path()),
            capture_manifest_path: self.artifact_path(&self.output_layout.capture_manifest),
            output_checksums_path: self.artifact_path(&self.output_layout.output_checksums),
        })
        .map_err(|source| CaptureWriteError::Json {
            path: direct_exec_status_template_path.clone(),
            source,
        })?;
        write_file(
            &direct_exec_status_template_path,
            &direct_exec_status_payload,
        )?;

        let execution_receipt_payload = serde_json::to_vec_pretty(&CaptureExecutionReceipt {
            expected_capture_outputs: execution_plan.pending_capture_outputs.clone(),
            command_results: execution_plan
                .commands
                .iter()
                .map(|command| CaptureCommandResult {
                    name: command.name.clone(),
                    status: "pending".to_string(),
                    exit_code: None,
                    detail: command.notes.clone(),
                })
                .collect(),
            manual_step_results: execution_plan
                .manual_steps
                .iter()
                .map(|step| CaptureManualStepResult {
                    name: step.name.clone(),
                    status: "pending".to_string(),
                    detail: Some(step.instructions.clone()),
                })
                .collect(),
            ..CaptureExecutionReceipt::default()
        })
        .map_err(|source| CaptureWriteError::Json {
            path: execution_receipt_template_path.clone(),
            source,
        })?;
        write_file(&execution_receipt_template_path, &execution_receipt_payload)?;

        Ok(CaptureExecutionBundleWriteResult {
            script_path,
            direct_exec_launcher_path,
            execution_plan_path,
            direct_exec_status_template_path,
            execution_receipt_template_path,
        })
    }

    pub fn complete_execution_bundle(
        &self,
        host_root: impl AsRef<Path>,
    ) -> Result<CaptureBundleCompletionResult, CaptureBundleCompletionError> {
        self.complete_execution_bundle_with_receipt_path(host_root, None::<&Path>)
    }

    pub fn complete_execution_bundle_with_receipt_path(
        &self,
        host_root: impl AsRef<Path>,
        execution_receipt_path: Option<&Path>,
    ) -> Result<CaptureBundleCompletionResult, CaptureBundleCompletionError> {
        let host_root = host_root.as_ref();
        let execution_receipt_path =
            execution_receipt_path
                .map(PathBuf::from)
                .unwrap_or_else(|| {
                    host_root.join(relative_path_for_host(execution_receipt_relative_path()))
                });
        let execution_receipt = CaptureExecutionReceipt::from_json_path(&execution_receipt_path)?;
        let artifacts = self.read_materialized_artifacts(host_root)?;
        let write_result = self
            .write_artifacts_with_receipt(host_root, &artifacts, Some(&execution_receipt))
            .map_err(CaptureBundleCompletionError::Write)?;

        Ok(CaptureBundleCompletionResult {
            execution_receipt_path,
            write_result,
        })
    }

    pub fn run_execution_bundle(
        &self,
        host_root: impl AsRef<Path>,
    ) -> Result<CaptureDirectExecRunResult, CaptureDirectExecError> {
        let host_root = host_root.as_ref();
        let launcher_path =
            host_root.join(relative_path_for_host(direct_exec_launcher_relative_path()));
        if !launcher_path.exists() {
            return Err(CaptureDirectExecError::MissingArtifact {
                logical_name: "direct_exec_launcher",
                path: launcher_path,
            });
        }

        let execution_script_path =
            host_root.join(relative_path_for_host(execution_script_relative_path()));
        if !execution_script_path.exists() {
            return Err(CaptureDirectExecError::MissingArtifact {
                logical_name: "execution_script",
                path: execution_script_path,
            });
        }

        if !cfg!(target_os = "windows") {
            return Err(CaptureDirectExecError::UnsupportedHostOs {
                actual: std::env::consts::OS.to_string(),
            });
        }

        let direct_exec_status_path =
            host_root.join(relative_path_for_host(direct_exec_status_relative_path()));
        let execution_receipt_path =
            host_root.join(relative_path_for_host(execution_receipt_relative_path()));
        let capture_manifest_path =
            host_root.join(relative_path_for_host(&self.output_layout.capture_manifest));
        let output_checksums_path =
            host_root.join(relative_path_for_host(&self.output_layout.output_checksums));

        for stale_path in [
            &direct_exec_status_path,
            &execution_receipt_path,
            &capture_manifest_path,
            &output_checksums_path,
        ] {
            if stale_path.exists() {
                fs::remove_file(stale_path).map_err(|source| CaptureDirectExecError::Io {
                    action: "remove stale artifact",
                    path: stale_path.to_path_buf(),
                    source,
                })?;
            }
        }

        let output = Command::new("cmd")
            .arg("/C")
            .arg(&launcher_path)
            .current_dir(host_root)
            .output()
            .map_err(|source| CaptureDirectExecError::Io {
                action: "launch",
                path: launcher_path.clone(),
                source,
            })?;

        let launcher_stdout = String::from_utf8_lossy(&output.stdout).into_owned();
        let launcher_stderr = String::from_utf8_lossy(&output.stderr).into_owned();
        let direct_exec_status = CaptureDirectExecStatus::from_json_path(&direct_exec_status_path)?;
        if !output.status.success() || direct_exec_status.status != "completed" {
            return Err(CaptureDirectExecError::LauncherFailed {
                exit_code: output.status.code(),
                status_path: direct_exec_status_path,
                direct_exec_status,
                launcher_stdout,
                launcher_stderr,
            });
        }

        let completion_result = self
            .complete_execution_bundle_with_receipt_path(host_root, Some(&execution_receipt_path))
            .map_err(CaptureDirectExecError::Completion)?;

        Ok(CaptureDirectExecRunResult {
            launcher_path,
            direct_exec_status_path,
            direct_exec_status,
            launcher_exit_code: output.status.code().unwrap_or(0),
            launcher_stdout,
            launcher_stderr,
            completion_result,
        })
    }

    pub fn execution_plan(&self) -> CaptureExecutionPlan {
        let generated_interop_relative = format!(
            "work/generated/{}.generated.dll",
            sanitize_file_component(&self.template.excel.pia_assembly)
        );
        let generated_interop_assembly = self.artifact_path(&generated_interop_relative);

        CaptureExecutionPlan {
            script_path: self.execution_script_path(),
            direct_exec_launcher_path: self.direct_exec_launcher_path(),
            direct_exec_status_path: self.artifact_path(direct_exec_status_relative_path()),
            generated_interop_assembly: generated_interop_assembly.clone(),
            pending_capture_outputs: self.summary().pending_capture_outputs,
            manual_steps: vec![CaptureManualStep {
                name: "oleview_snapshot_export".to_string(),
                instructions: format!(
                    "Launch {} and export the reconstructed Excel type library snapshot to the pinned IDL/ODL paths before running Step 3 normalization.",
                    self.normalized_paths.oleview
                ),
                outputs: vec![
                    self.artifact_path(&self.output_layout.excel_typelib_snapshot_idl),
                    self.artifact_path(&self.output_layout.excel_typelib_snapshot_odl),
                ],
            }],
            commands: vec![
                CaptureCommand {
                    name: "tlbimp_fallback".to_string(),
                    program: self.normalized_paths.tlbimp.clone(),
                    args: vec![
                        self.normalized_paths.typelib_container.clone(),
                        format!("/out:{generated_interop_assembly}"),
                        format!("/namespace:{}", self.template.excel.pia_assembly),
                        "/nologo".to_string(),
                    ],
                    outputs: vec![generated_interop_assembly.clone()],
                    condition: Some(format!(
                        "only when {} is unavailable",
                        self.normalized_paths.pia_path
                    )),
                    notes: Some(
                        "Fallback managed interop assembly used for reflection-driven identity and public surface capture."
                            .to_string(),
                    ),
                },
                CaptureCommand {
                    name: "powershell_capture_reflection".to_string(),
                    program: "powershell.exe".to_string(),
                    args: vec![
                        "-NoProfile".to_string(),
                        "-ExecutionPolicy".to_string(),
                        "Bypass".to_string(),
                        "-Command".to_string(),
                        "<generated by office-capture --powershell-script>".to_string(),
                    ],
                    outputs: vec![
                        self.artifact_path(&self.output_layout.capture_log),
                        self.artifact_path(&self.output_layout.raw_typelib_identity),
                        self.artifact_path(&self.output_layout.excel_pia_identity),
                        self.artifact_path(&self.output_layout.excel_pia_public_surface),
                    ],
                    condition: None,
                    notes: Some(
                        "Automates interop-derived identity capture and PIA public surface JSON emission; IDL/ODL snapshots remain a manual oleview step."
                            .to_string(),
                    ),
                },
            ],
        }
    }

    pub fn render_direct_exec_launcher(&self) -> String {
        format!(
            "@echo off\r\n\
setlocal\r\n\
set \"SCRIPT_DIR=%~dp0\"\r\n\
set \"CAPTURE_SCRIPT=%SCRIPT_DIR%capture.ps1\"\r\n\
set \"STATUS_PATH=%SCRIPT_DIR%..\\manifest\\direct_exec_status.json\"\r\n\
set \"RECEIPT_PATH=%SCRIPT_DIR%..\\manifest\\execution_receipt.json\"\r\n\
set \"MANIFEST_PATH=%SCRIPT_DIR%..\\manifest\\capture_manifest.json\"\r\n\
set \"CHECKSUMS_PATH=%SCRIPT_DIR%..\\manifest\\output_checksums.json\"\r\n\
set \"STARTED_AT=%DATE% %TIME%\"\r\n\
powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"%CAPTURE_SCRIPT%\"\r\n\
set \"EXIT_CODE=%ERRORLEVEL%\"\r\n\
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command \"$status = [ordered]@{{ startedAtUtc = $env:STARTED_AT; completedAtUtc = (Get-Date).ToUniversalTime().ToString('o'); status = $(if ($env:EXIT_CODE -eq '0') {{ 'completed' }} else {{ 'failed' }}); exitCode = [int]$env:EXIT_CODE; launcherPath = {launcher_path}; captureScriptPath = {capture_script_path}; executionReceiptPath = {execution_receipt_path}; captureManifestPath = {capture_manifest_path}; outputChecksumsPath = {output_checksums_path} }}; Set-Content -LiteralPath $env:STATUS_PATH -Value ($status | ConvertTo-Json -Depth 8) -NoNewline\"\r\n\
exit /b %EXIT_CODE%\r\n",
            launcher_path = ps_quote(&self.direct_exec_launcher_path()),
            capture_script_path = ps_quote(&self.execution_script_path()),
            execution_receipt_path =
                ps_quote(&self.artifact_path(execution_receipt_relative_path())),
            capture_manifest_path =
                ps_quote(&self.artifact_path(&self.output_layout.capture_manifest)),
            output_checksums_path =
                ps_quote(&self.artifact_path(&self.output_layout.output_checksums)),
        )
    }

    fn read_materialized_artifacts(
        &self,
        host_root: &Path,
    ) -> Result<CaptureArtifacts, CaptureBundleCompletionError> {
        Ok(CaptureArtifacts {
            capture_log: read_required_utf8(
                host_root.join(relative_path_for_host(&self.output_layout.capture_log)),
                "capture_log",
            )?,
            raw_typelib_identity: read_required_utf8(
                host_root.join(relative_path_for_host(
                    &self.output_layout.raw_typelib_identity,
                )),
                "raw_typelib_identity",
            )?,
            excel_typelib_snapshot_idl: read_required_utf8(
                host_root.join(relative_path_for_host(
                    &self.output_layout.excel_typelib_snapshot_idl,
                )),
                "excel_typelib_snapshot_idl",
            )?,
            excel_typelib_snapshot_odl: read_required_utf8(
                host_root.join(relative_path_for_host(
                    &self.output_layout.excel_typelib_snapshot_odl,
                )),
                "excel_typelib_snapshot_odl",
            )?,
            excel_pia_identity: read_required_utf8(
                host_root.join(relative_path_for_host(
                    &self.output_layout.excel_pia_identity,
                )),
                "excel_pia_identity",
            )?,
            excel_pia_public_surface: read_required_utf8(
                host_root.join(relative_path_for_host(
                    &self.output_layout.excel_pia_public_surface,
                )),
                "excel_pia_public_surface",
            )?,
        })
    }

    pub fn render_powershell_script(&self) -> String {
        let execution_plan = self.execution_plan();
        let generated_interop = execution_plan.generated_interop_assembly.clone();
        let capture_log = self.artifact_path(&self.output_layout.capture_log);
        let raw_typelib_identity = self.artifact_path(&self.output_layout.raw_typelib_identity);
        let pia_identity = self.artifact_path(&self.output_layout.excel_pia_identity);
        let pia_public_surface = self.artifact_path(&self.output_layout.excel_pia_public_surface);
        let snapshot_idl = self.artifact_path(&self.output_layout.excel_typelib_snapshot_idl);
        let snapshot_odl = self.artifact_path(&self.output_layout.excel_typelib_snapshot_odl);
        let output_checksums = self.artifact_path(&self.output_layout.output_checksums);
        let capture_manifest = self.artifact_path(&self.output_layout.capture_manifest);
        let execution_receipt = self.artifact_path(execution_receipt_relative_path());

        format!(
            r#"$ErrorActionPreference = 'Stop'

$CaptureName = {capture_name}
$CaptureRoot = {capture_root}
$OutputDir = {output_dir}
$ExcelInstallRoot = {excel_install_root}
$TypelibContainer = {typelib_container}
$VendorPiaPath = {pia_path}
$OleviewPath = {oleview_path}
$TlbimpPath = {tlbimp_path}
$RegasmPath = {regasm_path}
$CaptureLogPath = {capture_log}
$RawTypelibIdentityPath = {raw_typelib_identity}
$ExcelPiaIdentityPath = {pia_identity}
$ExcelPiaPublicSurfacePath = {pia_public_surface}
$SnapshotIdlPath = {snapshot_idl}
$SnapshotOdlPath = {snapshot_odl}
$OutputChecksumsPath = {output_checksums}
$CaptureManifestPath = {capture_manifest}
$ExecutionReceiptPath = {execution_receipt}
$GeneratedInteropAssembly = {generated_interop}
$FailIfOutputDirExists = {fail_if_output_dir_exists}
$FailIfToolingMissing = {fail_if_tooling_missing}
$FailIfPathsMissing = {fail_if_paths_missing}
$WriteCaptureManifest = {write_capture_manifest}
$WriteOutputChecksums = {write_output_checksums}

function Ensure-Directory {{
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {{
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }}
}}

function Ensure-RequiredPath {{
    param([string]$Label, [string]$Path, [bool]$Required)
    if ($Required -and -not (Test-Path -LiteralPath $Path)) {{
        throw "missing required path: $Label => $Path"
    }}
}}

function Get-AttributeValue {{
    param([object[]]$Attributes, [string]$AttributeTypeName, [string]$PropertyName)
    foreach ($attribute in $Attributes) {{
        if ($attribute.GetType().FullName -eq $AttributeTypeName) {{
            return $attribute.$PropertyName
        }}
    }}
    return $null
}}

function Get-TypeRef {{
    param([Type]$Type)
    if ($null -eq $Type) {{
        return $null
    }}

    $item = [ordered]@{{ name = $Type.Name }}
    if ($Type.Namespace) {{
        $item['namespace'] = $Type.Namespace
    }}
    if ($Type.FullName) {{
        $item['aliasOf'] = $Type.FullName
    }}
    return $item
}}

function Get-ParameterJson {{
    param([System.Reflection.ParameterInfo]$Parameter)
    $item = [ordered]@{{
        name = $Parameter.Name
        typeRef = (Get-TypeRef $Parameter.ParameterType)
    }}
    if ($Parameter.IsOptional) {{
        $item['optional'] = $true
    }}
    return $item
}}

function Get-InterfaceKind {{
    param([Type]$Type)
    $interfaceType = Get-AttributeValue $Type.GetCustomAttributes($true) 'System.Runtime.InteropServices.InterfaceTypeAttribute' 'Value'
    if ($null -eq $interfaceType) {{
        return 'dispatch'
    }}
    switch ($interfaceType.ToString()) {{
        'InterfaceIsDual' {{ return 'dual' }}
        'InterfaceIsIUnknown' {{ return 'unknown' }}
        default {{ return 'dispatch' }}
    }}
}}

function Get-SafeGuid {{
    param([Type]$Type)
    $value = Get-AttributeValue $Type.GetCustomAttributes($true) 'System.Runtime.InteropServices.GuidAttribute' 'Value'
    if ($null -eq $value -or [string]::IsNullOrWhiteSpace($value)) {{
        return $null
    }}
    if ($value.StartsWith('{{')) {{
        return $value
    }}
    return '{{' + $value + '}}'
}}

if ($FailIfOutputDirExists -and (Test-Path -LiteralPath $CaptureRoot)) {{
    throw "capture_root already exists: $CaptureRoot"
}}

Ensure-RequiredPath 'paths.excel_install_root' $ExcelInstallRoot $FailIfPathsMissing
Ensure-RequiredPath 'paths.typelib_container' $TypelibContainer $FailIfPathsMissing
Ensure-RequiredPath 'tooling.oleview' $OleviewPath $FailIfToolingMissing
Ensure-RequiredPath 'tooling.tlbimp' $TlbimpPath $FailIfToolingMissing
Ensure-RequiredPath 'tooling.regasm' $RegasmPath $FailIfToolingMissing

Ensure-Directory $CaptureRoot
Ensure-Directory (Split-Path -Parent $CaptureLogPath)
Ensure-Directory (Split-Path -Parent $RawTypelibIdentityPath)
Ensure-Directory (Split-Path -Parent $ExcelPiaPublicSurfacePath)
Ensure-Directory (Split-Path -Parent $CaptureManifestPath)
Ensure-Directory (Split-Path -Parent $GeneratedInteropAssembly)

$InteropAssemblyPath = $VendorPiaPath
if (-not (Test-Path -LiteralPath $InteropAssemblyPath)) {{
    & $TlbimpPath $TypelibContainer ('/out:' + $GeneratedInteropAssembly) ('/namespace:' + {pia_namespace}) '/nologo'
    $InteropAssemblyPath = $GeneratedInteropAssembly
}}

$CaptureLog = @"
office-capture execution layer
- interop assembly: $InteropAssemblyPath
- manual oleview step still required for:
  - $SnapshotIdlPath
  - $SnapshotOdlPath
"@
Set-Content -LiteralPath $CaptureLogPath -Value $CaptureLog -NoNewline

Write-Host "Manual step required: launch oleview.exe and export Excel IDL/ODL snapshots."
Write-Host "  oleview: $OleviewPath"
Write-Host "  idl => $SnapshotIdlPath"
Write-Host "  odl => $SnapshotOdlPath"

$Assembly = [System.Reflection.Assembly]::LoadFrom($InteropAssemblyPath)
$ExportedTypes = $Assembly.GetExportedTypes() | Sort-Object FullName

$PiaIdentity = [ordered]@{{
    assembly = $Assembly.GetName().Name
    version = $Assembly.GetName().Version.ToString()
    fullName = $Assembly.FullName
    namespace = {pia_namespace}
    path = $InteropAssemblyPath
    source = $(if ($InteropAssemblyPath -eq $VendorPiaPath) {{ 'vendor_pia' }} else {{ 'tlbimp_fallback' }})
}}

$PublicSurface = [ordered]@{{
    library = 'Excel'
    version = '{typelib_version}'
    namespace = {pia_namespace}
    enums = @()
    interfaces = @()
    classes = @()
}}

$TypelibIdentity = [ordered]@{{
    library = 'Excel'
    version = '{typelib_version}'
    namespace = {pia_namespace}
    typeLibraryGuid = (Get-AttributeValue $Assembly.GetCustomAttributes($true) 'System.Runtime.InteropServices.GuidAttribute' 'Value')
    interfaces = @()
    coclasses = @()
}}

foreach ($Type in $ExportedTypes) {{
    if ($Type.IsEnum) {{
        $EnumValues = @()
        foreach ($Name in [Enum]::GetNames($Type)) {{
            $EnumValues += [ordered]@{{
                name = $Name
                value = [int64][Enum]::Parse($Type, $Name)
            }}
        }}
        $PublicSurface.enums += [ordered]@{{
            name = $Type.Name
            values = $EnumValues
        }}
        continue
    }}

    if ($Type.IsInterface) {{
        $Members = @()
        $Properties = $Type.GetProperties([System.Reflection.BindingFlags]'Instance, Public, DeclaredOnly') | Sort-Object Name
        foreach ($Property in $Properties) {{
            $DispId = Get-AttributeValue $Property.GetCustomAttributes($true) 'System.Runtime.InteropServices.DispIdAttribute' 'Value'
            $IndexParams = @($Property.GetIndexParameters() | ForEach-Object {{ Get-ParameterJson $_ }})
            if ($Property.CanRead) {{
                $Members += [ordered]@{{
                    name = $Property.Name
                    memberKind = 'property_get'
                    returnType = (Get-TypeRef $Property.PropertyType)
                    params = $IndexParams
                    dispId = $DispId
                }}
            }}
            if ($Property.CanWrite) {{
                $SetParams = @($IndexParams)
                $SetParams += [ordered]@{{
                    name = 'value'
                    typeRef = (Get-TypeRef $Property.PropertyType)
                }}
                $Members += [ordered]@{{
                    name = $Property.Name
                    memberKind = 'property_set'
                    returnType = $null
                    params = $SetParams
                    dispId = $DispId
                }}
            }}
        }}

        $Methods = $Type.GetMethods([System.Reflection.BindingFlags]'Instance, Public, DeclaredOnly') | Where-Object {{ -not $_.IsSpecialName }} | Sort-Object Name
        foreach ($Method in $Methods) {{
            $Members += [ordered]@{{
                name = $Method.Name
                memberKind = 'method'
                returnType = (Get-TypeRef $Method.ReturnType)
                params = @($Method.GetParameters() | ForEach-Object {{ Get-ParameterJson $_ }})
                dispId = (Get-AttributeValue $Method.GetCustomAttributes($true) 'System.Runtime.InteropServices.DispIdAttribute' 'Value')
            }}
        }}

        $Events = $Type.GetEvents([System.Reflection.BindingFlags]'Instance, Public, DeclaredOnly') | Sort-Object Name
        foreach ($Event in $Events) {{
            $Members += [ordered]@{{
                name = $Event.Name
                memberKind = 'event'
                returnType = (Get-TypeRef $Event.EventHandlerType)
                params = @()
                dispId = (Get-AttributeValue $Event.GetCustomAttributes($true) 'System.Runtime.InteropServices.DispIdAttribute' 'Value')
            }}
        }}

        $PublicSurface.interfaces += [ordered]@{{
            name = $Type.Name
            kind = (Get-InterfaceKind $Type)
            inherits = @($Type.GetInterfaces() | ForEach-Object {{ $_.Name }} | Sort-Object -Unique)
            members = $Members
        }}
        $TypelibIdentity.interfaces += [ordered]@{{
            name = $Type.Name
            iid = (Get-SafeGuid $Type)
            kind = (Get-InterfaceKind $Type)
        }}
        continue
    }}

    if ($Type.IsClass) {{
        $DefaultInterface = Get-AttributeValue $Type.GetCustomAttributes($true) 'System.Runtime.InteropServices.ComDefaultInterfaceAttribute' 'Value'
        $PublicSurface.classes += [ordered]@{{
            name = $Type.Name
            implements = @($Type.GetInterfaces() | ForEach-Object {{ $_.Name }} | Sort-Object -Unique)
            defaultInterface = $(if ($null -ne $DefaultInterface) {{ $DefaultInterface.Name }} else {{ $null }})
        }}
        $TypelibIdentity.coclasses += [ordered]@{{
            name = $Type.Name
            clsid = (Get-SafeGuid $Type)
            defaultInterface = $(if ($null -ne $DefaultInterface) {{ $DefaultInterface.Name }} else {{ $null }})
        }}
    }}
}}

Set-Content -LiteralPath $RawTypelibIdentityPath -Value ($TypelibIdentity | ConvertTo-Json -Depth 12) -NoNewline
Set-Content -LiteralPath $ExcelPiaIdentityPath -Value ($PiaIdentity | ConvertTo-Json -Depth 8) -NoNewline
Set-Content -LiteralPath $ExcelPiaPublicSurfacePath -Value ($PublicSurface | ConvertTo-Json -Depth 16) -NoNewline

if ($WriteOutputChecksums) {{
    Write-Host "Checksums are written by office-capture after artifact materialization: $OutputChecksumsPath"
}}
if ($WriteCaptureManifest) {{
    Write-Host "Capture manifest is written by office-capture after artifact materialization: $CaptureManifestPath"
}}

$ExecutionReceipt = [ordered]@{{
    startedAtUtc = [DateTime]::UtcNow.ToString('o')
    completedAtUtc = [DateTime]::UtcNow.ToString('o')
    host = [ordered]@{{
        computerName = $env:COMPUTERNAME
        userName = $env:USERNAME
        hostOs = 'windows'
        hostArch = $env:PROCESSOR_ARCHITECTURE
        powershellVersion = $PSVersionTable.PSVersion.ToString()
    }}
    commandResults = @(
        [ordered]@{{
            name = 'powershell_capture_reflection'
            status = 'completed'
            exitCode = 0
            detail = 'interop reflection capture finished'
        }}
    )
    manualStepResults = @(
        [ordered]@{{
            name = 'oleview_snapshot_export'
            status = 'pending'
            detail = 'manual export still required for Excel IDL/ODL snapshot files'
        }}
    )
}}
Set-Content -LiteralPath $ExecutionReceiptPath -Value ($ExecutionReceipt | ConvertTo-Json -Depth 8) -NoNewline
"#,
            capture_name = ps_quote(&self.template.capture.name),
            capture_root = ps_quote(&self.normalized_paths.capture_root),
            output_dir = ps_quote(&self.normalized_paths.output_dir),
            excel_install_root = ps_quote(&self.normalized_paths.excel_install_root),
            typelib_container = ps_quote(&self.normalized_paths.typelib_container),
            pia_path = ps_quote(&self.normalized_paths.pia_path),
            oleview_path = ps_quote(&self.normalized_paths.oleview),
            tlbimp_path = ps_quote(&self.normalized_paths.tlbimp),
            regasm_path = ps_quote(&self.normalized_paths.regasm),
            capture_log = ps_quote(&capture_log),
            raw_typelib_identity = ps_quote(&raw_typelib_identity),
            pia_identity = ps_quote(&pia_identity),
            pia_public_surface = ps_quote(&pia_public_surface),
            snapshot_idl = ps_quote(&snapshot_idl),
            snapshot_odl = ps_quote(&snapshot_odl),
            output_checksums = ps_quote(&output_checksums),
            capture_manifest = ps_quote(&capture_manifest),
            execution_receipt = ps_quote(&execution_receipt),
            generated_interop = ps_quote(&generated_interop),
            fail_if_output_dir_exists = if self.template.validation.fail_if_output_dir_exists {
                "$true"
            } else {
                "$false"
            },
            fail_if_tooling_missing = if self.template.validation.fail_if_tooling_missing {
                "$true"
            } else {
                "$false"
            },
            fail_if_paths_missing = if self.template.validation.fail_if_paths_missing {
                "$true"
            } else {
                "$false"
            },
            write_capture_manifest = if self.template.validation.write_capture_manifest {
                "$true"
            } else {
                "$false"
            },
            write_output_checksums = if self.template.validation.write_output_checksums {
                "$true"
            } else {
                "$false"
            },
            pia_namespace = ps_quote(&self.template.excel.pia_assembly),
            typelib_version = format!(
                "{}.{}",
                self.template.excel.type_library_major, self.template.excel.type_library_minor
            ),
        )
    }

    pub fn summary(&self) -> CapturePlanSummary {
        let output_paths = self
            .writable_output_map()
            .into_iter()
            .map(|(_, relative_path)| self.artifact_path(&relative_path))
            .collect();
        let pending_capture_outputs = [
            &self.output_layout.raw_typelib_identity,
            &self.output_layout.excel_typelib_snapshot_idl,
            &self.output_layout.excel_typelib_snapshot_odl,
            &self.output_layout.excel_pia_identity,
            &self.output_layout.excel_pia_public_surface,
        ]
        .into_iter()
        .map(|relative_path| {
            Path::new(relative_path)
                .file_name()
                .and_then(|name| name.to_str())
                .unwrap_or(relative_path)
                .to_string()
        })
        .collect();

        CapturePlanSummary {
            capture_name: self.template.capture.name.clone(),
            capture_workspace: self.template.capture.workspace.clone(),
            target_host_os: self.template.target.host_os.clone(),
            target_product_family: self.template.target.product_family.clone(),
            target_channel: self.template.target.channel.clone(),
            target_version: self.template.target.version.clone(),
            target_build: self.template.target.build.clone(),
            target_arch: self.template.target.arch.clone(),
            target_locale: self.template.target.locale.clone(),
            output_dir: self.normalized_paths.output_dir.clone(),
            capture_root: self.normalized_paths.capture_root.clone(),
            output_paths,
            pending_capture_outputs,
            downstream_path: self.artifact_path(&self.output_layout.office_idl_excel_om),
            unresolved_fields: self.unresolved_fields.clone(),
            ready_to_run: self.unresolved_fields.is_empty(),
        }
    }

    fn writable_output_map(&self) -> [(&'static str, String); 8] {
        [
            (
                "capture_manifest",
                self.output_layout.capture_manifest.clone(),
            ),
            ("capture_log", self.output_layout.capture_log.clone()),
            (
                "raw_typelib_identity",
                self.output_layout.raw_typelib_identity.clone(),
            ),
            (
                "excel_typelib_snapshot_idl",
                self.output_layout.excel_typelib_snapshot_idl.clone(),
            ),
            (
                "excel_typelib_snapshot_odl",
                self.output_layout.excel_typelib_snapshot_odl.clone(),
            ),
            (
                "excel_pia_identity",
                self.output_layout.excel_pia_identity.clone(),
            ),
            (
                "excel_pia_public_surface",
                self.output_layout.excel_pia_public_surface.clone(),
            ),
            (
                "output_checksums",
                self.output_layout.output_checksums.clone(),
            ),
        ]
    }
}

impl fmt::Display for CapturePlanSummary {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        writeln!(f, "capture: {}", self.capture_name)?;
        writeln!(f, "workspace: {}", self.capture_workspace)?;
        writeln!(
            f,
            "target: {} / {} / {} / {} / {} / {} / {}",
            self.target_host_os,
            self.target_product_family,
            self.target_channel,
            self.target_version,
            self.target_build,
            self.target_arch,
            self.target_locale,
        )?;
        writeln!(f, "output_dir: {}", self.output_dir)?;
        writeln!(f, "capture_root: {}", self.capture_root)?;
        writeln!(f, "ready_to_run: {}", self.ready_to_run)?;
        if self.unresolved_fields.is_empty() {
            writeln!(f, "unresolved_fields: none")?;
        } else {
            writeln!(
                f,
                "unresolved_fields: {}",
                self.unresolved_fields.join(", ")
            )?;
        }
        writeln!(f, "output_paths:")?;
        for output_path in &self.output_paths {
            writeln!(f, "  - {}", output_path)?;
        }
        writeln!(f, "pending_capture_outputs:")?;
        for output_name in &self.pending_capture_outputs {
            writeln!(f, "  - {}", output_name)?;
        }
        writeln!(f, "downstream_path: {}", self.downstream_path)?;
        Ok(())
    }
}

impl fmt::Display for CapturePlanError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CapturePlanError::Io { path, source } => {
                write!(f, "failed to read {}: {}", path.display(), source)
            }
            CapturePlanError::Toml { path, source } => {
                write!(f, "failed to parse {}: {}", path.display(), source)
            }
            CapturePlanError::InvalidField {
                field,
                value,
                message,
            } => {
                write!(f, "invalid {}={:?}: {}", field, value, message)
            }
        }
    }
}

impl std::error::Error for CapturePlanError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            CapturePlanError::Io { source, .. } => Some(source),
            CapturePlanError::Toml { source, .. } => Some(source),
            CapturePlanError::InvalidField { .. } => None,
        }
    }
}

impl fmt::Display for CaptureWriteError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CaptureWriteError::CreateDir { path, source } => {
                write!(f, "failed to create {}: {}", path.display(), source)
            }
            CaptureWriteError::Write { path, source } => {
                write!(f, "failed to write {}: {}", path.display(), source)
            }
            CaptureWriteError::Json { path, source } => {
                write!(f, "failed to serialize {}: {}", path.display(), source)
            }
        }
    }
}

impl std::error::Error for CaptureWriteError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            CaptureWriteError::CreateDir { source, .. } => Some(source),
            CaptureWriteError::Write { source, .. } => Some(source),
            CaptureWriteError::Json { source, .. } => Some(source),
        }
    }
}

impl fmt::Display for CaptureBundleCompletionError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CaptureBundleCompletionError::Io { path, source } => {
                write!(f, "failed to read {}: {}", path.display(), source)
            }
            CaptureBundleCompletionError::Json { path, source } => {
                write!(f, "failed to parse {}: {}", path.display(), source)
            }
            CaptureBundleCompletionError::MissingArtifact { logical_name, path } => {
                write!(
                    f,
                    "missing required artifact {} at {}",
                    logical_name,
                    path.display()
                )
            }
            CaptureBundleCompletionError::Write(source) => write!(f, "{source}"),
        }
    }
}

impl std::error::Error for CaptureBundleCompletionError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            CaptureBundleCompletionError::Io { source, .. } => Some(source),
            CaptureBundleCompletionError::Json { source, .. } => Some(source),
            CaptureBundleCompletionError::MissingArtifact { .. } => None,
            CaptureBundleCompletionError::Write(source) => Some(source),
        }
    }
}

impl fmt::Display for CaptureDirectExecError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CaptureDirectExecError::Io {
                action,
                path,
                source,
            } => write!(f, "failed to {action} {}: {}", path.display(), source),
            CaptureDirectExecError::Json { path, source } => {
                write!(f, "failed to parse {}: {}", path.display(), source)
            }
            CaptureDirectExecError::MissingArtifact { logical_name, path } => {
                write!(
                    f,
                    "missing required artifact {} at {}",
                    logical_name,
                    path.display()
                )
            }
            CaptureDirectExecError::UnsupportedHostOs { actual } => write!(
                f,
                "direct execution requires windows host, current host is {}",
                actual
            ),
            CaptureDirectExecError::LauncherFailed {
                exit_code,
                status_path,
                direct_exec_status,
                ..
            } => write!(
                f,
                "launcher failed with exit code {:?} and status {:?} from {}",
                exit_code,
                direct_exec_status.status,
                status_path.display()
            ),
            CaptureDirectExecError::Completion(source) => write!(f, "{source}"),
        }
    }
}

impl std::error::Error for CaptureDirectExecError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            CaptureDirectExecError::Io { source, .. } => Some(source),
            CaptureDirectExecError::Json { source, .. } => Some(source),
            CaptureDirectExecError::MissingArtifact { .. } => None,
            CaptureDirectExecError::UnsupportedHostOs { .. } => None,
            CaptureDirectExecError::LauncherFailed { .. } => None,
            CaptureDirectExecError::Completion(source) => Some(source),
        }
    }
}

fn validate_capture_section(capture: &CaptureSection) -> Result<(), CapturePlanError> {
    if capture.name.trim().is_empty() {
        return Err(CapturePlanError::InvalidField {
            field: "capture.name",
            value: capture.name.clone(),
            message: "must not be empty",
        });
    }
    if capture.status.trim().is_empty() {
        return Err(CapturePlanError::InvalidField {
            field: "capture.status",
            value: capture.status.clone(),
            message: "must not be empty",
        });
    }
    if capture.workspace.trim().is_empty() {
        return Err(CapturePlanError::InvalidField {
            field: "capture.workspace",
            value: capture.workspace.clone(),
            message: "must not be empty",
        });
    }
    Ok(())
}

fn validate_target_section(target: &CaptureTarget) -> Result<(), CapturePlanError> {
    if !target.host_os.eq_ignore_ascii_case("windows") {
        return Err(CapturePlanError::InvalidField {
            field: "target.host_os",
            value: target.host_os.clone(),
            message: "must be windows",
        });
    }
    Ok(())
}

fn validate_excel_section(excel: &CaptureExcel) -> Result<(), CapturePlanError> {
    if excel.type_library_major == 0 {
        return Err(CapturePlanError::InvalidField {
            field: "excel.type_library_major",
            value: excel.type_library_major.to_string(),
            message: "must be greater than zero",
        });
    }
    Ok(())
}

fn validate_capture_root(output_dir: &str, capture_root: &str) -> Result<(), CapturePlanError> {
    let output_dir = output_dir.trim_end_matches('\\').to_ascii_lowercase();
    let capture_root = capture_root.trim_end_matches('\\').to_ascii_lowercase();
    if capture_root == output_dir || capture_root.starts_with(&(output_dir + "\\")) {
        return Ok(());
    }
    Err(CapturePlanError::InvalidField {
        field: "paths.capture_root",
        value: capture_root,
        message: "must be located under paths.output_dir",
    })
}

fn collect_unresolved_fields(template: &CaptureTemplate) -> Vec<String> {
    let mut unresolved_fields = Vec::new();

    if contains_placeholder(&template.target.version) {
        unresolved_fields.push("target.version".to_string());
    }
    if contains_placeholder(&template.target.build) {
        unresolved_fields.push("target.build".to_string());
    }
    if contains_placeholder(&template.paths.pia_path) {
        unresolved_fields.push("paths.pia_path".to_string());
    }
    if contains_placeholder(&template.paths.capture_root) {
        unresolved_fields.push("paths.capture_root".to_string());
    }

    unresolved_fields
}

fn normalize_windows_path(field: &'static str, value: &str) -> Result<String, CapturePlanError> {
    let normalized = normalize_windows_path_string(value);
    if normalized.trim().is_empty() {
        return Err(CapturePlanError::InvalidField {
            field,
            value: value.to_string(),
            message: "must not be empty",
        });
    }
    if !is_windows_absolute_path(&normalized) {
        return Err(CapturePlanError::InvalidField {
            field,
            value: value.to_string(),
            message: "must be an absolute Windows path",
        });
    }
    Ok(normalized)
}

fn normalize_output_path(
    field: &'static str,
    value: &str,
    default_dir: Option<&str>,
) -> Result<String, CapturePlanError> {
    let normalized = value.trim().replace('\\', "/");
    if normalized.is_empty() {
        return Err(CapturePlanError::InvalidField {
            field,
            value: value.to_string(),
            message: "must not be empty",
        });
    }
    if normalized.starts_with('/') || is_windows_absolute_path(&normalized.replace('/', "\\")) {
        return Err(CapturePlanError::InvalidField {
            field,
            value: value.to_string(),
            message: "must be relative to capture_root",
        });
    }

    let mut parts = Vec::new();
    for part in normalized.split('/') {
        let part = part.trim();
        if part.is_empty() {
            continue;
        }
        if part == "." || part == ".." {
            return Err(CapturePlanError::InvalidField {
                field,
                value: value.to_string(),
                message: "must not contain relative traversal segments",
            });
        }
        parts.push(part.to_string());
    }

    if parts.is_empty() {
        return Err(CapturePlanError::InvalidField {
            field,
            value: value.to_string(),
            message: "must not be empty",
        });
    }

    let needs_default_dir = default_dir.is_some() && parts.len() == 1;
    let mut relative_path = String::new();
    if let Some(default_dir) = default_dir.filter(|_| needs_default_dir) {
        relative_path.push_str(default_dir);
        relative_path.push('/');
    }
    relative_path.push_str(&parts.join("/"));
    Ok(relative_path)
}

fn normalize_windows_path_string(value: &str) -> String {
    let trimmed = value.trim();
    let replaced = trimmed.replace('/', "\\");

    if let Some(rest) = replaced.strip_prefix("\\\\") {
        return format!("\\\\{}", collapse_backslashes(rest));
    }

    if replaced.len() >= 2 && replaced.as_bytes()[1] == b':' {
        let mut normalized = String::from(&replaced[..2]);
        let rest = &replaced[2..];
        if let Some(rest) = rest.strip_prefix('\\') {
            normalized.push('\\');
            normalized.push_str(&collapse_backslashes(rest));
        } else {
            normalized.push_str(&collapse_backslashes(rest));
        }
        return normalized;
    }

    collapse_backslashes(&replaced)
}

fn collapse_backslashes(value: &str) -> String {
    let mut collapsed = String::with_capacity(value.len());
    let mut previous_was_backslash = false;

    for ch in value.chars() {
        if ch == '\\' {
            if !previous_was_backslash {
                collapsed.push(ch);
            }
            previous_was_backslash = true;
        } else {
            previous_was_backslash = false;
            collapsed.push(ch);
        }
    }

    collapsed
}

fn is_windows_absolute_path(value: &str) -> bool {
    let bytes = value.as_bytes();
    if value.starts_with("\\\\") {
        return value.len() > 2;
    }
    bytes.len() >= 3 && bytes[0].is_ascii_alphabetic() && bytes[1] == b':' && bytes[2] == b'\\'
}

fn contains_placeholder(value: &str) -> bool {
    value.contains("...")
        || value.contains("x.y")
        || value.contains("<pending>")
        || value.contains("TODO")
}

fn join_windows_path(base: &str, leaf: &str) -> String {
    let base = base.trim_end_matches(['\\', '/']);
    let leaf = leaf.trim_start_matches(['\\', '/']).replace('/', "\\");
    if base.is_empty() {
        leaf
    } else {
        format!("{base}\\{leaf}")
    }
}

fn execution_script_relative_path() -> &'static str {
    "scripts/capture.ps1"
}

fn direct_exec_launcher_relative_path() -> &'static str {
    "scripts/run_capture.cmd"
}

fn execution_plan_relative_path() -> &'static str {
    "manifest/execution_plan.json"
}

fn direct_exec_status_relative_path() -> &'static str {
    "manifest/direct_exec_status.json"
}

fn direct_exec_status_template_relative_path() -> &'static str {
    "manifest/direct_exec_status.template.json"
}

fn execution_receipt_relative_path() -> &'static str {
    "manifest/execution_receipt.json"
}

fn execution_receipt_template_relative_path() -> &'static str {
    "manifest/execution_receipt.template.json"
}

fn relative_path_for_host(relative_path: &str) -> PathBuf {
    let mut path = PathBuf::new();
    for part in relative_path.split('/') {
        if !part.is_empty() {
            path.push(part);
        }
    }
    path
}

fn write_file(path: &Path, contents: &[u8]) -> Result<(), CaptureWriteError> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|source| CaptureWriteError::CreateDir {
            path: parent.to_path_buf(),
            source,
        })?;
    }
    fs::write(path, contents).map_err(|source| CaptureWriteError::Write {
        path: path.to_path_buf(),
        source,
    })
}

fn read_required_utf8(
    path: PathBuf,
    logical_name: &'static str,
) -> Result<String, CaptureBundleCompletionError> {
    let bytes = fs::read(&path).map_err(|source| {
        if source.kind() == std::io::ErrorKind::NotFound {
            CaptureBundleCompletionError::MissingArtifact {
                logical_name,
                path: path.clone(),
            }
        } else {
            CaptureBundleCompletionError::Io {
                path: path.clone(),
                source,
            }
        }
    })?;
    String::from_utf8(bytes).map_err(|source| CaptureBundleCompletionError::Io {
        path,
        source: std::io::Error::new(std::io::ErrorKind::InvalidData, source),
    })
}

fn sanitize_file_component(value: &str) -> String {
    value
        .chars()
        .map(|ch| {
            if ch.is_ascii_alphanumeric() || matches!(ch, '.' | '_' | '-') {
                ch
            } else {
                '_'
            }
        })
        .collect()
}

fn ps_quote(value: &str) -> String {
    format!("'{}'", value.replace('\'', "''"))
}

fn sha256_hex(contents: &[u8]) -> String {
    let digest = Sha256::digest(contents);
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        output.push_str(&format!("{byte:02x}"));
    }
    output
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    fn repo_root() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .canonicalize()
            .expect("repo root")
    }

    fn resolved_template() -> &'static str {
        r#"
            [capture]
            name = "excel_om_windows_capture"
            status = "ready"
            workspace = "ootd"

            [target]
            host_os = "windows"
            product_family = "microsoft_365_apps"
            channel = "Current"
            version = "16.0.1234.5678"
            build = "16.0.1234.5678"
            arch = "x64"
            locale = "en-us"

            [excel]
            type_library_major = 16
            type_library_minor = 0
            pia_assembly = "Microsoft.Office.Interop.Excel"
            pia_file = "Microsoft.Office.Interop.Excel.dll"

            [paths]
            excel_install_root = 'C:/Program Files/Microsoft Office/root/Office16'
            typelib_container = 'C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE'
            pia_path = 'C:/Windows/assembly/GAC_MSIL/Microsoft.Office.Interop.Excel/Microsoft.Office.Interop.Excel.dll'
            output_dir = 'C:/capture/excel-om'
            capture_root = 'C:/capture/excel-om/excel_om_windows_capture'

            [tooling]
            oleview = 'C:/Program Files (x86)/Windows Kits/10/bin/x64/oleview.exe'
            tlbimp = 'C:/Windows/Microsoft.NET/Framework64/v4.0.30319/Tlbimp.exe'
            regasm = 'C:/Windows/Microsoft.NET/Framework64/v4.0.30319/RegAsm.exe'

            [validation]
            fail_if_output_dir_exists = true
            fail_if_tooling_missing = true
            fail_if_paths_missing = true
            fail_if_typelib_version_mismatch = true
            write_capture_manifest = true
            write_output_checksums = true

            [outputs]
            capture_manifest = "manifest/capture_manifest.json"
            capture_log = "logs/capture.log"
            raw_typelib_identity = "raw_typelib_identity.json"
            excel_typelib_snapshot_idl = "excel_typelib_snapshot.idl"
            excel_typelib_snapshot_odl = "excel_typelib_snapshot.odl"
            excel_pia_identity = "excel_pia_identity.json"
            excel_pia_public_surface = "excel_pia_public_surface.json"
            output_checksums = "manifest/output_checksums.json"

            [downstream]
            office_idl_excel_om = "office_idl_excel_om.json"
        "#
    }

    fn sample_artifacts() -> CaptureArtifacts {
        CaptureArtifacts {
            capture_log: "capture log".to_string(),
            raw_typelib_identity: r#"{"library":"Excel"}"#.to_string(),
            excel_typelib_snapshot_idl: "library Excel {}".to_string(),
            excel_typelib_snapshot_odl: "odl Excel {}".to_string(),
            excel_pia_identity: r#"{"assembly":"Microsoft.Office.Interop.Excel"}"#.to_string(),
            excel_pia_public_surface: r#"{"library":"Excel","interfaces":[]}"#.to_string(),
        }
    }

    fn sample_execution_receipt() -> CaptureExecutionReceipt {
        CaptureExecutionReceipt {
            started_at_utc: Some("2026-03-27T01:02:03Z".to_string()),
            completed_at_utc: Some("2026-03-27T01:03:04Z".to_string()),
            expected_capture_outputs: vec![
                "raw_typelib_identity.json".to_string(),
                "excel_typelib_snapshot.idl".to_string(),
                "excel_typelib_snapshot.odl".to_string(),
                "excel_pia_identity.json".to_string(),
                "excel_pia_public_surface.json".to_string(),
            ],
            host: CaptureHostIdentity {
                computer_name: Some("WIN-EXCEL".to_string()),
                user_name: Some("runner".to_string()),
                host_os: Some("Windows 11".to_string()),
                host_arch: Some("x64".to_string()),
                powershell_version: Some("5.1".to_string()),
            },
            command_results: vec![CaptureCommandResult {
                name: "powershell_capture_reflection".to_string(),
                status: "completed".to_string(),
                exit_code: Some(0),
                detail: Some("interop reflection capture finished".to_string()),
            }],
            manual_step_results: vec![CaptureManualStepResult {
                name: "oleview_snapshot_export".to_string(),
                status: "completed".to_string(),
                detail: Some("IDL/ODL snapshots exported".to_string()),
            }],
        }
    }

    #[test]
    fn loads_pinned_template_and_reports_unresolved_placeholder_fields() {
        let plan = CapturePlan::from_toml_path(
            repo_root().join("specs/pinned/windows_capture.template.toml"),
        )
        .expect("template");
        let summary = plan.summary();

        assert_eq!(plan.template.capture.name, "excel_om_windows_capture");
        assert_eq!(
            plan.normalized_paths.output_dir,
            "C:\\ootd-capture\\excel-om"
        );
        assert_eq!(
            plan.normalized_paths.capture_root,
            "C:\\ootd-capture\\excel-om\\excel_om_windows_capture"
        );
        assert_eq!(
            plan.output_layout.raw_typelib_identity,
            "raw/raw_typelib_identity.json"
        );
        assert_eq!(
            plan.output_layout.excel_pia_public_surface,
            "snapshots/excel_pia_public_surface.json"
        );
        assert_eq!(
            summary.output_paths,
            vec![
                "C:\\ootd-capture\\excel-om\\excel_om_windows_capture\\manifest\\capture_manifest.json".to_string(),
                "C:\\ootd-capture\\excel-om\\excel_om_windows_capture\\logs\\capture.log".to_string(),
                "C:\\ootd-capture\\excel-om\\excel_om_windows_capture\\raw\\raw_typelib_identity.json".to_string(),
                "C:\\ootd-capture\\excel-om\\excel_om_windows_capture\\snapshots\\excel_typelib_snapshot.idl".to_string(),
                "C:\\ootd-capture\\excel-om\\excel_om_windows_capture\\snapshots\\excel_typelib_snapshot.odl".to_string(),
                "C:\\ootd-capture\\excel-om\\excel_om_windows_capture\\raw\\excel_pia_identity.json".to_string(),
                "C:\\ootd-capture\\excel-om\\excel_om_windows_capture\\snapshots\\excel_pia_public_surface.json".to_string(),
                "C:\\ootd-capture\\excel-om\\excel_om_windows_capture\\manifest\\output_checksums.json".to_string(),
            ]
        );
        assert_eq!(
            summary.downstream_path,
            "C:\\ootd-capture\\excel-om\\excel_om_windows_capture\\office_idl_excel_om.json"
        );
        assert_eq!(
            summary.unresolved_fields,
            vec![
                "target.version".to_string(),
                "target.build".to_string(),
                "paths.pia_path".to_string()
            ]
        );
        assert!(!summary.ready_to_run);
        assert!(
            summary
                .to_string()
                .contains("raw\\raw_typelib_identity.json")
        );
    }

    #[test]
    fn normalizes_mixed_windows_separators_and_builds_output_layout() {
        let template = r#"
            [capture]
            name = "excel_om_windows_capture"
            status = "template"
            workspace = "ootd"

            [target]
            host_os = "windows"
            product_family = "microsoft_365_apps"
            channel = "Current"
            version = "16.0.1234.5678"
            build = "16.0.1234.5678"
            arch = "x64"
            locale = "en-us"

            [excel]
            type_library_major = 16
            type_library_minor = 0
            pia_assembly = "Microsoft.Office.Interop.Excel"
            pia_file = "Microsoft.Office.Interop.Excel.dll"

            [paths]
            excel_install_root = 'C:/Program Files/Microsoft Office/root/Office16'
            typelib_container = 'C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE'
            pia_path = 'C:/Windows/assembly/GAC_MSIL/Microsoft.Office.Interop.Excel/Microsoft.Office.Interop.Excel.dll'
            output_dir = 'C:/capture//excel-om'
            capture_root = 'C:/capture//excel-om/excel_om_windows_capture'

            [tooling]
            oleview = 'C:/Program Files (x86)/Windows Kits/10/bin/x64/oleview.exe'
            tlbimp = 'C:/Windows/Microsoft.NET/Framework64/v4.0.30319/Tlbimp.exe'
            regasm = 'C:/Windows/Microsoft.NET/Framework64/v4.0.30319/RegAsm.exe'

            [validation]
            fail_if_output_dir_exists = true
            fail_if_tooling_missing = true
            fail_if_paths_missing = true
            fail_if_typelib_version_mismatch = true
            write_capture_manifest = true
            write_output_checksums = true

            [outputs]
            capture_manifest = "manifest/capture_manifest.json"
            capture_log = "logs/capture.log"
            raw_typelib_identity = "raw_typelib_identity.json"
            excel_typelib_snapshot_idl = "excel_typelib_snapshot.idl"
            excel_typelib_snapshot_odl = "excel_typelib_snapshot.odl"
            excel_pia_identity = "excel_pia_identity.json"
            excel_pia_public_surface = "excel_pia_public_surface.json"
            output_checksums = "manifest/output_checksums.json"

            [downstream]
            office_idl_excel_om = "office_idl_excel_om.json"
        "#;

        let plan = CapturePlan::from_toml_str(template).expect("plan");

        assert_eq!(
            plan.normalized_paths.excel_install_root,
            "C:\\Program Files\\Microsoft Office\\root\\Office16"
        );
        assert_eq!(plan.normalized_paths.output_dir, "C:\\capture\\excel-om");
        assert_eq!(
            plan.artifact_path(&plan.output_layout.excel_pia_public_surface),
            "C:\\capture\\excel-om\\excel_om_windows_capture\\snapshots\\excel_pia_public_surface.json"
        );
        assert!(plan.summary().unresolved_fields.is_empty());
        assert!(plan.summary().ready_to_run);
    }

    #[test]
    fn rejects_relative_windows_paths() {
        let template = resolved_template().replace(
            "excel_install_root = 'C:/Program Files/Microsoft Office/root/Office16'",
            "excel_install_root = 'relative/install/root'",
        );

        let error = CapturePlan::from_toml_str(&template).expect_err("relative path should fail");
        match error {
            CapturePlanError::InvalidField { field, .. } => {
                assert_eq!(field, "paths.excel_install_root");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn rejects_empty_capture_name() {
        let template =
            resolved_template().replace(r#"name = "excel_om_windows_capture""#, r#"name = """#);

        let error =
            CapturePlan::from_toml_str(&template).expect_err("empty capture name should fail");
        match error {
            CapturePlanError::InvalidField { field, .. } => {
                assert_eq!(field, "capture.name");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn rejects_non_windows_target_host_os() {
        let template =
            resolved_template().replace(r#"host_os = "windows""#, r#"host_os = "linux""#);

        let error =
            CapturePlan::from_toml_str(&template).expect_err("non-windows target should fail");
        match error {
            CapturePlanError::InvalidField { field, .. } => {
                assert_eq!(field, "target.host_os");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn rejects_zero_type_library_major() {
        let template =
            resolved_template().replace("type_library_major = 16", "type_library_major = 0");

        let error =
            CapturePlan::from_toml_str(&template).expect_err("zero type library major should fail");
        match error {
            CapturePlanError::InvalidField { field, .. } => {
                assert_eq!(field, "excel.type_library_major");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn rejects_capture_root_outside_output_dir() {
        let template = resolved_template().replace(
            "capture_root = 'C:/capture/excel-om/excel_om_windows_capture'",
            "capture_root = 'D:/foreign/excel_om_windows_capture'",
        );

        let error =
            CapturePlan::from_toml_str(&template).expect_err("capture_root should be rejected");
        match error {
            CapturePlanError::InvalidField { field, .. } => {
                assert_eq!(field, "paths.capture_root");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn rejects_output_path_traversal() {
        let template = resolved_template().replace(
            "excel_pia_public_surface = \"excel_pia_public_surface.json\"",
            "excel_pia_public_surface = \"../excel_pia_public_surface.json\"",
        );

        let error =
            CapturePlan::from_toml_str(&template).expect_err("output traversal should fail");
        match error {
            CapturePlanError::InvalidField { field, .. } => {
                assert_eq!(field, "outputs.excel_pia_public_surface");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn rejects_absolute_output_paths() {
        let template = resolved_template().replace(
            "excel_pia_public_surface = \"excel_pia_public_surface.json\"",
            "excel_pia_public_surface = \"C:/capture/excel_pia_public_surface.json\"",
        );

        let error =
            CapturePlan::from_toml_str(&template).expect_err("absolute output path should fail");
        match error {
            CapturePlanError::InvalidField { field, .. } => {
                assert_eq!(field, "outputs.excel_pia_public_surface");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn normalizes_unc_capture_paths() {
        let template = resolved_template()
            .replace(
                "output_dir = 'C:/capture/excel-om'",
                r"output_dir = '\\server\capture\excel-om'",
            )
            .replace(
                "capture_root = 'C:/capture/excel-om/excel_om_windows_capture'",
                r"capture_root = '\\server\capture\excel-om\excel_om_windows_capture'",
            )
            .replace(
                "excel_install_root = 'C:/Program Files/Microsoft Office/root/Office16'",
                r"excel_install_root = '\\server\office\Office16'",
            );

        let plan = CapturePlan::from_toml_str(&template).expect("plan");

        assert_eq!(
            plan.normalized_paths.output_dir,
            "\\\\server\\capture\\excel-om"
        );
        assert_eq!(
            plan.normalized_paths.capture_root,
            "\\\\server\\capture\\excel-om\\excel_om_windows_capture"
        );
        assert_eq!(
            plan.normalized_paths.excel_install_root,
            "\\\\server\\office\\Office16"
        );
    }

    #[test]
    fn writes_bundle_with_manifest_and_checksums() {
        let plan = CapturePlan::from_toml_str(resolved_template()).expect("plan");
        let tempdir = TempDir::new().expect("tempdir");
        let result = plan
            .write_artifacts(tempdir.path(), &sample_artifacts())
            .expect("write bundle");

        let raw_typelib_path = tempdir.path().join("raw/raw_typelib_identity.json");
        let public_surface_path = tempdir
            .path()
            .join("snapshots/excel_pia_public_surface.json");
        let manifest_path = tempdir.path().join("manifest/capture_manifest.json");
        let checksums_path = tempdir.path().join("manifest/output_checksums.json");

        assert!(raw_typelib_path.exists());
        assert!(public_surface_path.exists());
        assert!(manifest_path.exists());
        assert!(checksums_path.exists());
        assert!(!result.downstream_path.exists());

        let checksums: BTreeMap<String, String> =
            serde_json::from_slice(&fs::read(&checksums_path).expect("checksums file"))
                .expect("checksums json");
        assert_eq!(
            checksums.get("raw/raw_typelib_identity.json"),
            Some(&sha256_hex(br#"{"library":"Excel"}"#))
        );

        let manifest: serde_json::Value =
            serde_json::from_slice(&fs::read(&manifest_path).expect("manifest file"))
                .expect("manifest json");
        assert_eq!(
            manifest["writableOutputs"]["excel_pia_public_surface"],
            "C:\\capture\\excel-om\\excel_om_windows_capture\\snapshots\\excel_pia_public_surface.json"
        );
        assert_eq!(
            manifest["downstreamOutput"],
            "C:\\capture\\excel-om\\excel_om_windows_capture\\office_idl_excel_om.json"
        );
        assert_eq!(
            manifest["expectedCaptureOutputs"],
            serde_json::json!([
                "raw_typelib_identity.json",
                "excel_typelib_snapshot.idl",
                "excel_typelib_snapshot.odl",
                "excel_pia_identity.json",
                "excel_pia_public_surface.json"
            ])
        );
        assert_eq!(result.written_paths.len(), 8);
        assert!(result.manifest_path.is_some());
        assert!(result.output_checksums_path.is_some());
    }

    #[test]
    fn writes_manifest_with_execution_receipt_when_provided() {
        let plan = CapturePlan::from_toml_str(resolved_template()).expect("plan");
        let tempdir = TempDir::new().expect("tempdir");
        let receipt = sample_execution_receipt();
        plan.write_artifacts_with_receipt(tempdir.path(), &sample_artifacts(), Some(&receipt))
            .expect("write bundle");

        let manifest_path = tempdir.path().join("manifest/capture_manifest.json");
        let manifest: serde_json::Value =
            serde_json::from_slice(&fs::read(&manifest_path).expect("manifest file"))
                .expect("manifest json");
        assert_eq!(
            manifest["executionReceipt"]["host"]["computerName"],
            "WIN-EXCEL"
        );
        assert_eq!(
            manifest["executionReceipt"]["manualStepResults"][0]["status"],
            "completed"
        );
    }

    #[test]
    fn skips_optional_manifest_and_checksums_when_disabled() {
        let template = resolved_template()
            .replace(
                "write_capture_manifest = true",
                "write_capture_manifest = false",
            )
            .replace(
                "write_output_checksums = true",
                "write_output_checksums = false",
            );
        let plan = CapturePlan::from_toml_str(&template).expect("plan");
        let tempdir = TempDir::new().expect("tempdir");
        let result = plan
            .write_artifacts(tempdir.path(), &sample_artifacts())
            .expect("write bundle");

        assert!(
            !tempdir
                .path()
                .join("manifest/capture_manifest.json")
                .exists()
        );
        assert!(
            !tempdir
                .path()
                .join("manifest/output_checksums.json")
                .exists()
        );
        assert!(tempdir.path().join("raw/excel_pia_identity.json").exists());
        assert!(result.manifest_path.is_none());
        assert!(result.output_checksums_path.is_none());
        assert_eq!(result.written_paths.len(), 6);
    }

    #[test]
    fn materializes_execution_bundle_with_script_plan_and_receipt_template() {
        let plan = CapturePlan::from_toml_str(resolved_template()).expect("plan");
        let tempdir = TempDir::new().expect("tempdir");
        let result = plan
            .materialize_execution_bundle(tempdir.path())
            .expect("materialize execution bundle");

        assert_eq!(
            result.script_path,
            tempdir.path().join("scripts/capture.ps1")
        );
        assert_eq!(
            result.direct_exec_launcher_path,
            tempdir.path().join("scripts/run_capture.cmd")
        );
        assert_eq!(
            result.execution_plan_path,
            tempdir.path().join("manifest/execution_plan.json")
        );
        assert_eq!(
            result.direct_exec_status_template_path,
            tempdir
                .path()
                .join("manifest/direct_exec_status.template.json")
        );
        assert_eq!(
            result.execution_receipt_template_path,
            tempdir
                .path()
                .join("manifest/execution_receipt.template.json")
        );

        let script = fs::read_to_string(&result.script_path).expect("script file");
        assert!(script.contains("$CaptureManifestPath"));
        assert!(script.contains("$ExecutionReceiptPath"));
        assert!(script.contains("Manual step required: launch oleview.exe"));

        let launcher = fs::read_to_string(&result.direct_exec_launcher_path).expect("launcher");
        assert!(launcher.contains("capture.ps1"));
        assert!(launcher.contains("direct_exec_status.json"));
        assert!(launcher.contains("exit /b %EXIT_CODE%"));

        let execution_plan: serde_json::Value = serde_json::from_slice(
            &fs::read(&result.execution_plan_path).expect("execution plan file"),
        )
        .expect("execution plan json");
        assert_eq!(
            execution_plan["script_path"],
            "C:\\capture\\excel-om\\excel_om_windows_capture\\scripts\\capture.ps1"
        );
        assert_eq!(
            execution_plan["direct_exec_launcher_path"],
            "C:\\capture\\excel-om\\excel_om_windows_capture\\scripts\\run_capture.cmd"
        );
        assert_eq!(
            execution_plan["direct_exec_status_path"],
            "C:\\capture\\excel-om\\excel_om_windows_capture\\manifest\\direct_exec_status.json"
        );
        assert_eq!(execution_plan["commands"][0]["name"], "tlbimp_fallback");
        assert_eq!(
            execution_plan["pending_capture_outputs"],
            serde_json::json!([
                "raw_typelib_identity.json",
                "excel_typelib_snapshot.idl",
                "excel_typelib_snapshot.odl",
                "excel_pia_identity.json",
                "excel_pia_public_surface.json"
            ])
        );

        let direct_exec_status_template: serde_json::Value = serde_json::from_slice(
            &fs::read(&result.direct_exec_status_template_path)
                .expect("direct exec status template file"),
        )
        .expect("direct exec status template json");
        assert_eq!(direct_exec_status_template["status"], "pending");
        assert_eq!(
            direct_exec_status_template["launcherPath"],
            "C:\\capture\\excel-om\\excel_om_windows_capture\\scripts\\run_capture.cmd"
        );

        let receipt_template: serde_json::Value = serde_json::from_slice(
            &fs::read(&result.execution_receipt_template_path).expect("receipt template file"),
        )
        .expect("receipt template json");
        assert_eq!(
            receipt_template["expectedCaptureOutputs"],
            serde_json::json!([
                "raw_typelib_identity.json",
                "excel_typelib_snapshot.idl",
                "excel_typelib_snapshot.odl",
                "excel_pia_identity.json",
                "excel_pia_public_surface.json"
            ])
        );
        assert_eq!(receipt_template["commandResults"][0]["status"], "pending");
        assert_eq!(
            receipt_template["manualStepResults"][0]["status"],
            "pending"
        );
    }

    #[test]
    fn completes_execution_bundle_from_materialized_receipt_and_artifacts() {
        let plan = CapturePlan::from_toml_str(resolved_template()).expect("plan");
        let tempdir = TempDir::new().expect("tempdir");
        plan.materialize_execution_bundle(tempdir.path())
            .expect("materialize execution bundle");

        fs::write(tempdir.path().join("logs/capture.log"), "capture log").expect("capture log");
        fs::write(
            tempdir.path().join("raw/raw_typelib_identity.json"),
            r#"{"library":"Excel"}"#,
        )
        .expect("raw typelib identity");
        fs::write(
            tempdir.path().join("snapshots/excel_typelib_snapshot.idl"),
            "library Excel {}",
        )
        .expect("snapshot idl");
        fs::write(
            tempdir.path().join("snapshots/excel_typelib_snapshot.odl"),
            "odl Excel {}",
        )
        .expect("snapshot odl");
        fs::write(
            tempdir.path().join("raw/excel_pia_identity.json"),
            r#"{"assembly":"Microsoft.Office.Interop.Excel"}"#,
        )
        .expect("pia identity");
        fs::write(
            tempdir
                .path()
                .join("snapshots/excel_pia_public_surface.json"),
            r#"{"library":"Excel","interfaces":[]}"#,
        )
        .expect("pia public surface");

        let receipt_path = tempdir.path().join("manifest/execution_receipt.json");
        fs::write(
            &receipt_path,
            serde_json::to_vec_pretty(&sample_execution_receipt()).expect("receipt payload"),
        )
        .expect("execution receipt");

        let result = plan
            .complete_execution_bundle(tempdir.path())
            .expect("complete execution bundle");

        assert_eq!(result.execution_receipt_path, receipt_path);
        assert!(
            tempdir
                .path()
                .join("manifest/capture_manifest.json")
                .exists()
        );
        assert!(
            tempdir
                .path()
                .join("manifest/output_checksums.json")
                .exists()
        );

        let manifest: serde_json::Value = serde_json::from_slice(
            &fs::read(tempdir.path().join("manifest/capture_manifest.json"))
                .expect("manifest file"),
        )
        .expect("manifest json");
        assert_eq!(
            manifest["executionReceipt"]["commandResults"][0]["status"],
            "completed"
        );
        assert_eq!(
            manifest["executionReceipt"]["host"]["computerName"],
            "WIN-EXCEL"
        );
        assert_eq!(
            manifest["expectedCaptureOutputs"],
            manifest["executionReceipt"]["expectedCaptureOutputs"]
        );
        assert!(
            result
                .write_result
                .checksums
                .contains_key("raw/raw_typelib_identity.json")
        );
    }

    #[test]
    fn completes_execution_bundle_from_explicit_receipt_path() {
        let plan = CapturePlan::from_toml_str(resolved_template()).expect("plan");
        let tempdir = TempDir::new().expect("tempdir");
        plan.materialize_execution_bundle(tempdir.path())
            .expect("materialize execution bundle");

        fs::write(tempdir.path().join("logs/capture.log"), "capture log").expect("capture log");
        fs::write(
            tempdir.path().join("raw/raw_typelib_identity.json"),
            r#"{"library":"Excel"}"#,
        )
        .expect("raw typelib identity");
        fs::write(
            tempdir.path().join("snapshots/excel_typelib_snapshot.idl"),
            "library Excel {}",
        )
        .expect("snapshot idl");
        fs::write(
            tempdir.path().join("snapshots/excel_typelib_snapshot.odl"),
            "odl Excel {}",
        )
        .expect("snapshot odl");
        fs::write(
            tempdir.path().join("raw/excel_pia_identity.json"),
            r#"{"assembly":"Microsoft.Office.Interop.Excel"}"#,
        )
        .expect("pia identity");
        fs::write(
            tempdir
                .path()
                .join("snapshots/excel_pia_public_surface.json"),
            r#"{"library":"Excel","interfaces":[]}"#,
        )
        .expect("pia public surface");

        let receipt_dir = tempdir.path().join("custom");
        fs::create_dir_all(&receipt_dir).expect("receipt dir");
        let receipt_path = receipt_dir.join("execution_receipt.json");
        fs::write(
            &receipt_path,
            serde_json::to_vec_pretty(&sample_execution_receipt()).expect("receipt payload"),
        )
        .expect("execution receipt");

        let result = plan
            .complete_execution_bundle_with_receipt_path(tempdir.path(), Some(&receipt_path))
            .expect("complete execution bundle");

        assert_eq!(result.execution_receipt_path, receipt_path);

        let manifest: serde_json::Value = serde_json::from_slice(
            &fs::read(tempdir.path().join("manifest/capture_manifest.json"))
                .expect("manifest file"),
        )
        .expect("manifest json");
        assert_eq!(
            manifest["executionReceipt"]["host"]["computerName"],
            "WIN-EXCEL"
        );
    }

    #[test]
    fn completion_requires_execution_receipt_file() {
        let plan = CapturePlan::from_toml_str(resolved_template()).expect("plan");
        let tempdir = TempDir::new().expect("tempdir");

        let error = plan
            .complete_execution_bundle(tempdir.path())
            .expect_err("receipt should be required");
        match error {
            CaptureBundleCompletionError::MissingArtifact { logical_name, .. } => {
                assert_eq!(logical_name, "execution_receipt");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn completion_requires_materialized_capture_log_artifact() {
        let plan = CapturePlan::from_toml_str(resolved_template()).expect("plan");
        let tempdir = TempDir::new().expect("tempdir");
        plan.materialize_execution_bundle(tempdir.path())
            .expect("materialize execution bundle");

        let receipt_path = tempdir.path().join("manifest/execution_receipt.json");
        fs::write(
            &receipt_path,
            serde_json::to_vec_pretty(&sample_execution_receipt()).expect("receipt payload"),
        )
        .expect("execution receipt");

        let error = plan
            .complete_execution_bundle(tempdir.path())
            .expect_err("capture log should be required");
        match error {
            CaptureBundleCompletionError::MissingArtifact { logical_name, .. } => {
                assert_eq!(logical_name, "capture_log");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn run_execution_bundle_requires_materialized_launcher() {
        let plan = CapturePlan::from_toml_str(resolved_template()).expect("plan");
        let tempdir = TempDir::new().expect("tempdir");

        let error = plan
            .run_execution_bundle(tempdir.path())
            .expect_err("launcher should be required");
        match error {
            CaptureDirectExecError::MissingArtifact { logical_name, .. } => {
                assert_eq!(logical_name, "direct_exec_launcher");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn run_execution_bundle_requires_materialized_capture_script() {
        let plan = CapturePlan::from_toml_str(resolved_template()).expect("plan");
        let tempdir = TempDir::new().expect("tempdir");
        plan.materialize_execution_bundle(tempdir.path())
            .expect("materialize execution bundle");
        fs::remove_file(tempdir.path().join("scripts/capture.ps1")).expect("remove capture script");

        let error = plan
            .run_execution_bundle(tempdir.path())
            .expect_err("capture script should be required");
        match error {
            CaptureDirectExecError::MissingArtifact { logical_name, .. } => {
                assert_eq!(logical_name, "execution_script");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[cfg(not(target_os = "windows"))]
    #[test]
    fn run_execution_bundle_rejects_non_windows_hosts() {
        let plan = CapturePlan::from_toml_str(resolved_template()).expect("plan");
        let tempdir = TempDir::new().expect("tempdir");
        plan.materialize_execution_bundle(tempdir.path())
            .expect("materialize execution bundle");

        let error = plan
            .run_execution_bundle(tempdir.path())
            .expect_err("non-windows host should be rejected");
        match error {
            CaptureDirectExecError::UnsupportedHostOs { actual } => {
                assert_eq!(actual, std::env::consts::OS);
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parses_execution_receipt_from_json_file() {
        let tempdir = TempDir::new().expect("tempdir");
        let receipt_path = tempdir.path().join("execution_receipt.json");
        fs::write(
            &receipt_path,
            serde_json::to_vec_pretty(&sample_execution_receipt()).expect("receipt payload"),
        )
        .expect("receipt file");

        let receipt =
            CaptureExecutionReceipt::from_json_path(&receipt_path).expect("execution receipt");
        assert_eq!(receipt.host.computer_name.as_deref(), Some("WIN-EXCEL"));
        assert_eq!(
            receipt.expected_capture_outputs,
            vec![
                "raw_typelib_identity.json".to_string(),
                "excel_typelib_snapshot.idl".to_string(),
                "excel_typelib_snapshot.odl".to_string(),
                "excel_pia_identity.json".to_string(),
                "excel_pia_public_surface.json".to_string(),
            ]
        );
        assert_eq!(receipt.command_results[0].status, "completed");
        assert_eq!(
            receipt.manual_step_results[0].name,
            "oleview_snapshot_export"
        );
    }

    #[test]
    fn rejects_malformed_execution_receipt_json_file() {
        let tempdir = TempDir::new().expect("tempdir");
        let receipt_path = tempdir.path().join("execution_receipt.json");
        fs::write(&receipt_path, "{").expect("receipt file");

        let error = CaptureExecutionReceipt::from_json_path(&receipt_path)
            .expect_err("malformed execution receipt should fail");
        match error {
            CaptureBundleCompletionError::Json { path, .. } => {
                assert_eq!(path, receipt_path);
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parses_direct_exec_status_from_json_file() {
        let tempdir = TempDir::new().expect("tempdir");
        let status_path = tempdir.path().join("direct_exec_status.json");
        fs::write(
            &status_path,
            r#"{
  "startedAtUtc":"2026-03-27T01:02:03Z",
  "completedAtUtc":"2026-03-27T01:03:04Z",
  "status":"completed",
  "exitCode":0,
  "launcherPath":"C:\\capture\\scripts\\run_capture.cmd",
  "captureScriptPath":"C:\\capture\\scripts\\capture.ps1",
  "executionReceiptPath":"C:\\capture\\manifest\\execution_receipt.json",
  "captureManifestPath":"C:\\capture\\manifest\\capture_manifest.json",
  "outputChecksumsPath":"C:\\capture\\manifest\\output_checksums.json"
}"#,
        )
        .expect("status file");

        let status =
            CaptureDirectExecStatus::from_json_path(&status_path).expect("direct exec status");
        assert_eq!(status.status, "completed");
        assert_eq!(status.exit_code, Some(0));
        assert_eq!(
            status.launcher_path,
            "C:\\capture\\scripts\\run_capture.cmd"
        );
    }

    #[test]
    fn reports_missing_direct_exec_status_file() {
        let tempdir = TempDir::new().expect("tempdir");
        let status_path = tempdir.path().join("direct_exec_status.json");

        let error = CaptureDirectExecStatus::from_json_path(&status_path)
            .expect_err("missing direct exec status should fail");
        match error {
            CaptureDirectExecError::MissingArtifact { logical_name, path } => {
                assert_eq!(logical_name, "direct_exec_status");
                assert_eq!(path, status_path);
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn rejects_malformed_direct_exec_status_json_file() {
        let tempdir = TempDir::new().expect("tempdir");
        let status_path = tempdir.path().join("direct_exec_status.json");
        fs::write(&status_path, "{").expect("status file");

        let error = CaptureDirectExecStatus::from_json_path(&status_path)
            .expect_err("malformed direct exec status should fail");
        match error {
            CaptureDirectExecError::Json { path, .. } => {
                assert_eq!(path, status_path);
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn builds_execution_plan_with_manual_oleview_and_reflection_steps() {
        let plan = CapturePlan::from_toml_str(resolved_template()).expect("plan");
        let execution_plan = plan.execution_plan();

        assert_eq!(
            execution_plan.script_path,
            "C:\\capture\\excel-om\\excel_om_windows_capture\\scripts\\capture.ps1"
        );
        assert_eq!(
            execution_plan.direct_exec_launcher_path,
            "C:\\capture\\excel-om\\excel_om_windows_capture\\scripts\\run_capture.cmd"
        );
        assert_eq!(
            execution_plan.direct_exec_status_path,
            "C:\\capture\\excel-om\\excel_om_windows_capture\\manifest\\direct_exec_status.json"
        );
        assert_eq!(
            execution_plan.pending_capture_outputs,
            vec![
                "raw_typelib_identity.json".to_string(),
                "excel_typelib_snapshot.idl".to_string(),
                "excel_typelib_snapshot.odl".to_string(),
                "excel_pia_identity.json".to_string(),
                "excel_pia_public_surface.json".to_string(),
            ]
        );
        assert_eq!(execution_plan.manual_steps.len(), 1);
        assert_eq!(
            execution_plan.manual_steps[0].name,
            "oleview_snapshot_export"
        );
        assert_eq!(
            execution_plan.manual_steps[0].outputs,
            vec![
                "C:\\capture\\excel-om\\excel_om_windows_capture\\snapshots\\excel_typelib_snapshot.idl".to_string(),
                "C:\\capture\\excel-om\\excel_om_windows_capture\\snapshots\\excel_typelib_snapshot.odl".to_string(),
            ]
        );
        assert_eq!(execution_plan.commands.len(), 2);
        assert_eq!(execution_plan.commands[0].name, "tlbimp_fallback");
        assert!(
            execution_plan.commands[0]
                .args
                .iter()
                .any(|arg| arg.contains("/namespace:Microsoft.Office.Interop.Excel"))
        );
        assert_eq!(
            execution_plan.commands[1].outputs,
            vec![
                "C:\\capture\\excel-om\\excel_om_windows_capture\\logs\\capture.log".to_string(),
                "C:\\capture\\excel-om\\excel_om_windows_capture\\raw\\raw_typelib_identity.json".to_string(),
                "C:\\capture\\excel-om\\excel_om_windows_capture\\raw\\excel_pia_identity.json".to_string(),
                "C:\\capture\\excel-om\\excel_om_windows_capture\\snapshots\\excel_pia_public_surface.json".to_string(),
            ]
        );
    }

    #[test]
    fn renders_powershell_script_without_step_three_output() {
        let script = CapturePlan::from_toml_str(resolved_template())
            .expect("plan")
            .render_powershell_script();

        assert!(script.contains("Manual step required: launch oleview.exe"));
        assert!(script.contains("ConvertTo-Json -Depth 16"));
        assert!(script.contains("$GeneratedInteropAssembly"));
        assert!(script.contains("$ExecutionReceiptPath"));
        assert!(script.contains("raw\\raw_typelib_identity.json"));
        assert!(script.contains("snapshots\\excel_typelib_snapshot.idl"));
        assert!(!script.contains("office_idl_excel_om.json"));
    }

    #[test]
    fn renders_direct_exec_launcher_with_status_capture() {
        let launcher = CapturePlan::from_toml_str(resolved_template())
            .expect("plan")
            .render_direct_exec_launcher();

        assert!(launcher.contains("@echo off"));
        assert!(launcher.contains("capture.ps1"));
        assert!(launcher.contains("direct_exec_status.json"));
        assert!(launcher.contains("execution_receipt.json"));
        assert!(launcher.contains("exit /b %EXIT_CODE%"));
    }

    #[test]
    fn sanitizes_file_components_for_generated_artifact_names() {
        assert_eq!(
            sanitize_file_component("Microsoft Office/Interop:Excel?*"),
            "Microsoft_Office_Interop_Excel__"
        );
        assert_eq!(
            sanitize_file_component("Microsoft.Office-Interop_Excel"),
            "Microsoft.Office-Interop_Excel"
        );
    }

    #[test]
    fn reports_ready_state_for_fully_resolved_plan() {
        let summary = CapturePlan::from_toml_str(resolved_template())
            .expect("plan")
            .summary();
        assert!(summary.ready_to_run);
        assert_eq!(summary.unresolved_fields, Vec::<String>::new());
    }
}
