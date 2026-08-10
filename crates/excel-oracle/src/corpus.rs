use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};

use crate::{
    CaseSpec, ComparisonPolicy, EngineIdentity, EngineKind, OracleContractError,
    OracleSuiteManifest, RunBundle, RunCaseStatus, RunManifest,
    compare_repeated_excel_observations,
};

pub const SUITE_MANIFEST_PATH: &str = "manifest/suite_manifest.json";
pub const RUN_MANIFEST_PATH: &str = "manifest/run_manifest.json";

const MAX_MANIFEST_BYTES: u64 = 16 * 1024 * 1024;
const MAX_CASE_BYTES: u64 = 16 * 1024 * 1024;
const MAX_OBSERVATION_BYTES: u64 = 64 * 1024 * 1024;
const MAX_INPUT_BYTES: u64 = 512 * 1024 * 1024;

#[derive(Debug, Clone, PartialEq)]
pub struct PinnedSuiteArtifacts {
    pub manifest: OracleSuiteManifest,
    pub cases: BTreeMap<String, Vec<u8>>,
    pub inputs: BTreeMap<String, Vec<u8>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RepeatedExcelRunEvidence {
    pub engine: EngineIdentity,
    pub run_ids: [String; 2],
    pub verified_case_ids: Vec<String>,
}

impl PinnedSuiteArtifacts {
    pub fn load(root: impl AsRef<Path>) -> Result<Self, OracleContractError> {
        let root = validate_artifact_root(root.as_ref(), "suite artifact root")?;
        let manifest_bytes = read_regular_artifact(
            &root,
            SUITE_MANIFEST_PATH,
            "suite manifest",
            MAX_MANIFEST_BYTES,
        )?;
        let manifest_json = std::str::from_utf8(&manifest_bytes).map_err(|error| {
            OracleContractError::new(format!("suite manifest was not UTF-8: {error}"))
        })?;
        let manifest = OracleSuiteManifest::from_json_str(manifest_json)?;

        let mut cases = BTreeMap::new();
        let mut inputs = BTreeMap::new();
        for artifact in &manifest.cases {
            let case_bytes =
                read_regular_artifact(&root, &artifact.path, "case artifact", MAX_CASE_BYTES)?;
            let case = manifest.load_case(&artifact.case_id, &case_bytes)?;
            let input_bytes =
                read_regular_artifact(&root, &case.input.path, "case input", MAX_INPUT_BYTES)?;
            artifact.validate_input(&case, &input_bytes)?;
            cases.insert(artifact.case_id.clone(), case_bytes);
            inputs.insert(artifact.case_id.clone(), input_bytes);
        }

        Ok(Self {
            manifest,
            cases,
            inputs,
        })
    }

    pub fn load_run(&self, root: impl AsRef<Path>) -> Result<RunBundle, OracleContractError> {
        let root = validate_artifact_root(root.as_ref(), "run artifact root")?;
        let manifest_bytes =
            read_regular_artifact(&root, RUN_MANIFEST_PATH, "run manifest", MAX_MANIFEST_BYTES)?;
        let manifest_json = std::str::from_utf8(&manifest_bytes).map_err(|error| {
            OracleContractError::new(format!("run manifest was not UTF-8: {error}"))
        })?;
        let manifest = RunManifest::from_json_str(&self.manifest, manifest_json)?;

        let mut observations = BTreeMap::new();
        for record in &manifest.cases {
            if record.status != RunCaseStatus::Completed {
                continue;
            }
            let observation_path = record
                .observation_path
                .as_deref()
                .expect("validated completed run record observation path");
            let observation_bytes = read_regular_artifact(
                &root,
                observation_path,
                "run observation",
                MAX_OBSERVATION_BYTES,
            )?;
            let case = self.load_case(&record.case_id)?;
            record.load_observation(&case, &manifest.engine, &observation_bytes)?;
            observations.insert(record.case_id.clone(), observation_bytes);
        }

        Ok(RunBundle {
            manifest,
            observations,
        })
    }

    pub fn verify_repeated_excel_runs(
        &self,
        first: &RunBundle,
        second: &RunBundle,
    ) -> Result<RepeatedExcelRunEvidence, OracleContractError> {
        first.validate(&self.manifest)?;
        second.validate(&self.manifest)?;
        first
            .manifest
            .validate_required_completeness(&self.manifest)?;
        second
            .manifest
            .validate_required_completeness(&self.manifest)?;
        if first.manifest.engine.kind != EngineKind::Excel
            || second.manifest.engine.kind != EngineKind::Excel
        {
            return Err(OracleContractError::new(
                "repeated evidence requires two desktop Excel runs",
            ));
        }
        if first
            .manifest
            .run_id
            .eq_ignore_ascii_case(&second.manifest.run_id)
        {
            return Err(OracleContractError::new(
                "repeated Excel runs must use distinct run ids",
            ));
        }

        let mut verified_case_ids = Vec::new();
        for artifact in &self.manifest.cases {
            let case = self.load_case(&artifact.case_id)?;
            let first_record = first
                .manifest
                .cases
                .iter()
                .find(|record| record.case_id == artifact.case_id)
                .expect("validated first run coverage");
            let second_record = second
                .manifest
                .cases
                .iter()
                .find(|record| record.case_id == artifact.case_id)
                .expect("validated second run coverage");
            if first_record.status != second_record.status {
                return Err(OracleContractError::new(format!(
                    "repeated Excel run status diverged for case {}",
                    artifact.case_id,
                )));
            }
            match first_record.status {
                RunCaseStatus::Completed => {
                    let first_observation = first_record.load_observation(
                        &case,
                        &first.manifest.engine,
                        first
                            .observations
                            .get(&artifact.case_id)
                            .expect("validated first observation coverage"),
                    )?;
                    let second_observation = second_record.load_observation(
                        &case,
                        &second.manifest.engine,
                        second
                            .observations
                            .get(&artifact.case_id)
                            .expect("validated second observation coverage"),
                    )?;
                    let comparison = compare_repeated_excel_observations(
                        &case,
                        &first_observation,
                        &second_observation,
                        ComparisonPolicy::default(),
                    )?;
                    if let Some(mismatch) = comparison.mismatches.first() {
                        return Err(OracleContractError::new(format!(
                            "repeated Excel runs diverged for case {} at {}",
                            artifact.case_id, mismatch.path,
                        )));
                    }
                    verified_case_ids.push(artifact.case_id.clone());
                }
                RunCaseStatus::Failed => {
                    return Err(OracleContractError::new(format!(
                        "failed case {} cannot become repeated Excel evidence",
                        artifact.case_id,
                    )));
                }
                RunCaseStatus::Unsupported | RunCaseStatus::Skipped => {}
            }
        }
        if verified_case_ids.is_empty() {
            return Err(OracleContractError::new(
                "repeated Excel runs did not complete any cases",
            ));
        }

        Ok(RepeatedExcelRunEvidence {
            engine: first.manifest.engine.clone(),
            run_ids: [
                first.manifest.run_id.clone(),
                second.manifest.run_id.clone(),
            ],
            verified_case_ids,
        })
    }

    fn load_case(&self, case_id: &str) -> Result<CaseSpec, OracleContractError> {
        let bytes = self.cases.get(case_id).ok_or_else(|| {
            OracleContractError::new(format!("suite case artifact {case_id} was not loaded"))
        })?;
        self.manifest.load_case(case_id, bytes)
    }
}

fn validate_artifact_root(root: &Path, label: &str) -> Result<PathBuf, OracleContractError> {
    let metadata = fs::symlink_metadata(root)
        .map_err(|error| OracleContractError::new(format!("failed to inspect {label}: {error}")))?;
    if metadata.file_type().is_symlink() {
        return Err(OracleContractError::new(format!(
            "{label} must not be a symbolic link"
        )));
    }
    if !metadata.is_dir() {
        return Err(OracleContractError::new(format!(
            "{label} must be a directory"
        )));
    }
    Ok(root.to_path_buf())
}

fn read_regular_artifact(
    root: &Path,
    relative: &str,
    label: &str,
    max_bytes: u64,
) -> Result<Vec<u8>, OracleContractError> {
    validate_portable_artifact_path(relative, label)?;
    let segments = relative.split('/').collect::<Vec<_>>();
    let mut path = root.to_path_buf();
    for (index, segment) in segments.iter().enumerate() {
        path.push(segment);
        let metadata = fs::symlink_metadata(&path).map_err(|error| {
            OracleContractError::new(format!("failed to inspect {label} {relative}: {error}"))
        })?;
        if metadata.file_type().is_symlink() {
            return Err(OracleContractError::new(format!(
                "{label} must not be a symbolic link: {relative}"
            )));
        }
        if index + 1 == segments.len() {
            if !metadata.is_file() {
                return Err(OracleContractError::new(format!(
                    "{label} must be a regular file: {relative}"
                )));
            }
            if metadata.len() > max_bytes {
                return Err(OracleContractError::new(format!(
                    "{label} exceeded the {max_bytes}-byte limit: {relative}"
                )));
            }
        } else if !metadata.is_dir() {
            return Err(OracleContractError::new(format!(
                "{label} parent must be a directory: {relative}"
            )));
        }
    }

    let bytes = fs::read(&path).map_err(|error| {
        OracleContractError::new(format!("failed to read {label} {relative}: {error}"))
    })?;
    if u64::try_from(bytes.len()).unwrap_or(u64::MAX) > max_bytes {
        return Err(OracleContractError::new(format!(
            "{label} exceeded the {max_bytes}-byte limit while reading: {relative}"
        )));
    }
    Ok(bytes)
}

fn validate_portable_artifact_path(relative: &str, label: &str) -> Result<(), OracleContractError> {
    crate::validate_safe_relative_path(label, relative)?;
    for segment in relative.split('/') {
        if segment.chars().any(|character| {
            character.is_control() || matches!(character, '<' | '>' | ':' | '"' | '|' | '?' | '*')
        }) || segment.ends_with([' ', '.'])
        {
            return Err(OracleContractError::new(format!(
                "{label} must use portable path segments"
            )));
        }
        let device_stem = segment
            .split_once('.')
            .map_or(segment, |(stem, _)| stem)
            .to_ascii_uppercase();
        if matches!(device_stem.as_str(), "CON" | "PRN" | "AUX" | "NUL")
            || device_stem
                .strip_prefix("COM")
                .or_else(|| device_stem.strip_prefix("LPT"))
                .is_some_and(|suffix| {
                    suffix.len() == 1 && matches!(suffix.as_bytes()[0], b'1'..=b'9')
                })
        {
            return Err(OracleContractError::new(format!(
                "{label} must not use a Windows device-name segment"
            )));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::validate_portable_artifact_path;

    #[test]
    fn portable_artifact_paths_reject_windows_aliases() {
        for path in [
            "C:relative.json",
            "cases/name?.json",
            "cases/name. ",
            "cases/CON.json",
            "observations/com1/result.json",
        ] {
            assert!(
                validate_portable_artifact_path(path, "artifact").is_err(),
                "{path} must be rejected",
            );
        }
        validate_portable_artifact_path("cases/application.name.json", "artifact")
            .expect("portable path");
    }
}
