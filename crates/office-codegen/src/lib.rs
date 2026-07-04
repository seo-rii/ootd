use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::{Component, Path, PathBuf};

use office_idl::{
    AccessMode, CaptureMetadata, CaptureOrigin, CaptureOriginKind, ClassDef, EnumDef, EnumValue,
    InterfaceDef, InterfaceKind, Member, MemberKind, OfficeIdlDocument, Parameter, SidecarMetadata,
    SupportState, TypeRef, TypeRefKind,
};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CodegenSummary {
    pub library: String,
    pub version: String,
    pub enum_count: usize,
    pub interface_count: usize,
    pub class_count: usize,
    pub member_count: usize,
    pub implemented_member_count: usize,
    pub partial_member_count: usize,
    pub stub_member_count: usize,
    pub unsupported_member_count: usize,
}

#[derive(Debug)]
pub enum OmSourcesLoadError {
    Io(std::io::Error),
    Toml(toml::de::Error),
}

#[derive(Debug)]
pub enum PiaCaptureLoadError {
    Io(std::io::Error),
    Json(serde_json::Error),
}

#[derive(Debug)]
pub enum DifferentialReportLoadError {
    Io(std::io::Error),
    Json(serde_json::Error),
    Contract { message: String },
}

#[derive(Debug)]
pub enum OmCaptureBundleError {
    LibraryMismatch { typelib: String, pia: String },
    VersionMismatch { typelib: String, pia: String },
    NamespaceMismatch { typelib: String, pia: String },
}

#[derive(Debug)]
pub enum CanonicalOmGenerationError {
    Io {
        action: &'static str,
        path: PathBuf,
        source: std::io::Error,
    },
    Json {
        path: PathBuf,
        source: serde_json::Error,
    },
    Normalize(OmCaptureBundleError),
    CaptureBundleContract {
        message: String,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CaptureBundlePaths {
    pub bundle_root_path: PathBuf,
    pub raw_typelib_identity_path: PathBuf,
    pub excel_pia_public_surface_path: PathBuf,
    pub capture_manifest_path: PathBuf,
    pub output_checksums_path: PathBuf,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CanonicalOmGenerationResult {
    pub output_path: PathBuf,
    pub bundle_paths: CaptureBundlePaths,
    pub summary: OmCaptureBundleSummary,
}

pub const DIFFERENTIAL_REPORT_ARTIFACT_NAME: &str = "differential_report.json";
pub const DIFFERENTIAL_GATE_SUMMARY_ARTIFACT_NAME: &str = "differential_gate_summary.json";

pub const PRIORITY_OM_SURFACES: [&str; 58] = [
    "Application",
    "WorksheetFunction",
    "Workbooks",
    "Workbook",
    "Worksheets",
    "Sheets",
    "Charts",
    "Worksheet",
    "Range",
    "Areas",
    "Names",
    "Name",
    "ChartObjects",
    "ChartObject",
    "ShapeRange",
    "Chart",
    "ChartArea",
    "PlotArea",
    "ChartTitle",
    "Legend",
    "LegendEntries",
    "LegendEntry",
    "LegendKey",
    "DataTable",
    "ChartFormat",
    "Adjustments",
    "FillFormat",
    "GlowFormat",
    "LineFormat",
    "PictureFormat",
    "Crop",
    "ShadowFormat",
    "SoftEdgeFormat",
    "TextFrame2",
    "ThreeDFormat",
    "ChartGroups",
    "ChartGroup",
    "CategoryCollection",
    "ChartCategory",
    "SeriesLines",
    "DropLines",
    "HiLoLines",
    "UpBars",
    "DownBars",
    "Axes",
    "Axis",
    "TickLabels",
    "Gridlines",
    "DisplayUnitLabel",
    "AxisTitle",
    "SeriesCollection",
    "Series",
    "LeaderLines",
    "Border",
    "DataLabels",
    "DataLabel",
    "Points",
    "Point",
];

#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SupportStateCounts {
    pub generated_only: usize,
    pub stub: usize,
    pub partial: usize,
    pub implemented: usize,
    pub oracle_verified: usize,
    pub unsupported: usize,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OmTypeMetadataEntry {
    pub kind: TypeRefKind,
    pub name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alias_of: Option<String>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub nullable: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OmParameterMetadataEntry {
    pub name: String,
    #[serde(rename = "type")]
    pub type_ref: OmTypeMetadataEntry,
    #[serde(default, skip_serializing_if = "is_false")]
    pub optional: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub by_ref: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OmMemberMetadataEntry {
    pub name: String,
    pub member_kind: MemberKind,
    pub access: AccessMode,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub disp_id: Option<i32>,
    pub support: SupportState,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub return_type: Option<OmTypeMetadataEntry>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub params: Vec<OmParameterMetadataEntry>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub capture_origin_kinds: Vec<CaptureOriginKind>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OmFocusSurfaceRegistryEntry {
    pub name: String,
    pub kind: InterfaceKind,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub iid: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub inherits: Vec<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub implementing_classes: Vec<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub default_coclasses: Vec<String>,
    pub member_count: usize,
    pub support_counts: SupportStateCounts,
    pub members: Vec<OmMemberMetadataEntry>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OmFocusSurfaceRegistry {
    pub library: String,
    pub version: String,
    pub focus_surfaces: Vec<OmFocusSurfaceRegistryEntry>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub missing_focus_surfaces: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OmFocusSurfaceCoverageEntry {
    pub name: String,
    pub member_count: usize,
    pub support_counts: SupportStateCounts,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub generated_only_members: Vec<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub stub_members: Vec<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub partial_members: Vec<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub implemented_members: Vec<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub oracle_verified_members: Vec<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub unsupported_members: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OmCoverageReport {
    pub library: String,
    pub version: String,
    pub enum_count: usize,
    pub interface_count: usize,
    pub class_count: usize,
    pub member_count: usize,
    pub support_counts: SupportStateCounts,
    pub focus_surfaces: Vec<OmFocusSurfaceCoverageEntry>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub missing_focus_surfaces: Vec<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum DifferentialCaseStatus {
    Passed,
    Failed,
    MissingOracle,
    MissingRuntime,
    Unsupported,
    Skipped,
}

#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DifferentialStatusCounts {
    pub passed: usize,
    pub failed: usize,
    pub missing_oracle: usize,
    pub missing_runtime: usize,
    pub unsupported: usize,
    pub skipped: usize,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DifferentialCaseResult {
    pub name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub surface: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub member: Option<String>,
    pub status: DifferentialCaseStatus,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub expected: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actual: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub message: Option<String>,
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub artifacts: BTreeMap<String, String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DifferentialReportContext {
    pub project_name: String,
    pub default_profile: String,
    pub default_mode: String,
    pub primary_om_artifact: String,
    pub primary_ooxml_source: String,
    pub enabled_corpus_groups: Vec<String>,
    pub enabled_corpus_source_count: usize,
    pub validation_modes: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DifferentialReport {
    pub library: String,
    pub version: String,
    pub profile: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub context: Option<DifferentialReportContext>,
    pub case_count: usize,
    pub status_counts: DifferentialStatusCounts,
    pub cases: Vec<DifferentialCaseResult>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DifferentialGateSummary {
    pub passed: bool,
    pub blocking_case_count: usize,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub blocking_cases: Vec<String>,
    pub incomplete_oracle_count: usize,
    pub missing_runtime_count: usize,
    pub failed_case_count: usize,
    pub unsupported_case_count: usize,
    pub skipped_case_count: usize,
}

impl From<std::io::Error> for OmSourcesLoadError {
    fn from(value: std::io::Error) -> Self {
        Self::Io(value)
    }
}

impl From<toml::de::Error> for OmSourcesLoadError {
    fn from(value: toml::de::Error) -> Self {
        Self::Toml(value)
    }
}

impl From<std::io::Error> for PiaCaptureLoadError {
    fn from(value: std::io::Error) -> Self {
        Self::Io(value)
    }
}

impl From<serde_json::Error> for PiaCaptureLoadError {
    fn from(value: serde_json::Error) -> Self {
        Self::Json(value)
    }
}

impl From<std::io::Error> for DifferentialReportLoadError {
    fn from(value: std::io::Error) -> Self {
        Self::Io(value)
    }
}

impl From<serde_json::Error> for DifferentialReportLoadError {
    fn from(value: serde_json::Error) -> Self {
        Self::Json(value)
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct OmSourcesManifest {
    pub manifest: OmSourcesHeader,
    pub contract: OmSourcesContract,
    pub artifacts: OmSourceArtifacts,
    #[serde(default)]
    pub docs: OmSourceDocs,
    pub target_capture: OmTargetCapture,
    pub acquisition: OmAcquisition,
    #[serde(default)]
    pub references: BTreeMap<String, String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct OmSourcesHeader {
    pub schema_version: u32,
    pub name: String,
    pub status: String,
    pub owner: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct OmSourcesContract {
    pub domain: String,
    pub primary_artifact: String,
    pub secondary_artifact: String,
    pub behavior_docs: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct OmSourceArtifacts {
    pub excel_type_library: OmTypeLibraryArtifact,
    pub excel_primary_interop_assembly: OmPrimaryInteropArtifact,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct OmTypeLibraryArtifact {
    pub kind: String,
    pub platform: String,
    pub app: String,
    pub major: u32,
    pub minor: u32,
    pub machine_readable: bool,
    pub acquisition_priority: u32,
    pub notes: String,
    pub capture: OmArtifactCapture,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct OmPrimaryInteropArtifact {
    pub kind: String,
    pub assembly: String,
    pub assembly_file: String,
    pub namespace: String,
    pub machine_readable: bool,
    pub acquisition_priority: u32,
    pub notes: String,
    pub capture: OmArtifactCapture,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct OmArtifactCapture {
    pub status: String,
    pub tooling: Vec<String>,
    pub required_outputs: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize, Default)]
pub struct OmSourceDocs {
    pub excel_vba_reference: Option<OmReferenceDoc>,
    pub office_library_reference: Option<OmReferenceDoc>,
    pub interop_namespace: Option<OmReferenceDoc>,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct OmReferenceDoc {
    pub kind: String,
    pub url: String,
    pub acquisition_priority: u32,
    pub notes: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct OmTargetCapture {
    pub status: String,
    pub product_family: String,
    pub channel: String,
    pub version: String,
    pub build: String,
    pub arch: String,
    pub locale: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct OmAcquisition {
    pub host_os: String,
    pub requires_installed_excel: bool,
    pub requires_windows_sdk: bool,
    pub requires_dotnet_framework_tooling: bool,
    pub normalization_target: String,
    pub normalization_output: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct SourceRegistryManifest {
    pub project: SourceRegistryProject,
    pub om_contract: SourceRegistryOmContract,
    pub ooxml: SourceRegistryOoxml,
    #[serde(default)]
    pub binary_formats: BTreeMap<String, String>,
    #[serde(default)]
    pub behavior: BTreeMap<String, String>,
    pub test_corpus: SourceRegistryTestCorpus,
    pub validation: SourceRegistryValidation,
    #[serde(default)]
    pub profiles: BTreeMap<String, SourceRegistryProfile>,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct SourceRegistryProject {
    pub name: String,
    pub default_profile: String,
    pub default_mode: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct SourceRegistryOmContract {
    pub primary: String,
    pub secondary: String,
    pub docs_primary: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct SourceRegistryOoxml {
    pub primary: String,
    pub packaging: String,
    pub implementation_notes: String,
    pub excel_extensions: String,
    pub shared_structures: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct SourceRegistryTestCorpus {
    #[serde(default)]
    pub synthetic: bool,
    #[serde(default)]
    pub real_world: bool,
    #[serde(default)]
    pub official_ms: SourceRegistryOfficialMsCorpus,
    #[serde(default)]
    pub open_source: SourceRegistryOpenSourceCorpus,
}

#[derive(Debug, Clone, Default, PartialEq, Eq, Deserialize)]
pub struct SourceRegistryOfficialMsCorpus {
    #[serde(default)]
    pub enabled: bool,
    #[serde(default)]
    pub office_scripts_samples: bool,
    #[serde(default)]
    pub data_validation_examples: bool,
    #[serde(default)]
    pub power_bi_financial_sample: bool,
    #[serde(default)]
    pub mos_excel_course_materials: bool,
    #[serde(default)]
    pub mos_excel_expert_course_materials: bool,
}

#[derive(Debug, Clone, Default, PartialEq, Eq, Deserialize)]
pub struct SourceRegistryOpenSourceCorpus {
    #[serde(default)]
    pub enabled: bool,
    #[serde(default)]
    pub open_xml_sdk: bool,
    #[serde(default)]
    pub apache_poi_test_data: bool,
    #[serde(default)]
    pub libreoffice_sc_qa_unit_data: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct SourceRegistryValidation {
    pub openxml_validator: bool,
    pub excel_oracle: bool,
    pub render_snapshot: bool,
    pub fuzz: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct SourceRegistryProfile {
    pub dynamic_arrays: bool,
    pub implicit_intersection_at: bool,
    pub strict_write: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceRegistrySummary {
    pub project_name: String,
    pub default_profile: String,
    pub default_mode: String,
    pub primary_om_artifact: String,
    pub secondary_om_artifact: String,
    pub primary_docs_source: String,
    pub primary_ooxml_source: String,
    pub enabled_corpus_groups: Vec<String>,
    pub official_ms_corpus_sources: Vec<String>,
    pub open_source_corpus_sources: Vec<String>,
    pub enabled_corpus_source_count: usize,
    pub validation_modes: Vec<String>,
    pub profile_count: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DifferentialArtifactContract {
    pub report_artifact: String,
    pub gate_summary_artifact: String,
    pub artifact_names: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DifferentialArtifactPaths {
    pub output_root_path: PathBuf,
    pub report_path: PathBuf,
    pub gate_summary_path: PathBuf,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DifferentialArtifactBundle {
    pub paths: DifferentialArtifactPaths,
    pub report: DifferentialReport,
    pub gate_summary: DifferentialGateSummary,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OmCaptureSummary {
    pub primary_artifact: String,
    pub secondary_artifact: String,
    pub ready_for_windows_capture: bool,
    pub machine_readable_artifact_count: usize,
    pub pending_outputs: Vec<String>,
    pub pending_output_count: usize,
    pub behavior_doc_count: usize,
    pub unresolved_target_fields: Vec<&'static str>,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TypelibIdentityCapture {
    pub library: String,
    pub version: String,
    pub namespace: String,
    pub type_library_guid: String,
    #[serde(default)]
    pub interfaces: Vec<TypelibInterfaceIdentity>,
    #[serde(default)]
    pub coclasses: Vec<TypelibCoclassIdentity>,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TypelibInterfaceIdentity {
    pub name: String,
    pub iid: String,
    pub kind: String,
    #[serde(default)]
    pub inherits: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TypelibCoclassIdentity {
    pub name: String,
    pub clsid: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_interface: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OmCaptureBundleSummary {
    pub library: String,
    pub version: String,
    pub type_library_guid: String,
    pub interface_iid_count: usize,
    pub coclass_clsid_count: usize,
    pub missing_pia_interfaces: Vec<String>,
    pub missing_pia_classes: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PiaPublicSurfaceCapture {
    pub library: String,
    pub version: String,
    pub namespace: String,
    #[serde(default)]
    pub enums: Vec<PiaCaptureEnum>,
    #[serde(default)]
    pub interfaces: Vec<PiaCaptureInterface>,
    #[serde(default)]
    pub classes: Vec<PiaCaptureClass>,
}

#[derive(Debug, Clone, PartialEq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PiaCaptureEnum {
    pub name: String,
    pub values: Vec<PiaCaptureEnumValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub metadata: Option<SidecarMetadata>,
}

#[derive(Debug, Clone, PartialEq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PiaCaptureEnumValue {
    pub name: String,
    pub value: i64,
}

#[derive(Debug, Clone, PartialEq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PiaCaptureInterface {
    pub name: String,
    pub kind: InterfaceKind,
    #[serde(default)]
    pub inherits: Vec<String>,
    #[serde(default)]
    pub members: Vec<PiaCaptureMember>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub metadata: Option<SidecarMetadata>,
}

#[derive(Debug, Clone, PartialEq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PiaCaptureClass {
    pub name: String,
    #[serde(default)]
    pub implements: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_interface: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub metadata: Option<SidecarMetadata>,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum PiaCaptureMemberKind {
    PropertyGet,
    PropertySet,
    Method,
    Event,
}

#[derive(Debug, Clone, PartialEq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PiaCaptureMember {
    pub name: String,
    pub member_kind: PiaCaptureMemberKind,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub return_type: Option<TypeRef>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub params: Vec<Parameter>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub disp_id: Option<i32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub notes: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub metadata: Option<SidecarMetadata>,
}

impl CodegenSummary {
    pub fn from_document(document: &OfficeIdlDocument) -> Self {
        let mut summary = Self {
            library: document.library.clone(),
            version: document.version.clone(),
            enum_count: document.enums.len(),
            interface_count: document.interfaces.len(),
            class_count: document.classes.len(),
            member_count: 0,
            implemented_member_count: 0,
            partial_member_count: 0,
            stub_member_count: 0,
            unsupported_member_count: 0,
        };

        for interface in &document.interfaces {
            for member in &interface.members {
                summary.member_count += 1;
                match member.support {
                    SupportState::Implemented => summary.implemented_member_count += 1,
                    SupportState::Partial => summary.partial_member_count += 1,
                    SupportState::Stub | SupportState::GeneratedOnly => {
                        summary.stub_member_count += 1
                    }
                    SupportState::OracleVerified => summary.implemented_member_count += 1,
                    SupportState::Unsupported => summary.unsupported_member_count += 1,
                }
            }
        }

        summary
    }

    pub fn from_json_str(input: &str) -> Result<Self, office_idl::IdlLoadError> {
        let document = OfficeIdlDocument::from_json_str(input)?;
        Ok(Self::from_document(&document))
    }
}

impl OmSourcesManifest {
    pub fn from_toml_str(input: &str) -> Result<Self, OmSourcesLoadError> {
        Ok(toml::from_str(input)?)
    }

    pub fn from_toml_path(path: impl AsRef<Path>) -> Result<Self, OmSourcesLoadError> {
        let input = fs::read_to_string(path)?;
        Self::from_toml_str(&input)
    }
}

impl SourceRegistryManifest {
    pub fn from_toml_str(input: &str) -> Result<Self, OmSourcesLoadError> {
        Ok(toml::from_str(input)?)
    }

    pub fn from_toml_path(path: impl AsRef<Path>) -> Result<Self, OmSourcesLoadError> {
        let input = fs::read_to_string(path)?;
        Self::from_toml_str(&input)
    }
}

impl TypelibIdentityCapture {
    pub fn from_json_str(input: &str) -> Result<Self, PiaCaptureLoadError> {
        Ok(serde_json::from_str(input)?)
    }

    pub fn from_json_path(path: impl AsRef<Path>) -> Result<Self, PiaCaptureLoadError> {
        let input = fs::read_to_string(path)?;
        Self::from_json_str(&input)
    }
}

impl PiaPublicSurfaceCapture {
    pub fn from_json_str(input: &str) -> Result<Self, PiaCaptureLoadError> {
        Ok(serde_json::from_str(input)?)
    }

    pub fn from_json_path(path: impl AsRef<Path>) -> Result<Self, PiaCaptureLoadError> {
        let input = fs::read_to_string(path)?;
        Self::from_json_str(&input)
    }

    pub fn to_office_idl_document(&self) -> OfficeIdlDocument {
        OfficeIdlDocument {
            library: self.library.clone(),
            version: self.version.clone(),
            metadata: Some(SidecarMetadata {
                namespace: Some(self.namespace.clone()),
                type_library_guid: None,
                iid: None,
                clsid: None,
                source_inherits: vec!["IDispatch".to_string()],
                source_default_interface: None,
                capture: None,
            }),
            enums: self
                .enums
                .iter()
                .map(|enumeration| EnumDef {
                    name: enumeration.name.clone(),
                    values: enumeration
                        .values
                        .iter()
                        .map(|value| EnumValue {
                            name: value.name.clone(),
                            value: value.value,
                        })
                        .collect(),
                    metadata: enumeration.metadata.clone(),
                })
                .collect(),
            interfaces: self
                .interfaces
                .iter()
                .map(|interface| {
                    let mut normalized_members = Vec::<Member>::new();
                    let mut property_slots = BTreeMap::<(String, Option<i32>), usize>::new();

                    for capture_member in &interface.members {
                        let mut member_metadata = capture_member
                            .metadata
                            .clone()
                            .unwrap_or_else(SidecarMetadata::default);
                        let source_disp_id = capture_member.disp_id.or_else(|| {
                            member_metadata
                                .capture
                                .as_ref()
                                .and_then(|capture| capture.origins.first())
                                .and_then(|origin| origin.disp_id)
                        });
                        let origin_kind = match capture_member.member_kind {
                            PiaCaptureMemberKind::PropertyGet => CaptureOriginKind::PropertyGet,
                            PiaCaptureMemberKind::PropertySet => CaptureOriginKind::PropertySet,
                            PiaCaptureMemberKind::Method => CaptureOriginKind::Method,
                            PiaCaptureMemberKind::Event => CaptureOriginKind::Event,
                        };
                        let origin = CaptureOrigin {
                            kind: origin_kind,
                            source_interface: Some(interface.name.clone()),
                            source_member: Some(capture_member.name.clone()),
                            disp_id: source_disp_id,
                        };
                        let type_info = match capture_member.member_kind {
                            PiaCaptureMemberKind::PropertyGet
                            | PiaCaptureMemberKind::Method
                            | PiaCaptureMemberKind::Event => capture_member.return_type.clone(),
                            PiaCaptureMemberKind::PropertySet => capture_member
                                .params
                                .last()
                                .map(|parameter| parameter.type_ref.clone()),
                        };
                        let member_capture = member_metadata
                            .capture
                            .get_or_insert_with(CaptureMetadata::default);
                        if !member_capture.origins.contains(&origin) {
                            member_capture.origins.push(origin.clone());
                        }
                        if member_capture.type_info.is_none() {
                            member_capture.type_info = type_info.clone();
                        }
                        if member_metadata.source_inherits.is_empty() {
                            member_metadata.source_inherits = interface.inherits.clone();
                        }

                        match capture_member.member_kind {
                            PiaCaptureMemberKind::Method => normalized_members.push(Member {
                                name: capture_member.name.clone(),
                                member_kind: MemberKind::Method,
                                access: AccessMode::Read,
                                return_type: capture_member.return_type.clone(),
                                params: capture_member.params.clone(),
                                disp_id: source_disp_id,
                                support: SupportState::Stub,
                                notes: capture_member.notes.clone(),
                                metadata: Some(member_metadata.clone()),
                            }),
                            PiaCaptureMemberKind::Event => normalized_members.push(Member {
                                name: capture_member.name.clone(),
                                member_kind: MemberKind::Event,
                                access: AccessMode::Read,
                                return_type: capture_member.return_type.clone(),
                                params: capture_member.params.clone(),
                                disp_id: source_disp_id,
                                support: SupportState::Stub,
                                notes: capture_member.notes.clone(),
                                metadata: Some(member_metadata.clone()),
                            }),
                            PiaCaptureMemberKind::PropertyGet
                            | PiaCaptureMemberKind::PropertySet => {
                                let key = (capture_member.name.clone(), source_disp_id);
                                let property_type = if matches!(
                                    capture_member.member_kind,
                                    PiaCaptureMemberKind::PropertyGet
                                ) {
                                    capture_member.return_type.clone()
                                } else {
                                    capture_member
                                        .params
                                        .last()
                                        .map(|parameter| parameter.type_ref.clone())
                                };
                                let property_params = if matches!(
                                    capture_member.member_kind,
                                    PiaCaptureMemberKind::PropertyGet
                                ) {
                                    capture_member.params.clone()
                                } else {
                                    capture_member
                                        .params
                                        .split_last()
                                        .map(|(_, params)| params.to_vec())
                                        .unwrap_or_default()
                                };

                                if let Some(index) = property_slots.get(&key).copied() {
                                    let property = &mut normalized_members[index];
                                    property.access =
                                        match (property.access, &capture_member.member_kind) {
                                            (
                                                AccessMode::Read,
                                                PiaCaptureMemberKind::PropertySet,
                                            ) => AccessMode::Readwrite,
                                            (
                                                AccessMode::Write,
                                                PiaCaptureMemberKind::PropertyGet,
                                            ) => AccessMode::Readwrite,
                                            (existing, _) => existing,
                                        };
                                    if property.return_type.is_none() {
                                        property.return_type = property_type;
                                    }
                                    if property.params.is_empty() {
                                        property.params = property_params;
                                    }
                                    if property.notes.is_none() {
                                        property.notes = capture_member.notes.clone();
                                    }
                                    if let Some(existing_metadata) = property.metadata.as_mut() {
                                        if existing_metadata.source_inherits.is_empty() {
                                            existing_metadata.source_inherits =
                                                interface.inherits.clone();
                                        }
                                        let capture = existing_metadata
                                            .capture
                                            .get_or_insert_with(CaptureMetadata::default);
                                        if !capture.origins.contains(&origin) {
                                            capture.origins.push(origin.clone());
                                        }
                                        if capture.type_info.is_none() {
                                            capture.type_info = type_info;
                                        }
                                    } else {
                                        property.metadata = Some(member_metadata.clone());
                                    }
                                    continue;
                                }

                                let access = if matches!(
                                    capture_member.member_kind,
                                    PiaCaptureMemberKind::PropertyGet
                                ) {
                                    AccessMode::Read
                                } else {
                                    AccessMode::Write
                                };
                                property_slots.insert(key, normalized_members.len());
                                normalized_members.push(Member {
                                    name: capture_member.name.clone(),
                                    member_kind: MemberKind::Property,
                                    access,
                                    return_type: property_type,
                                    params: property_params,
                                    disp_id: source_disp_id,
                                    support: SupportState::Stub,
                                    notes: capture_member.notes.clone(),
                                    metadata: Some(member_metadata),
                                });
                            }
                        }
                    }

                    InterfaceDef {
                        name: interface.name.clone(),
                        kind: interface.kind.clone(),
                        inherits: interface.inherits.clone(),
                        members: normalized_members,
                        metadata: Some({
                            let mut metadata = interface.metadata.clone().unwrap_or_default();
                            if metadata.source_inherits.is_empty() {
                                metadata.source_inherits = interface.inherits.clone();
                            }
                            metadata
                        }),
                    }
                })
                .collect(),
            classes: self
                .classes
                .iter()
                .map(|class| ClassDef {
                    name: class.name.clone(),
                    implements: class.implements.clone(),
                    default_interface: class.default_interface.clone(),
                    metadata: Some({
                        let mut metadata = class.metadata.clone().unwrap_or_default();
                        if metadata.source_default_interface.is_none() {
                            metadata.source_default_interface = class.default_interface.clone();
                        }
                        metadata
                    }),
                })
                .collect(),
        }
    }
}

impl OmCaptureBundleSummary {
    pub fn from_capture(typelib: &TypelibIdentityCapture, pia: &PiaPublicSurfaceCapture) -> Self {
        let typelib_interfaces = typelib
            .interfaces
            .iter()
            .map(|interface| interface.name.as_str())
            .collect::<std::collections::BTreeSet<_>>();
        let typelib_coclasses = typelib
            .coclasses
            .iter()
            .map(|class| class.name.as_str())
            .collect::<std::collections::BTreeSet<_>>();
        let missing_pia_interfaces = pia
            .interfaces
            .iter()
            .filter(|interface| !typelib_interfaces.contains(interface.name.as_str()))
            .map(|interface| interface.name.clone())
            .collect::<Vec<_>>();
        let missing_pia_classes = pia
            .classes
            .iter()
            .filter(|class| !typelib_coclasses.contains(class.name.as_str()))
            .map(|class| class.name.clone())
            .collect::<Vec<_>>();

        Self {
            library: pia.library.clone(),
            version: pia.version.clone(),
            type_library_guid: typelib.type_library_guid.clone(),
            interface_iid_count: typelib.interfaces.len(),
            coclass_clsid_count: typelib.coclasses.len(),
            missing_pia_interfaces,
            missing_pia_classes,
        }
    }
}

impl CaptureBundlePaths {
    pub fn from_bundle_root(bundle_root: impl AsRef<Path>) -> Self {
        let bundle_root = bundle_root.as_ref();
        Self {
            bundle_root_path: bundle_root.to_path_buf(),
            raw_typelib_identity_path: bundle_root.join("raw/raw_typelib_identity.json"),
            excel_pia_public_surface_path: bundle_root
                .join("snapshots/excel_pia_public_surface.json"),
            capture_manifest_path: bundle_root.join("manifest/capture_manifest.json"),
            output_checksums_path: bundle_root.join("manifest/output_checksums.json"),
        }
    }
}

impl SupportStateCounts {
    pub fn record(&mut self, support: &SupportState) {
        match support {
            SupportState::GeneratedOnly => self.generated_only += 1,
            SupportState::Stub => self.stub += 1,
            SupportState::Partial => self.partial += 1,
            SupportState::Implemented => self.implemented += 1,
            SupportState::OracleVerified => self.oracle_verified += 1,
            SupportState::Unsupported => self.unsupported += 1,
        }
    }
}

impl OmCaptureSummary {
    pub fn from_manifest(manifest: &OmSourcesManifest) -> Self {
        let unresolved_target_fields = [
            (
                "product_family",
                manifest.target_capture.product_family.as_str(),
            ),
            ("channel", manifest.target_capture.channel.as_str()),
            ("version", manifest.target_capture.version.as_str()),
            ("build", manifest.target_capture.build.as_str()),
            ("arch", manifest.target_capture.arch.as_str()),
            ("locale", manifest.target_capture.locale.as_str()),
        ]
        .into_iter()
        .filter_map(|(name, value)| (value == "unresolved").then_some(name))
        .collect::<Vec<_>>();
        let machine_readable_artifact_count =
            usize::from(manifest.artifacts.excel_type_library.machine_readable)
                + usize::from(
                    manifest
                        .artifacts
                        .excel_primary_interop_assembly
                        .machine_readable,
                );
        let mut pending_outputs = manifest
            .artifacts
            .excel_type_library
            .capture
            .required_outputs
            .clone();
        pending_outputs.extend(
            manifest
                .artifacts
                .excel_primary_interop_assembly
                .capture
                .required_outputs
                .iter()
                .cloned(),
        );
        let pending_output_count = pending_outputs.len();
        let behavior_doc_count = usize::from(manifest.docs.excel_vba_reference.is_some())
            + usize::from(manifest.docs.office_library_reference.is_some())
            + usize::from(manifest.docs.interop_namespace.is_some());

        Self {
            primary_artifact: manifest.contract.primary_artifact.clone(),
            secondary_artifact: manifest.contract.secondary_artifact.clone(),
            ready_for_windows_capture: manifest.acquisition.host_os == "windows"
                && manifest.acquisition.requires_installed_excel
                && manifest.acquisition.requires_windows_sdk
                && manifest.acquisition.requires_dotnet_framework_tooling,
            machine_readable_artifact_count,
            pending_outputs,
            pending_output_count,
            behavior_doc_count,
            unresolved_target_fields,
        }
    }
}

impl SourceRegistrySummary {
    pub fn from_manifest(manifest: &SourceRegistryManifest) -> Self {
        let mut enabled_corpus_groups = Vec::new();
        if manifest.test_corpus.official_ms.enabled {
            enabled_corpus_groups.push("official_ms".to_string());
        }
        if manifest.test_corpus.open_source.enabled {
            enabled_corpus_groups.push("open_source".to_string());
        }
        if manifest.test_corpus.synthetic {
            enabled_corpus_groups.push("synthetic".to_string());
        }
        if manifest.test_corpus.real_world {
            enabled_corpus_groups.push("real_world".to_string());
        }

        let official_ms = &manifest.test_corpus.official_ms;
        let mut official_ms_corpus_sources = Vec::new();
        if official_ms.office_scripts_samples {
            official_ms_corpus_sources.push("office_scripts_samples".to_string());
        }
        if official_ms.data_validation_examples {
            official_ms_corpus_sources.push("data_validation_examples".to_string());
        }
        if official_ms.power_bi_financial_sample {
            official_ms_corpus_sources.push("power_bi_financial_sample".to_string());
        }
        if official_ms.mos_excel_course_materials {
            official_ms_corpus_sources.push("mos_excel_course_materials".to_string());
        }
        if official_ms.mos_excel_expert_course_materials {
            official_ms_corpus_sources.push("mos_excel_expert_course_materials".to_string());
        }

        let open_source = &manifest.test_corpus.open_source;
        let mut open_source_corpus_sources = Vec::new();
        if open_source.open_xml_sdk {
            open_source_corpus_sources.push("open_xml_sdk".to_string());
        }
        if open_source.apache_poi_test_data {
            open_source_corpus_sources.push("apache_poi_test_data".to_string());
        }
        if open_source.libreoffice_sc_qa_unit_data {
            open_source_corpus_sources.push("libreoffice_sc_qa_unit_data".to_string());
        }

        let mut validation_modes = Vec::new();
        if manifest.validation.openxml_validator {
            validation_modes.push("openxml_validator".to_string());
        }
        if manifest.validation.excel_oracle {
            validation_modes.push("excel_oracle".to_string());
        }
        if manifest.validation.render_snapshot {
            validation_modes.push("render_snapshot".to_string());
        }
        if manifest.validation.fuzz {
            validation_modes.push("fuzz".to_string());
        }

        let enabled_corpus_source_count = official_ms_corpus_sources.len()
            + open_source_corpus_sources.len()
            + usize::from(manifest.test_corpus.synthetic)
            + usize::from(manifest.test_corpus.real_world);

        Self {
            project_name: manifest.project.name.clone(),
            default_profile: manifest.project.default_profile.clone(),
            default_mode: manifest.project.default_mode.clone(),
            primary_om_artifact: manifest.om_contract.primary.clone(),
            secondary_om_artifact: manifest.om_contract.secondary.clone(),
            primary_docs_source: manifest.om_contract.docs_primary.clone(),
            primary_ooxml_source: manifest.ooxml.primary.clone(),
            enabled_corpus_groups,
            official_ms_corpus_sources,
            open_source_corpus_sources,
            enabled_corpus_source_count,
            validation_modes,
            profile_count: manifest.profiles.len(),
        }
    }
}

impl DifferentialArtifactContract {
    pub fn canonical() -> Self {
        Self {
            report_artifact: DIFFERENTIAL_REPORT_ARTIFACT_NAME.to_string(),
            gate_summary_artifact: DIFFERENTIAL_GATE_SUMMARY_ARTIFACT_NAME.to_string(),
            artifact_names: vec![
                DIFFERENTIAL_REPORT_ARTIFACT_NAME.to_string(),
                DIFFERENTIAL_GATE_SUMMARY_ARTIFACT_NAME.to_string(),
            ],
        }
    }

    pub fn paths_under(&self, output_root: impl AsRef<Path>) -> DifferentialArtifactPaths {
        let output_root_path = output_root.as_ref().to_path_buf();
        DifferentialArtifactPaths {
            report_path: output_root_path.join(&self.report_artifact),
            gate_summary_path: output_root_path.join(&self.gate_summary_artifact),
            output_root_path,
        }
    }
}

impl DifferentialArtifactPaths {
    pub fn canonical(output_root: impl AsRef<Path>) -> Self {
        DifferentialArtifactContract::canonical().paths_under(output_root)
    }
}

impl DifferentialArtifactBundle {
    pub fn passed(&self) -> bool {
        self.gate_summary.passed
    }

    pub fn blocking_case_count(&self) -> usize {
        self.gate_summary.blocking_case_count
    }

    pub fn blocking_cases(&self) -> &[String] {
        &self.gate_summary.blocking_cases
    }
}

impl DifferentialStatusCounts {
    pub fn record(&mut self, status: DifferentialCaseStatus) {
        match status {
            DifferentialCaseStatus::Passed => self.passed += 1,
            DifferentialCaseStatus::Failed => self.failed += 1,
            DifferentialCaseStatus::MissingOracle => self.missing_oracle += 1,
            DifferentialCaseStatus::MissingRuntime => self.missing_runtime += 1,
            DifferentialCaseStatus::Unsupported => self.unsupported += 1,
            DifferentialCaseStatus::Skipped => self.skipped += 1,
        }
    }
}

impl DifferentialReport {
    pub fn from_cases(
        library: impl Into<String>,
        version: impl Into<String>,
        profile: impl Into<String>,
        cases: Vec<DifferentialCaseResult>,
    ) -> Self {
        let mut status_counts = DifferentialStatusCounts::default();
        for case in &cases {
            status_counts.record(case.status);
        }
        Self {
            library: library.into(),
            version: version.into(),
            profile: profile.into(),
            context: None,
            case_count: cases.len(),
            status_counts,
            cases,
        }
    }

    pub fn with_context(mut self, context: DifferentialReportContext) -> Self {
        self.context = Some(context);
        self
    }

    pub fn validate(&self) -> Result<(), DifferentialReportLoadError> {
        if self.case_count != self.cases.len() {
            return Err(DifferentialReportLoadError::Contract {
                message: format!(
                    "differential report caseCount {} did not match cases length {}",
                    self.case_count,
                    self.cases.len()
                ),
            });
        }
        let mut seen_case_names = BTreeSet::new();
        for case in &self.cases {
            if !seen_case_names.insert(case.name.clone()) {
                return Err(DifferentialReportLoadError::Contract {
                    message: format!("differential report duplicate case name {}", case.name),
                });
            }
            for (artifact_key, artifact_path) in &case.artifacts {
                if artifact_key.is_empty() {
                    return Err(DifferentialReportLoadError::Contract {
                        message: format!(
                            "differential report case {} has empty artifact key",
                            case.name
                        ),
                    });
                }
                if artifact_path.is_empty() {
                    return Err(DifferentialReportLoadError::Contract {
                        message: format!(
                            "differential report case {} artifact {} has empty path",
                            case.name, artifact_key
                        ),
                    });
                }
                let path = Path::new(artifact_path);
                if path.is_absolute()
                    || path.components().any(|component| {
                        matches!(
                            component,
                            Component::ParentDir | Component::RootDir | Component::Prefix(_)
                        )
                    })
                {
                    return Err(DifferentialReportLoadError::Contract {
                        message: format!(
                            "differential report case {} artifact {} path {} must be relative and stay within the output root",
                            case.name, artifact_key, artifact_path
                        ),
                    });
                }
            }
        }
        let mut expected_counts = DifferentialStatusCounts::default();
        for case in &self.cases {
            expected_counts.record(case.status);
        }
        if self.status_counts != expected_counts {
            return Err(DifferentialReportLoadError::Contract {
                message: format!(
                    "differential report statusCounts {:?} did not match cases {:?}",
                    self.status_counts, expected_counts
                ),
            });
        }
        Ok(())
    }

    pub fn validate_source_context(
        &self,
        summary: &SourceRegistrySummary,
    ) -> Result<(), DifferentialReportLoadError> {
        let Some(context) = self.context.as_ref() else {
            return Err(DifferentialReportLoadError::Contract {
                message: "differential report missing source registry context".to_string(),
            });
        };
        let expected_context = DifferentialReportContext::from_source_registry_summary(summary);
        if context != &expected_context {
            return Err(DifferentialReportLoadError::Contract {
                message: format!(
                    "differential report source context {:?} did not match registry {:?}",
                    context, expected_context
                ),
            });
        }
        if self.profile != summary.default_profile {
            return Err(DifferentialReportLoadError::Contract {
                message: format!(
                    "differential report profile {} did not match registry default profile {}",
                    self.profile, summary.default_profile
                ),
            });
        }
        Ok(())
    }

    pub fn from_json_str(input: &str) -> Result<Self, DifferentialReportLoadError> {
        let report = serde_json::from_str::<Self>(input)?;
        report.validate()?;
        Ok(report)
    }

    pub fn from_json_path(path: impl AsRef<Path>) -> Result<Self, DifferentialReportLoadError> {
        let input = fs::read_to_string(path)?;
        Self::from_json_str(&input)
    }

    pub fn write_json_path(
        &self,
        path: impl AsRef<Path>,
    ) -> Result<(), DifferentialReportLoadError> {
        self.validate()?;
        let payload = serde_json::to_vec_pretty(self)?;
        fs::write(path, payload)?;
        Ok(())
    }

    pub fn gate_summary(&self) -> DifferentialGateSummary {
        let blocking_cases = self
            .cases
            .iter()
            .filter(|case| {
                matches!(
                    case.status,
                    DifferentialCaseStatus::Failed
                        | DifferentialCaseStatus::MissingOracle
                        | DifferentialCaseStatus::MissingRuntime
                )
            })
            .map(|case| case.name.clone())
            .collect::<Vec<_>>();
        DifferentialGateSummary {
            passed: blocking_cases.is_empty(),
            blocking_case_count: blocking_cases.len(),
            blocking_cases,
            incomplete_oracle_count: self.status_counts.missing_oracle,
            missing_runtime_count: self.status_counts.missing_runtime,
            failed_case_count: self.status_counts.failed,
            unsupported_case_count: self.status_counts.unsupported,
            skipped_case_count: self.status_counts.skipped,
        }
    }
}

impl DifferentialGateSummary {
    pub fn validate(&self) -> Result<(), DifferentialReportLoadError> {
        if self.blocking_case_count != self.blocking_cases.len() {
            return Err(DifferentialReportLoadError::Contract {
                message: format!(
                    "differential gate blockingCaseCount {} did not match blockingCases length {}",
                    self.blocking_case_count,
                    self.blocking_cases.len()
                ),
            });
        }
        if self.passed != (self.blocking_case_count == 0) {
            return Err(DifferentialReportLoadError::Contract {
                message: format!(
                    "differential gate passed {} did not match blockingCaseCount {}",
                    self.passed, self.blocking_case_count
                ),
            });
        }
        Ok(())
    }

    pub fn from_json_str(input: &str) -> Result<Self, DifferentialReportLoadError> {
        let gate = serde_json::from_str::<Self>(input)?;
        gate.validate()?;
        Ok(gate)
    }

    pub fn from_json_path(path: impl AsRef<Path>) -> Result<Self, DifferentialReportLoadError> {
        let input = fs::read_to_string(path)?;
        Self::from_json_str(&input)
    }

    pub fn write_json_path(
        &self,
        path: impl AsRef<Path>,
    ) -> Result<(), DifferentialReportLoadError> {
        self.validate()?;
        let payload = serde_json::to_vec_pretty(self)?;
        fs::write(path, payload)?;
        Ok(())
    }
}

impl DifferentialReportContext {
    pub fn from_source_registry_summary(summary: &SourceRegistrySummary) -> Self {
        Self {
            project_name: summary.project_name.clone(),
            default_profile: summary.default_profile.clone(),
            default_mode: summary.default_mode.clone(),
            primary_om_artifact: summary.primary_om_artifact.clone(),
            primary_ooxml_source: summary.primary_ooxml_source.clone(),
            enabled_corpus_groups: summary.enabled_corpus_groups.clone(),
            enabled_corpus_source_count: summary.enabled_corpus_source_count,
            validation_modes: summary.validation_modes.clone(),
        }
    }
}

pub fn summarize(document: &OfficeIdlDocument) -> CodegenSummary {
    CodegenSummary::from_document(document)
}

pub fn build_focus_surface_registry(document: &OfficeIdlDocument) -> OmFocusSurfaceRegistry {
    let mut focus_surfaces = Vec::new();
    let mut missing_focus_surfaces = Vec::new();

    for focus_name in PRIORITY_OM_SURFACES {
        let Some(interface) = document
            .interfaces
            .iter()
            .find(|item| item.name == focus_name)
        else {
            missing_focus_surfaces.push(focus_name.to_string());
            continue;
        };

        let mut support_counts = SupportStateCounts::default();
        let members = interface
            .members
            .iter()
            .map(|member| {
                support_counts.record(&member.support);
                OmMemberMetadataEntry {
                    name: member.name.clone(),
                    member_kind: member.member_kind.clone(),
                    access: member.access,
                    disp_id: member.disp_id.or_else(|| {
                        member
                            .metadata
                            .as_ref()
                            .and_then(|metadata| metadata.capture.as_ref())
                            .and_then(|capture| capture.origins.first())
                            .and_then(|origin| origin.disp_id)
                    }),
                    support: member.support.clone(),
                    return_type: member.return_type.as_ref().map(to_type_metadata_entry),
                    params: member
                        .params
                        .iter()
                        .map(to_parameter_metadata_entry)
                        .collect(),
                    capture_origin_kinds: member
                        .metadata
                        .as_ref()
                        .and_then(|metadata| metadata.capture.as_ref())
                        .map(|capture| {
                            capture
                                .origins
                                .iter()
                                .map(|origin| origin.kind.clone())
                                .collect::<Vec<_>>()
                        })
                        .unwrap_or_default(),
                }
            })
            .collect::<Vec<_>>();

        let implementing_classes = document
            .classes
            .iter()
            .filter(|class| class.implements.iter().any(|name| name == &interface.name))
            .map(|class| class.name.clone())
            .collect::<Vec<_>>();
        let default_coclasses = document
            .classes
            .iter()
            .filter(|class| class.default_interface.as_deref() == Some(interface.name.as_str()))
            .map(|class| class.name.clone())
            .collect::<Vec<_>>();

        focus_surfaces.push(OmFocusSurfaceRegistryEntry {
            name: interface.name.clone(),
            kind: interface.kind.clone(),
            iid: interface
                .metadata
                .as_ref()
                .and_then(|metadata| metadata.iid.clone()),
            inherits: interface.inherits.clone(),
            implementing_classes,
            default_coclasses,
            member_count: interface.members.len(),
            support_counts,
            members,
        });
    }

    OmFocusSurfaceRegistry {
        library: document.library.clone(),
        version: document.version.clone(),
        focus_surfaces,
        missing_focus_surfaces,
    }
}

pub fn build_focus_surface_registry_from_path(
    path: impl AsRef<Path>,
) -> Result<OmFocusSurfaceRegistry, office_idl::IdlLoadError> {
    let document = OfficeIdlDocument::from_path(path)?;
    Ok(build_focus_surface_registry(&document))
}

pub fn build_focus_surface_registry_from_json(
    input: &str,
) -> Result<OmFocusSurfaceRegistry, office_idl::IdlLoadError> {
    let document = OfficeIdlDocument::from_json_str(input)?;
    Ok(build_focus_surface_registry(&document))
}

pub fn build_coverage_report(document: &OfficeIdlDocument) -> OmCoverageReport {
    let mut support_counts = SupportStateCounts::default();
    let mut member_count = 0usize;
    for interface in &document.interfaces {
        for member in &interface.members {
            member_count += 1;
            support_counts.record(&member.support);
        }
    }

    let focus_registry = build_focus_surface_registry(document);
    let focus_surfaces = focus_registry
        .focus_surfaces
        .iter()
        .map(|surface| {
            let mut entry = OmFocusSurfaceCoverageEntry {
                name: surface.name.clone(),
                member_count: surface.member_count,
                support_counts: surface.support_counts.clone(),
                generated_only_members: Vec::new(),
                stub_members: Vec::new(),
                partial_members: Vec::new(),
                implemented_members: Vec::new(),
                oracle_verified_members: Vec::new(),
                unsupported_members: Vec::new(),
            };

            for member in &surface.members {
                match member.support {
                    SupportState::GeneratedOnly => {
                        entry.generated_only_members.push(member.name.clone())
                    }
                    SupportState::Stub => entry.stub_members.push(member.name.clone()),
                    SupportState::Partial => entry.partial_members.push(member.name.clone()),
                    SupportState::Implemented => {
                        entry.implemented_members.push(member.name.clone())
                    }
                    SupportState::OracleVerified => {
                        entry.oracle_verified_members.push(member.name.clone())
                    }
                    SupportState::Unsupported => {
                        entry.unsupported_members.push(member.name.clone())
                    }
                }
            }

            entry
        })
        .collect::<Vec<_>>();

    OmCoverageReport {
        library: document.library.clone(),
        version: document.version.clone(),
        enum_count: document.enums.len(),
        interface_count: document.interfaces.len(),
        class_count: document.classes.len(),
        member_count,
        support_counts,
        focus_surfaces,
        missing_focus_surfaces: focus_registry.missing_focus_surfaces,
    }
}

pub fn build_coverage_report_from_path(
    path: impl AsRef<Path>,
) -> Result<OmCoverageReport, office_idl::IdlLoadError> {
    let document = OfficeIdlDocument::from_path(path)?;
    Ok(build_coverage_report(&document))
}

pub fn build_coverage_report_from_json(
    input: &str,
) -> Result<OmCoverageReport, office_idl::IdlLoadError> {
    let document = OfficeIdlDocument::from_json_str(input)?;
    Ok(build_coverage_report(&document))
}

pub fn summarize_json(input: &str) -> Result<CodegenSummary, office_idl::IdlLoadError> {
    CodegenSummary::from_json_str(input)
}

pub fn summarize_om_sources(manifest: &OmSourcesManifest) -> OmCaptureSummary {
    OmCaptureSummary::from_manifest(manifest)
}

pub fn summarize_om_sources_toml(input: &str) -> Result<OmCaptureSummary, OmSourcesLoadError> {
    let manifest = OmSourcesManifest::from_toml_str(input)?;
    Ok(OmCaptureSummary::from_manifest(&manifest))
}

pub fn summarize_source_registry(manifest: &SourceRegistryManifest) -> SourceRegistrySummary {
    SourceRegistrySummary::from_manifest(manifest)
}

pub fn summarize_source_registry_toml(
    input: &str,
) -> Result<SourceRegistrySummary, OmSourcesLoadError> {
    let manifest = SourceRegistryManifest::from_toml_str(input)?;
    Ok(SourceRegistrySummary::from_manifest(&manifest))
}

pub fn differential_artifact_contract() -> DifferentialArtifactContract {
    DifferentialArtifactContract::canonical()
}

pub fn differential_artifact_paths(output_root: impl AsRef<Path>) -> DifferentialArtifactPaths {
    DifferentialArtifactPaths::canonical(output_root)
}

pub fn build_differential_report(
    library: impl Into<String>,
    version: impl Into<String>,
    profile: impl Into<String>,
    cases: Vec<DifferentialCaseResult>,
) -> DifferentialReport {
    DifferentialReport::from_cases(library, version, profile, cases)
}

pub fn build_differential_report_with_source_context(
    library: impl Into<String>,
    version: impl Into<String>,
    source_summary: &SourceRegistrySummary,
    cases: Vec<DifferentialCaseResult>,
) -> DifferentialReport {
    DifferentialReport::from_cases(
        library,
        version,
        source_summary.default_profile.clone(),
        cases,
    )
    .with_context(DifferentialReportContext::from_source_registry_summary(
        source_summary,
    ))
}

pub fn load_differential_report_from_json(
    input: &str,
) -> Result<DifferentialReport, DifferentialReportLoadError> {
    DifferentialReport::from_json_str(input)
}

pub fn load_differential_report_from_path(
    path: impl AsRef<Path>,
) -> Result<DifferentialReport, DifferentialReportLoadError> {
    DifferentialReport::from_json_path(path)
}

pub fn summarize_differential_gate(report: &DifferentialReport) -> DifferentialGateSummary {
    report.gate_summary()
}

pub fn summarize_differential_gate_with_source_context(
    report: &DifferentialReport,
    source_summary: &SourceRegistrySummary,
) -> Result<DifferentialGateSummary, DifferentialReportLoadError> {
    report.validate_source_context(source_summary)?;
    Ok(report.gate_summary())
}

pub fn load_differential_gate_from_path_with_source_context(
    path: impl AsRef<Path>,
    source_summary: &SourceRegistrySummary,
) -> Result<DifferentialGateSummary, DifferentialReportLoadError> {
    let report = DifferentialReport::from_json_path(path)?;
    summarize_differential_gate_with_source_context(&report, source_summary)
}

pub fn load_differential_gate_from_json(
    input: &str,
) -> Result<DifferentialGateSummary, DifferentialReportLoadError> {
    DifferentialGateSummary::from_json_str(input)
}

pub fn load_differential_gate_from_path(
    path: impl AsRef<Path>,
) -> Result<DifferentialGateSummary, DifferentialReportLoadError> {
    DifferentialGateSummary::from_json_path(path)
}

pub fn write_differential_gate_to_path(
    gate: &DifferentialGateSummary,
    path: impl AsRef<Path>,
) -> Result<(), DifferentialReportLoadError> {
    gate.write_json_path(path)
}

pub fn write_differential_gate_from_report_path_with_source_context(
    report_path: impl AsRef<Path>,
    source_summary: &SourceRegistrySummary,
    gate_path: impl AsRef<Path>,
) -> Result<DifferentialGateSummary, DifferentialReportLoadError> {
    let gate = load_differential_gate_from_path_with_source_context(report_path, source_summary)?;
    gate.write_json_path(gate_path)?;
    Ok(gate)
}

pub fn write_differential_report_and_gate_to_output_root(
    report: &DifferentialReport,
    source_summary: &SourceRegistrySummary,
    output_root: impl AsRef<Path>,
) -> Result<(DifferentialArtifactPaths, DifferentialGateSummary), DifferentialReportLoadError> {
    let paths = differential_artifact_paths(output_root);
    report.validate()?;
    report.validate_source_context(source_summary)?;
    fs::create_dir_all(&paths.output_root_path)?;
    report.write_json_path(&paths.report_path)?;
    let gate = write_differential_gate_from_report_path_with_source_context(
        &paths.report_path,
        source_summary,
        &paths.gate_summary_path,
    )?;
    Ok((paths, gate))
}

pub fn load_differential_artifacts_from_output_root(
    output_root: impl AsRef<Path>,
    source_summary: &SourceRegistrySummary,
) -> Result<DifferentialArtifactBundle, DifferentialReportLoadError> {
    let paths = differential_artifact_paths(output_root);
    let report = load_differential_report_from_path(&paths.report_path)?;
    report.validate_source_context(source_summary)?;
    let gate_summary = load_differential_gate_from_path(&paths.gate_summary_path)?;
    let expected_gate_summary =
        summarize_differential_gate_with_source_context(&report, source_summary)?;
    if gate_summary != expected_gate_summary {
        return Err(DifferentialReportLoadError::Contract {
            message: format!(
                "differential gate summary {:?} did not match report-derived {:?}",
                gate_summary, expected_gate_summary
            ),
        });
    }
    Ok(DifferentialArtifactBundle {
        paths,
        report,
        gate_summary,
    })
}

pub fn write_differential_report_to_path(
    report: &DifferentialReport,
    path: impl AsRef<Path>,
) -> Result<(), DifferentialReportLoadError> {
    report.write_json_path(path)
}

pub fn validate_differential_report_source_context(
    report: &DifferentialReport,
    source_summary: &SourceRegistrySummary,
) -> Result<(), DifferentialReportLoadError> {
    report.validate_source_context(source_summary)
}

pub fn normalize_pia_capture_json(input: &str) -> Result<OfficeIdlDocument, PiaCaptureLoadError> {
    let capture = PiaPublicSurfaceCapture::from_json_str(input)?;
    Ok(capture.to_office_idl_document())
}

pub fn summarize_capture_bundle(
    typelib: &TypelibIdentityCapture,
    pia: &PiaPublicSurfaceCapture,
) -> Result<OmCaptureBundleSummary, OmCaptureBundleError> {
    if typelib.library != pia.library {
        return Err(OmCaptureBundleError::LibraryMismatch {
            typelib: typelib.library.clone(),
            pia: pia.library.clone(),
        });
    }
    if typelib.version != pia.version {
        return Err(OmCaptureBundleError::VersionMismatch {
            typelib: typelib.version.clone(),
            pia: pia.version.clone(),
        });
    }
    if typelib.namespace != pia.namespace {
        return Err(OmCaptureBundleError::NamespaceMismatch {
            typelib: typelib.namespace.clone(),
            pia: pia.namespace.clone(),
        });
    }
    Ok(OmCaptureBundleSummary::from_capture(typelib, pia))
}

pub fn normalize_capture_bundle(
    typelib: &TypelibIdentityCapture,
    pia: &PiaPublicSurfaceCapture,
) -> Result<(OfficeIdlDocument, OmCaptureBundleSummary), OmCaptureBundleError> {
    let summary = summarize_capture_bundle(typelib, pia)?;
    let mut document = pia.to_office_idl_document();
    let document_metadata = document
        .metadata
        .get_or_insert_with(SidecarMetadata::default);
    document_metadata.type_library_guid = Some(typelib.type_library_guid.clone());

    let interface_iids = typelib
        .interfaces
        .iter()
        .map(|interface| (interface.name.as_str(), interface.iid.as_str()))
        .collect::<BTreeMap<_, _>>();
    for interface in &mut document.interfaces {
        if let Some(iid) = interface_iids.get(interface.name.as_str()) {
            let metadata = interface
                .metadata
                .get_or_insert_with(SidecarMetadata::default);
            metadata.iid = Some((*iid).to_string());
        }
        if let Some(typelib_interface) = typelib
            .interfaces
            .iter()
            .find(|candidate| candidate.name == interface.name)
        {
            let metadata = interface
                .metadata
                .get_or_insert_with(SidecarMetadata::default);
            if metadata.source_inherits.is_empty() {
                metadata.source_inherits = typelib_interface.inherits.clone();
            }
        }
    }

    let class_clsids = typelib
        .coclasses
        .iter()
        .map(|class| (class.name.as_str(), class.clsid.as_str()))
        .collect::<BTreeMap<_, _>>();
    for class in &mut document.classes {
        if let Some(clsid) = class_clsids.get(class.name.as_str()) {
            let metadata = class.metadata.get_or_insert_with(SidecarMetadata::default);
            metadata.clsid = Some((*clsid).to_string());
        }
    }

    Ok((document, summary))
}

pub fn load_capture_bundle(
    bundle_root: impl AsRef<Path>,
) -> Result<
    (
        TypelibIdentityCapture,
        PiaPublicSurfaceCapture,
        CaptureBundlePaths,
    ),
    CanonicalOmGenerationError,
> {
    let bundle_paths = CaptureBundlePaths::from_bundle_root(bundle_root);
    validate_capture_bundle_contract(&bundle_paths)?;
    let typelib = load_typelib_identity_capture(&bundle_paths.raw_typelib_identity_path)?;
    let pia = load_pia_public_surface_capture(&bundle_paths.excel_pia_public_surface_path)?;
    Ok((typelib, pia, bundle_paths))
}

pub fn normalize_capture_bundle_from_dir(
    bundle_root: impl AsRef<Path>,
) -> Result<
    (
        OfficeIdlDocument,
        OmCaptureBundleSummary,
        CaptureBundlePaths,
    ),
    CanonicalOmGenerationError,
> {
    let (typelib, pia, bundle_paths) = load_capture_bundle(bundle_root)?;
    let (document, summary) =
        normalize_capture_bundle(&typelib, &pia).map_err(CanonicalOmGenerationError::Normalize)?;
    Ok((document, summary, bundle_paths))
}

pub fn generate_canonical_office_idl_from_dir(
    bundle_root: impl AsRef<Path>,
    output_path: impl AsRef<Path>,
) -> Result<CanonicalOmGenerationResult, CanonicalOmGenerationError> {
    let (document, summary, bundle_paths) = normalize_capture_bundle_from_dir(bundle_root)?;
    let output_path = output_path.as_ref();
    if let Some(parent) = output_path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
    {
        fs::create_dir_all(parent).map_err(|source| CanonicalOmGenerationError::Io {
            action: "create canonical output directory",
            path: parent.to_path_buf(),
            source,
        })?;
    }
    let payload = document
        .to_json_pretty()
        .map_err(|source| CanonicalOmGenerationError::Json {
            path: output_path.to_path_buf(),
            source,
        })?;
    fs::write(output_path, payload).map_err(|source| CanonicalOmGenerationError::Io {
        action: "write canonical office-idl json",
        path: output_path.to_path_buf(),
        source,
    })?;
    Ok(CanonicalOmGenerationResult {
        output_path: output_path.to_path_buf(),
        bundle_paths,
        summary,
    })
}

fn to_type_metadata_entry(type_ref: &TypeRef) -> OmTypeMetadataEntry {
    OmTypeMetadataEntry {
        kind: type_ref.kind.clone(),
        name: type_ref.name.clone(),
        alias_of: type_ref.alias_of.clone(),
        nullable: type_ref.nullable,
    }
}

fn to_parameter_metadata_entry(parameter: &Parameter) -> OmParameterMetadataEntry {
    OmParameterMetadataEntry {
        name: parameter.name.clone(),
        type_ref: to_type_metadata_entry(&parameter.type_ref),
        optional: parameter.optional,
        by_ref: parameter.by_ref,
    }
}

fn load_typelib_identity_capture(
    path: &Path,
) -> Result<TypelibIdentityCapture, CanonicalOmGenerationError> {
    let input = fs::read_to_string(path).map_err(|source| CanonicalOmGenerationError::Io {
        action: "read typelib identity capture",
        path: path.to_path_buf(),
        source,
    })?;
    serde_json::from_str(&input).map_err(|source| CanonicalOmGenerationError::Json {
        path: path.to_path_buf(),
        source,
    })
}

fn validate_capture_bundle_contract(
    bundle_paths: &CaptureBundlePaths,
) -> Result<(), CanonicalOmGenerationError> {
    if !bundle_paths.capture_manifest_path.exists() {
        return Ok(());
    }

    let manifest_input =
        fs::read_to_string(&bundle_paths.capture_manifest_path).map_err(|source| {
            CanonicalOmGenerationError::Io {
                action: "read capture manifest",
                path: bundle_paths.capture_manifest_path.clone(),
                source,
            }
        })?;
    let manifest: serde_json::Value = serde_json::from_str(&manifest_input).map_err(|source| {
        CanonicalOmGenerationError::Json {
            path: bundle_paths.capture_manifest_path.clone(),
            source,
        }
    })?;
    let expected_output_names = collect_string_set_without_duplicates(
        manifest.get("expectedCaptureOutputs").ok_or_else(|| {
            CanonicalOmGenerationError::CaptureBundleContract {
                message: "capture_manifest.json missing expectedCaptureOutputs array".to_string(),
            }
        })?,
        "capture_manifest.json expectedCaptureOutputs",
    )?;
    let required_output_names = [
        "raw_typelib_identity.json",
        "excel_typelib_snapshot.idl",
        "excel_typelib_snapshot.odl",
        "excel_pia_identity.json",
        "excel_pia_public_surface.json",
    ]
    .into_iter()
    .map(str::to_string)
    .collect::<BTreeSet<_>>();
    if expected_output_names != required_output_names {
        return Err(CanonicalOmGenerationError::CaptureBundleContract {
            message: format!(
                "capture_manifest.json expectedCaptureOutputs {:?} did not match required {:?}",
                expected_output_names, required_output_names
            ),
        });
    }

    let writable_outputs = manifest
        .get("writableOutputs")
        .and_then(|value| value.as_object())
        .ok_or_else(|| CanonicalOmGenerationError::CaptureBundleContract {
            message: "capture_manifest.json missing writableOutputs object".to_string(),
        })?;
    let writable_output_logical_names = writable_outputs.keys().cloned().collect::<BTreeSet<_>>();
    let required_writable_output_logical_names = required_writable_output_paths()
        .keys()
        .map(|key| (*key).to_string())
        .collect::<BTreeSet<_>>();
    let allowed_writable_output_logical_names = allowed_writable_output_logical_names();
    if !allowed_writable_output_logical_names.is_superset(&writable_output_logical_names) {
        return Err(CanonicalOmGenerationError::CaptureBundleContract {
            message: format!(
                "capture_manifest.json writableOutputs keys {:?} contained entries outside allowed {:?}",
                writable_output_logical_names, allowed_writable_output_logical_names
            ),
        });
    }
    if !writable_output_logical_names.is_superset(&required_writable_output_logical_names) {
        return Err(CanonicalOmGenerationError::CaptureBundleContract {
            message: format!(
                "capture_manifest.json writableOutputs keys {:?} did not cover required payload keys {:?}",
                writable_output_logical_names, required_writable_output_logical_names
            ),
        });
    }
    let required_writable_output_paths = required_writable_output_paths();
    let writable_output_names = required_writable_output_paths
        .iter()
        .map(|(logical_name, expected_relative_path)| {
            writable_outputs
                .get(*logical_name)
                .expect("required writable output key coverage was already validated")
                .as_str()
                .ok_or_else(|| CanonicalOmGenerationError::CaptureBundleContract {
                    message: format!(
                        "capture_manifest.json writableOutputs.{logical_name} was not a path string"
                    ),
                })
                .and_then(|path| {
                    if !path_has_relative_suffix(path, expected_relative_path) {
                        return Err(CanonicalOmGenerationError::CaptureBundleContract {
                            message: format!(
                                "capture_manifest.json writableOutputs.{logical_name} path {path} did not end with required {expected_relative_path}"
                            ),
                        });
                    }
                    path.rsplit(['\\', '/'])
                        .next()
                        .map(str::to_string)
                        .ok_or_else(|| CanonicalOmGenerationError::CaptureBundleContract {
                            message: format!(
                                "capture_manifest.json writableOutputs.{logical_name} was not a path string"
                            ),
                        })
                })
        })
        .collect::<Result<BTreeSet<_>, _>>()?;
    if writable_output_names != expected_output_names {
        return Err(CanonicalOmGenerationError::CaptureBundleContract {
            message: format!(
                "capture_manifest.json writableOutputs payload names {:?} did not match expectedCaptureOutputs {:?}",
                writable_output_names, expected_output_names
            ),
        });
    }

    let checksums_input =
        fs::read_to_string(&bundle_paths.output_checksums_path).map_err(|source| {
            CanonicalOmGenerationError::Io {
                action: "read output checksums",
                path: bundle_paths.output_checksums_path.clone(),
                source,
            }
        })?;
    let checksums: BTreeMap<String, String> =
        serde_json::from_str(&checksums_input).map_err(|source| {
            CanonicalOmGenerationError::Json {
                path: bundle_paths.output_checksums_path.clone(),
                source,
            }
        })?;
    let checksum_output_names = checksums
        .keys()
        .filter_map(|relative_path| relative_path.rsplit(['\\', '/']).next())
        .map(str::to_string)
        .collect::<BTreeSet<_>>();
    let allowed_checksum_output_names = allowed_checksum_output_names();
    if !allowed_checksum_output_names.is_superset(&checksum_output_names) {
        return Err(CanonicalOmGenerationError::CaptureBundleContract {
            message: format!(
                "output_checksums.json payload names {:?} contained entries outside allowed {:?}",
                checksum_output_names, allowed_checksum_output_names
            ),
        });
    }
    if !checksum_output_names.is_superset(&expected_output_names) {
        return Err(CanonicalOmGenerationError::CaptureBundleContract {
            message: format!(
                "output_checksums.json payload names {:?} did not cover expectedCaptureOutputs {:?}",
                checksum_output_names, expected_output_names
            ),
        });
    }
    let required_output_relative_paths = required_capture_payload_paths();
    for expected_output_name in &expected_output_names {
        let matching_relative_paths = checksums
            .keys()
            .filter(|relative_path| {
                relative_path
                    .rsplit(['\\', '/'])
                    .next()
                    .is_some_and(|file_name| file_name == expected_output_name)
            })
            .collect::<Vec<_>>();
        if matching_relative_paths.len() != 1 {
            return Err(CanonicalOmGenerationError::CaptureBundleContract {
                message: format!(
                    "output_checksums.json expected payload {expected_output_name} matched {} checksum paths {:?}, expected exactly one",
                    matching_relative_paths.len(),
                    matching_relative_paths
                ),
            });
        }
        let relative_path = matching_relative_paths[0];
        let artifact_path = bundle_relative_path(&bundle_paths.bundle_root_path, relative_path)
            .ok_or_else(|| CanonicalOmGenerationError::CaptureBundleContract {
                message: format!(
                    "output_checksums.json path {relative_path} was not bundle-relative"
                ),
            })?;
        let expected_relative_path = required_output_relative_paths
            .get(expected_output_name.as_str())
            .expect("required output names and paths are in sync");
        if relative_path != expected_relative_path {
            return Err(CanonicalOmGenerationError::CaptureBundleContract {
                message: format!(
                    "output_checksums.json path {relative_path} for {expected_output_name} did not match required {expected_relative_path}"
                ),
            });
        }
        if !artifact_path.exists() {
            return Err(CanonicalOmGenerationError::CaptureBundleContract {
                message: format!(
                    "output_checksums.json path {relative_path} for {expected_output_name} did not exist at {}",
                    artifact_path.display()
                ),
            });
        }
        let artifact =
            fs::read(&artifact_path).map_err(|source| CanonicalOmGenerationError::Io {
                action: "read checksum-listed artifact",
                path: artifact_path.clone(),
                source,
            })?;
        let actual_checksum = sha256_hex(&artifact);
        let expected_checksum = checksums.get(relative_path).ok_or_else(|| {
            CanonicalOmGenerationError::CaptureBundleContract {
                message: format!(
                    "output_checksums.json path {relative_path} disappeared during validation"
                ),
            }
        })?;
        if &actual_checksum != expected_checksum {
            return Err(CanonicalOmGenerationError::CaptureBundleContract {
                message: format!(
                    "output_checksums.json checksum for {relative_path} was {expected_checksum}, actual {actual_checksum}"
                ),
            });
        }
    }

    if let Some(receipt) = manifest
        .get("executionReceipt")
        .and_then(|value| value.as_object())
    {
        if let Some(receipt_expected_outputs) = receipt
            .get("expectedCaptureOutputs")
            .filter(|value| !value.is_null())
        {
            let receipt_expected_output_names = collect_string_set_without_duplicates(
                receipt_expected_outputs,
                "capture_manifest.json executionReceipt.expectedCaptureOutputs",
            )?;
            if receipt_expected_output_names != expected_output_names {
                return Err(CanonicalOmGenerationError::CaptureBundleContract {
                    message: format!(
                        "capture_manifest.json executionReceipt expected outputs {:?} did not match manifest {:?}",
                        receipt_expected_output_names, expected_output_names
                    ),
                });
            }

            let command_results = receipt
                .get("commandResults")
                .and_then(|value| value.as_array())
                .ok_or_else(|| CanonicalOmGenerationError::CaptureBundleContract {
                    message: "capture_manifest.json executionReceipt missing commandResults array"
                        .to_string(),
                })?;
            let command_names = validate_receipt_results(
                "commandResults",
                command_results,
                &["tlbimp_fallback", "powershell_capture_reflection"],
            )?;
            if !command_names.contains("powershell_capture_reflection") {
                return Err(CanonicalOmGenerationError::CaptureBundleContract {
                    message:
                        "capture_manifest.json executionReceipt commandResults missing powershell_capture_reflection"
                            .to_string(),
                });
            }

            let manual_step_results = receipt
                .get("manualStepResults")
                .and_then(|value| value.as_array())
                .ok_or_else(|| CanonicalOmGenerationError::CaptureBundleContract {
                    message:
                        "capture_manifest.json executionReceipt missing manualStepResults array"
                            .to_string(),
                })?;
            let manual_step_names = validate_receipt_results(
                "manualStepResults",
                manual_step_results,
                &["oleview_snapshot_export"],
            )?;
            let expected_manual_step_names = ["oleview_snapshot_export".to_string()]
                .into_iter()
                .collect::<BTreeSet<_>>();
            if manual_step_names != expected_manual_step_names {
                return Err(CanonicalOmGenerationError::CaptureBundleContract {
                    message: format!(
                        "capture_manifest.json executionReceipt manualStepResults {:?} did not match expected oleview snapshot export",
                        manual_step_names
                    ),
                });
            }
        }
    }

    Ok(())
}

fn collect_string_set_without_duplicates(
    value: &serde_json::Value,
    label: &'static str,
) -> Result<BTreeSet<String>, CanonicalOmGenerationError> {
    let values =
        value
            .as_array()
            .ok_or_else(|| CanonicalOmGenerationError::CaptureBundleContract {
                message: format!("{label} was not an array"),
            })?;
    let mut names = BTreeSet::new();
    for value in values {
        let name =
            value
                .as_str()
                .ok_or_else(|| CanonicalOmGenerationError::CaptureBundleContract {
                    message: format!("{label} contains non-string entry"),
                })?;
        if !names.insert(name.to_string()) {
            return Err(CanonicalOmGenerationError::CaptureBundleContract {
                message: format!("{label} contained duplicate entry {name}"),
            });
        }
    }
    Ok(names)
}

fn validate_receipt_results(
    section: &'static str,
    results: &[serde_json::Value],
    known_names: &[&str],
) -> Result<BTreeSet<String>, CanonicalOmGenerationError> {
    let known_names = known_names.iter().copied().collect::<BTreeSet<_>>();
    let mut result_names = BTreeSet::new();
    for result in results {
        let result = result.as_object().ok_or_else(|| {
            CanonicalOmGenerationError::CaptureBundleContract {
                message: format!(
                    "capture_manifest.json executionReceipt.{section} contains non-object entry"
                ),
            }
        })?;
        let name = result
            .get("name")
            .and_then(|value| value.as_str())
            .ok_or_else(|| CanonicalOmGenerationError::CaptureBundleContract {
                message: format!(
                    "capture_manifest.json executionReceipt.{section} entry missing string name"
                ),
            })?;
        if !known_names.contains(name) {
            return Err(CanonicalOmGenerationError::CaptureBundleContract {
                message: format!(
                    "capture_manifest.json executionReceipt.{section} contained unknown result {name}"
                ),
            });
        }
        let status = result
            .get("status")
            .and_then(|value| value.as_str())
            .ok_or_else(|| CanonicalOmGenerationError::CaptureBundleContract {
                message: format!(
                    "capture_manifest.json executionReceipt.{section}.{name} missing string status"
                ),
            })?;
        if status != "completed" {
            return Err(CanonicalOmGenerationError::CaptureBundleContract {
                message: format!(
                    "capture_manifest.json executionReceipt.{section}.{name} was {status}, not completed"
                ),
            });
        }
        if !result_names.insert(name.to_string()) {
            return Err(CanonicalOmGenerationError::CaptureBundleContract {
                message: format!(
                    "capture_manifest.json executionReceipt.{section} contained duplicate result {name}"
                ),
            });
        }
    }
    Ok(result_names)
}

fn required_capture_payload_paths() -> BTreeMap<&'static str, &'static str> {
    [
        ("raw_typelib_identity.json", "raw/raw_typelib_identity.json"),
        (
            "excel_typelib_snapshot.idl",
            "snapshots/excel_typelib_snapshot.idl",
        ),
        (
            "excel_typelib_snapshot.odl",
            "snapshots/excel_typelib_snapshot.odl",
        ),
        ("excel_pia_identity.json", "raw/excel_pia_identity.json"),
        (
            "excel_pia_public_surface.json",
            "snapshots/excel_pia_public_surface.json",
        ),
    ]
    .into_iter()
    .collect()
}

fn required_writable_output_paths() -> BTreeMap<&'static str, &'static str> {
    [
        ("raw_typelib_identity", "raw/raw_typelib_identity.json"),
        (
            "excel_typelib_snapshot_idl",
            "snapshots/excel_typelib_snapshot.idl",
        ),
        (
            "excel_typelib_snapshot_odl",
            "snapshots/excel_typelib_snapshot.odl",
        ),
        ("excel_pia_identity", "raw/excel_pia_identity.json"),
        (
            "excel_pia_public_surface",
            "snapshots/excel_pia_public_surface.json",
        ),
    ]
    .into_iter()
    .collect()
}

fn allowed_writable_output_logical_names() -> BTreeSet<String> {
    [
        "raw_typelib_identity",
        "excel_typelib_snapshot_idl",
        "excel_typelib_snapshot_odl",
        "excel_pia_identity",
        "excel_pia_public_surface",
        "capture_log",
        "capture_manifest",
        "output_checksums",
    ]
    .into_iter()
    .map(str::to_string)
    .collect()
}

fn allowed_checksum_output_names() -> BTreeSet<String> {
    [
        "raw_typelib_identity.json",
        "excel_typelib_snapshot.idl",
        "excel_typelib_snapshot.odl",
        "excel_pia_identity.json",
        "excel_pia_public_surface.json",
        "capture.log",
    ]
    .into_iter()
    .map(str::to_string)
    .collect()
}

fn path_has_relative_suffix(path: &str, expected_relative_path: &str) -> bool {
    let path = path.replace('\\', "/");
    let expected_relative_path = expected_relative_path.replace('\\', "/");
    path == expected_relative_path || path.ends_with(&format!("/{expected_relative_path}"))
}

fn bundle_relative_path(bundle_root: &Path, relative_path: &str) -> Option<PathBuf> {
    if relative_path.starts_with(['/', '\\']) {
        return None;
    }
    let mut path = bundle_root.to_path_buf();
    for component in relative_path.split(['\\', '/']) {
        if component.is_empty() || component == "." {
            continue;
        }
        if component == ".." || component.contains(':') {
            return None;
        }
        path.push(component);
    }
    Some(path)
}

fn sha256_hex(contents: &[u8]) -> String {
    let digest = Sha256::digest(contents);
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        output.push_str(&format!("{byte:02x}"));
    }
    output
}

fn load_pia_public_surface_capture(
    path: &Path,
) -> Result<PiaPublicSurfaceCapture, CanonicalOmGenerationError> {
    let input = fs::read_to_string(path).map_err(|source| CanonicalOmGenerationError::Io {
        action: "read pia public surface capture",
        path: path.to_path_buf(),
        source,
    })?;
    serde_json::from_str(&input).map_err(|source| CanonicalOmGenerationError::Json {
        path: path.to_path_buf(),
        source,
    })
}

fn is_false(value: &bool) -> bool {
    !*value
}
