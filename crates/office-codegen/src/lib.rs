use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};

use office_idl::{
    AccessMode, CaptureMetadata, CaptureOrigin, CaptureOriginKind, ClassDef, EnumDef, EnumValue,
    InterfaceDef, InterfaceKind, Member, MemberKind, OfficeIdlDocument, Parameter, SidecarMetadata,
    SupportState, TypeRef, TypeRefKind,
};
use serde::{Deserialize, Serialize};

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
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CaptureBundlePaths {
    pub raw_typelib_identity_path: PathBuf,
    pub excel_pia_public_surface_path: PathBuf,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CanonicalOmGenerationResult {
    pub output_path: PathBuf,
    pub bundle_paths: CaptureBundlePaths,
    pub summary: OmCaptureBundleSummary,
}

pub const PRIORITY_OM_SURFACES: [&str; 37] = [
    "Application",
    "Workbook",
    "Worksheet",
    "Range",
    "ChartObjects",
    "ChartObject",
    "Chart",
    "ChartArea",
    "PlotArea",
    "ChartTitle",
    "Legend",
    "DataTable",
    "ChartFormat",
    "Adjustments",
    "FillFormat",
    "GlowFormat",
    "LineFormat",
    "PictureFormat",
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

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OmCaptureSummary {
    pub primary_artifact: String,
    pub secondary_artifact: String,
    pub ready_for_windows_capture: bool,
    pub machine_readable_artifact_count: usize,
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
            raw_typelib_identity_path: bundle_root.join("raw/raw_typelib_identity.json"),
            excel_pia_public_surface_path: bundle_root
                .join("snapshots/excel_pia_public_surface.json"),
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
        let pending_output_count = manifest
            .artifacts
            .excel_type_library
            .capture
            .required_outputs
            .len()
            + manifest
                .artifacts
                .excel_primary_interop_assembly
                .capture
                .required_outputs
                .len();
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
            pending_output_count,
            behavior_doc_count,
            unresolved_target_fields,
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
