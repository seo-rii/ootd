use std::fs;
use std::io::Read;
use std::path::Path;

use serde::{Deserialize, Serialize};
use thiserror::Error;

#[derive(Debug, Error)]
pub enum IdlLoadError {
    #[error("failed to read idl json: {0}")]
    Io(#[from] std::io::Error),
    #[error("failed to parse idl json: {0}")]
    Json(#[from] serde_json::Error),
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OfficeIdlDocument {
    pub library: String,
    pub version: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub metadata: Option<SidecarMetadata>,
    #[serde(default)]
    pub enums: Vec<EnumDef>,
    pub interfaces: Vec<InterfaceDef>,
    #[serde(default)]
    pub classes: Vec<ClassDef>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SidecarMetadata {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub namespace: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub type_library_guid: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub iid: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub clsid: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub source_inherits: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_default_interface: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub capture: Option<CaptureMetadata>,
}

#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CaptureMetadata {
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub origins: Vec<CaptureOrigin>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub type_info: Option<TypeRef>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CaptureOrigin {
    pub kind: CaptureOriginKind,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_interface: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_member: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub disp_id: Option<i32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum CaptureOriginKind {
    PropertyGet,
    PropertySet,
    Method,
    Event,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum SupportState {
    GeneratedOnly,
    Stub,
    Partial,
    Implemented,
    OracleVerified,
    Unsupported,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TypeRef {
    pub kind: TypeRefKind,
    pub name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub item_type: Option<Box<TypeRef>>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub nullable: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alias_of: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub metadata: Option<Box<SidecarMetadata>>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum TypeRefKind {
    Primitive,
    Enum,
    Interface,
    Class,
    Alias,
    Variant,
    Array,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Parameter {
    pub name: String,
    #[serde(rename = "type")]
    pub type_ref: TypeRef,
    #[serde(default, skip_serializing_if = "is_false")]
    pub optional: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub by_ref: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_value: Option<serde_json::Value>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum MemberKind {
    Property,
    Method,
    Event,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum AccessMode {
    Read,
    Write,
    Readwrite,
}

impl Default for AccessMode {
    fn default() -> Self {
        Self::Read
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Member {
    pub name: String,
    pub member_kind: MemberKind,
    #[serde(default, skip_serializing_if = "is_default_access")]
    pub access: AccessMode,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub return_type: Option<TypeRef>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub params: Vec<Parameter>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub disp_id: Option<i32>,
    pub support: SupportState,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub notes: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub metadata: Option<SidecarMetadata>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum InterfaceKind {
    Dispatch,
    Dual,
    Vtable,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct InterfaceDef {
    pub name: String,
    pub kind: InterfaceKind,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub inherits: Vec<String>,
    pub members: Vec<Member>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub metadata: Option<SidecarMetadata>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ClassDef {
    pub name: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub implements: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_interface: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub metadata: Option<SidecarMetadata>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct EnumValue {
    pub name: String,
    pub value: i64,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct EnumDef {
    pub name: String,
    pub values: Vec<EnumValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub metadata: Option<SidecarMetadata>,
}

impl OfficeIdlDocument {
    pub fn from_json_str(input: &str) -> Result<Self, IdlLoadError> {
        Ok(serde_json::from_str(input)?)
    }

    pub fn from_json_slice(input: &[u8]) -> Result<Self, IdlLoadError> {
        Ok(serde_json::from_slice(input)?)
    }

    pub fn from_reader(mut reader: impl Read) -> Result<Self, IdlLoadError> {
        let mut buffer = Vec::new();
        reader.read_to_end(&mut buffer)?;
        Self::from_json_slice(&buffer)
    }

    pub fn from_path(path: impl AsRef<Path>) -> Result<Self, IdlLoadError> {
        let bytes = fs::read(path)?;
        Self::from_json_slice(&bytes)
    }

    pub fn to_json_pretty(&self) -> Result<String, serde_json::Error> {
        serde_json::to_string_pretty(self)
    }
}

pub fn load_json_str(input: &str) -> Result<OfficeIdlDocument, IdlLoadError> {
    OfficeIdlDocument::from_json_str(input)
}

pub fn load_json_slice(input: &[u8]) -> Result<OfficeIdlDocument, IdlLoadError> {
    OfficeIdlDocument::from_json_slice(input)
}

fn is_false(value: &bool) -> bool {
    !*value
}

fn is_default_access(value: &AccessMode) -> bool {
    matches!(value, AccessMode::Read)
}
