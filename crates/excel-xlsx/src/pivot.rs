use std::collections::{BTreeMap, VecDeque};

use office_common::{OmError, OmErrorCode, OmResult};
use office_opc::{CompressionMethod, OpcPackage};

use super::relationships::{RelationshipEntry, parse_relationship_entries_with_options};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotPackagePartKind {
    PivotTableDefinition,
    PivotCacheDefinition,
    PivotCacheRecords,
    Slicer,
    SlicerCache,
    Timeline,
    TimelineCache,
    OpaqueRelated,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PreservedPivotPart {
    pub kind: PivotPackagePartKind,
    pub content_type: Option<String>,
    pub compression: CompressionMethod,
    pub source_bytes: Vec<u8>,
    pub relationships_part_uri: Option<String>,
    pub relationships_part_source_bytes: Option<Vec<u8>>,
    pub relationships_part_compression: Option<CompressionMethod>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotPackageRelationship {
    pub source_part_uri: Option<String>,
    pub relationships_part_uri: String,
    pub relationship_id: String,
    pub relationship_type: String,
    pub target: String,
    pub target_mode: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PivotPackageInventory {
    pub parts: BTreeMap<String, PreservedPivotPart>,
    pub relationships: Vec<PivotPackageRelationship>,
}

impl PivotPackageInventory {
    pub fn is_empty(&self) -> bool {
        self.parts.is_empty()
    }
}

#[derive(Debug)]
struct PackageRelationshipOwner {
    source_part_uri: Option<String>,
    relationships_part_uri: String,
    source_bytes: Vec<u8>,
    compression: CompressionMethod,
    relationships: Vec<RelationshipEntry>,
}

pub(crate) fn collect_pivot_package_inventory(
    package: &OpcPackage,
) -> OmResult<PivotPackageInventory> {
    let mut content_type_overrides = BTreeMap::new();
    let mut default_content_types = BTreeMap::new();
    if let Some(content_types_part) = package.part("[Content_Types].xml") {
        let summary = super::parse_content_types_part_summary(content_types_part.bytes.as_slice())?;
        for attributes in summary.override_attr_maps {
            let Some(part_name) = attributes.get("PartName") else {
                continue;
            };
            let Some(content_type) = attributes.get("ContentType") else {
                continue;
            };
            content_type_overrides.insert(
                part_name
                    .strip_prefix('/')
                    .unwrap_or(part_name)
                    .to_ascii_lowercase(),
                content_type.clone(),
            );
        }
        for attributes in summary.default_attr_maps {
            let Some(extension) = attributes.get("Extension") else {
                continue;
            };
            let Some(content_type) = attributes.get("ContentType") else {
                continue;
            };
            default_content_types.insert(extension.to_ascii_lowercase(), content_type.clone());
        }
    }
    let mut current_content_types = BTreeMap::new();
    for part in package.parts() {
        let content_type = content_type_overrides
            .get(part.name.to_ascii_lowercase().as_str())
            .cloned()
            .or_else(|| {
                part.name
                    .rsplit_once('.')
                    .and_then(|(_, extension)| {
                        default_content_types.get(&extension.to_ascii_lowercase())
                    })
                    .cloned()
            })
            .or_else(|| part.content_type.clone());
        current_content_types.insert(part.name.clone(), content_type);
    }

    let mut relationship_owners = Vec::new();
    if let Some(relationships_part) = package.part("_rels/.rels") {
        relationship_owners.push(PackageRelationshipOwner {
            source_part_uri: None,
            relationships_part_uri: "_rels/.rels".to_string(),
            source_bytes: relationships_part.bytes.clone(),
            compression: relationships_part.compression,
            relationships: parse_relationship_entries_with_options(
                relationships_part.bytes.as_slice(),
                &[],
                true,
            )?,
        });
    }
    for part in package.parts() {
        if part.name == "[Content_Types].xml"
            || part.name == "_rels/.rels"
            || part.name.ends_with(".rels")
        {
            continue;
        }
        let relationships_part_uri = match part.name.rsplit_once('/') {
            Some((parent, file_name)) => format!("{parent}/_rels/{file_name}.rels"),
            None => format!("_rels/{}.rels", part.name),
        };
        let Some(relationships_part) = package.part(relationships_part_uri.as_str()) else {
            continue;
        };
        let parent_segments = part
            .name
            .rsplit_once('/')
            .map(|(parent, _)| parent.split('/').collect::<Vec<_>>())
            .unwrap_or_default();
        relationship_owners.push(PackageRelationshipOwner {
            source_part_uri: Some(part.name.clone()),
            relationships_part_uri,
            source_bytes: relationships_part.bytes.clone(),
            compression: relationships_part.compression,
            relationships: parse_relationship_entries_with_options(
                relationships_part.bytes.as_slice(),
                parent_segments.as_slice(),
                true,
            )?,
        });
    }

    let content_type_kind = |content_type: &str| {
        let normalized = content_type.to_ascii_lowercase();
        match normalized.as_str() {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivottable+xml" => {
                Some(PivotPackagePartKind::PivotTableDefinition)
            }
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotcachedefinition+xml" => {
                Some(PivotPackagePartKind::PivotCacheDefinition)
            }
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotcacherecords+xml" => {
                Some(PivotPackagePartKind::PivotCacheRecords)
            }
            "application/vnd.ms-excel.slicer+xml" => Some(PivotPackagePartKind::Slicer),
            "application/vnd.ms-excel.slicercache+xml" => Some(PivotPackagePartKind::SlicerCache),
            "application/vnd.ms-excel.timeline+xml" => Some(PivotPackagePartKind::Timeline),
            "application/vnd.ms-excel.timelinecache+xml" => {
                Some(PivotPackagePartKind::TimelineCache)
            }
            _ => None,
        }
    };
    let relationship_type_kind = |relationship_type: &str| match relationship_type
        .rsplit('/')
        .next()
        .unwrap_or(relationship_type)
        .to_ascii_lowercase()
        .as_str()
    {
        "pivottable" => Some(PivotPackagePartKind::PivotTableDefinition),
        "pivotcachedefinition" => Some(PivotPackagePartKind::PivotCacheDefinition),
        "pivotcacherecords" => Some(PivotPackagePartKind::PivotCacheRecords),
        "slicer" => Some(PivotPackagePartKind::Slicer),
        "slicercache" => Some(PivotPackagePartKind::SlicerCache),
        "timeline" => Some(PivotPackagePartKind::Timeline),
        "timelinecache" => Some(PivotPackagePartKind::TimelineCache),
        _ => None,
    };

    let mut part_kinds = BTreeMap::new();
    for part in package.parts() {
        let Some(kind) = current_content_types
            .get(part.name.as_str())
            .and_then(Option::as_deref)
            .and_then(&content_type_kind)
        else {
            continue;
        };
        part_kinds.insert(part.name.clone(), kind);
    }
    for owner in &relationship_owners {
        for relationship in &owner.relationships {
            let Some(kind) = relationship_type_kind(relationship.relationship_type.as_str()) else {
                continue;
            };
            if relationship
                .target_mode
                .as_deref()
                .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
            {
                continue;
            }
            if package.part(relationship.target.as_str()).is_none() {
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    format!(
                        "pivot relationship target is missing: {}",
                        relationship.target
                    ),
                ));
            }
            if let Some(existing) = part_kinds.insert(relationship.target.clone(), kind)
                && existing != kind
            {
                return Err(OmError::new(
                    OmErrorCode::Parse,
                    format!("conflicting pivot part kinds for {}", relationship.target),
                ));
            }
        }
    }

    let mut pending = part_kinds.keys().cloned().collect::<VecDeque<_>>();
    while let Some(source_part_uri) = pending.pop_front() {
        for owner in relationship_owners
            .iter()
            .filter(|owner| owner.source_part_uri.as_deref() == Some(source_part_uri.as_str()))
        {
            for relationship in &owner.relationships {
                if relationship
                    .target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
                {
                    continue;
                }
                if package.part(relationship.target.as_str()).is_none() {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        format!(
                            "pivot package graph target is missing: {}",
                            relationship.target
                        ),
                    ));
                }
                if !part_kinds.contains_key(relationship.target.as_str()) {
                    part_kinds.insert(
                        relationship.target.clone(),
                        relationship_type_kind(relationship.relationship_type.as_str())
                            .unwrap_or(PivotPackagePartKind::OpaqueRelated),
                    );
                    pending.push_back(relationship.target.clone());
                }
            }
        }
    }

    let mut parts = BTreeMap::new();
    for (part_uri, kind) in &part_kinds {
        let part = package.part(part_uri.as_str()).ok_or_else(|| {
            OmError::new(
                OmErrorCode::InvalidState,
                format!("inventoried pivot part is missing: {part_uri}"),
            )
        })?;
        let relationship_owner = relationship_owners
            .iter()
            .find(|owner| owner.source_part_uri.as_deref() == Some(part_uri.as_str()));
        parts.insert(
            part_uri.clone(),
            PreservedPivotPart {
                kind: *kind,
                content_type: current_content_types
                    .get(part_uri.as_str())
                    .cloned()
                    .flatten(),
                compression: part.compression,
                source_bytes: part.bytes.clone(),
                relationships_part_uri: relationship_owner
                    .map(|owner| owner.relationships_part_uri.clone()),
                relationships_part_source_bytes: relationship_owner
                    .map(|owner| owner.source_bytes.clone()),
                relationships_part_compression: relationship_owner.map(|owner| owner.compression),
            },
        );
    }

    let mut relationships = relationship_owners
        .iter()
        .flat_map(|owner| {
            owner.relationships.iter().filter_map(|relationship| {
                let source_is_inventoried = owner
                    .source_part_uri
                    .as_deref()
                    .is_some_and(|source| part_kinds.contains_key(source));
                let target_is_inventoried = relationship
                    .target_mode
                    .as_deref()
                    .is_none_or(|mode| !mode.eq_ignore_ascii_case("External"))
                    && part_kinds.contains_key(relationship.target.as_str());
                (source_is_inventoried || target_is_inventoried).then(|| PivotPackageRelationship {
                    source_part_uri: owner.source_part_uri.clone(),
                    relationships_part_uri: owner.relationships_part_uri.clone(),
                    relationship_id: relationship.id.clone(),
                    relationship_type: relationship.relationship_type.clone(),
                    target: relationship.target.clone(),
                    target_mode: relationship.target_mode.clone(),
                })
            })
        })
        .collect::<Vec<_>>();
    relationships.sort_by(|left, right| {
        (
            left.source_part_uri.as_deref(),
            left.relationships_part_uri.as_str(),
            left.relationship_id.as_str(),
            left.relationship_type.as_str(),
            left.target.as_str(),
            left.target_mode.as_deref(),
        )
            .cmp(&(
                right.source_part_uri.as_deref(),
                right.relationships_part_uri.as_str(),
                right.relationship_id.as_str(),
                right.relationship_type.as_str(),
                right.target.as_str(),
                right.target_mode.as_deref(),
            ))
    });

    Ok(PivotPackageInventory {
        parts,
        relationships,
    })
}

pub(crate) fn ensure_pivot_package_inventory_preserved(
    package: &OpcPackage,
    expected: &PivotPackageInventory,
) -> OmResult<()> {
    if expected.is_empty() {
        return Ok(());
    }
    let actual = collect_pivot_package_inventory(package)?;
    for (part_uri, expected_part) in &expected.parts {
        let Some(actual_part) = actual.parts.get(part_uri) else {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!("preserved pivot part is missing: {part_uri}"),
            ));
        };
        if actual_part != expected_part {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!("preserved pivot part changed: {part_uri}"),
            ));
        }
    }
    if actual.parts.len() != expected.parts.len() {
        return Err(OmError::new(
            OmErrorCode::InvalidState,
            "preserved pivot part inventory changed",
        ));
    }
    if actual.relationships != expected.relationships {
        return Err(OmError::new(
            OmErrorCode::InvalidState,
            "preserved pivot relationship inventory changed",
        ));
    }
    Ok(())
}
