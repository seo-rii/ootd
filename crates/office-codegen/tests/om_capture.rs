use std::{
    fs,
    path::PathBuf,
    time::{SystemTime, UNIX_EPOCH},
};

use office_codegen::{
    CanonicalOmGenerationError, CodegenSummary, OmCaptureBundleError, OmSourcesManifest,
    PiaCaptureClass, PiaCaptureInterface, PiaPublicSurfaceCapture, TypelibIdentityCapture,
    build_coverage_report, build_coverage_report_from_json, build_coverage_report_from_path,
    build_focus_surface_registry, build_focus_surface_registry_from_json,
    build_focus_surface_registry_from_path, generate_canonical_office_idl_from_dir,
    normalize_capture_bundle, normalize_capture_bundle_from_dir, normalize_pia_capture_json,
    summarize_capture_bundle, summarize_om_sources, summarize_om_sources_toml,
};
use office_idl::{AccessMode, CaptureOriginKind, InterfaceKind, OfficeIdlDocument};

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()
        .expect("repo root")
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
    assert_eq!(summary.enum_count, 1);
    assert_eq!(summary.interface_count, 6);
    assert_eq!(summary.class_count, 3);
    assert_eq!(summary.member_count, 69);
    assert_eq!(summary.stub_member_count, 69);
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
        document.interfaces[4]
            .metadata
            .as_ref()
            .map(|metadata| metadata.source_inherits.clone()),
        Some(vec![
            "IDispatch".to_string(),
            "Excel._Worksheet".to_string()
        ])
    );
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
    assert_eq!(add_member.params.len(), 0);
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
    assert_eq!(application_intersect.params.len(), 2);
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
    assert_eq!(application_union.params.len(), 2);
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
    assert_eq!(range_address.params.len(), 2);
    assert_eq!(range_address.params[0].name, "RowAbsolute");
    assert!(range_address.params[0].optional);
    assert_eq!(range_address.params[1].name, "ColumnAbsolute");
    assert!(range_address.params[1].optional);
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
            "Workbook".to_string(),
            "Worksheet".to_string(),
            "Range".to_string()
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
            "Workbook".to_string(),
            "Worksheet".to_string(),
            "Range".to_string()
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
    assert_eq!(report_from_path.missing_focus_surfaces.len(), 3);

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
    assert_eq!(
        generation.bundle_paths.raw_typelib_identity_path,
        raw_dir.join("raw_typelib_identity.json")
    );
    assert_eq!(
        generation.bundle_paths.excel_pia_public_surface_path,
        snapshots_dir.join("excel_pia_public_surface.json")
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
    assert_eq!(round_trip.interfaces.len(), template.interfaces.len());
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
    assert_eq!(registry.focus_surfaces.len(), 4);
    assert!(registry.missing_focus_surfaces.is_empty());

    let application = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Application")
        .expect("Application");
    let workbook = registry
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Workbook")
        .expect("Workbook");
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

    assert_eq!(application.member_count, 16);
    assert_eq!(workbook.member_count, 10);
    assert_eq!(worksheet.member_count, 13);
    assert_eq!(range.member_count, 22);
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
    assert_eq!(application_intersect.params.len(), 2);
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
    assert_eq!(application_union.params.len(), 2);
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
    assert_eq!(range_address.params.len(), 2);
    assert_eq!(range_address.params[0].name, "RowAbsolute");
    assert!(range_address.params[0].optional);
    assert_eq!(range_address.params[1].name, "ColumnAbsolute");
    assert!(range_address.params[1].optional);
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
    assert_eq!(coverage.member_count, 69);
    assert_eq!(coverage.support_counts.stub, 69);
    assert!(coverage.missing_focus_surfaces.is_empty());

    let application_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Application")
        .expect("Application coverage");
    let workbook_coverage = coverage
        .focus_surfaces
        .iter()
        .find(|entry| entry.name == "Workbook")
        .expect("Workbook coverage");
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

    assert_eq!(application_coverage.member_count, 16);
    assert_eq!(application_coverage.support_counts.stub, 16);
    assert_eq!(
        application_coverage.stub_members,
        vec![
            "Workbooks".to_string(),
            "Worksheets".to_string(),
            "ActiveWorkbook".to_string(),
            "ActiveSheet".to_string(),
            "ActiveCell".to_string(),
            "Selection".to_string(),
            "DisplayAlerts".to_string(),
            "Calculation".to_string(),
            "Cells".to_string(),
            "Rows".to_string(),
            "Columns".to_string(),
            "CalculateFullRebuild".to_string(),
            "Goto".to_string(),
            "Range".to_string(),
            "Intersect".to_string(),
            "Union".to_string()
        ]
    );

    assert_eq!(workbook_coverage.member_count, 10);
    assert_eq!(workbook_coverage.support_counts.stub, 10);
    assert_eq!(
        workbook_coverage.stub_members,
        vec![
            "Worksheets".to_string(),
            "Name".to_string(),
            "Parent".to_string(),
            "Path".to_string(),
            "FullName".to_string(),
            "ReadOnly".to_string(),
            "Saved".to_string(),
            "Save".to_string(),
            "SaveAs".to_string(),
            "Close".to_string()
        ]
    );

    assert_eq!(worksheet_coverage.member_count, 13);
    assert_eq!(worksheet_coverage.support_counts.stub, 13);
    assert_eq!(
        worksheet_coverage.stub_members,
        vec![
            "Name".to_string(),
            "Parent".to_string(),
            "Index".to_string(),
            "Type".to_string(),
            "Range".to_string(),
            "UsedRange".to_string(),
            "Cells".to_string(),
            "Rows".to_string(),
            "Columns".to_string(),
            "Activate".to_string(),
            "Delete".to_string(),
            "Move".to_string(),
            "Copy".to_string()
        ]
    );

    assert_eq!(range_coverage.member_count, 22);
    assert_eq!(range_coverage.support_counts.stub, 22);
    assert_eq!(
        range_coverage.stub_members,
        vec![
            "Value".to_string(),
            "Value2".to_string(),
            "Formula".to_string(),
            "Text".to_string(),
            "HasFormula".to_string(),
            "Address".to_string(),
            "Parent".to_string(),
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
            "Select".to_string(),
            "ClearContents".to_string()
        ]
    );
}
