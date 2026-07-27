use std::fs;
use std::path::Path;

#[test]
fn unit_tests_are_externalized_from_library_root() {
    let crate_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let library = fs::read_to_string(crate_root.join("src/lib.rs")).expect("read library root");

    assert!(crate_root.join("src/tests.rs").is_file());
    assert!(library.contains("#[cfg(test)]\nmod tests;"));
    assert!(!library.contains("#[cfg(test)]\nmod tests {"));
}

#[test]
fn calculation_engine_is_isolated_from_library_root() {
    let crate_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let library = fs::read_to_string(crate_root.join("src/lib.rs")).expect("read library root");

    assert!(crate_root.join("src/calc/mod.rs").is_file());
    assert!(library.contains("mod calc;"));
    assert!(!library.contains("enum FormulaEvalError"));
}

#[test]
fn calculation_engine_declares_its_parent_dependencies() {
    let crate_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let calculation =
        fs::read_to_string(crate_root.join("src/calc/mod.rs")).expect("read calculation module");

    assert!(!calculation.contains("use super::*;"));
    assert!(calculation.contains("use super::{"));
    assert!(calculation.contains("use office_common::{"));
}

#[test]
fn recalculation_writeback_is_isolated_from_library_root() {
    let crate_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let library = fs::read_to_string(crate_root.join("src/lib.rs")).expect("read library root");

    assert!(crate_root.join("src/recalculation.rs").is_file());
    assert!(library.contains("mod recalculation;"));
    assert!(!library.contains("fn calculate_sheet_formulas("));
}

#[test]
fn application_dispatch_is_grouped_by_object_surface() {
    let crate_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let library = fs::read_to_string(crate_root.join("src/lib.rs")).expect("read library root");
    let application = fs::read_to_string(crate_root.join("src/dispatch/application.rs"))
        .expect("read Application dispatch module");

    assert!(library.contains("mod dispatch;"));
    assert!(application.contains("fn dispatch_get_application("));
    assert!(application.contains("fn dispatch_invoke_application("));
    assert!(!application.contains("use super::super::*;"));
    assert!(!library.contains("fn dispatch_get_application("));
    assert!(!library.contains("fn dispatch_invoke_application("));
}

#[test]
fn workbook_dispatch_is_grouped_by_object_surface() {
    let crate_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let library = fs::read_to_string(crate_root.join("src/lib.rs")).expect("read library root");
    let workbook = fs::read_to_string(crate_root.join("src/dispatch/workbook.rs"))
        .expect("read Workbook dispatch module");

    for method in [
        "dispatch_get_workbook",
        "dispatch_get_workbooks",
        "dispatch_invoke_workbook",
        "dispatch_invoke_workbooks",
    ] {
        assert!(workbook.contains(&format!("fn {method}(")));
        assert!(!library.contains(&format!("fn {method}(")));
    }
    assert!(!workbook.contains("use super::super::*;"));
}

#[test]
fn worksheet_function_dispatch_is_grouped_by_object_surface() {
    let crate_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let library = fs::read_to_string(crate_root.join("src/lib.rs")).expect("read library root");
    let worksheet_function =
        fs::read_to_string(crate_root.join("src/dispatch/worksheet_function.rs"))
            .expect("read WorksheetFunction dispatch module");

    for method in [
        "dispatch_get_worksheet_function",
        "dispatch_invoke_worksheet_function",
        "worksheet_function_formula_arg",
        "worksheet_function_array_literal",
        "worksheet_function_range_reference_text",
    ] {
        assert!(worksheet_function.contains(&format!("fn {method}(")));
        assert!(!library.contains(&format!("fn {method}(")));
    }
    assert!(!worksheet_function.contains("use super::super::*;"));
}

#[test]
fn worksheet_dispatch_is_grouped_by_object_surface() {
    let crate_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let library = fs::read_to_string(crate_root.join("src/lib.rs")).expect("read library root");
    let worksheet = fs::read_to_string(crate_root.join("src/dispatch/worksheet.rs"))
        .expect("read Worksheet dispatch module");

    for method in [
        "dispatch_get_sheet_collection",
        "dispatch_get_worksheet",
        "dispatch_invoke_sheet_collection",
        "dispatch_invoke_worksheet",
    ] {
        assert!(worksheet.contains(&format!("fn {method}(")));
        assert!(!library.contains(&format!("fn {method}(")));
    }
    assert!(!worksheet.contains("use super::super::*;"));
}

#[test]
fn names_dispatch_is_grouped_by_object_surface() {
    let crate_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let library = fs::read_to_string(crate_root.join("src/lib.rs")).expect("read library root");
    let names = fs::read_to_string(crate_root.join("src/dispatch/names.rs"))
        .expect("read Names dispatch module");

    for method in [
        "dispatch_get_names",
        "dispatch_invoke_names",
        "dispatch_get_name",
        "dispatch_invoke_name",
    ] {
        assert!(names.contains(&format!("fn {method}(")));
        assert!(!library.contains(&format!("fn {method}(")));
    }
    assert!(!names.contains("use super::super::*;"));
}
