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
