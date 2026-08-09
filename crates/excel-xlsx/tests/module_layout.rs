use std::fs;
use std::path::Path;

#[test]
fn unit_tests_are_externalized_from_library_root() {
    let crate_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let library = fs::read_to_string(crate_root.join("src/lib.rs"))
        .expect("read library root")
        .replace("\r\n", "\n");

    assert!(crate_root.join("src/tests.rs").is_file());
    assert!(library.contains("#[cfg(test)]\nmod tests;"));
    assert!(!library.contains("#[cfg(test)]\nmod tests {"));
}

#[test]
fn shared_string_parser_is_isolated_from_library_root() {
    let crate_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let library = fs::read_to_string(crate_root.join("src/lib.rs")).expect("read library root");

    assert!(crate_root.join("src/shared_strings.rs").is_file());
    assert!(library.contains("mod shared_strings;"));
    assert!(!library.contains("fn parse_shared_strings("));
}

#[test]
fn relationship_primitives_are_isolated_from_library_root() {
    let crate_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let library = fs::read_to_string(crate_root.join("src/lib.rs")).expect("read library root");

    assert!(crate_root.join("src/relationships.rs").is_file());
    assert!(library.contains("mod relationships;"));
    assert!(!library.contains("struct RelationshipEntry"));
    assert!(!library.contains("fn normalize_relationship_target("));
}

#[test]
fn worksheet_cell_codec_is_isolated_from_library_root() {
    let crate_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let library = fs::read_to_string(crate_root.join("src/lib.rs")).expect("read library root");

    assert!(crate_root.join("src/worksheet/cells.rs").is_file());
    assert!(library.contains("mod worksheet;"));
    assert!(!library.contains("enum RowContentSegment"));
    assert!(!library.contains("fn parse_cell_reference("));
}

#[test]
fn worksheet_table_owner_parser_is_isolated_from_library_root() {
    let crate_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let library = fs::read_to_string(crate_root.join("src/lib.rs")).expect("read library root");
    let worksheet_module =
        fs::read_to_string(crate_root.join("src/worksheet/mod.rs")).expect("read worksheet module");

    assert!(crate_root.join("src/worksheet/table.rs").is_file());
    assert!(worksheet_module.contains("mod table;"));
    assert!(!library.contains("fn parse_table_structural_owner("));
    assert!(!library.contains("fn resolve_table_structural_owners("));
}
