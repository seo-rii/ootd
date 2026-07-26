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
