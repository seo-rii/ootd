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
