use std::{path::PathBuf, process::Command};

use tempfile::TempDir;

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()
        .expect("repo root")
}

fn capture_binary() -> PathBuf {
    PathBuf::from(env!("CARGO_BIN_EXE_office-capture"))
}

fn run_capture(args: &[&str], current_dir: &std::path::Path) -> std::process::Output {
    Command::new(capture_binary())
        .args(args)
        .current_dir(current_dir)
        .output()
        .expect("run office-capture")
}

#[test]
fn defaults_to_repo_template_when_invoked_outside_repo_root() {
    let tempdir = TempDir::new().expect("tempdir");
    let output = run_capture(&[], tempdir.path());

    assert!(output.status.success());
    let stdout = String::from_utf8(output.stdout).expect("stdout utf8");
    assert!(stdout.contains("capture: excel_om_windows_capture"));
    assert!(stdout.contains("capture_root: C:\\ootd-capture\\excel-om\\excel_om_windows_capture"));
}

#[test]
fn prints_json_summary_when_requested() {
    let tempdir = TempDir::new().expect("tempdir");
    let output = run_capture(&["--json"], tempdir.path());

    assert!(output.status.success());
    let stdout = String::from_utf8(output.stdout).expect("stdout utf8");
    let summary: serde_json::Value = serde_json::from_str(&stdout).expect("json summary");

    assert_eq!(summary["capture_name"], "excel_om_windows_capture");
    assert_eq!(summary["ready_to_run"], false);
    assert_eq!(
        summary["output_paths"][0],
        "C:\\ootd-capture\\excel-om\\excel_om_windows_capture\\manifest\\capture_manifest.json"
    );
}

#[test]
fn rejects_unknown_flag_with_exit_code_two() {
    let tempdir = TempDir::new().expect("tempdir");
    let output = run_capture(&["--bogus"], tempdir.path());

    assert!(!output.status.success());
    assert_eq!(output.status.code(), Some(2));
    let stderr = String::from_utf8(output.stderr).expect("stderr utf8");
    assert!(stderr.contains("unknown flag --bogus"));
}

#[test]
fn prints_help_and_exits_cleanly() {
    let tempdir = TempDir::new().expect("tempdir");
    let output = run_capture(&["--help"], tempdir.path());

    assert!(output.status.success());
    let stdout = String::from_utf8(output.stdout).expect("stdout utf8");
    assert!(stdout.contains("Usage: office-capture"));
    assert!(stdout.contains("--json"));
    assert!(stdout.contains("--powershell-script"));
    assert!(stdout.contains("--powershell-script-out"));
    assert!(stdout.contains("--materialize-execution-bundle"));
    assert!(stdout.contains("--complete-execution-bundle"));
    assert!(stdout.contains("--run-execution-bundle"));
}

#[test]
fn rejects_missing_run_execution_bundle_value() {
    let tempdir = TempDir::new().expect("tempdir");
    let output = run_capture(&["--run-execution-bundle"], tempdir.path());

    assert!(!output.status.success());
    assert_eq!(output.status.code(), Some(2));
    let stderr = String::from_utf8(output.stderr).expect("stderr utf8");
    assert!(stderr.contains("missing value for --run-execution-bundle"));
}

#[test]
fn accepts_explicit_template_path_from_an_arbitrary_working_directory() {
    let tempdir = TempDir::new().expect("tempdir");
    let template = repo_root().join("specs/pinned/windows_capture.template.toml");

    let output = run_capture(
        &["--template", template.to_str().expect("template path")],
        tempdir.path(),
    );

    assert!(output.status.success());
    let stdout = String::from_utf8(output.stdout).expect("stdout utf8");
    assert!(stdout.contains("workspace: ootd"));
}

#[test]
fn prints_powershell_script_when_requested() {
    let tempdir = TempDir::new().expect("tempdir");
    let output = run_capture(&["--powershell-script"], tempdir.path());

    assert!(output.status.success());
    let stdout = String::from_utf8(output.stdout).expect("stdout utf8");
    assert!(stdout.contains("$ErrorActionPreference = 'Stop'"));
    assert!(stdout.contains("Manual step required: launch oleview.exe"));
    assert!(stdout.contains("$GeneratedInteropAssembly"));
}

#[test]
fn rejects_conflicting_output_modes() {
    let tempdir = TempDir::new().expect("tempdir");
    let output = run_capture(
        &["--json", "--run-execution-bundle", "bundle"],
        tempdir.path(),
    );

    assert!(!output.status.success());
    assert_eq!(output.status.code(), Some(2));
    let stderr = String::from_utf8(output.stderr).expect("stderr utf8");
    assert!(stderr.contains("choose at most one output mode"));
}

#[cfg(not(windows))]
#[test]
fn reports_non_windows_failure_for_run_execution_bundle() {
    let tempdir = TempDir::new().expect("tempdir");
    let bundle_root = tempdir.path().join("bundle");
    let materialize_output = run_capture(
        &[
            "--materialize-execution-bundle",
            bundle_root.to_str().expect("bundle root"),
        ],
        tempdir.path(),
    );
    assert!(materialize_output.status.success());
    let output = run_capture(
        &[
            "--run-execution-bundle",
            bundle_root.to_str().expect("bundle root"),
        ],
        tempdir.path(),
    );

    assert!(!output.status.success());
    assert_eq!(output.status.code(), Some(1));
    let stderr = String::from_utf8(output.stderr).expect("stderr utf8");
    assert!(stderr.contains("direct execution requires windows host"));
}

#[test]
fn materializes_powershell_script_to_an_explicit_path() {
    let tempdir = TempDir::new().expect("tempdir");
    let script_path = tempdir.path().join("scripts").join("capture.ps1");
    let output = run_capture(
        &[
            "--powershell-script-out",
            script_path.to_str().expect("script path"),
        ],
        tempdir.path(),
    );

    assert!(output.status.success());
    assert!(output.stdout.is_empty());
    assert!(script_path.exists());
    let script = std::fs::read_to_string(&script_path).expect("script file");
    assert!(script.contains("$ErrorActionPreference = 'Stop'"));
    assert!(script.contains("Manual step required: launch oleview.exe"));
}

#[test]
fn materializes_execution_bundle_to_an_explicit_directory() {
    let tempdir = TempDir::new().expect("tempdir");
    let bundle_root = tempdir.path().join("bundle");
    let output = run_capture(
        &[
            "--materialize-execution-bundle",
            bundle_root.to_str().expect("bundle root"),
        ],
        tempdir.path(),
    );

    assert!(output.status.success());
    let stdout = String::from_utf8(output.stdout).expect("stdout utf8");
    assert!(stdout.contains("script_path:"));
    assert!(stdout.contains("launcher_path:"));
    assert!(stdout.contains("execution_plan_path:"));
    assert!(stdout.contains("launcher_plan_path:"));
    assert!(stdout.contains("execution_receipt_template_path:"));
    assert!(stdout.contains("launcher_status_template_path:"));

    assert!(bundle_root.join("scripts/capture.ps1").exists());
    assert!(bundle_root.join("scripts/run_capture.cmd").exists());
    assert!(bundle_root.join("manifest/execution_plan.json").exists());
    assert!(
        bundle_root
            .join("manifest/direct_exec_status.template.json")
            .exists()
    );
    assert!(
        bundle_root
            .join("manifest/execution_receipt.template.json")
            .exists()
    );
}

#[test]
fn completes_execution_bundle_from_a_receipt_file() {
    let tempdir = TempDir::new().expect("tempdir");
    let bundle_root = tempdir.path().join("bundle");

    std::fs::create_dir_all(bundle_root.join("logs")).expect("logs dir");
    std::fs::create_dir_all(bundle_root.join("raw")).expect("raw dir");
    std::fs::create_dir_all(bundle_root.join("snapshots")).expect("snapshots dir");
    std::fs::create_dir_all(bundle_root.join("manifest")).expect("manifest dir");

    std::fs::write(bundle_root.join("logs/capture.log"), "capture log").expect("capture log");
    std::fs::write(
        bundle_root.join("raw/raw_typelib_identity.json"),
        r#"{"library":"Excel"}"#,
    )
    .expect("raw typelib identity");
    std::fs::write(
        bundle_root.join("snapshots/excel_typelib_snapshot.idl"),
        "library Excel {}",
    )
    .expect("snapshot idl");
    std::fs::write(
        bundle_root.join("snapshots/excel_typelib_snapshot.odl"),
        "odl Excel {}",
    )
    .expect("snapshot odl");
    std::fs::write(
        bundle_root.join("raw/excel_pia_identity.json"),
        r#"{"assembly":"Microsoft.Office.Interop.Excel"}"#,
    )
    .expect("pia identity");
    std::fs::write(
        bundle_root.join("snapshots/excel_pia_public_surface.json"),
        r#"{"library":"Excel","interfaces":[]}"#,
    )
    .expect("pia public surface");
    std::fs::write(
        bundle_root.join("manifest/execution_receipt.json"),
        r#"{
  "startedAtUtc":"2026-03-27T01:02:03Z",
  "completedAtUtc":"2026-03-27T01:03:04Z",
  "host":{"computerName":"WIN-EXCEL","userName":"runner","hostOs":"Windows 11","hostArch":"x64","powershellVersion":"5.1"},
  "commandResults":[{"name":"powershell_capture_reflection","status":"completed","exitCode":0,"detail":"interop reflection capture finished"}],
  "manualStepResults":[{"name":"oleview_snapshot_export","status":"completed","detail":"IDL/ODL snapshots exported"}]
}"#,
    )
    .expect("execution receipt");

    let output = run_capture(
        &[
            "--complete-execution-bundle",
            bundle_root.to_str().expect("bundle root"),
        ],
        tempdir.path(),
    );

    assert!(output.status.success());
    let stdout = String::from_utf8(output.stdout).expect("stdout utf8");
    assert!(stdout.contains("manifest_path:"));
    assert!(stdout.contains("output_checksums_path:"));
    assert!(stdout.contains("launcher_status_path:"));
    assert!(stdout.contains("execution_receipt_path:"));
    assert!(stdout.contains("launcher_path:"));
    assert!(bundle_root.join("manifest/capture_manifest.json").exists());
    assert!(bundle_root.join("manifest/output_checksums.json").exists());
}
