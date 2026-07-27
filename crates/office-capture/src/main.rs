use std::{fs, path::PathBuf};

fn main() {
    let mut template_path = None::<PathBuf>;
    let mut emit_json = false;
    let mut emit_powershell_script = false;
    let mut powershell_script_out = None::<PathBuf>;
    let mut materialize_execution_bundle = None::<PathBuf>;
    let mut complete_execution_bundle = None::<PathBuf>;
    let mut run_execution_bundle = None::<PathBuf>;

    let mut args = std::env::args().skip(1);
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "-h" | "--help" => {
                println!(
                    "Usage: office-capture [--template PATH] [--json] [PATH]\n\n\
                     Options:\n\
                     \n\
                     -t, --template PATH  Path to windows_capture.template.toml\n\
                     --json               Print the capture plan as JSON\n\
                     --powershell-script  Print the Windows capture PowerShell script\n\
                     --powershell-script-out PATH\n\
                                        Write the PowerShell script to PATH\n\
                     --materialize-execution-bundle DIR\n\
                                        Write scripts/capture.ps1 and manifest helpers under DIR\n\
                     --complete-execution-bundle DIR\n\
                                        Finalize manifest/checksums from DIR and manifest/execution_receipt.json\n\
                     --run-execution-bundle DIR\n\
                                        Run the materialized direct-exec launcher under DIR\n\
                     -h, --help           Show this help message"
                );
                return;
            }
            "--json" => {
                emit_json = true;
            }
            "--powershell-script" => {
                emit_powershell_script = true;
            }
            "--powershell-script-out" => {
                let Some(value) = args.next() else {
                    eprintln!("office-capture: missing value for --powershell-script-out");
                    std::process::exit(2);
                };
                if powershell_script_out.is_some() {
                    eprintln!(
                        "office-capture: powershell script output path specified more than once"
                    );
                    std::process::exit(2);
                }
                powershell_script_out = Some(PathBuf::from(value));
            }
            "--materialize-execution-bundle" => {
                let Some(value) = args.next() else {
                    eprintln!("office-capture: missing value for --materialize-execution-bundle");
                    std::process::exit(2);
                };
                if materialize_execution_bundle.is_some() {
                    eprintln!(
                        "office-capture: execution bundle output path specified more than once"
                    );
                    std::process::exit(2);
                }
                materialize_execution_bundle = Some(PathBuf::from(value));
            }
            "--complete-execution-bundle" => {
                let Some(value) = args.next() else {
                    eprintln!("office-capture: missing value for --complete-execution-bundle");
                    std::process::exit(2);
                };
                if complete_execution_bundle.is_some() {
                    eprintln!("office-capture: completion bundle path specified more than once");
                    std::process::exit(2);
                }
                complete_execution_bundle = Some(PathBuf::from(value));
            }
            "--run-execution-bundle" => {
                let Some(value) = args.next() else {
                    eprintln!("office-capture: missing value for --run-execution-bundle");
                    std::process::exit(2);
                };
                if run_execution_bundle.is_some() {
                    eprintln!("office-capture: execution run bundle path specified more than once");
                    std::process::exit(2);
                }
                run_execution_bundle = Some(PathBuf::from(value));
            }
            "-t" | "--template" => {
                let Some(value) = args.next() else {
                    eprintln!("office-capture: missing value for --template");
                    std::process::exit(2);
                };
                if template_path.is_some() {
                    eprintln!("office-capture: template path specified more than once");
                    std::process::exit(2);
                }
                template_path = Some(PathBuf::from(value));
            }
            value if value.starts_with('-') => {
                eprintln!("office-capture: unknown flag {value}");
                std::process::exit(2);
            }
            value => {
                if template_path.is_some() {
                    eprintln!("office-capture: expected at most one template path");
                    std::process::exit(2);
                }
                template_path = Some(PathBuf::from(value));
            }
        }
    }

    let output_mode_count = usize::from(emit_json)
        + usize::from(emit_powershell_script)
        + usize::from(powershell_script_out.is_some())
        + usize::from(materialize_execution_bundle.is_some())
        + usize::from(complete_execution_bundle.is_some())
        + usize::from(run_execution_bundle.is_some());
    if output_mode_count > 1 {
        eprintln!(
            "office-capture: choose at most one output mode: --json, --powershell-script, --powershell-script-out, --materialize-execution-bundle, --complete-execution-bundle, or --run-execution-bundle"
        );
        std::process::exit(2);
    }

    let template_path = template_path
        .unwrap_or_else(|| PathBuf::from("specs/pinned/windows_capture.template.toml"));
    let repo_root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let cwd = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
    let resolved_template_path = if template_path.is_absolute() {
        template_path.clone()
    } else {
        let cwd_candidate = cwd.join(&template_path);
        if cwd_candidate.exists() {
            cwd_candidate
        } else {
            let repo_candidate = repo_root.join(&template_path);
            if repo_candidate.exists() {
                repo_candidate
            } else {
                cwd_candidate
            }
        }
    };

    match office_capture::CapturePlan::from_toml_path(&resolved_template_path) {
        Ok(plan) => {
            if let Some(bundle_root) = run_execution_bundle {
                let result = plan
                    .run_execution_bundle(&bundle_root)
                    .unwrap_or_else(|err| {
                        eprintln!("office-capture: {err}");
                        std::process::exit(1);
                    });
                if let Some(manifest_path) = result.completion_result.write_result.manifest_path {
                    println!("manifest_path: {}", manifest_path.display());
                }
                if let Some(checksums_path) =
                    result.completion_result.write_result.output_checksums_path
                {
                    println!("output_checksums_path: {}", checksums_path.display());
                }
                println!(
                    "launcher_status_path: {}",
                    result.direct_exec_status_path.display()
                );
                println!(
                    "execution_receipt_path: {}",
                    result.completion_result.execution_receipt_path.display()
                );
                println!("launcher_path: {}", result.launcher_path.display());
                println!("launcher_exit_code: {}", result.launcher_exit_code);
                println!("launcher_status: {}", result.direct_exec_status.status);
                return;
            }

            if let Some(bundle_root) = complete_execution_bundle {
                let result = plan
                    .complete_execution_bundle(&bundle_root)
                    .unwrap_or_else(|err| {
                        eprintln!("office-capture: {err}");
                        std::process::exit(1);
                    });
                if let Some(manifest_path) = result.write_result.manifest_path {
                    println!("manifest_path: {}", manifest_path.display());
                }
                if let Some(checksums_path) = result.write_result.output_checksums_path {
                    println!("output_checksums_path: {}", checksums_path.display());
                }
                println!(
                    "launcher_status_path: {}",
                    bundle_root
                        .join("manifest/direct_exec_status.json")
                        .display()
                );
                println!(
                    "execution_receipt_path: {}",
                    result.execution_receipt_path.display()
                );
                println!("launcher_path: {}", plan.direct_exec_launcher_path());
                return;
            }

            if let Some(bundle_root) = materialize_execution_bundle {
                let result = plan
                    .materialize_execution_bundle(&bundle_root)
                    .unwrap_or_else(|err| {
                        eprintln!("office-capture: {err}");
                        std::process::exit(1);
                    });
                println!(
                    "launcher_path: {}",
                    result.direct_exec_launcher_path.display()
                );
                println!("script_path: {}", result.script_path.display());
                println!(
                    "launcher_plan_path: {}",
                    result.execution_plan_path.display()
                );
                println!(
                    "execution_plan_path: {}",
                    result.execution_plan_path.display()
                );
                println!(
                    "launcher_status_template_path: {}",
                    result.direct_exec_status_template_path.display()
                );
                println!(
                    "execution_receipt_template_path: {}",
                    result.execution_receipt_template_path.display()
                );
                return;
            }

            if let Some(script_out_path) = powershell_script_out {
                let script = plan.render_powershell_script();
                if let Some(parent) = script_out_path.parent()
                    && !parent.as_os_str().is_empty()
                {
                    fs::create_dir_all(parent).unwrap_or_else(|err| {
                        eprintln!(
                            "office-capture: failed to create {}: {}",
                            parent.display(),
                            err
                        );
                        std::process::exit(1);
                    });
                }
                fs::write(&script_out_path, script).unwrap_or_else(|err| {
                    eprintln!(
                        "office-capture: failed to write {}: {}",
                        script_out_path.display(),
                        err
                    );
                    std::process::exit(1);
                });
                return;
            }

            if emit_powershell_script {
                print!("{}", plan.render_powershell_script());
            } else if emit_json {
                let summary = plan.summary();
                let rendered = serde_json::json!({
                    "capture_name": summary.capture_name,
                    "capture_workspace": summary.capture_workspace,
                    "target_host_os": summary.target_host_os,
                    "target_product_family": summary.target_product_family,
                    "target_channel": summary.target_channel,
                    "target_version": summary.target_version,
                    "target_build": summary.target_build,
                    "target_arch": summary.target_arch,
                    "target_locale": summary.target_locale,
                    "output_dir": summary.output_dir,
                    "capture_root": summary.capture_root,
                    "output_paths": summary.output_paths,
                    "pending_capture_outputs": summary.pending_capture_outputs,
                    "downstream_path": summary.downstream_path,
                    "unresolved_fields": summary.unresolved_fields,
                    "ready_to_run": summary.ready_to_run,
                });
                println!(
                    "{}",
                    serde_json::to_string_pretty(&rendered).expect("serialize capture summary")
                );
            } else {
                let summary = plan.summary();
                print!("{summary}");
            }
        }
        Err(err) => {
            eprintln!("office-capture: {err}");
            std::process::exit(1);
        }
    }
}
