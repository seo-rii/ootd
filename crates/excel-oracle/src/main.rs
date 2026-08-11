use std::collections::BTreeMap;
use std::path::PathBuf;

use excel_oracle::{ORACLE_SCHEMA_VERSION, PinnedSuiteArtifacts};

const MAX_FRAGMENT_ROOTS: usize = 4_096;
const HELP: &str = "Usage:\n  excel-oracle capture-plan --suite-root PATH\n  excel-oracle assemble-run --suite-root PATH --fragment-root PATH... --output-root PATH\n\nOptions:\n  --suite-root PATH     Root containing manifest/suite_manifest.json\n  --fragment-root PATH  Case-run fragment root; repeat for every fragment\n  --output-root PATH    Fresh destination for the complete run bundle\n  -h, --help            Show this help message";

fn main() {
    let mut args = std::env::args_os().skip(1);
    let Some(command) = args.next() else {
        eprintln!("excel-oracle: expected a command\n\n{HELP}");
        std::process::exit(2);
    };
    let command = match command.to_str() {
        Some("-h" | "--help") => {
            println!("{HELP}");
            return;
        }
        Some("capture-plan") => "capture-plan",
        Some("assemble-run") => "assemble-run",
        Some(value) => {
            eprintln!("excel-oracle: unknown command {value}\n\n{HELP}");
            std::process::exit(2);
        }
        None => {
            eprintln!("excel-oracle: command must be valid Unicode\n\n{HELP}");
            std::process::exit(2);
        }
    };

    if command == "capture-plan" {
        let mut suite_root = None::<PathBuf>;
        while let Some(arg) = args.next() {
            match arg.to_str() {
                Some("-h" | "--help") => {
                    println!("{HELP}");
                    return;
                }
                Some("--suite-root") => {
                    let Some(value) = args.next() else {
                        eprintln!("excel-oracle: missing value for --suite-root");
                        std::process::exit(2);
                    };
                    if suite_root.is_some() {
                        eprintln!("excel-oracle: --suite-root specified more than once");
                        std::process::exit(2);
                    }
                    suite_root = Some(PathBuf::from(value));
                }
                Some(value) if value.starts_with('-') => {
                    eprintln!("excel-oracle: unknown flag {value}");
                    std::process::exit(2);
                }
                Some(value) => {
                    eprintln!("excel-oracle: unexpected positional argument {value}");
                    std::process::exit(2);
                }
                None => {
                    eprintln!("excel-oracle: flags must be valid Unicode");
                    std::process::exit(2);
                }
            }
        }
        let Some(suite_root) = suite_root else {
            eprintln!("excel-oracle: missing required --suite-root");
            std::process::exit(2);
        };
        let suite = PinnedSuiteArtifacts::load(&suite_root).unwrap_or_else(|error| {
            eprintln!("excel-oracle: {error}");
            std::process::exit(1);
        });
        let cases = suite
            .manifest
            .cases
            .iter()
            .map(|artifact| {
                let case_bytes = suite
                    .cases
                    .get(&artifact.case_id)
                    .expect("validated suite case bytes");
                let case = suite
                    .manifest
                    .load_case(&artifact.case_id, case_bytes)
                    .expect("validated suite case metadata");
                serde_json::json!({
                    "caseId": artifact.case_id,
                    "caseVersion": artifact.case_version,
                    "tier": artifact.tier,
                    "casePath": artifact.path,
                    "caseSha256": artifact.sha256,
                    "inputPath": case.input.path,
                    "inputSha256": artifact.input_sha256,
                })
            })
            .collect::<Vec<_>>();
        let plan = serde_json::json!({
            "schemaVersion": ORACLE_SCHEMA_VERSION,
            "suiteId": suite.manifest.id,
            "profileId": suite.manifest.profile_id,
            "expectedEngine": suite.manifest.expected_engine,
            "caseCount": cases.len(),
            "cases": cases,
        });
        println!(
            "{}",
            serde_json::to_string_pretty(&plan).expect("JSON value serialization is infallible"),
        );
        return;
    }

    let mut suite_root = None::<PathBuf>;
    let mut fragment_roots = Vec::<PathBuf>::new();
    let mut output_root = None::<PathBuf>;
    while let Some(arg) = args.next() {
        match arg.to_str() {
            Some("-h" | "--help") => {
                println!("{HELP}");
                return;
            }
            Some("--suite-root") => {
                let Some(value) = args.next() else {
                    eprintln!("excel-oracle: missing value for --suite-root");
                    std::process::exit(2);
                };
                if suite_root.is_some() {
                    eprintln!("excel-oracle: --suite-root specified more than once");
                    std::process::exit(2);
                }
                suite_root = Some(PathBuf::from(value));
            }
            Some("--fragment-root") => {
                let Some(value) = args.next() else {
                    eprintln!("excel-oracle: missing value for --fragment-root");
                    std::process::exit(2);
                };
                if fragment_roots.len() == MAX_FRAGMENT_ROOTS {
                    eprintln!(
                        "excel-oracle: at most {MAX_FRAGMENT_ROOTS} fragment roots are allowed"
                    );
                    std::process::exit(2);
                }
                fragment_roots.push(PathBuf::from(value));
            }
            Some("--output-root") => {
                let Some(value) = args.next() else {
                    eprintln!("excel-oracle: missing value for --output-root");
                    std::process::exit(2);
                };
                if output_root.is_some() {
                    eprintln!("excel-oracle: --output-root specified more than once");
                    std::process::exit(2);
                }
                output_root = Some(PathBuf::from(value));
            }
            Some(value) if value.starts_with('-') => {
                eprintln!("excel-oracle: unknown flag {value}");
                std::process::exit(2);
            }
            Some(value) => {
                eprintln!("excel-oracle: unexpected positional argument {value}");
                std::process::exit(2);
            }
            None => {
                eprintln!("excel-oracle: flags must be valid Unicode");
                std::process::exit(2);
            }
        }
    }

    let Some(suite_root) = suite_root else {
        eprintln!("excel-oracle: missing required --suite-root");
        std::process::exit(2);
    };
    if fragment_roots.is_empty() {
        eprintln!("excel-oracle: at least one --fragment-root is required");
        std::process::exit(2);
    }
    let Some(output_root) = output_root else {
        eprintln!("excel-oracle: missing required --output-root");
        std::process::exit(2);
    };

    let suite = PinnedSuiteArtifacts::load(&suite_root).unwrap_or_else(|error| {
        eprintln!("excel-oracle: {error}");
        std::process::exit(1);
    });
    let mut fragments = Vec::with_capacity(fragment_roots.len());
    for fragment_root in &fragment_roots {
        fragments.push(
            suite
                .load_run_fragment(fragment_root)
                .unwrap_or_else(|error| {
                    eprintln!("excel-oracle: {error}");
                    std::process::exit(1);
                }),
        );
    }
    let assembled = suite
        .assemble_run_fragments(&fragments)
        .unwrap_or_else(|error| {
            eprintln!("excel-oracle: {error}");
            std::process::exit(1);
        });
    let run_id = assembled.manifest.run_id.clone();
    let case_count = assembled.manifest.cases.len();
    let completed_observation_count = assembled.observations.len();
    let written = suite
        .write_run_bundle(&assembled, &output_root)
        .unwrap_or_else(|error| {
            eprintln!("excel-oracle: {error}");
            std::process::exit(1);
        });
    let observation_paths = written
        .observation_paths
        .iter()
        .map(|(case_id, path)| {
            (
                case_id.clone(),
                path.as_os_str().to_string_lossy().into_owned(),
            )
        })
        .collect::<BTreeMap<_, _>>();
    let receipt = serde_json::json!({
        "schemaVersion": ORACLE_SCHEMA_VERSION,
        "runId": run_id,
        "caseCount": case_count,
        "completedObservationCount": completed_observation_count,
        "outputRoot": written.root.as_os_str().to_string_lossy(),
        "manifestPath": written.manifest_path.as_os_str().to_string_lossy(),
        "observationPaths": observation_paths,
    });
    println!(
        "{}",
        serde_json::to_string_pretty(&receipt).unwrap_or_else(|error| {
            eprintln!("excel-oracle: failed to serialize assembly receipt: {error}");
            std::process::exit(1);
        }),
    );
}
