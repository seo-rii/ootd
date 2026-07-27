#!/usr/bin/env python3
"""Run rustfmt with a bounded, reviewed per-file policy."""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from collections.abc import Callable, Mapping, Sequence
from pathlib import Path


MAX_RUSTFMT_FILE_BYTES = 1_000_000

# These files predate the bounded format gate. Their individual ceilings make
# growth an explicit review event while module extraction continues.
REVIEWED_OVERSIZED_RUST_FILES: dict[Path, int] = {
    Path("crates/excel-runtime/src/lib.rs"): 1_650_000,
    Path("crates/excel-runtime/src/tests.rs"): 4_900_000,
    Path("crates/excel-xlsx/src/lib.rs"): 1_425_000,
    Path("crates/excel-xlsx/src/tests.rs"): 4_800_000,
}


class PolicyError(RuntimeError):
    """The repository no longer matches the reviewed size policy."""


class RustfmtError(RuntimeError):
    """One or more bounded rustfmt checks failed."""


def discover_tracked_rust_files(repository_root: Path) -> list[Path]:
    """Return tracked Rust paths in deterministic repository-relative order."""
    result = subprocess.run(
        [
            "git",
            "-C",
            str(repository_root),
            "ls-files",
            "-z",
            "--",
            "*.rs",
        ],
        check=True,
        stdout=subprocess.PIPE,
    )
    paths = [
        Path(os.fsdecode(raw_path))
        for raw_path in result.stdout.split(b"\0")
        if raw_path
    ]
    return sorted(paths, key=Path.as_posix)


def validate_size_policy(
    repository_root: Path,
    tracked_files: Sequence[Path],
    *,
    threshold_bytes: int = MAX_RUSTFMT_FILE_BYTES,
    reviewed_exceptions: Mapping[str | Path, int] | None = None,
) -> list[Path]:
    """Validate oversized exceptions and return files rustfmt should check."""
    if threshold_bytes <= 0:
        raise PolicyError("Rust file size threshold must be positive")

    configured_exceptions = (
        REVIEWED_OVERSIZED_RUST_FILES
        if reviewed_exceptions is None
        else reviewed_exceptions
    )
    exceptions = {
        Path(path): ceiling for path, ceiling in configured_exceptions.items()
    }
    tracked = sorted(set(tracked_files), key=Path.as_posix)
    tracked_set = set(tracked)
    sizes: dict[Path, int] = {}
    errors: list[str] = []

    for path in tracked:
        if path.is_absolute() or ".." in path.parts:
            errors.append(f"tracked Rust path is not repository-relative: {path}")
            continue
        absolute_path = repository_root / path
        if not absolute_path.is_file():
            errors.append(f"tracked Rust file is missing: {path.as_posix()}")
            continue
        sizes[path] = absolute_path.stat().st_size

    for path, ceiling in sorted(
        exceptions.items(),
        key=lambda item: item[0].as_posix(),
    ):
        display_path = path.as_posix()
        if path.is_absolute() or ".." in path.parts or path.suffix != ".rs":
            errors.append(f"invalid reviewed exception path: {display_path}")
            continue
        if ceiling <= threshold_bytes:
            errors.append(
                f"reviewed exception ceiling must exceed {threshold_bytes} bytes: "
                f"{display_path} ({ceiling} bytes)"
            )
            continue
        if path not in tracked_set:
            errors.append(f"reviewed exception is not tracked: {display_path}")
            continue
        size = sizes.get(path)
        if size is None:
            continue
        if size <= threshold_bytes:
            errors.append(
                f"reviewed exception is no longer oversized: {display_path} "
                f"({size} <= {threshold_bytes} bytes)"
            )
        elif size > ceiling:
            errors.append(
                f"reviewed exception exceeds its reviewed ceiling: {display_path} "
                f"({size} > {ceiling} bytes)"
            )

    for path in tracked:
        size = sizes.get(path)
        if (
            size is not None
            and size > threshold_bytes
            and path not in exceptions
        ):
            errors.append(
                f"unreviewed oversized Rust file: {path.as_posix()} "
                f"({size} > {threshold_bytes} bytes)"
            )

    if errors:
        raise PolicyError("\n".join(errors))

    return [path for path in tracked if path not in exceptions]


def run_rustfmt(
    repository_root: Path,
    files: Sequence[Path],
    *,
    rustfmt: str = "rustfmt",
    runner: Callable[..., subprocess.CompletedProcess] = subprocess.run,
) -> None:
    """Run rustfmt sequentially without recursively formatting child modules."""
    failed: list[Path] = []
    for path in sorted(files, key=Path.as_posix):
        print(f"rustfmt --check: {path.as_posix()}", flush=True)
        command = [
            rustfmt,
            "--check",
            "--edition",
            "2024",
            "--config",
            "skip_children=true",
            str(repository_root / path),
        ]
        result = runner(command, check=False)
        if result.returncode != 0:
            failed.append(path)

    if failed:
        joined_paths = ", ".join(path.as_posix() for path in failed)
        raise RustfmtError(f"rustfmt failed for: {joined_paths}")


def parse_arguments(arguments: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Check tracked Rust files one at a time while guarding reviewed "
            "oversized-file exceptions."
        )
    )
    parser.add_argument(
        "--repo-root",
        type=Path,
        default=Path(__file__).resolve().parents[1],
        help="Git worktree root (defaults to this script's repository)",
    )
    parser.add_argument(
        "--rustfmt",
        default="rustfmt",
        help="rustfmt executable to invoke",
    )
    parser.add_argument(
        "--policy-only",
        action="store_true",
        help="validate tracked paths and size limits without invoking rustfmt",
    )
    return parser.parse_args(arguments)


def main(arguments: Sequence[str] | None = None) -> int:
    options = parse_arguments(sys.argv[1:] if arguments is None else arguments)
    repository_root = options.repo_root.resolve()

    try:
        tracked_files = discover_tracked_rust_files(repository_root)
        format_targets = validate_size_policy(repository_root, tracked_files)
        if not options.policy_only:
            run_rustfmt(
                repository_root,
                format_targets,
                rustfmt=options.rustfmt,
            )
    except (OSError, PolicyError, RustfmtError, subprocess.CalledProcessError) as error:
        print(f"bounded rustfmt check failed: {error}", file=sys.stderr)
        return 1

    exception_count = len(REVIEWED_OVERSIZED_RUST_FILES)
    print(
        f"bounded rustfmt check passed: {len(format_targets)} files checked, "
        f"{exception_count} reviewed oversized files skipped"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
