from __future__ import annotations

import importlib.util
import subprocess
import tempfile
import unittest
from pathlib import Path


REPOSITORY_ROOT = Path(__file__).resolve().parents[2]
HELPER_PATH = REPOSITORY_ROOT / "scripts" / "check_rustfmt.py"

SPEC = importlib.util.spec_from_file_location("check_rustfmt", HELPER_PATH)
if SPEC is None or SPEC.loader is None:
    raise RuntimeError(f"cannot load {HELPER_PATH}")
CHECK_RUSTFMT = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(CHECK_RUSTFMT)


class TrackedRustFileDiscoveryTests(unittest.TestCase):
    def test_discovers_only_tracked_rust_files_in_stable_order(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            subprocess.run(
                ["git", "init", "--quiet", str(root)],
                check=True,
            )
            (root / "nested").mkdir()
            (root / "z.rs").write_text("fn z() {}\n", encoding="utf-8")
            (root / "nested" / "a.rs").write_text("fn a() {}\n", encoding="utf-8")
            (root / "ignored.rs").write_text("fn ignored() {}\n", encoding="utf-8")
            (root / "notes.txt").write_text("not Rust\n", encoding="utf-8")
            subprocess.run(
                ["git", "-C", str(root), "add", "z.rs", "nested/a.rs", "notes.txt"],
                check=True,
            )

            files = CHECK_RUSTFMT.discover_tracked_rust_files(root)

            self.assertEqual(
                [Path("nested/a.rs"), Path("z.rs")],
                files,
            )


class OversizedFilePolicyTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary_directory.name)

    def tearDown(self) -> None:
        self.temporary_directory.cleanup()

    def write_bytes(self, path: str, size: int) -> Path:
        destination = self.root / path
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(b"x" * size)
        return Path(path)

    def test_accepts_a_reviewed_oversized_file_within_its_ceiling(self) -> None:
        small = self.write_bytes("small.rs", 4)
        legacy = self.write_bytes("legacy.rs", 12)

        targets = CHECK_RUSTFMT.validate_size_policy(
            self.root,
            [legacy, small],
            threshold_bytes=10,
            reviewed_exceptions={"legacy.rs": 20},
        )

        self.assertEqual([small], targets)

    def test_rejects_a_new_unreviewed_oversized_file(self) -> None:
        oversized = self.write_bytes("new.rs", 11)

        with self.assertRaisesRegex(
            CHECK_RUSTFMT.PolicyError,
            "unreviewed oversized Rust file",
        ):
            CHECK_RUSTFMT.validate_size_policy(
                self.root,
                [oversized],
                threshold_bytes=10,
                reviewed_exceptions={},
            )

    def test_rejects_a_missing_reviewed_exception(self) -> None:
        tracked = self.write_bytes("small.rs", 4)

        with self.assertRaisesRegex(
            CHECK_RUSTFMT.PolicyError,
            "reviewed exception is not tracked",
        ):
            CHECK_RUSTFMT.validate_size_policy(
                self.root,
                [tracked],
                threshold_bytes=10,
                reviewed_exceptions={"missing.rs": 20},
            )

    def test_rejects_a_reviewed_exception_that_is_no_longer_oversized(self) -> None:
        legacy = self.write_bytes("legacy.rs", 10)

        with self.assertRaisesRegex(
            CHECK_RUSTFMT.PolicyError,
            "reviewed exception is no longer oversized",
        ):
            CHECK_RUSTFMT.validate_size_policy(
                self.root,
                [legacy],
                threshold_bytes=10,
                reviewed_exceptions={"legacy.rs": 20},
            )

    def test_rejects_growth_past_a_reviewed_exception_ceiling(self) -> None:
        legacy = self.write_bytes("legacy.rs", 21)

        with self.assertRaisesRegex(
            CHECK_RUSTFMT.PolicyError,
            "exceeds its reviewed ceiling",
        ):
            CHECK_RUSTFMT.validate_size_policy(
                self.root,
                [legacy],
                threshold_bytes=10,
                reviewed_exceptions={"legacy.rs": 20},
            )


class RustfmtInvocationTests(unittest.TestCase):
    def test_runs_rustfmt_one_file_at_a_time_without_following_modules(self) -> None:
        commands: list[list[str]] = []

        def record(command: list[str], *, check: bool) -> subprocess.CompletedProcess:
            commands.append(command)
            return subprocess.CompletedProcess(command, 0)

        CHECK_RUSTFMT.run_rustfmt(
            REPOSITORY_ROOT,
            [Path("z.rs"), Path("nested/a.rs")],
            rustfmt="reviewed-rustfmt",
            runner=record,
        )

        self.assertEqual(
            [
                [
                    "reviewed-rustfmt",
                    "--check",
                    "--edition",
                    "2024",
                    "--config",
                    "skip_children=true",
                    str(REPOSITORY_ROOT / "nested/a.rs"),
                ],
                [
                    "reviewed-rustfmt",
                    "--check",
                    "--edition",
                    "2024",
                    "--config",
                    "skip_children=true",
                    str(REPOSITORY_ROOT / "z.rs"),
                ],
            ],
            commands,
        )


class CurrentRepositoryPolicyTests(unittest.TestCase):
    def test_reviewed_exception_inventory_matches_the_repository(self) -> None:
        tracked = CHECK_RUSTFMT.discover_tracked_rust_files(REPOSITORY_ROOT)

        targets = CHECK_RUSTFMT.validate_size_policy(REPOSITORY_ROOT, tracked)

        self.assertGreater(len(targets), 0)
        self.assertEqual(
            set(CHECK_RUSTFMT.REVIEWED_OVERSIZED_RUST_FILES),
            set(tracked) - set(targets),
        )


if __name__ == "__main__":
    unittest.main()
