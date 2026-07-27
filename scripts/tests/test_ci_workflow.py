from __future__ import annotations

import re
import unittest
from pathlib import Path


REPOSITORY_ROOT = Path(__file__).resolve().parents[2]
WORKFLOW_PATH = REPOSITORY_ROOT / ".github" / "workflows" / "ci.yml"


def workflow_job(workflow: str, name: str) -> str:
    marker = f"  {name}:\n"
    start = workflow.find(marker)
    if start < 0:
        raise AssertionError(f"CI job is missing: {name}")
    following_job = re.search(r"^  [A-Za-z0-9_-]+:\n", workflow[start + len(marker) :], re.MULTILINE)
    if following_job is None:
        return workflow[start:]
    return workflow[start : start + len(marker) + following_job.start()]


class CiWorkflowContractTests(unittest.TestCase):
    def test_quality_job_gates_bounded_formatting_and_foundational_clippy(self) -> None:
        workflow = WORKFLOW_PATH.read_text(encoding="utf-8")

        quality = workflow_job(workflow, "quality")

        self.assertIn("rustfmt --component clippy", quality)
        self.assertIn("python3 -m unittest scripts/tests/test_", quality)
        self.assertIn("./scripts/check_rustfmt.py", quality)
        self.assertIn("cargo clippy", quality)
        for package in [
            "office-idl",
            "office-common",
            "office-codegen",
            "office-capture",
            "office-opc",
            "excel-model",
        ]:
            self.assertIn(f"-p {package}", quality)
        self.assertIn("--all-targets --all-features -- -D warnings", quality)


if __name__ == "__main__":
    unittest.main()
