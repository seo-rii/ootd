from __future__ import annotations

import unittest
from pathlib import Path


REPOSITORY_ROOT = Path(__file__).resolve().parents[2]
SCRIPT_PATH = REPOSITORY_ROOT / "tools" / "excel-oracle-win" / "scripts" / "run-suite.ps1"


class ExcelOracleSuiteScriptContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.script = SCRIPT_PATH.read_text(encoding="utf-8")

    def test_requires_explicit_tools_inputs_and_fresh_roots(self) -> None:
        self.assertIn("#requires -Version 7.4", self.script)
        self.assertIn("Set-StrictMode -Version Latest", self.script)
        for parameter in [
            "$RunnerPath",
            "$OracleCliPath",
            "$RunId",
            "$SuiteRoot",
            "$CaptureRoot",
            "$OutputRoot",
            "$TimeoutSeconds",
        ]:
            self.assertIn(parameter, self.script)
        self.assertIn("CaptureRoot must not exist before launch.", self.script)
        self.assertIn("OutputRoot must not exist before launch.", self.script)

    def test_validates_the_complete_suite_before_creating_capture_output(self) -> None:
        preflight = self.script.index("'capture-plan'")
        validate_exit = self.script.index("capture-plan failed")
        create_capture_root = self.script.index("CreateDirectory($captureRoot)")
        start_case = self.script.index("foreach ($case in $plan.cases)")

        self.assertLess(preflight, validate_exit)
        self.assertLess(validate_exit, create_capture_root)
        self.assertLess(create_capture_root, start_case)
        self.assertIn("$plan.caseCount -ne $plan.cases.Count", self.script)
        self.assertIn("$plan.caseCount -gt 4096", self.script)

    def test_runs_only_hash_verified_private_case_and_input_copies(self) -> None:
        self.assertIn("Get-FileHash -LiteralPath $sourceCasePath -Algorithm SHA256", self.script)
        self.assertIn("Get-FileHash -LiteralPath $sourceInputPath -Algorithm SHA256", self.script)
        self.assertIn("[System.IO.File]::Copy($sourceCasePath, $verifiedCasePath, $false)", self.script)
        self.assertIn("[System.IO.File]::Copy($sourceInputPath, $verifiedInputPath, $false)", self.script)
        self.assertIn("Get-FileHash -LiteralPath $verifiedCasePath -Algorithm SHA256", self.script)
        self.assertIn("Get-FileHash -LiteralPath $verifiedInputPath -Algorithm SHA256", self.script)
        self.assertIn("'-CasePath', $verifiedCasePath", self.script)
        self.assertIn("'-InputPath', $verifiedInputPath", self.script)
        self.assertIn("$caseStartInfo.RedirectStandardOutput = $true", self.script)
        self.assertIn("$caseStartInfo.RedirectStandardError = $true", self.script)
        self.assertIn("watchdog.stdout.log", self.script)
        self.assertIn("watchdog.stderr.log", self.script)

    def test_assembles_only_after_every_case_watchdog_succeeds(self) -> None:
        case_loop = self.script.index("foreach ($case in $plan.cases)")
        case_failure = self.script.index("case watchdog failed")
        assembly = self.script.index("'assemble-run'")

        self.assertLess(case_loop, case_failure)
        self.assertLess(case_failure, assembly)
        self.assertIn("'-NoProfile'", self.script)
        self.assertIn("'-NonInteractive'", self.script)
        self.assertIn("$caseProcess.ExitCode -ne 0", self.script)
        self.assertIn("foreach ($fragmentRoot in $fragmentRoots)", self.script)
        self.assertIn("assembly failed", self.script)
        self.assertIn("assembly_receipt.json", self.script)


if __name__ == "__main__":
    unittest.main()
