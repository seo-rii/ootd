# Desktop Excel Behavioral Oracle

This tool executes one versioned behavioral case in one isolated desktop Excel process. It is
separate from `office-capture`: the capture crate acquires OM metadata, while this runner observes
runtime values, errors, mutations, save/reopen behavior, and repair evidence.

The runner targets .NET 10 without Office PIA or NuGet dependencies. Its contracts and fake-backed
lifecycle tests build on Linux; COM execution requires a dedicated Windows host with desktop Excel.

## Build And Test

```powershell
dotnet run --project tests/ExcelOracle.Contracts.Tests/ExcelOracle.Contracts.Tests.csproj
dotnet run --project tests/ExcelOracle.Win.Tests/ExcelOracle.Win.Tests.csproj
dotnet publish src/ExcelOracle.Win/ExcelOracle.Win.csproj -c Release -r win-x64 --self-contained false
```

Run a case through the bounded PowerShell watchdog:

```powershell
pwsh ./scripts/run-case.ps1 `
  -RunnerPath ./src/ExcelOracle.Win/bin/Release/net10.0/ExcelOracle.Win.exe `
  -RunId application-name-a `
  -CasePath C:/ootd/oracle/cases/application.name.json `
  -InputPath C:/ootd/oracle/inputs/application-name.xlsx `
  -RunRoot C:/ootd/oracle/runs/application-name-a `
  -Channel Current `
  -Locale en-US `
  -Timezone UTC
```

`RunRoot` must not exist. The launcher refuses to start while another `EXCEL.EXE` exists, applies a
hard timeout, and only force-stops Excel PIDs whose process id and start time were recorded by the
runner, including a separate normal-open verification process. A forced termination is an
infrastructure failure, never a semantic pass.

## Safety Boundary

- Run only SHA-256-pinned corpus inputs on an offline, disposable Windows profile.
- Only `.xlsx` and `.xltx` are accepted. VBA projects, macro sheets, and dialog sheets are rejected
  before Excel activation.
- Excel is hidden with alerts, events, link updates, and macros disabled. Workbooks open with
  `UpdateLinks=0` and `CorruptLoad=xlNormalLoad`; there is no automatic repair fallback.
- Each case records every Excel process it activates. All COM objects are released in reverse order after
  `Workbook.Close(false)` and `Application.Quit()`.
- Network isolation is a host-provisioning requirement. `AutomationSecurity=ForceDisable` alone is
  not a complete sandbox, particularly for legacy XLM behavior.

The current repository has not executed this runner against real Excel yet. Do not label a case
Oracle-verified until its pinned observation and run manifest have been captured twice on the same
declared host profile.
