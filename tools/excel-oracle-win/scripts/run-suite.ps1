#requires -Version 7.4
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$RunnerPath,
    [Parameter(Mandatory = $true)][string]$OracleCliPath,
    [Parameter(Mandatory = $true)][string]$RunId,
    [Parameter(Mandatory = $true)][string]$SuiteRoot,
    [Parameter(Mandatory = $true)][string]$CaptureRoot,
    [Parameter(Mandatory = $true)][string]$OutputRoot,
    [ValidateRange(10, 3600)][int]$TimeoutSeconds = 300
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false

if ($RunId -notmatch '^[A-Za-z0-9._-]+$') {
    throw 'RunId must be a trimmed ASCII identifier.'
}

$runnerPath = [System.IO.Path]::GetFullPath($RunnerPath)
$oracleCliPath = [System.IO.Path]::GetFullPath($OracleCliPath)
$suiteRoot = [System.IO.Path]::GetFullPath($SuiteRoot)
$captureRoot = [System.IO.Path]::GetFullPath($CaptureRoot)
$outputRoot = [System.IO.Path]::GetFullPath($OutputRoot)
$runCaseScript = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'run-case.ps1'))

if (-not (Test-Path -LiteralPath $runnerPath -PathType Leaf)) {
    throw 'RunnerPath must identify the Excel Oracle runner executable.'
}
if (-not (Test-Path -LiteralPath $oracleCliPath -PathType Leaf)) {
    throw 'OracleCliPath must identify the excel-oracle executable.'
}
if (-not (Test-Path -LiteralPath $suiteRoot -PathType Container)) {
    throw 'SuiteRoot must identify an existing directory.'
}
if (-not (Test-Path -LiteralPath $runCaseScript -PathType Leaf)) {
    throw 'The case watchdog script is missing.'
}
if (Test-Path -LiteralPath $captureRoot) {
    throw 'CaptureRoot must not exist before launch.'
}
if (Test-Path -LiteralPath $outputRoot) {
    throw 'OutputRoot must not exist before launch.'
}
if ([string]::Equals($captureRoot, $outputRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'CaptureRoot and OutputRoot must be different paths.'
}
$outputParent = [System.IO.Path]::GetDirectoryName($outputRoot)
if ([string]::IsNullOrWhiteSpace($outputParent)
    -or -not (Test-Path -LiteralPath $outputParent -PathType Container)) {
    throw 'OutputRoot parent must be an existing directory.'
}

$planStartInfo = [System.Diagnostics.ProcessStartInfo]::new()
$planStartInfo.FileName = $oracleCliPath
$planStartInfo.UseShellExecute = $false
$planStartInfo.RedirectStandardOutput = $true
$planStartInfo.RedirectStandardError = $true
foreach ($argument in @('capture-plan', '--suite-root', $suiteRoot)) {
    [void]$planStartInfo.ArgumentList.Add($argument)
}
$planProcess = [System.Diagnostics.Process]::new()
$planProcess.StartInfo = $planStartInfo
if (-not $planProcess.Start()) {
    throw 'Failed to start the capture-plan preflight.'
}
$planStdoutTask = $planProcess.StandardOutput.ReadToEndAsync()
$planStderrTask = $planProcess.StandardError.ReadToEndAsync()
if (-not $planProcess.WaitForExit($TimeoutSeconds * 1000)) {
    $planProcess.Kill($true)
    $planProcess.WaitForExit()
    $planProcess.Dispose()
    throw 'capture-plan timed out.'
}
$planStdout = $planStdoutTask.GetAwaiter().GetResult()
$planStderr = $planStderrTask.GetAwaiter().GetResult()
$planExitCode = $planProcess.ExitCode
$planProcess.Dispose()
if ($planExitCode -ne 0) {
    throw "capture-plan failed with exit code $planExitCode`: $($planStderr.Trim())"
}
if ([string]::IsNullOrWhiteSpace($planStdout)) {
    throw 'capture-plan returned an empty document.'
}
try {
    $plan = $planStdout | ConvertFrom-Json
}
catch {
    throw "capture-plan returned invalid JSON: $($_.Exception.Message)"
}
if ($plan.schemaVersion -ne 1) {
    throw 'capture-plan schemaVersion was not supported.'
}
if ($plan.caseCount -ne $plan.cases.Count) {
    throw 'capture-plan caseCount did not match its cases.'
}
if ($plan.caseCount -lt 1 -or $plan.caseCount -gt 4096) {
    throw 'capture-plan case count must be between 1 and 4096.'
}
if ($plan.expectedEngine.kind -ne 'excel') {
    throw 'capture-plan expectedEngine must identify desktop Excel.'
}
foreach ($profileValue in @(
    $plan.profileId,
    $plan.expectedEngine.channel,
    $plan.expectedEngine.locale,
    $plan.expectedEngine.timezone
)) {
    if ([string]::IsNullOrWhiteSpace([string]$profileValue)) {
        throw 'capture-plan returned incomplete profile metadata.'
    }
}
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel Oracle requires a host without a pre-existing EXCEL.EXE process.'
}

[System.IO.Directory]::CreateDirectory($captureRoot) | Out-Null
$captureManifestRoot = Join-Path $captureRoot 'manifest'
[System.IO.Directory]::CreateDirectory($captureManifestRoot) | Out-Null
$planPath = Join-Path $captureManifestRoot 'capture_plan.json'
$temporaryPlanPath = "$planPath.tmp"
[System.IO.File]::WriteAllText(
    $temporaryPlanPath,
    $planStdout,
    [System.Text.UTF8Encoding]::new($false))
Move-Item -LiteralPath $temporaryPlanPath -Destination $planPath

$suitePrefix = $suiteRoot.TrimEnd(
    [char[]]@([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
) + [System.IO.Path]::DirectorySeparatorChar
$fragmentRoots = [System.Collections.Generic.List[string]]::new()
$startedAt = [DateTime]::UtcNow
$completedCaseCount = 0
$suiteStatus = 'failed'
$failureMessage = $null
$assemblyStdout = $null

try {
    foreach ($case in $plan.cases) {
        $caseId = [string]$case.caseId
        if ($caseId -notmatch '^[A-Za-z0-9._-]+$') {
            throw 'capture-plan returned an invalid case identifier.'
        }
        $sourceCasePath = [System.IO.Path]::GetFullPath(
            [System.IO.Path]::Combine($suiteRoot, [string]$case.casePath))
        $sourceInputPath = [System.IO.Path]::GetFullPath(
            [System.IO.Path]::Combine($suiteRoot, [string]$case.inputPath))
        if (-not $sourceCasePath.StartsWith($suitePrefix, [StringComparison]::OrdinalIgnoreCase)
            -or -not $sourceInputPath.StartsWith($suitePrefix, [StringComparison]::OrdinalIgnoreCase)) {
            throw "capture-plan path escaped SuiteRoot for case $caseId."
        }
        if (-not (Test-Path -LiteralPath $sourceCasePath -PathType Leaf)
            -or -not (Test-Path -LiteralPath $sourceInputPath -PathType Leaf)) {
            throw "capture-plan artifact disappeared for case $caseId."
        }

        $sourceCaseHash = (Get-FileHash -LiteralPath $sourceCasePath -Algorithm SHA256).Hash.ToLowerInvariant()
        $sourceInputHash = (Get-FileHash -LiteralPath $sourceInputPath -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($sourceCaseHash -cne [string]$case.caseSha256
            -or $sourceInputHash -cne [string]$case.inputSha256) {
            throw "capture-plan artifact changed before copy for case $caseId."
        }

        $verifiedRoot = Join-Path $captureRoot ("verified/{0}" -f $caseId)
        [System.IO.Directory]::CreateDirectory($verifiedRoot) | Out-Null
        $verifiedCasePath = Join-Path $verifiedRoot 'case.json'
        $inputExtension = [System.IO.Path]::GetExtension($sourceInputPath).ToLowerInvariant()
        $verifiedInputPath = Join-Path $verifiedRoot ("input{0}" -f $inputExtension)
        [System.IO.File]::Copy($sourceCasePath, $verifiedCasePath, $false)
        [System.IO.File]::Copy($sourceInputPath, $verifiedInputPath, $false)
        $verifiedCaseHash = (Get-FileHash -LiteralPath $verifiedCasePath -Algorithm SHA256).Hash.ToLowerInvariant()
        $verifiedInputHash = (Get-FileHash -LiteralPath $verifiedInputPath -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($verifiedCaseHash -cne [string]$case.caseSha256
            -or $verifiedInputHash -cne [string]$case.inputSha256) {
            throw "verified artifact copy changed for case $caseId."
        }

        $fragmentRoot = Join-Path $captureRoot ("fragments/{0}" -f $caseId)
        $caseStartInfo = [System.Diagnostics.ProcessStartInfo]::new()
        $caseStartInfo.FileName = [System.Environment]::ProcessPath
        $caseStartInfo.UseShellExecute = $false
        $caseStartInfo.RedirectStandardOutput = $true
        $caseStartInfo.RedirectStandardError = $true
        foreach ($argument in @(
            '-NoLogo',
            '-NoProfile',
            '-NonInteractive',
            '-File', $runCaseScript,
            '-RunnerPath', $runnerPath,
            '-RunId', $RunId,
            '-CasePath', $verifiedCasePath,
            '-InputPath', $verifiedInputPath,
            '-RunRoot', $fragmentRoot,
            '-Channel', [string]$plan.expectedEngine.channel,
            '-Locale', [string]$plan.expectedEngine.locale,
            '-Timezone', [string]$plan.expectedEngine.timezone,
            '-TimeoutSeconds', [string]$TimeoutSeconds
        )) {
            [void]$caseStartInfo.ArgumentList.Add($argument)
        }
        $caseProcess = [System.Diagnostics.Process]::new()
        $caseProcess.StartInfo = $caseStartInfo
        if (-not $caseProcess.Start()) {
            throw "Failed to start case watchdog for $caseId."
        }
        $caseStdoutTask = $caseProcess.StandardOutput.ReadToEndAsync()
        $caseStderrTask = $caseProcess.StandardError.ReadToEndAsync()
        $caseProcess.WaitForExit()
        $caseStdout = $caseStdoutTask.GetAwaiter().GetResult()
        $caseStderr = $caseStderrTask.GetAwaiter().GetResult()
        $caseLauncherLogRoot = Join-Path $captureRoot ("logs/{0}" -f $caseId)
        [System.IO.Directory]::CreateDirectory($caseLauncherLogRoot) | Out-Null
        [System.IO.File]::WriteAllText(
            (Join-Path $caseLauncherLogRoot 'watchdog.stdout.log'),
            $caseStdout,
            [System.Text.UTF8Encoding]::new($false))
        [System.IO.File]::WriteAllText(
            (Join-Path $caseLauncherLogRoot 'watchdog.stderr.log'),
            $caseStderr,
            [System.Text.UTF8Encoding]::new($false))
        if ($caseProcess.ExitCode -ne 0) {
            $caseExitCode = $caseProcess.ExitCode
            $caseProcess.Dispose()
            throw "case watchdog failed for $caseId with exit code $caseExitCode."
        }
        $caseProcess.Dispose()
        $fragmentRoots.Add($fragmentRoot)
        $completedCaseCount++
    }

    $assemblyStartInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $assemblyStartInfo.FileName = $oracleCliPath
    $assemblyStartInfo.UseShellExecute = $false
    $assemblyStartInfo.RedirectStandardOutput = $true
    $assemblyStartInfo.RedirectStandardError = $true
    foreach ($argument in @('assemble-run', '--suite-root', $suiteRoot)) {
        [void]$assemblyStartInfo.ArgumentList.Add($argument)
    }
    foreach ($fragmentRoot in $fragmentRoots) {
        [void]$assemblyStartInfo.ArgumentList.Add('--fragment-root')
        [void]$assemblyStartInfo.ArgumentList.Add($fragmentRoot)
    }
    [void]$assemblyStartInfo.ArgumentList.Add('--output-root')
    [void]$assemblyStartInfo.ArgumentList.Add($outputRoot)

    $assemblyProcess = [System.Diagnostics.Process]::new()
    $assemblyProcess.StartInfo = $assemblyStartInfo
    if (-not $assemblyProcess.Start()) {
        throw 'Failed to start suite assembly.'
    }
    $assemblyStdoutTask = $assemblyProcess.StandardOutput.ReadToEndAsync()
    $assemblyStderrTask = $assemblyProcess.StandardError.ReadToEndAsync()
    $assemblyProcess.WaitForExit()
    $assemblyStdout = $assemblyStdoutTask.GetAwaiter().GetResult()
    $assemblyStderr = $assemblyStderrTask.GetAwaiter().GetResult()
    $assemblyExitCode = $assemblyProcess.ExitCode
    $assemblyProcess.Dispose()
    if ($assemblyExitCode -ne 0) {
        throw "assembly failed with exit code $assemblyExitCode`: $($assemblyStderr.Trim())"
    }
    try {
        [void]($assemblyStdout | ConvertFrom-Json)
    }
    catch {
        throw "assembly returned invalid JSON: $($_.Exception.Message)"
    }
    $assemblyReceiptPath = Join-Path $captureManifestRoot 'assembly_receipt.json'
    $temporaryAssemblyReceiptPath = "$assemblyReceiptPath.tmp"
    [System.IO.File]::WriteAllText(
        $temporaryAssemblyReceiptPath,
        $assemblyStdout,
        [System.Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temporaryAssemblyReceiptPath -Destination $assemblyReceiptPath
    $suiteStatus = 'completed'
}
catch {
    $failureMessage = $_.Exception.Message
    throw
}
finally {
    $status = [ordered]@{
        schemaVersion = 1
        runId = $RunId
        suiteId = [string]$plan.suiteId
        startedAtUtc = $startedAt.ToString('O')
        finishedAtUtc = [DateTime]::UtcNow.ToString('O')
        plannedCaseCount = [int]$plan.caseCount
        completedCaseCount = $completedCaseCount
        status = $suiteStatus
        message = $failureMessage
    }
    $statusPath = Join-Path $captureManifestRoot 'suite_status.json'
    $temporaryStatusPath = "$statusPath.tmp"
    [System.IO.File]::WriteAllText(
        $temporaryStatusPath,
        ($status | ConvertTo-Json -Depth 4),
        [System.Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temporaryStatusPath -Destination $statusPath
}

Write-Output ($assemblyStdout.TrimEnd())
