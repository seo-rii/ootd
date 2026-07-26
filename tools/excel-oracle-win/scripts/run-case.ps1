#requires -Version 7.4
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$RunnerPath,
    [Parameter(Mandatory = $true)][string]$RunId,
    [Parameter(Mandatory = $true)][string]$CasePath,
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$RunRoot,
    [Parameter(Mandatory = $true)][string]$Channel,
    [Parameter(Mandatory = $true)][string]$Locale,
    [Parameter(Mandatory = $true)][string]$Timezone,
    [ValidateRange(10, 3600)][int]$TimeoutSeconds = 300
)

$ErrorActionPreference = 'Stop'
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel Oracle requires a host without a pre-existing EXCEL.EXE process.'
}
if (Test-Path -LiteralPath $RunRoot) {
    throw 'RunRoot must not exist before launch.'
}

$temporaryStdout = Join-Path ([System.IO.Path]::GetTempPath()) ("ootd-oracle-{0}.stdout" -f [guid]::NewGuid().ToString('N'))
$temporaryStderr = Join-Path ([System.IO.Path]::GetTempPath()) ("ootd-oracle-{0}.stderr" -f [guid]::NewGuid().ToString('N'))
$startedAt = [DateTime]::UtcNow
$timedOut = $false
$forcedTermination = $false
$exitCode = 5
$runnerProcess = $null

try {
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = [System.IO.Path]::GetFullPath($RunnerPath)
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    foreach ($argument in @(
        'observe',
        '--run-id', $RunId,
        '--case', [System.IO.Path]::GetFullPath($CasePath),
        '--input', [System.IO.Path]::GetFullPath($InputPath),
        '--output-root', [System.IO.Path]::GetFullPath($RunRoot),
        '--observation', (Join-Path ([System.IO.Path]::GetFullPath($RunRoot)) 'observations/oracle.json'),
        '--channel', $Channel,
        '--locale', $Locale,
        '--timezone', $Timezone
    )) {
        [void]$startInfo.ArgumentList.Add($argument)
    }

    $runnerProcess = [System.Diagnostics.Process]::new()
    $runnerProcess.StartInfo = $startInfo
    if (-not $runnerProcess.Start()) {
        throw 'Failed to start the Excel Oracle runner.'
    }
    $stdoutTask = $runnerProcess.StandardOutput.ReadToEndAsync()
    $stderrTask = $runnerProcess.StandardError.ReadToEndAsync()
    if (-not $runnerProcess.WaitForExit($TimeoutSeconds * 1000)) {
        $timedOut = $true
        $runnerProcess.Kill($true)
        $runnerProcess.WaitForExit()
        $exitCode = 124
    } else {
        $exitCode = $runnerProcess.ExitCode
    }
    [System.IO.File]::WriteAllText($temporaryStdout, $stdoutTask.GetAwaiter().GetResult())
    [System.IO.File]::WriteAllText($temporaryStderr, $stderrTask.GetAwaiter().GetResult())

    if ($timedOut) {
        $ownedProcessPath = Join-Path $RunRoot 'manifest/owned_processes.json'
        if (Test-Path -LiteralPath $ownedProcessPath) {
            $owned = Get-Content -LiteralPath $ownedProcessPath -Raw | ConvertFrom-Json
            foreach ($ownedProcess in $owned.processes) {
                $excel = Get-Process -Id ([int]$ownedProcess.processId) -ErrorAction SilentlyContinue
                if ($null -ne $excel
                    -and $excel.ProcessName -eq 'EXCEL'
                    -and $excel.StartTime.ToUniversalTime().ToString('O') -eq [string]$ownedProcess.startTimeUtc) {
                    Stop-Process -Id $excel.Id -Force
                    $forcedTermination = $true
                }
            }
        }
    }
}
finally {
    $logs = Join-Path $RunRoot 'logs'
    $manifest = Join-Path $RunRoot 'manifest'
    [System.IO.Directory]::CreateDirectory($logs) | Out-Null
    [System.IO.Directory]::CreateDirectory($manifest) | Out-Null
    if (Test-Path -LiteralPath $temporaryStdout) {
        Move-Item -LiteralPath $temporaryStdout -Destination (Join-Path $logs 'stdout.log') -Force
    }
    if (Test-Path -LiteralPath $temporaryStderr) {
        Move-Item -LiteralPath $temporaryStderr -Destination (Join-Path $logs 'stderr.log') -Force
    }
    $status = [ordered]@{
        schemaVersion = 1
        startedAtUtc = $startedAt.ToString('O')
        finishedAtUtc = [DateTime]::UtcNow.ToString('O')
        exitCode = $exitCode
        timedOut = $timedOut
        forcedTermination = $forcedTermination
    }
    $statusPath = Join-Path $manifest 'launcher_status.json'
    $temporaryStatus = "$statusPath.tmp"
    [System.IO.File]::WriteAllText($temporaryStatus, ($status | ConvertTo-Json -Depth 4))
    Move-Item -LiteralPath $temporaryStatus -Destination $statusPath -Force
    if ($null -ne $runnerProcess) {
        $runnerProcess.Dispose()
    }
}

exit $exitCode
