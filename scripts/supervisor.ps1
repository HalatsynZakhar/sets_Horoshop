[CmdletBinding()]
param(
    [string]$Branch = "main",
    [ValidateRange(1, 1439)]
    [int]$CheckIntervalMinutes = 5
)

$ErrorActionPreference = "Stop"
$AppDir = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$PythonExe = Join-Path $AppDir ".venv\Scripts\python.exe"
$ServerScript = Join-Path $AppDir "sets_server.py"
$LogsDir = Join-Path $AppDir "logs"
$PidFile = Join-Path $LogsDir "horoshop_sets.pid"
$SupervisorLog = Join-Path $LogsDir "supervisor.log"
$WorkerOutputLog = Join-Path $LogsDir "server-output.log"
$WorkerErrorLog = Join-Path $LogsDir "server-error.log"
$Worker = $null
$NextUpdateCheck = [DateTime]::MinValue

function Write-Log {
    param([string]$Message)

    New-Item -ItemType Directory -Path $LogsDir -Force | Out-Null
    $Line = "[{0:yyyy-MM-ddTHH:mm:ssK}] {1}" -f (Get-Date), $Message
    Add-Content -LiteralPath $SupervisorLog -Value $Line -Encoding UTF8
    Write-Output $Line

    if ((Test-Path $SupervisorLog) -and (Get-Item $SupervisorLog).Length -gt 2MB) {
        Get-Content -LiteralPath $SupervisorLog -Tail 1000 |
            Set-Content -LiteralPath $SupervisorLog -Encoding UTF8
    }
}

function Test-WorkerRunning {
    return $null -ne $script:Worker -and !$script:Worker.HasExited
}

function Restore-WorkerFromPidFile {
    if (!(Test-Path $PidFile)) { return }

    $savedPid = 0
    if (![int]::TryParse((Get-Content $PidFile -Raw).Trim(), [ref]$savedPid)) {
        Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
        return
    }
    $processInfo = Get-CimInstance Win32_Process -Filter "ProcessId = $savedPid" -ErrorAction SilentlyContinue
    if ($null -eq $processInfo -or [string]$processInfo.CommandLine -notlike "*$ServerScript*") {
        Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
        return
    }
    $script:Worker = Get-Process -Id $savedPid -ErrorAction SilentlyContinue
    if (Test-WorkerRunning) { Write-Log "Found running web server. PID: $savedPid" }
}

function Start-Worker {
    if (Test-WorkerRunning) { return }
    if (!(Test-Path $PythonExe)) { throw "Virtual environment Python was not found: $PythonExe" }
    if (!(Test-Path $ServerScript)) { throw "Server file was not found: $ServerScript" }

    New-Item -ItemType Directory -Path $LogsDir -Force | Out-Null
    $script:Worker = Start-Process `
        -FilePath $PythonExe `
        -ArgumentList "`"$ServerScript`"" `
        -WorkingDirectory $AppDir `
        -WindowStyle Hidden `
        -RedirectStandardOutput $WorkerOutputLog `
        -RedirectStandardError $WorkerErrorLog `
        -PassThru
    Set-Content -LiteralPath $PidFile -Value $script:Worker.Id -Encoding ASCII
    Start-Sleep -Seconds 2
    $script:Worker.Refresh()
    if ($script:Worker.HasExited) {
        $exitCode = $script:Worker.ExitCode
        $script:Worker = $null
        Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
        $details = if (Test-Path $WorkerErrorLog) {
            (Get-Content -LiteralPath $WorkerErrorLog -Tail 10) -join " | "
        } else { "No details in server-error.log." }
        throw "Web server stopped immediately (code $exitCode): $details"
    }
    Write-Log "Web server started. PID: $($script:Worker.Id)"
}

function Stop-Worker {
    if (!(Test-WorkerRunning)) { return }
    $workerId = $script:Worker.Id
    Stop-Process -Id $workerId -Force -ErrorAction SilentlyContinue
    $script:Worker.WaitForExit()
    $script:Worker = $null
    Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
    Write-Log "Web server stopped. PID: $workerId"
}

function Update-ProjectIfNeeded {
    if ((Get-Date) -lt $script:NextUpdateCheck) { return $false }
    $script:NextUpdateCheck = (Get-Date).AddMinutes($CheckIntervalMinutes)

    git -C $AppDir fetch origin $Branch --quiet
    if ($LASTEXITCODE -ne 0) { throw "Could not fetch origin/$Branch." }

    $changes = git -C $AppDir status --porcelain
    if ($LASTEXITCODE -ne 0) { throw "Could not read Git status." }
    if ($changes) {
        Write-Log "Local Git changes detected; automatic update was skipped."
        return $false
    }

    $local = (git -C $AppDir rev-parse HEAD).Trim()
    $remote = (git -C $AppDir rev-parse "origin/$Branch").Trim()
    if ($local -eq $remote) { return $false }

    Write-Log "Update found. Applying fast-forward update from origin/$Branch."
    Stop-Worker
    git -C $AppDir pull --ff-only origin $Branch
    if ($LASTEXITCODE -ne 0) { throw "Git fast-forward update failed." }
    & $PythonExe -m pip install -r (Join-Path $AppDir "requirements.txt")
    if ($LASTEXITCODE -ne 0) { throw "Could not install dependencies after update." }
    return $true
}

if (!(Get-Command git -ErrorAction SilentlyContinue)) { throw "Git was not found in PATH." }
if (!(Test-Path (Join-Path $AppDir ".git"))) { throw "$AppDir is not a Git repository." }

try {
    Restore-WorkerFromPidFile
    Write-Log "Supervisor started. Branch: $Branch; Git check every $CheckIntervalMinutes min."
    while ($true) {
        try {
            if (Update-ProjectIfNeeded) {
                Write-Log "Update completed. Restarting supervisor to load its new version."
                exit 75
            }
            Start-Worker
        }
        catch {
            Write-Log "Supervisor error: $($_.Exception.Message)"
        }
        Start-Sleep -Seconds 10
    }
}
finally {
    Stop-Worker
}
