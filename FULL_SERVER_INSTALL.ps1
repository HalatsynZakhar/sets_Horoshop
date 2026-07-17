[CmdletBinding()]
param(
    [string]$Repository = "https://github.com/HalatsynZakhar/sets_Horoshop.git",
    [string]$InstallDir = "C:\HoroshopSets",
    [string]$Branch = "main",
    [ValidateRange(1, 1439)]
    [int]$CheckIntervalMinutes = 5,
    [switch]$SkipConfigEdit,
    [switch]$NoPause
)

$ErrorActionPreference = "Stop"
$TaskName = "HoroshopSets"

function Assert-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    if (!$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Run PowerShell as Administrator."
    }
}

function Ensure-Command {
    param([string]$Name, [string]$WingetId)
    if (Get-Command $Name -ErrorAction SilentlyContinue) { return }
    if (!(Get-Command winget -ErrorAction SilentlyContinue)) {
        throw "$Name is not installed and winget is unavailable. Install it, then run this script again."
    }
    Write-Output "Installing $Name..."
    winget install --id $WingetId --exact --silent --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) { throw "winget could not install $Name." }
    Update-CurrentPath
}

function Update-CurrentPath {
    $machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$machinePath;$userPath"
}

function Get-PythonCommand {
    if (Get-Command py -ErrorAction SilentlyContinue) { return @("py", "-3") }
    if (Get-Command python -ErrorAction SilentlyContinue) { return @("python") }
    throw "Python 3 was not found after installation."
}

function Invoke-TaskCommand {
    param([string[]]$Arguments)
    & schtasks.exe @Arguments
    if ($LASTEXITCODE -ne 0) { throw "schtasks failed: $($Arguments -join ' ')" }
}

try {
    Assert-Administrator
    Ensure-Command -Name "git" -WingetId "Git.Git"
    Ensure-Command -Name "py" -WingetId "Python.Python.3.13"
    Update-CurrentPath

    $safeDirectory = $InstallDir.Replace("\", "/")
    $knownSafeDirectories = @(& git config --system --get-all safe.directory 2>$null)
    if ($knownSafeDirectories -notcontains $safeDirectory) {
        & git config --system --add safe.directory $safeDirectory
        if ($LASTEXITCODE -ne 0) { throw "Could not mark $InstallDir as a safe Git directory." }
    }

    if (Test-Path (Join-Path $InstallDir ".git")) {
        git -C $InstallDir fetch origin $Branch
        if ($LASTEXITCODE -ne 0) { throw "Could not fetch origin/$Branch." }
        git -C $InstallDir reset --hard "origin/$Branch"
        if ($LASTEXITCODE -ne 0) { throw "Could not reset the server copy to origin/$Branch." }
        git -C $InstallDir clean -fd
        if ($LASTEXITCODE -ne 0) { throw "Could not clean untracked deployment files." }
    }
    elseif (Test-Path $InstallDir) {
        throw "$InstallDir exists but is not a Git repository."
    }
    else {
        git clone --branch $Branch $Repository $InstallDir
        if ($LASTEXITCODE -ne 0) { throw "Could not clone the repository." }
    }

    $configPath = Join-Path $InstallDir "config.json"
    if (!(Test-Path $configPath)) {
        Copy-Item (Join-Path $InstallDir "config.example.json") $configPath
    }
    if (!$SkipConfigEdit) {
        Write-Host "Set the real Horoshop domain in config.json, then close Notepad." -ForegroundColor Yellow
        Start-Process notepad.exe -ArgumentList "`"$configPath`"" -Wait
    }
    $config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if (![string]$config.horoshop.domain -or [string]$config.horoshop.domain -match "example\.com") {
        throw "Set the real horoshop.domain in $configPath before continuing."
    }

    Set-Location $InstallDir
    $venvDir = Join-Path $InstallDir ".venv"
    $venvPython = Join-Path $InstallDir ".venv\Scripts\python.exe"
    if (!(Test-Path $venvPython)) {
        $python = Get-PythonCommand
        if ($python.Count -eq 2) {
            & $python[0] $python[1] -m venv $venvDir
        }
        else {
            & $python[0] -m venv $venvDir
        }
        if ($LASTEXITCODE -ne 0) { throw "Could not create virtual environment." }
    }
    & $venvPython -m pip install --upgrade pip
    & $venvPython -m pip install -r (Join-Path $InstallDir "requirements.txt")
    if ($LASTEXITCODE -ne 0) { throw "Could not install Python dependencies." }
    & $venvPython -m py_compile (Join-Path $InstallDir "horoshop_sets.py") (Join-Path $InstallDir "sets_server.py")
    if ($LASTEXITCODE -ne 0) { throw "Python syntax check failed." }

    $port = if ($config.server.port) { [int]$config.server.port } else { 8093 }
    $ruleName = "HoroshopSets-$port"
    Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue | Remove-NetFirewallRule
    New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort $port | Out-Null

    $launcher = Join-Path $InstallDir "scripts\auto_update.bat"
    $taskAction = "`"$launcher`" `"$Branch`" `"$CheckIntervalMinutes`""
    Invoke-TaskCommand -Arguments @("/Create", "/TN", $TaskName, "/SC", "ONSTART", "/TR", $taskAction, "/RU", "SYSTEM", "/RL", "HIGHEST", "/F")
    Invoke-TaskCommand -Arguments @("/Run", "/TN", $TaskName)

    Write-Host "Installation completed. Web panel: http://<server>:$port" -ForegroundColor Green
}
catch {
    Write-Host "Installation failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    if (!$NoPause) { Read-Host "Press Enter to close" }
}
