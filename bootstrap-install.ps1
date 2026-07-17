[CmdletBinding()]
param(
  [string]$RepoUrl = 'https://github.com/chakshugoel1/cygnus-playwright-poc.git',
  # Empty = auto-detect the remote's actual default branch below. Passing
  # -Branch explicitly always wins. (A hardcoded wrong default here is exactly
  # what silently breaks a fresh-machine clone if the repo's default branch is
  # ever renamed - e.g. main vs master - so this only hardcodes a fallback,
  # never the primary path.)
  [string]$Branch = '',
  [string]$Destination = "$env:USERPROFILE\cygnus-playwright-poc",
  [ValidateSet('full', 'fast', 'verify', 'repair')]
  [string]$Mode = 'full',
  [switch]$SkipDesktop,
  [switch]$InstallDesktop,
  [switch]$SkipAuthSetup,
  [switch]$SkipParity,
  [switch]$NonInteractive
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Section {
  param([string]$Message)
  Write-Host "`n=== $Message ===" -ForegroundColor Cyan
}

function Write-Info {
  param([string]$Message)
  Write-Host "[info] $Message" -ForegroundColor Gray
}

function Write-Ok {
  param([string]$Message)
  Write-Host "[ok] $Message" -ForegroundColor Green
}

function Assert-Command {
  param([string]$Name)
  if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
    throw "Required command missing: $Name"
  }
}

function Invoke-External {
  param(
    [Parameter(Mandatory = $true)][string]$FilePath,
    [string[]]$ArgumentList = @(),
    [string]$WorkingDirectory = ''
  )

  try {
    if ($WorkingDirectory) { Push-Location $WorkingDirectory }
    & $FilePath @ArgumentList
    if ($LASTEXITCODE -ne 0) {
      throw "Command failed ($LASTEXITCODE): $FilePath $($ArgumentList -join ' ')"
    }
  }
  finally {
    if ($WorkingDirectory) { Pop-Location }
  }
}

Write-Section 'Cygnus bootstrap installer'
Assert-Command 'git'
Assert-Command 'powershell'

if (-not $Branch) {
  Write-Info 'No -Branch specified - detecting the repository default branch...'
  $Branch = 'master' # fallback used only if detection below fails
  try {
    $headLine = (& git ls-remote --symref $RepoUrl HEAD 2>$null) | Where-Object { $_ -match '^ref:' } | Select-Object -First 1
    if ($headLine -match 'refs/heads/(\S+)') {
      $Branch = $Matches[1]
    }
  } catch {
    Write-Info "Could not auto-detect the default branch ($($_.Exception.Message)) - using fallback '$Branch'."
  }
  Write-Info "Using branch: $Branch"
}

if (-not (Test-Path -LiteralPath $Destination)) {
  Write-Info "Cloning repository to: $Destination"
  Invoke-External -FilePath 'git' -ArgumentList @('clone', '--branch', $Branch, $RepoUrl, $Destination)
  Write-Ok 'Repository cloned'
} else {
  Write-Info "Repository exists. Pulling latest in: $Destination"
  Invoke-External -FilePath 'git' -ArgumentList @('fetch', 'origin') -WorkingDirectory $Destination
  Invoke-External -FilePath 'git' -ArgumentList @('checkout', $Branch) -WorkingDirectory $Destination
  Invoke-External -FilePath 'git' -ArgumentList @('pull', '--ff-only', 'origin', $Branch) -WorkingDirectory $Destination
  Write-Ok 'Repository updated'
}

$installScript = Join-Path $Destination 'install.ps1'
if (-not (Test-Path -LiteralPath $installScript)) {
  throw "Installer not found: $installScript"
}

$installerArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $installScript, '-Mode', $Mode)
if ($SkipDesktop) { $installerArgs += '-SkipDesktop' }
if ($InstallDesktop) { $installerArgs += '-InstallDesktop' }
if ($SkipAuthSetup) { $installerArgs += '-SkipAuthSetup' }
if ($SkipParity) { $installerArgs += '-SkipParity' }
if ($NonInteractive) { $installerArgs += '-NonInteractive' }

Write-Section 'Running project installer'
Invoke-External -FilePath 'powershell' -ArgumentList $installerArgs

Write-Section 'Done'
Write-Ok 'Machine bootstrap + project installer completed successfully.'
