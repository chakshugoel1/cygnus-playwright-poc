[CmdletBinding()]
param(
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

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $repoRoot 'scripts\setup-lib.ps1')

$isVerify = $Mode -eq 'verify'
$isFast = $Mode -eq 'fast'
$isRepair = $Mode -eq 'repair'

Write-Section "Cygnus Playwright POC Installer ($Mode)"
Write-Info "Repo: $repoRoot"

$policy = Get-ExecutionPolicy -Scope CurrentUser
if ($policy -eq 'Restricted') {
  Write-Warn 'CurrentUser execution policy is Restricted.'
  Write-Warn 'Run once: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser'
}

Write-Section '1) Prerequisites'
Ensure-ToolWithWinget -ToolName 'node' -WingetId 'OpenJS.NodeJS.LTS' -VerifyOnly:$isVerify
Ensure-ToolWithWinget -ToolName 'git' -WingetId 'Git.Git' -VerifyOnly:$isVerify
Assert-Command 'npm' | Out-Null
Write-Ok "npm is available"

if (-not $isVerify -and -not $isFast) {
  Write-Step 'Install root dependencies (npm install)...'
  Invoke-External -FilePath 'npm' -ArgumentList @('install') -WorkingDirectory $repoRoot

  Write-Step 'Install Playwright browser cache (chromium)...'
  Invoke-External -FilePath 'npm' -ArgumentList @('run', 'install-browsers') -WorkingDirectory $repoRoot
} elseif ($isFast) {
  Write-Info 'Fast mode: skipped npm install and browser install.'
} else {
  Write-Info 'Verify mode: skipped npm install and browser install.'
}

Write-Section '2) Secrets setup'
$userHome = if ($env:USERPROFILE) { $env:USERPROFILE } else { $env:HOME }
if (-not $userHome) {
  throw 'Could not resolve user home directory from USERPROFILE/HOME.'
}

$secretsDir = Join-Path $userHome '.askme-poc-secrets'
$envFile = Join-Path $secretsDir '.env'
$spFile = Join-Path $secretsDir 'pbi-service-principal.json'
Ensure-Directory -Path $secretsDir
Write-Ok "Secrets directory: $secretsDir"

$envTemplate = Join-Path $repoRoot '.env.template'
$spTemplate = Join-Path $repoRoot 'pbi-service-principal.template.json'

if (Test-Path -LiteralPath $envFile) {
  Write-Ok '.env already exists'
} elseif (-not $isVerify) {
  $username = $env:E2E_USERNAME
  $password = $env:E2E_PASSWORD

  if (-not $username -and -not $NonInteractive) {
    $username = Read-Host 'Enter E2E_USERNAME (test account email)'
  }
  if (-not $password -and -not $NonInteractive) {
    $password = Read-PlainTextSecret 'Enter E2E_PASSWORD'
  }

  if ($username -and $password) {
    @(
      "E2E_USERNAME=$username"
      "E2E_PASSWORD=$password"
    ) | Set-Content -Path $envFile -Encoding UTF8
    Write-Ok '.env created from provided credentials'
  } else {
    Copy-TemplateIfMissing -DestinationPath $envFile -TemplatePath $envTemplate | Out-Null
    Write-Warn '.env template created. Fill credentials before running auth setup.'
  }
} else {
  Write-Warn '.env missing'
}

if (Test-Path -LiteralPath $spFile) {
  Write-Ok 'pbi-service-principal.json already exists'
} elseif (-not $isVerify) {
  Copy-TemplateIfMissing -DestinationPath $spFile -TemplatePath $spTemplate | Out-Null
  Write-Info 'Created pbi-service-principal.json template (OPTIONAL - the parity flow works without it, falling back to your signed-in user token). Only fill in tenant/client details if you also need the DAX-based main:run flow.'
} else {
  Write-Info 'pbi-service-principal.json not present (OPTIONAL - not needed for the parity flow).'
}

Write-Section '3) Auth session and parity check'
$authFile = Join-Path $secretsDir '.auth\cygnus.user.json'
$hasEnvCreds = $false
if (Test-Path -LiteralPath $envFile) {
  $envText = Get-Content -LiteralPath $envFile -Raw
  $hasUser = $envText -match '(?m)^E2E_USERNAME\s*=\s*\S+'
  $hasPass = $envText -match '(?m)^E2E_PASSWORD\s*=\s*\S+'
  $hasEnvCreds = ($hasUser -and $hasPass)
}

if (Test-Path -LiteralPath $authFile) {
  Write-Ok 'Saved browser auth session already exists'
} elseif ($SkipAuthSetup -or $isVerify) {
  Write-Warn 'Auth session missing (setup skipped in this mode).'
} elseif ($hasEnvCreds) {
  Write-Step 'Running one-time auth setup (npm run test:setup)...'
  Invoke-External -FilePath 'npm' -ArgumentList @('run', 'test:setup') -WorkingDirectory $repoRoot
  if (Test-Path -LiteralPath $authFile) {
    Write-Ok 'Auth session created'
  } else {
    Write-Warn 'Auth setup finished but session file was not found. Check setup output.'
  }
} else {
  Write-Warn 'Skipping auth setup: .env credentials are missing.'
}

if (-not $SkipParity -and -not $isVerify -and -not $isFast) {
  Write-Step 'Running parity smoke check (npm run parity)...'
  Invoke-External -FilePath 'npm' -ArgumentList @('run', 'parity') -WorkingDirectory $repoRoot
  Write-Ok 'Parity run completed'
} elseif ($SkipParity) {
  Write-Info 'Parity smoke check skipped by flag.'
} elseif ($isFast) {
  Write-Info 'Fast mode: parity smoke check skipped.'
}

$runDesktop = $false
if (-not $SkipDesktop -and -not $isVerify) {
  if ($InstallDesktop) {
    $runDesktop = $true
  } elseif (-not $NonInteractive) {
    $answer = Read-Host 'Install app-desktop dependencies too? (y/N)'
    $runDesktop = Test-Yes $answer
  }
}

if ($runDesktop) {
  Write-Section '4) Optional desktop app dependencies'
  Invoke-External -FilePath 'npm' -ArgumentList @('install') -WorkingDirectory (Join-Path $repoRoot 'app-desktop')
  Write-Ok 'app-desktop dependencies installed'
}

Write-Section 'Done'
Write-Info 'Installer finished.'
Write-Info 'Useful flags:'
Write-Info '  .\install.ps1 -Mode verify'
Write-Info '  .\install.ps1 -Mode fast -SkipParity'
Write-Info '  .\install.ps1 -Mode full -NonInteractive'
Write-Info '  .\install.ps1 -Mode repair -SkipParity'
