[CmdletBinding()]
param(
  [ValidateSet('full', 'fast', 'verify', 'repair')]
  [string]$Mode = 'full',
  # Both OFF by default: the installer prepares everything (deps, browser,
  # desktop app, shortcut, saved credentials if provided) but does NOT open a
  # real browser to log in or run a ~10-20 min two-report parity check on its
  # own. Pass these to opt into full automation (useful for unattended/CI
  # runs where nobody is going to double-click the desktop shortcut).
  [switch]$RunAuthSetup,
  [switch]$RunParity,
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

  # Always installed, not optional - the desktop shortcut created at the end
  # of this script needs app-desktop/node_modules (specifically Electron) to
  # already be there, and a new user should never have to do this themselves.
  Write-Step 'Install desktop app dependencies (npm install)...'
  Invoke-External -FilePath 'npm' -ArgumentList @('install') -WorkingDirectory (Join-Path $repoRoot 'app-desktop')
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

$secretsDir = Join-Path $userHome 'Power_BI_report_validation_credentials'
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
    Write-Info '.env template created. You can also enter credentials later from the desktop app.'
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

Write-Section '3) Auth session'
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
} elseif ($isVerify) {
  Write-Warn 'Auth session missing (setup skipped in this mode).'
} elseif ($hasEnvCreds -and $RunAuthSetup) {
  Write-Step 'Running one-time auth setup (npm run test:setup)...'
  Invoke-External -FilePath 'npm' -ArgumentList @('run', 'test:setup') -WorkingDirectory $repoRoot
  if (Test-Path -LiteralPath $authFile) {
    Write-Ok 'Auth session created'
  } else {
    Write-Warn 'Auth setup finished but session file was not found. Check setup output.'
  }
} elseif ($hasEnvCreds) {
  # Not run automatically by default - this opens a REAL browser and signs in
  # for real. The user should choose when that happens (from the desktop app,
  # or via -RunAuthSetup) rather than have it fire unannounced during install.
  Write-Ok 'Credentials saved. Sign in from the desktop app when you are ready (or pass -RunAuthSetup here).'
} else {
  Write-Info 'No credentials yet - enter them from the desktop app when you open it.'
}

$hasAuthSession = Test-Path -LiteralPath $authFile

Write-Section '4) Parity smoke check'
if ($isVerify -or $isFast) {
  Write-Info 'Parity smoke check skipped in this mode.'
} elseif ($RunParity -and $hasAuthSession) {
  Write-Step 'Running parity smoke check (npm run parity)...'
  Invoke-External -FilePath 'npm' -ArgumentList @('run', 'parity') -WorkingDirectory $repoRoot
  Write-Ok 'Parity run completed'
} elseif ($RunParity) {
  Write-Warn 'Parity smoke check skipped: no saved auth session. Sign in first (desktop app, or -RunAuthSetup).'
} else {
  Write-Info 'Not run automatically. Open the desktop app to enter your report details and run Parity from there.'
}

Write-Section '5) Desktop shortcut'
$launchBat = Join-Path $repoRoot 'app-desktop\launch-desktop.bat'
$shortcutPath = $null
if (-not $isVerify) {
  if (-not (Test-Path -LiteralPath $launchBat)) {
    Write-Warn "launch-desktop.bat not found at $launchBat - skipping shortcut creation."
  } else {
    $defaultShortcutDir = [Environment]::GetFolderPath('Desktop')
    $shortcutDirInput = if ($NonInteractive) { '' } else { Read-Host "Where should the shortcut go? [$defaultShortcutDir]" }
    $shortcutDir = if ([string]::IsNullOrWhiteSpace($shortcutDirInput)) { $defaultShortcutDir } else { $shortcutDirInput }
    Ensure-Directory -Path $shortcutDir

    $shortcutName = 'Power BI Report Validator.lnk'
    $shortcutPath = Join-Path $shortcutDir $shortcutName

    $wshShell = New-Object -ComObject WScript.Shell
    $shortcut = $wshShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $launchBat
    $shortcut.WorkingDirectory = Split-Path -Parent $launchBat
    $shortcut.Description = 'Power BI Report Validator'
    $shortcut.Save()

    Write-Ok "Shortcut created: $shortcutPath"
  }
} else {
  Write-Info 'Shortcut creation skipped in verify mode.'
}

Write-Section 'Done'
Write-Info 'Installer finished - dependencies, browser, desktop app, and credentials are set up.'
Write-Host ''
if ($shortcutPath) {
  $shortcutFileName = Split-Path -Leaf $shortcutPath
  $shortcutFolder = Split-Path -Parent $shortcutPath
  Write-Info "Double-click `"$shortcutFileName`" in `"$shortcutFolder`" to get started."
} else {
  Write-Info 'Double-click app-desktop\launch-desktop.bat to get started.'
}
Write-Host ''
Write-Info 'Useful flags:'
Write-Info '  .\install.ps1 -Mode verify'
Write-Info '  .\install.ps1 -Mode fast'
Write-Info '  .\install.ps1 -Mode full -NonInteractive'
Write-Info '  .\install.ps1 -RunAuthSetup -RunParity   (opt into full automation, no desktop app needed)'
