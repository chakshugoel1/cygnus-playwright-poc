Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Section {
  param([string]$Message)
  Write-Host "`n=== $Message ===" -ForegroundColor Cyan
}

function Write-Step {
  param([string]$Message)
  Write-Host "[step] $Message" -ForegroundColor DarkCyan
}

function Write-Ok {
  param([string]$Message)
  Write-Host "[ok] $Message" -ForegroundColor Green
}

function Write-Warn {
  param([string]$Message)
  Write-Host "[warn] $Message" -ForegroundColor Yellow
}

function Write-Info {
  param([string]$Message)
  Write-Host "[info] $Message" -ForegroundColor Gray
}

function Assert-Command {
  param([string]$CommandName)
  $cmd = Get-Command $CommandName -ErrorAction SilentlyContinue
  if (-not $cmd) {
    throw "Required command not found in PATH: $CommandName"
  }
  return $true
}

function Test-CommandExists {
  param([string]$CommandName)
  return [bool](Get-Command $CommandName -ErrorAction SilentlyContinue)
}

function Invoke-External {
  param(
    [Parameter(Mandatory = $true)][string]$FilePath,
    [string[]]$ArgumentList = @(),
    [string]$WorkingDirectory = ''
  )

  $old = Get-Location
  try {
    if ($WorkingDirectory) {
      Push-Location $WorkingDirectory
    }
    & $FilePath @ArgumentList
    if ($LASTEXITCODE -ne 0) {
      throw "Command failed ($LASTEXITCODE): $FilePath $($ArgumentList -join ' ')"
    }
  }
  finally {
    if ($WorkingDirectory) {
      Pop-Location
    }
  }
}

function Ensure-Directory {
  param([Parameter(Mandatory = $true)][string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
  }
}

function Copy-TemplateIfMissing {
  param(
    [Parameter(Mandatory = $true)][string]$DestinationPath,
    [Parameter(Mandatory = $true)][string]$TemplatePath
  )

  if (Test-Path -LiteralPath $DestinationPath) {
    return $false
  }

  if (-not (Test-Path -LiteralPath $TemplatePath)) {
    throw "Template file not found: $TemplatePath"
  }

  Copy-Item -LiteralPath $TemplatePath -Destination $DestinationPath -Force
  return $true
}

function Read-PlainTextSecret {
  param([Parameter(Mandatory = $true)][string]$Prompt)

  $secure = Read-Host -AsSecureString -Prompt $Prompt
  $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
  try {
    return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
  }
  finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
  }
}

function Ensure-ToolWithWinget {
  param(
    [Parameter(Mandatory = $true)][string]$ToolName,
    [Parameter(Mandatory = $true)][string]$WingetId,
    [switch]$VerifyOnly
  )

  if (Test-CommandExists $ToolName) {
    Write-Ok "$ToolName is available"
    return
  }

  if ($VerifyOnly) {
    throw "$ToolName is missing"
  }

  if (-not (Test-CommandExists 'winget')) {
    throw "winget is required to install $ToolName automatically. Install $ToolName manually and rerun."
  }

  Write-Step "Installing $ToolName via winget ($WingetId)..."
  Invoke-External -FilePath 'winget' -ArgumentList @(
    'install', '--id', $WingetId, '-e',
    '--accept-source-agreements', '--accept-package-agreements'
  )

  if (-not (Test-CommandExists $ToolName)) {
    throw "$ToolName still not found after installation. Reopen PowerShell and rerun installer."
  }

  Write-Ok "$ToolName installed"
}

function Test-Yes {
  param([string]$Value)
  if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
  $v = $Value.Trim().ToLowerInvariant()
  return ($v -eq 'y' -or $v -eq 'yes' -or $v -eq 'true' -or $v -eq '1')
}
