$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop 'Cygnus Desktop Runner.lnk'
$targetPath = Join-Path $scriptDir 'launch-desktop.bat'

$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "$env:WINDIR\System32\cmd.exe"
$shortcut.Arguments = "/c `"$targetPath`""
$shortcut.WorkingDirectory = $scriptDir
$shortcut.WindowStyle = 1
$shortcut.Description = 'Launch Cygnus Desktop Runner'
$shortcut.Save()

Write-Host "Shortcut created: $shortcutPath"
