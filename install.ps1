# Copy office-read and office-edit into a skills directory (overwrites previous copies).
# Usage: .\install.ps1 [destination]   default: ~/.claude/skills
$ErrorActionPreference = 'Stop'
$src = $PSScriptRoot
$dst = if ($args.Count -ge 1) { $args[0] } else { Join-Path $HOME '.claude\skills' }
New-Item -ItemType Directory -Force $dst | Out-Null
foreach ($s in 'office-read', 'office-edit') {
    $target = Join-Path $dst $s
    if (Test-Path $target) { Remove-Item -Recurse -Force $target }
    Copy-Item -Recurse (Join-Path $src $s) $target
    Get-ChildItem -Recurse -Directory -Filter '__pycache__' $target | Remove-Item -Recurse -Force
    Write-Host "installed $s -> $target"
}
Write-Host "Start a new agent session to pick up the skills."
