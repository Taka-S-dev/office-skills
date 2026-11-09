# Copy office-read and office-edit into skills directories (overwrites previous copies).
# Usage: .\install.ps1 [destination]
#   default: ~/.claude/skills, and ~/.copilot/skills when GitHub Copilot's folder exists
$ErrorActionPreference = 'Stop'
$src = $PSScriptRoot
if ($args.Count -ge 1) {
    $dsts = @($args[0])
} else {
    $dsts = @(Join-Path $HOME '.claude\skills')
    if (Test-Path (Join-Path $HOME '.copilot')) { $dsts += Join-Path $HOME '.copilot\skills' }
}
foreach ($dst in $dsts) {
    New-Item -ItemType Directory -Force $dst | Out-Null
    foreach ($s in 'office-read', 'office-edit') {
        $target = Join-Path $dst $s
        if (Test-Path $target) { Remove-Item -Recurse -Force $target }
        Copy-Item -Recurse (Join-Path $src $s) $target
        Get-ChildItem -Recurse -Directory -Filter '__pycache__' $target | Remove-Item -Recurse -Force
        Write-Host "installed $s -> $target"
    }
}
Write-Host "Start a new agent session to pick up the skills."
