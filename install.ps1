# Install this skill into the agent's skill directory (one line):
#
#   irm https://raw.githubusercontent.com/C1jian6/ai-ppt-diagram-skill/main/install.ps1 | iex
#
# Override the target with:  .\install.ps1 D:\skills\dot2pptx
# Override the source with:  $env:DOT2PPTX_REPO = "https://github.com/you/fork.git"

param(
    [string]$Destination = (Join-Path $env:USERPROFILE ".dsh\skills\dot2pptx")
)

$ErrorActionPreference = "Stop"
$repo = if ($env:DOT2PPTX_REPO) { $env:DOT2PPTX_REPO } else { "https://github.com/C1jian6/ai-ppt-diagram-skill.git" }

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Error "git is required: https://git-scm.com/downloads"
}

if (Test-Path (Join-Path $Destination ".git")) {
    Write-Host "updating $Destination"
    git -C $Destination pull --ff-only
} elseif (Test-Path $Destination) {
    Write-Error "$Destination exists but is not a git checkout; remove it first"
} else {
    Write-Host "cloning into $Destination"
    New-Item -ItemType Directory -Force -Path (Split-Path $Destination) | Out-Null
    git clone --depth 1 $repo $Destination
}

Write-Host "`nchecking the environment:"
python (Join-Path $Destination "scripts\bootstrap.py") --check

Write-Host "`ndone. restart the agent session so the skill is picked up."
