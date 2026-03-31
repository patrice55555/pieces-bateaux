param(
    [string]$CommitMessage
)

$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
Set-Location $repoRoot

$gitCommand = Get-Command git -ErrorAction SilentlyContinue
if (-not $gitCommand) {
    throw 'Git est introuvable dans le PATH.'
}

$insideWorkTree = git rev-parse --is-inside-work-tree 2>$null
if ($LASTEXITCODE -ne 0 -or $insideWorkTree.Trim() -ne 'true') {
    throw "Le dossier $repoRoot n'est pas un depot Git valide."
}

$statusOutput = git status --short
if ([string]::IsNullOrWhiteSpace(($statusOutput | Out-String))) {
    Write-Host 'Aucun changement a pousser vers GitHub.'
    exit 0
}

if ([string]::IsNullOrWhiteSpace($CommitMessage)) {
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $CommitMessage = "Mise a jour scraping $timestamp"
}

git add -A
git commit -m $CommitMessage
git push origin main

Write-Host "Push GitHub termine avec le message : $CommitMessage"