$ErrorActionPreference = "Stop"

$githubRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $githubRoot
Set-Location $projectRoot

$repoFolder = Join-Path $githubRoot "Repositorio"
$releaseFolder = Join-Path $githubRoot "Release"

if (Test-Path $repoFolder) { Remove-Item $repoFolder -Recurse -Force }
if (Test-Path $releaseFolder) { Remove-Item $releaseFolder -Recurse -Force }

New-Item -ItemType Directory -Path $repoFolder | Out-Null
New-Item -ItemType Directory -Path $releaseFolder | Out-Null

$repoFiles = @(
    "launcher_clientes_onedrive.pyw",
    "build_instalador_buscador.ps1",
    "gerar_icone_buscador.ps1",
    "buscador_cliente_headcargo.iss",
    "publicar_atualizacao_buscador.ps1",
    "github_update_manifest.json",
    "COMO_ATUALIZAR.md",
    "logo_buscador.png",
    "logo_buscador.ico",
    "Buscador Cliente HeadCargo.spec",
    ".gitignore"
)

foreach ($file in $repoFiles) {
    if (Test-Path $file) {
        Copy-Item $file (Join-Path $repoFolder (Split-Path $file -Leaf)) -Force
    }
}

$setupFile = Join-Path $projectRoot "dist\Setup Buscador Cliente HeadCargo.exe"
if (Test-Path $setupFile) {
    Copy-Item $setupFile (Join-Path $releaseFolder "Setup Buscador Cliente HeadCargo.exe") -Force
}

Write-Host "Pasta GitHub preparada com sucesso."
Write-Host "Repositorio: $repoFolder"
Write-Host "Release: $releaseFolder"
