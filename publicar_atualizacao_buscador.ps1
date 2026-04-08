param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [string]$Notes = "Nova versao publicada no GitHub."
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$setupExe = Join-Path $projectRoot "dist\Setup.Buscador.Cliente.HeadCargo.exe"
$manifestPath = Join-Path $projectRoot "github_update_manifest.json"
$installerUrl = "https://github.com/kauanlauer/Buscador-Cliente/releases/latest/download/Setup.Buscador.Cliente.HeadCargo.exe"

if (-not (Test-Path $setupExe)) {
    throw "Nao encontrei o instalador em dist\Setup.Buscador.Cliente.HeadCargo.exe"
}

$manifest = [ordered]@{
    version = $Version
    installer_url = $installerUrl
    notes = $Notes
} | ConvertTo-Json -Depth 3

[System.IO.File]::WriteAllText(
    $manifestPath,
    $manifest,
    [System.Text.UTF8Encoding]::new($false)
)

Write-Host "Manifesto atualizado com sucesso."
Write-Host "Arquivo: $manifestPath"
Write-Host "Version: $Version"
Write-Host ""
Write-Host "Proximo passo:"
Write-Host "1. Suba o instalador em um GitHub Release como: Setup.Buscador.Cliente.HeadCargo.exe"
Write-Host "2. Commit e push do github_update_manifest.json"
