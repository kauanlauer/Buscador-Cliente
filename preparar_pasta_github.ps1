$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$githubFolder = Join-Path $projectRoot "Github"

if (Test-Path $githubFolder) {
    Get-ChildItem -Path $githubFolder -Force | Remove-Item -Recurse -Force
} else {
    New-Item -ItemType Directory -Path $githubFolder | Out-Null
}

$repoFiles = @(
    ".gitignore",
    "build_instalador_buscador.ps1",
    "Buscador Cliente HeadCargo.spec",
    "buscador_cliente_headcargo.iss",
    "buscador_cliente_headcargo_version_info.txt",
    "CHECKLIST_NOVA_VERSAO.md",
    "COMO_ATUALIZAR.md",
    "gerar_icone_buscador.ps1",
    "github_update_manifest.json",
    "launcher_clientes_onedrive.pyw",
    "logo_buscador.ico",
    "logo_buscador.png",
    "publicar_atualizacao_buscador.ps1"
)

foreach ($file in $repoFiles) {
    $source = Join-Path $projectRoot $file
    if (Test-Path $source) {
        Copy-Item $source (Join-Path $githubFolder (Split-Path $file -Leaf)) -Force
    }
}

Write-Host "Pasta Github preparada com sucesso."
Write-Host "Envie para o repositorio apenas os arquivos que estao em: $githubFolder"
Write-Host "O instalador continua em dist\\Setup.Buscador.Cliente.HeadCargo.exe para anexar no GitHub Release."
