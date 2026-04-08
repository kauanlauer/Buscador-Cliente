$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$mainScript = "launcher_clientes_onedrive.pyw"
$mainExeName = "Buscador Cliente HeadCargo"
$setupScript = "buscador_cliente_headcargo.iss"
$setupExeName = "Setup.Buscador.Cliente.HeadCargo.exe"
$iconScript = ".\gerar_icone_buscador.ps1"
$iconFile = "logo_buscador.ico"
$versionInfoFile = "buscador_cliente_headcargo_version_info.txt"

Write-Host "Gerando icone do Windows..."
powershell -NoProfile -ExecutionPolicy Bypass -File $iconScript

Write-Host "Verificando PyInstaller..."
python -m PyInstaller --version *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Host "PyInstaller nao encontrado. Instalando..."
    python -m pip install pyinstaller
}

Write-Host "Verificando Inno Setup..."
$iscc = Get-Command iscc -ErrorAction SilentlyContinue
if (-not $iscc) {
    $candidate = Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"
    if (Test-Path $candidate) {
        $iscc = @{ Source = $candidate }
    }
}
if (-not $iscc) {
    Write-Host "Inno Setup nao encontrado. Instalando..."
    winget install --id JRSoftware.InnoSetup --exact --accept-package-agreements --accept-source-agreements --disable-interactivity
    $candidate = Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"
    if (Test-Path $candidate) {
        $iscc = @{ Source = $candidate }
    } else {
        $iscc = Get-Command iscc -ErrorAction SilentlyContinue
    }
}
if (-not $iscc) {
    throw "Nao foi possivel localizar o compilador do Inno Setup (iscc)."
}

Write-Host "Limpando build anterior..."
if (Test-Path ".\build") { Remove-Item ".\build" -Recurse -Force }
if (Test-Path ".\dist\$mainExeName.exe") { Remove-Item ".\dist\$mainExeName.exe" -Force }
if (Test-Path ".\dist\$setupExeName") { Remove-Item ".\dist\$setupExeName" -Force }

Write-Host "Gerando executavel principal..."
python -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --windowed `
  --add-data "logo_buscador.png;." `
  --add-data "$iconFile;." `
  --icon "$iconFile" `
  --version-file "$versionInfoFile" `
  --name "$mainExeName" `
  $mainScript

Write-Host "Gerando instalador padrao do Windows..."
& $iscc.Source $setupScript
if ($LASTEXITCODE -ne 0) {
    throw "Falha ao compilar o instalador do Inno Setup."
}

Write-Host ""
Write-Host "Build concluido."
Write-Host "App: dist\$mainExeName.exe"
Write-Host "Instalador: dist\$setupExeName"
