$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$pngPath = Join-Path $projectRoot "logo_buscador.png"
$icoPath = Join-Path $projectRoot "logo_buscador.ico"

if (-not (Test-Path $pngPath)) {
    throw "Nao encontrei o arquivo logo_buscador.png"
}

$shouldGenerate = $true
if (Test-Path $icoPath) {
    $pngTime = (Get-Item $pngPath).LastWriteTimeUtc
    $icoTime = (Get-Item $icoPath).LastWriteTimeUtc
    $shouldGenerate = $pngTime -gt $icoTime
}

Write-Host "Verificando Pillow..."
python -m pip install pillow | Out-Null

if (-not $shouldGenerate) {
    Write-Host "Icone ja esta atualizado: $icoPath"
    exit 0
}

@'
from pathlib import Path
from PIL import Image

png_path = Path("logo_buscador.png")
ico_path = Path("logo_buscador.ico")
image = Image.open(png_path).convert("RGBA")
sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
image.save(ico_path, format="ICO", sizes=sizes)
print(f"Icone gerado em: {ico_path.resolve()}")
'@ | python -
