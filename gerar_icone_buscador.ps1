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
from collections import deque
from PIL import Image

png_path = Path("logo_buscador.png")
ico_path = Path("logo_buscador.ico")
image = Image.open(png_path).convert("RGBA")

# Remove apenas o fundo branco ligado às bordas, preservando os detalhes
# brancos internos do logo.
pixels = image.load()
width, height = image.size
queue = deque()
visited = set()

def is_background(x: int, y: int) -> bool:
    r, g, b, a = pixels[x, y]
    return a > 0 and r >= 245 and g >= 245 and b >= 245

for x in range(width):
    queue.append((x, 0))
    queue.append((x, height - 1))
for y in range(height):
    queue.append((0, y))
    queue.append((width - 1, y))

while queue:
    x, y = queue.popleft()
    if (x, y) in visited:
        continue
    visited.add((x, y))
    if not (0 <= x < width and 0 <= y < height):
        continue
    if not is_background(x, y):
        continue

    pixels[x, y] = (255, 255, 255, 0)
    queue.extend(
        [
            (x - 1, y),
            (x + 1, y),
            (x, y - 1),
            (x, y + 1),
        ]
    )

sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
image.save(ico_path, format="ICO", sizes=sizes)
print(f"Icone gerado em: {ico_path.resolve()}")
'@ | python -
