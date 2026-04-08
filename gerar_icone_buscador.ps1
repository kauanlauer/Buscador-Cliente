$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$pngPath = Join-Path $projectRoot "logo_buscador.png"
$icoPath = Join-Path $projectRoot "logo_buscador.ico"
$scriptPath = $MyInvocation.MyCommand.Path

if (-not (Test-Path $pngPath)) {
    throw "Nao encontrei o arquivo logo_buscador.png"
}

$shouldGenerate = $true
if (Test-Path $icoPath) {
    $pngTime = (Get-Item $pngPath).LastWriteTimeUtc
    $icoTime = (Get-Item $icoPath).LastWriteTimeUtc
    $scriptTime = (Get-Item $scriptPath).LastWriteTimeUtc
    $shouldGenerate = $pngTime -gt $icoTime -or $scriptTime -gt $icoTime
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
from collections import deque

png_path = Path("logo_buscador.png")
ico_path = Path("logo_buscador.ico")
image = Image.open(png_path).convert("RGBA")

pixels = image.load()
width, height = image.size
visited = [[False] * width for _ in range(height)]
queue = deque()

def near_white(pixel):
    red, green, blue, alpha = pixel
    return alpha > 0 and red >= 245 and green >= 245 and blue >= 245

for x in range(width):
    queue.append((x, 0))
    queue.append((x, height - 1))

for y in range(height):
    queue.append((0, y))
    queue.append((width - 1, y))

while queue:
    x, y = queue.popleft()
    if x < 0 or y < 0 or x >= width or y >= height or visited[y][x]:
        continue
    visited[y][x] = True
    if not near_white(pixels[x, y]):
        continue
    pixels[x, y] = (255, 255, 255, 0)
    queue.append((x + 1, y))
    queue.append((x - 1, y))
    queue.append((x, y + 1))
    queue.append((x, y - 1))

bbox = image.getbbox()
if bbox:
    image = image.crop(bbox)

canvas = Image.new("RGBA", (1024, 1024), (255, 255, 255, 0))
image.thumbnail((940, 940), Image.Resampling.LANCZOS)
offset = ((1024 - image.width) // 2, (1024 - image.height) // 2)
canvas.alpha_composite(image, offset)

sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (40, 40), (32, 32), (24, 24), (20, 20), (16, 16)]
canvas.save(ico_path, format="ICO", sizes=sizes)
print(f"Icone gerado em: {ico_path.resolve()}")
'@ | python -
