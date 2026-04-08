$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$mainScript = "launcher_clientes_onedrive.pyw"
$mainExeName = "Buscador Cliente HeadCargo"
$setupScript = "buscador_cliente_headcargo.iss"
$setupExeName = "Setup Buscador Cliente HeadCargo.exe"
$iconScript = ".\gerar_icone_buscador.ps1"
$iconFile = "logo_buscador.ico"
$versionInfoFile = "buscador_cliente_headcargo_version_info.txt"

function Get-SignToolPath {
    if ($env:BUSCADOR_SIGNTOOL_PATH -and (Test-Path $env:BUSCADOR_SIGNTOOL_PATH)) {
        return $env:BUSCADOR_SIGNTOOL_PATH
    }

    $command = Get-Command signtool -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $roots = @(
        (Join-Path ${env:ProgramFiles(x86)} "Windows Kits\10\bin"),
        (Join-Path $env:ProgramFiles "Windows Kits\10\bin")
    ) | Where-Object { $_ -and (Test-Path $_) }

    foreach ($root in $roots) {
        $match = Get-ChildItem -Path $root -Filter signtool.exe -File -Recurse -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending |
            Select-Object -First 1
        if ($match) {
            return $match.FullName
        }
    }

    return $null
}

function Invoke-CodeSigningIfConfigured {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    if (-not (Test-Path $TargetPath)) {
        throw "Arquivo para assinatura nao encontrado: $TargetPath"
    }

    $signTool = Get-SignToolPath
    $certFile = $env:BUSCADOR_SIGN_CERT_FILE
    $subjectName = $env:BUSCADOR_SIGN_SUBJECT_NAME

    if (-not $signTool) {
        Write-Host "Signtool nao encontrado. Build segue sem assinatura digital."
        return $false
    }

    if ($certFile -and -not (Test-Path $certFile)) {
        throw "Certificado informado em BUSCADOR_SIGN_CERT_FILE nao encontrado: $certFile"
    }

    if (-not $certFile -and -not $subjectName) {
        Write-Host "Assinatura digital nao configurada. Defina BUSCADOR_SIGN_CERT_FILE ou BUSCADOR_SIGN_SUBJECT_NAME para assinar."
        return $false
    }

    $timestampUrl = if ($env:BUSCADOR_SIGN_TIMESTAMP_URL) { $env:BUSCADOR_SIGN_TIMESTAMP_URL } else { "http://timestamp.digicert.com" }
    $arguments = @("sign", "/fd", "SHA256", "/td", "SHA256", "/tr", $timestampUrl)

    if ($certFile) {
        $arguments += @("/f", $certFile)
        if ($env:BUSCADOR_SIGN_CERT_PASSWORD) {
            $arguments += @("/p", $env:BUSCADOR_SIGN_CERT_PASSWORD)
        }
    } else {
        $arguments += @("/n", $subjectName)
    }

    $arguments += $TargetPath

    Write-Host "Assinando digitalmente: $TargetPath"
    & $signTool @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Falha ao assinar digitalmente: $TargetPath"
    }

    return $true
}

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

if (-not (Test-Path $versionInfoFile)) {
    throw "Arquivo de versao nao encontrado: $versionInfoFile"
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
  --noupx `
  --add-data "logo_buscador.png;." `
  --add-data "$iconFile;." `
  --icon "$iconFile" `
  --version-file "$versionInfoFile" `
  --name "$mainExeName" `
  $mainScript

Invoke-CodeSigningIfConfigured ".\dist\$mainExeName.exe" | Out-Null

Write-Host "Gerando instalador padrao do Windows..."
& $iscc.Source $setupScript
if ($LASTEXITCODE -ne 0) {
    throw "Falha ao compilar o instalador do Inno Setup."
}

$signedInstaller = Invoke-CodeSigningIfConfigured ".\dist\$setupExeName"

Write-Host ""
Write-Host "Build concluido."
Write-Host "App: dist\$mainExeName.exe"
Write-Host "Instalador: dist\$setupExeName"
if (-not $signedInstaller) {
    Write-Host "Observacao: sem assinatura digital valida o Windows ainda pode exibir aviso de reputacao/SmartScreen."
}
