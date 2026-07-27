[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ($env:OS -ne 'Windows_NT') {
    throw 'The bundled VBA write-back helper is built for Windows x64.'
}

$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$source = Join-Path $root 'native\vba-writeback\cli.py'
$requirements = Join-Path $root 'native\vba-writeback\requirements.lock'
$outputDirectory = Join-Path $root 'bin\win32-x64'
$outputPath = Join-Path $outputDirectory 'excel-ai-vba-writeback.exe'
$buildRoot = Join-Path ([IO.Path]::GetTempPath()) (
    'excel-ai-vba-writeback-build-' + [Guid]::NewGuid().ToString('N')
)
$venv = Join-Path $buildRoot 'venv'
$python = Join-Path $venv 'Scripts\python.exe'

try {
    [void][IO.Directory]::CreateDirectory($buildRoot)
    [void][IO.Directory]::CreateDirectory($outputDirectory)
    & py -3.11 -m venv $venv
    if ($LASTEXITCODE -ne 0) {
        throw 'Could not create the isolated Python 3.11 build environment.'
    }
    & $python -m pip install --disable-pip-version-check --require-virtualenv `
        -r $requirements
    if ($LASTEXITCODE -ne 0) {
        throw 'Could not install the pinned VBA write-back build dependencies.'
    }
    & $python -m PyInstaller `
        --noconfirm `
        --clean `
        --onefile `
        --name 'excel-ai-vba-writeback' `
        --distpath $outputDirectory `
        --workpath (Join-Path $buildRoot 'work') `
        --specpath (Join-Path $buildRoot 'spec') `
        $source
    if ($LASTEXITCODE -ne 0 -or -not [IO.File]::Exists($outputPath)) {
        throw 'PyInstaller did not create the expected executable.'
    }
    & $outputPath --help | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw 'The built VBA write-back executable failed its smoke test.'
    }
    $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $outputPath).Hash
    [Console]::Out.WriteLine("VBA_WRITEBACK_EXE|$outputPath")
    [Console]::Out.WriteLine("SHA256|$hash")
}
finally {
    if ([IO.Directory]::Exists($buildRoot)) {
        Remove-Item -LiteralPath $buildRoot -Recurse -Force
    }
}
