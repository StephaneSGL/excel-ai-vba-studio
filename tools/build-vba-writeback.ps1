[CmdletBinding()]
param(
    [string]$PythonPath = 'python'
)

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
$previousPythonHashSeed = $env:PYTHONHASHSEED
$previousSourceDateEpoch = $env:SOURCE_DATE_EPOCH

function Get-Sha256 {
    param([string]$Path)

    $stream = [IO.File]::OpenRead($Path)
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        return [BitConverter]::ToString(
            $sha256.ComputeHash($stream)
        ).Replace('-', '')
    }
    finally {
        $sha256.Dispose()
        $stream.Dispose()
    }
}

try {
    # PyInstaller documents both values as required inputs for reproducible
    # Windows builds. Keep them fixed so CI can compare the shipped helper
    # byte-for-byte with a clean rebuild from this repository.
    $env:PYTHONHASHSEED = '1'
    $env:SOURCE_DATE_EPOCH = '1704067200'
    [void][IO.Directory]::CreateDirectory($buildRoot)
    [void][IO.Directory]::CreateDirectory($outputDirectory)
    $bootstrapPython = (Get-Command $PythonPath -ErrorAction Stop).Source
    $pythonVersion = & $bootstrapPython -c (
        'import platform; print(platform.python_version())'
    )
    if ($LASTEXITCODE -ne 0 -or [string]$pythonVersion -cne '3.11.9') {
        throw (
            'The reproducible helper build requires Python 3.11.9 exactly; ' +
            "found $pythonVersion at $bootstrapPython."
        )
    }
    & $bootstrapPython -m venv $venv
    if ($LASTEXITCODE -ne 0) {
        throw 'Could not create the isolated Python 3.11 build environment.'
    }
    & $python -m pip install --disable-pip-version-check --require-virtualenv `
        --require-hashes `
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
    $hash = Get-Sha256 $outputPath
    [Console]::Out.WriteLine("VBA_WRITEBACK_EXE|$outputPath")
    [Console]::Out.WriteLine("SHA256|$hash")
}
finally {
    $env:PYTHONHASHSEED = $previousPythonHashSeed
    $env:SOURCE_DATE_EPOCH = $previousSourceDateEpoch
    if ([IO.Directory]::Exists($buildRoot)) {
        Remove-Item -LiteralPath $buildRoot -Recurse -Force
    }
}
