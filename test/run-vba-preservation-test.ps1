param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,
    [string]$BridgeScriptPath = ''
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-EntryHash {
    param(
        [System.IO.Compression.ZipArchiveEntry]$Entry
    )

    $stream = $Entry.Open()
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha256.ComputeHash($stream)
        return (($hash | ForEach-Object { $_.ToString('x2') }) -join '')
    }
    finally {
        $sha256.Dispose()
        $stream.Dispose()
    }
}

function Get-Sha256 {
    param([string]$Path)

    $stream = [IO.File]::OpenRead($Path)
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString(
            $sha256.ComputeHash($stream)
        ).Replace('-', '').ToLowerInvariant())
    }
    finally {
        $sha256.Dispose()
        $stream.Dispose()
    }
}

function Get-WorkbookEvidence {
    param(
        [string]$Path,
        [string]$VbaOutputPath
    )

    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $workbookEntry = $archive.Entries |
            Where-Object { $_.FullName -ieq 'xl/workbook.xml' } |
            Select-Object -First 1
        $vbaEntry = $archive.Entries |
            Where-Object { $_.FullName -ieq 'xl/vbaProject.bin' } |
            Select-Object -First 1
        if ($null -eq $workbookEntry) {
            throw 'Invalid workbook: xl/workbook.xml is missing.'
        }
        if ($null -eq $vbaEntry) {
            throw 'Fixture has no xl/vbaProject.bin. Supply a real macro-enabled XLSM.'
        }
        [IO.Compression.ZipFileExtensions]::ExtractToFile(
            $vbaEntry,
            $VbaOutputPath,
            $true
        )

        $stream = $workbookEntry.Open()
        $reader = New-Object IO.StreamReader($stream)
        try {
            [xml]$workbookXml = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
            $stream.Dispose()
        }

        $namespaceManager = New-Object Xml.XmlNamespaceManager($workbookXml.NameTable)
        $namespaceManager.AddNamespace(
            'main',
            'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        )
        $firstSheet = $workbookXml.SelectSingleNode(
            '/main:workbook/main:sheets/main:sheet[1]',
            $namespaceManager
        )
        if ($null -eq $firstSheet) {
            throw 'Invalid workbook: no worksheet declaration found.'
        }

        return @{
            SheetName = [string]$firstSheet.GetAttribute('name')
            VbaBytes = [long]$vbaEntry.Length
            VbaHash = Get-EntryHash $vbaEntry
            EntryNames = @($archive.Entries | ForEach-Object { $_.FullName } | Sort-Object)
        }
    }
    finally {
        $archive.Dispose()
    }
}

function Get-ExcelProcessIds {
    return @(
        Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty Id
    )
}

$fixturePath = [IO.Path]::GetFullPath($WorkbookPath)
if ([string]::IsNullOrWhiteSpace($BridgeScriptPath)) {
    $BridgeScriptPath = Join-Path $PSScriptRoot '..\scripts\office-ai-apply-edits.ps1'
}
$bridgePath = [IO.Path]::GetFullPath($BridgeScriptPath)
if (-not (Test-Path -LiteralPath $fixturePath -PathType Leaf)) {
    throw "Fixture missing: $fixturePath"
}
if ([IO.Path]::GetExtension($fixturePath) -ine '.xlsm') {
    throw 'Fixture must use the .xlsm extension.'
}
if (-not (Test-Path -LiteralPath $bridgePath -PathType Leaf)) {
    throw "Bridge script missing: $bridgePath"
}

$workingPath = Join-Path ([IO.Path]::GetTempPath()) (
    'excel-vba-preservation-' + [Guid]::NewGuid().ToString('N') + '.xlsm'
)
$payloadPath = Join-Path ([IO.Path]::GetTempPath()) (
    'excel-vba-preservation-' + [Guid]::NewGuid().ToString('N') + '.json'
)
$beforeVbaPath = Join-Path ([IO.Path]::GetTempPath()) (
    'excel-vba-preservation-before-' + [Guid]::NewGuid().ToString('N') + '.bin'
)
$afterVbaPath = Join-Path ([IO.Path]::GetTempPath()) (
    'excel-vba-preservation-after-' + [Guid]::NewGuid().ToString('N') + '.bin'
)
$comparisonScriptPath = Join-Path $PSScriptRoot 'compare-vba-project-streams.mjs'
if (-not (Test-Path -LiteralPath $comparisonScriptPath -PathType Leaf)) {
    throw "VBA comparison script missing: $comparisonScriptPath"
}
$beforeProcessIds = Get-ExcelProcessIds
$backupPath = $null

try {
    Copy-Item -LiteralPath $fixturePath -Destination $workingPath
    $before = Get-WorkbookEvidence $workingPath $beforeVbaPath
    $expectedWorkbookSha256 = Get-Sha256 $workingPath

    @{
        version = 2
        transactionId = [Guid]::NewGuid().ToString('D')
        expectedWorkbookSha256 = $expectedWorkbookSha256
        operations = @(
            @{
                sheetName = $before.SheetName
                row = 1
                column = 1
                value = @{
                    kind = 'text'
                    value = 'Office Workbench VBA preservation test'
                }
            }
        )
    } |
        ConvertTo-Json -Depth 8 -Compress |
        Set-Content -LiteralPath $payloadPath -Encoding UTF8

    $output = & powershell.exe `
        -NoLogo `
        -NoProfile `
        -NonInteractive `
        -ExecutionPolicy Bypass `
        -File $bridgePath `
        -WorkbookPath $workingPath `
        -OperationsPath $payloadPath 2>&1
    if ($LASTEXITCODE -ne 0 -or [string]$output -notmatch '"ok":true') {
        throw "Native bridge failed: $output"
    }
    $resultLine = @(
        $output | Where-Object { [string]$_ -match '^\s*\{' }
    ) | Select-Object -Last 1
    $result = [string]$resultLine | ConvertFrom-Json
    $backupPath = [string]$result.backupPath
    if (
        $result.ok -ne $true -or
        -not (Test-Path -LiteralPath $backupPath -PathType Leaf) -or
        (Get-Sha256 $backupPath) -cne $expectedWorkbookSha256
    ) {
        throw "Native bridge returned an invalid persistent backup: $resultLine"
    }

    $after = Get-WorkbookEvidence $workingPath $afterVbaPath
    $entrySetPreserved =
        [string]::Join("`n", $before.EntryNames) -ceq
        [string]::Join("`n", $after.EntryNames)
    $vbaComparison = & node $comparisonScriptPath $beforeVbaPath $afterVbaPath
    if ($LASTEXITCODE -ne 0) {
        throw "VBA project streams changed: $vbaComparison"
    }
    $vbaResult = [string]$vbaComparison | ConvertFrom-Json

    Start-Sleep -Milliseconds 750
    $afterProcessIds = Get-ExcelProcessIds
    $newProcessIds = @($afterProcessIds | Where-Object { $_ -notin $beforeProcessIds })

    if (-not $entrySetPreserved) {
        throw 'OOXML entry set changed after targeted native save.'
    }
    if (-not $vbaResult.ok) {
        throw "VBA project streams changed: $vbaComparison"
    }
    if ($newProcessIds.Count -gt 0) {
        throw "New orphan Excel processes detected: $($newProcessIds -join ', ')"
    }

    Write-Output (
        "[PASS] Real VBA project streams preserved " +
        "($($vbaResult.beforeStreamCount) streams, $($before.VbaBytes) container bytes)"
    )
    Write-Output '[PASS] OOXML entry set preserved'
    Write-Output '[PASS] Persistent pre-edit backup verified'
    Write-Output '[PASS] No new orphan Excel processes'
    Write-Output "RESULTS|passed=4|failed=0|sheet=$($before.SheetName)"
}
finally {
    if (Test-Path -LiteralPath $payloadPath -PathType Leaf) {
        Remove-Item -LiteralPath $payloadPath -Force
    }
    if (Test-Path -LiteralPath $workingPath -PathType Leaf) {
        Remove-Item -LiteralPath $workingPath -Force
    }
    if (Test-Path -LiteralPath $beforeVbaPath -PathType Leaf) {
        Remove-Item -LiteralPath $beforeVbaPath -Force
    }
    if (Test-Path -LiteralPath $afterVbaPath -PathType Leaf) {
        Remove-Item -LiteralPath $afterVbaPath -Force
    }
    if (
        -not [string]::IsNullOrWhiteSpace($backupPath) -and
        (Test-Path -LiteralPath $backupPath -PathType Leaf)
    ) {
        Remove-Item -LiteralPath $backupPath -Force
    }
    $backupDirectory = Join-Path (
        [IO.Path]::GetDirectoryName($workingPath)
    ) '.excel-ai-vba-backups'
    if (
        (Test-Path -LiteralPath $backupDirectory -PathType Container) -and
        @(Get-ChildItem -LiteralPath $backupDirectory -Force).Count -eq 0
    ) {
        Remove-Item -LiteralPath $backupDirectory -Force
    }
}
