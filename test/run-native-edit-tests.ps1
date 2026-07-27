param(
    [string]$WorkbookPath = "$PSScriptRoot\fixtures\NativeEditingSynthetic.xlsm",
    [string]$BridgeScriptPath = "$PSScriptRoot\..\scripts\office-ai-apply-edits.ps1"
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Release-ComObject {
    param([AllowNull()][object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Invoke-NativeEdit {
    param(
        [string]$TargetPath,
        [object[]]$Operations
    )
    $expectedHash = (Get-FileHash -LiteralPath $TargetPath -Algorithm SHA256).
        Hash.ToLowerInvariant()
    $transactionId = [Guid]::NewGuid().ToString('D')
    $payloadPath = Join-Path ([IO.Path]::GetTempPath()) (
        'excel-native-test-' + [Guid]::NewGuid().ToString('N') + '.json'
    )
    try {
        @{
            version = 2
            transactionId = $transactionId
            expectedWorkbookSha256 = $expectedHash
            operations = $Operations
        } |
            ConvertTo-Json -Depth 10 -Compress |
            Set-Content -LiteralPath $payloadPath -Encoding UTF8
        $output = & powershell.exe `
            -NoLogo `
            -NoProfile `
            -NonInteractive `
            -ExecutionPolicy Bypass `
            -File $BridgeScriptPath `
            -WorkbookPath $TargetPath `
            -OperationsPath $payloadPath 2>&1
        if ($LASTEXITCODE -ne 0 -or [string]$output -notmatch '"ok":true') {
            throw "Native bridge failed: $output"
        }
        $resultLine = @(
            $output | Where-Object { [string]$_ -match '^\s*\{' }
        ) | Select-Object -Last 1
        $result = [string]$resultLine | ConvertFrom-Json
        if (
            $result.ok -ne $true -or
            [string]::IsNullOrWhiteSpace([string]$result.backupPath) -or
            -not (Test-Path -LiteralPath $result.backupPath -PathType Leaf)
        ) {
            throw "Native bridge returned an invalid result: $resultLine"
        }
        return $result
    }
    finally {
        if (Test-Path -LiteralPath $payloadPath -PathType Leaf) {
            Remove-Item -LiteralPath $payloadPath -Force
        }
    }
}

function Read-WorkbookState {
    param([string]$Path)

    $excel = $null
    $workbooks = $null
    $workbook = $null
    $sheet = $null
    $conditionalRange = $null
    $conditions = $null
    $buttons = $null
    $topBorder = $null
    $cells = @()
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AskToUpdateLinks = $false
        $excel.AutomationSecurity = 3
        $workbooks = $excel.Workbooks
        $workbook = $workbooks.Open($Path, 0, $true)
        $sheet = $workbook.Worksheets.Item('Donnees')

        foreach ($address in @('A5', 'B5', 'C5', 'D5', 'F5')) {
            $cells += $sheet.Range($address)
        }
        $conditionalRange = $sheet.Range('B2:B5')
        $conditions = $conditionalRange.FormatConditions
        $buttons = $sheet.Buttons()
        $topBorder = $cells[3].Borders.Item(8)

        return [PSCustomObject]@{
            A5Text = [string]$cells[0].Text
            B5Value = $cells[1].Value2
            C5Formula = [string]$cells[2].Formula
            D5Align = $cells[3].HorizontalAlignment
            D5VAlign = $cells[3].VerticalAlignment
            D5Fill = $cells[3].Interior.Color
            D5FontColor = $cells[3].Font.Color
            D5Bold = $cells[3].Font.Bold
            D5Italic = $cells[3].Font.Italic
            D5Wrap = $cells[3].WrapText
            D5Border = $topBorder.LineStyle
            F5Pattern = $cells[4].Interior.Pattern
            F5FontColorIndex = $cells[4].Font.ColorIndex
            F5Bold = $cells[4].Font.Bold
            ConditionalFormattingCount = $conditions.Count
            ButtonCount = $buttons.Count
        }
    }
    finally {
        Release-ComObject $topBorder
        Release-ComObject $buttons
        Release-ComObject $conditions
        Release-ComObject $conditionalRange
        foreach ($cell in $cells) {
            Release-ComObject $cell
        }
        Release-ComObject $sheet
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch { }
        }
        Release-ComObject $workbook
        Release-ComObject $workbooks
        if ($null -ne $excel) {
            try { $excel.Quit() } catch { }
        }
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-ZipState {
    param([string]$Path)

    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    $hashes = @{}
    try {
        $entries = @($archive.Entries | ForEach-Object { $_.FullName } | Sort-Object)
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -notmatch '\.(bin|png|jpe?g|gif|emf|wmf|vml)$') {
                continue
            }
            $stream = $entry.Open()
            $sha = [Security.Cryptography.SHA256]::Create()
            try {
                $hashes[$entry.FullName] = [BitConverter]::ToString(
                    $sha.ComputeHash($stream)
                ).Replace('-', '')
            }
            finally {
                $sha.Dispose()
                $stream.Dispose()
            }
        }
        return [PSCustomObject]@{
            Entries = $entries
            OpaqueHashes = $hashes
        }
    }
    finally {
        $archive.Dispose()
    }
}

$fixturePath = [IO.Path]::GetFullPath($WorkbookPath)
if (-not (Test-Path -LiteralPath $fixturePath -PathType Leaf)) {
    throw "Fixture missing: $fixturePath. Run test/create-synthetic-xlsm.ps1."
}

$preexistingExcelIds = @(
    Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty Id
)
$workingPath = Join-Path ([IO.Path]::GetTempPath()) (
    'excel-native-integration-' + [Guid]::NewGuid().ToString('N') + '.xlsm'
)
$results = New-Object 'System.Collections.Generic.List[object]'
$backupPaths = New-Object 'System.Collections.Generic.List[string]'

function Test-Condition {
    param(
        [string]$Name,
        [bool]$Passed,
        [string]$Detail = ''
    )
    [void]$results.Add([PSCustomObject]@{
        Name = $Name
        Passed = $Passed
        Detail = $Detail
    })
    $status = if ($Passed) { 'PASS' } else { 'FAIL' }
    Write-Host "[$status] $Name $Detail"
}

try {
    Copy-Item -LiteralPath $fixturePath -Destination $workingPath
    $beforeWorkbook = Read-WorkbookState $workingPath
    $beforeZip = Get-ZipState $workingPath

    $firstResult = Invoke-NativeEdit $workingPath @(
        @{
            sheetName = 'Donnees'
            row = 5
            column = 1
            value = @{ kind = 'text'; value = 'IntegrationEdit' }
        },
        @{
            sheetName = 'Donnees'
            row = 5
            column = 2
            value = @{ kind = 'number'; value = 99.5 }
        },
        @{
            sheetName = 'Donnees'
            row = 5
            column = 3
            value = @{ kind = 'formula'; value = '=B5*2' }
        },
        @{
            sheetName = 'Donnees'
            row = 5
            column = 4
            style = @{
                align = 'center'
                valign = 'middle'
                bgcolor = '#123456'
                color = '#ffffff'
                format = 'number'
                textwrap = $true
                font = @{ name = 'Calibri'; size = 12; bold = $true; italic = $true }
                border = @{ top = @('thin', '#ff0000') }
            }
        },
        @{
            sheetName = 'Donnees'
            row = 5
            column = 6
            style = @{
                bgcolor = '#ff0000'
                color = '#00ff00'
                font = @{ bold = $true }
            }
        }
    )
    [void]$backupPaths.Add([string]$firstResult.backupPath)

    $hashBeforeSecondEdit = (Get-FileHash -LiteralPath $workingPath -Algorithm SHA256).
        Hash.ToLowerInvariant()
    $secondResult = Invoke-NativeEdit $workingPath @(
        @{
            sheetName = 'Donnees'
            row = 5
            column = 6
            style = @{
                bgcolor = $null
                color = $null
                font = @{ bold = $null }
            }
        }
    )
    [void]$backupPaths.Add([string]$secondResult.backupPath)

    $afterWorkbook = Read-WorkbookState $workingPath
    $afterZip = Get-ZipState $workingPath

    Test-Condition 'Text edit' ($afterWorkbook.A5Text -eq 'IntegrationEdit')
    Test-Condition 'Number edit' ($afterWorkbook.B5Value -eq 99.5)
    Test-Condition 'Formula edit' ($afterWorkbook.C5Formula -eq '=B5*2')
    Test-Condition 'Horizontal alignment edit' ($afterWorkbook.D5Align -eq -4108)
    Test-Condition 'Vertical alignment edit' ($afterWorkbook.D5VAlign -eq -4108)
    Test-Condition 'Fill edit' ($afterWorkbook.D5Fill -eq 5649426)
    Test-Condition 'Font color edit' ($afterWorkbook.D5FontColor -eq 16777215)
    Test-Condition 'Bold and italic edit' (
        $afterWorkbook.D5Bold -eq $true -and
        $afterWorkbook.D5Italic -eq $true
    )
    Test-Condition 'Wrap edit' ($afterWorkbook.D5Wrap -eq $true)
    Test-Condition 'Border edit' ($afterWorkbook.D5Border -ne -4142)
    Test-Condition 'Fill reset' ($afterWorkbook.F5Pattern -eq -4142)
    Test-Condition 'Font color reset' ($afterWorkbook.F5FontColorIndex -eq -4105)
    Test-Condition 'Font bold reset' ($afterWorkbook.F5Bold -eq $false)
    Test-Condition 'Conditional formatting preserved' (
        $afterWorkbook.ConditionalFormattingCount -eq
        $beforeWorkbook.ConditionalFormattingCount
    )
    Test-Condition 'Form button preserved' (
        $afterWorkbook.ButtonCount -eq $beforeWorkbook.ButtonCount
    )
    Test-Condition 'OOXML entry set preserved' (
        (Compare-Object $beforeZip.Entries $afterZip.Entries).Count -eq 0
    )

    $opaqueDifferences = @(
        foreach ($name in $beforeZip.OpaqueHashes.Keys) {
            if ($beforeZip.OpaqueHashes[$name] -ne $afterZip.OpaqueHashes[$name]) {
                $name
            }
        }
    )
    Test-Condition 'Opaque OOXML parts preserved' (
        $opaqueDifferences.Count -eq 0
    ) ($opaqueDifferences -join ', ')
    Test-Condition 'Persistent backups created' (
        $backupPaths.Count -eq 2 -and
        @(
            $backupPaths |
                Where-Object { Test-Path -LiteralPath $_ -PathType Leaf }
        ).Count -eq 2
    )
    Test-Condition 'Second backup matches pre-edit workbook' (
        (Get-FileHash -LiteralPath $secondResult.backupPath -Algorithm SHA256).
            Hash.ToLowerInvariant() -ceq $hashBeforeSecondEdit
    )
}
finally {
    if (Test-Path -LiteralPath $workingPath -PathType Leaf) {
        Remove-Item -LiteralPath $workingPath -Force
    }
    foreach ($backupPath in $backupPaths) {
        if (Test-Path -LiteralPath $backupPath -PathType Leaf) {
            Remove-Item -LiteralPath $backupPath -Force
        }
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

$cleanupDeadline = [DateTime]::UtcNow.AddSeconds(10)
do {
    Start-Sleep -Milliseconds 250
    $newExcelProcesses = @(
        Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
            Where-Object { $preexistingExcelIds -notcontains $_.Id }
    )
} while (
    $newExcelProcesses.Count -gt 0 -and
    [DateTime]::UtcNow -lt $cleanupDeadline
)
Test-Condition 'No new orphan Excel processes' ($newExcelProcesses.Count -eq 0) (
    ($newExcelProcesses | Select-Object -ExpandProperty Id) -join ', '
)

$failed = @($results | Where-Object { -not $_.Passed })
Write-Host "RESULTS|passed=$($results.Count - $failed.Count)|failed=$($failed.Count)"
if ($failed.Count -gt 0) {
    exit 1
}
