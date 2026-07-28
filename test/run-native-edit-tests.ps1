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

function Read-ConditionalRuleState {
    param(
        [object]$Sheet,
        [string]$Address
    )

    $range = $null
    $conditions = $null
    $condition = $null
    $interior = $null
    $font = $null
    $criteria = $null
    $criterion1 = $null
    $criterion2 = $null
    $criterion3 = $null
    $formatColor1 = $null
    $formatColor2 = $null
    $formatColor3 = $null
    $minPoint = $null
    $maxPoint = $null
    $barColor = $null
    $iconSet = $null
    try {
        $range = $Sheet.Range($Address)
        $conditions = $range.FormatConditions
        if ($conditions.Count -eq 0) {
            return [PSCustomObject]@{ Count = 0; Type = $null }
        }
        $condition = $conditions.Item(1)
        $state = [ordered]@{
            Count = [int]$conditions.Count
            Type = [int]$condition.Type
        }
        switch ([int]$condition.Type) {
            1 {
                $interior = $condition.Interior
                $font = $condition.Font
                $state.Operator = [int]$condition.Operator
                $state.Formula1 = [string]$condition.Formula1
                $state.FillColor = [int64]$interior.Color
                $state.FontColor = [int64]$font.Color
                $state.Bold = [bool]$font.Bold
            }
            9 {
                $interior = $condition.Interior
                $font = $condition.Font
                $state.Text = [string]$condition.Text
                $state.TextOperator = [int]$condition.TextOperator
                $state.FillColor = [int64]$interior.Color
                $state.FontColor = [int64]$font.Color
                $state.Bold = [bool]$font.Bold
            }
            3 {
                $criteria = $condition.ColorScaleCriteria
                $criterion1 = $criteria.Item(1)
                $criterion2 = $criteria.Item(2)
                $criterion3 = $criteria.Item(3)
                $formatColor1 = $criterion1.FormatColor
                $formatColor2 = $criterion2.FormatColor
                $formatColor3 = $criterion3.FormatColor
                $state.CriteriaTypes = @(
                    [int]$criterion1.Type,
                    [int]$criterion2.Type,
                    [int]$criterion3.Type
                )
                $state.MiddleValue = [double]$criterion2.Value
                $state.Colors = @(
                    [int64]$formatColor1.Color,
                    [int64]$formatColor2.Color,
                    [int64]$formatColor3.Color
                )
            }
            4 {
                $minPoint = $condition.MinPoint
                $maxPoint = $condition.MaxPoint
                $barColor = $condition.BarColor
                $state.MinType = [int]$minPoint.Type
                $state.MaxType = [int]$maxPoint.Type
                $state.Color = [int64]$barColor.Color
            }
            6 {
                $iconSet = $condition.IconSet
                $criteria = $condition.IconCriteria
                $criterion2 = $criteria.Item(2)
                $criterion3 = $criteria.Item(3)
                $state.IconSetId = [int]$iconSet.ID
                $state.Thresholds = @(
                    [double]$criterion2.Value,
                    [double]$criterion3.Value
                )
                $state.ThresholdTypes = @(
                    [int]$criterion2.Type,
                    [int]$criterion3.Type
                )
                $state.ThresholdOperators = @(
                    [int]$criterion2.Operator,
                    [int]$criterion3.Operator
                )
            }
        }
        return [PSCustomObject]$state
    }
    finally {
        Release-ComObject $iconSet
        Release-ComObject $barColor
        Release-ComObject $maxPoint
        Release-ComObject $minPoint
        Release-ComObject $formatColor3
        Release-ComObject $formatColor2
        Release-ComObject $formatColor1
        Release-ComObject $criterion3
        Release-ComObject $criterion2
        Release-ComObject $criterion1
        Release-ComObject $criteria
        Release-ComObject $font
        Release-ComObject $interior
        Release-ComObject $condition
        Release-ComObject $conditions
        Release-ComObject $range
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
    $allCells = $null
    $allConditions = $null
    $buttons = $null
    $topBorder = $null
    $columns = $null
    $rows = $null
    $columnT = $null
    $row10 = $null
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
        $allCells = $sheet.Cells
        $allConditions = $allCells.FormatConditions
        $buttons = $sheet.Buttons()
        $topBorder = $cells[3].Borders.Item(8)
        $columns = $sheet.Columns
        $rows = $sheet.Rows
        $columnT = $columns.Item(20)
        $row10 = $rows.Item(10)

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
            TotalConditionalFormattingCount = $allConditions.Count
            CellIsRule = Read-ConditionalRuleState $sheet 'G2:G10'
            ContainsTextRule = Read-ConditionalRuleState $sheet 'H2:H10'
            ColorScaleRule = Read-ConditionalRuleState $sheet 'I2:I10'
            DataBarRule = Read-ConditionalRuleState $sheet 'J2:J10'
            IconSetRule = Read-ConditionalRuleState $sheet 'K2:K10'
            ColumnTWidth = [double]$columnT.ColumnWidth
            Row10Height = [double]$row10.RowHeight
            ButtonCount = $buttons.Count
        }
    }
    finally {
        Release-ComObject $row10
        Release-ComObject $columnT
        Release-ComObject $rows
        Release-ComObject $columns
        Release-ComObject $topBorder
        Release-ComObject $buttons
        Release-ComObject $allConditions
        Release-ComObject $allCells
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

function Convert-TestHexToOleColor {
    param([string]$Hex)

    $normalized = $Hex.TrimStart('#')
    $red = [Convert]::ToInt32($normalized.Substring(0, 2), 16)
    $green = [Convert]::ToInt32($normalized.Substring(2, 2), 16)
    $blue = [Convert]::ToInt32($normalized.Substring(4, 2), 16)
    return $red + ($green * 256) + ($blue * 65536)
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
        },
        @{
            kind = 'columnWidth'
            sheetName = 'Donnees'
            column = 20
            widthPx = 145
        },
        @{
            kind = 'rowHeight'
            sheetName = 'Donnees'
            row = 10
            heightPx = 48
        },
        @{
            kind = 'addConditionalFormatting'
            sheetName = 'Donnees'
            rangeRef = 'G2:G10'
            rule = @{
                type = 'cellIs'
                operator = 'greaterThan'
                operand = 10
                fillColor = '#ffc7ce'
                fontColor = '#9c0006'
                bold = $true
            }
        },
        @{
            kind = 'addConditionalFormatting'
            sheetName = 'Donnees'
            rangeRef = 'H2:H10'
            rule = @{
                type = 'containsText'
                text = 'Budget'
                fillColor = '#ffc7ce'
                fontColor = '#9c0006'
                bold = $true
            }
        },
        @{
            kind = 'addConditionalFormatting'
            sheetName = 'Donnees'
            rangeRef = 'I2:I10'
            rule = @{
                type = 'colorScale'
                colors = @('#f8696b', '#ffeb84', '#63be7b')
            }
        },
        @{
            kind = 'addConditionalFormatting'
            sheetName = 'Donnees'
            rangeRef = 'J2:J10'
            rule = @{
                type = 'dataBar'
                color = '#5b9bd5'
            }
        },
        @{
            kind = 'addConditionalFormatting'
            sheetName = 'Donnees'
            rangeRef = 'K2:K10'
            rule = @{
                type = 'iconSet'
                iconSet = '3TrafficLights1'
                thresholds = @(33, 67)
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
    Test-Condition 'Column width edit' (
        [Math]::Abs($afterWorkbook.ColumnTWidth - 20) -lt 0.01
    ) ([string]$afterWorkbook.ColumnTWidth)
    Test-Condition 'Row height edit' (
        [Math]::Abs($afterWorkbook.Row10Height - 36) -lt 0.01
    ) ([string]$afterWorkbook.Row10Height)
    Test-Condition 'Five conditional rules appended' (
        $afterWorkbook.TotalConditionalFormattingCount -eq
        $beforeWorkbook.TotalConditionalFormattingCount + 5
    ) (
        "$($beforeWorkbook.TotalConditionalFormattingCount) -> " +
        "$($afterWorkbook.TotalConditionalFormattingCount)"
    )
    Test-Condition 'Cell-is rule applied' (
        $afterWorkbook.CellIsRule.Count -eq 1 -and
        $afterWorkbook.CellIsRule.Type -eq 1 -and
        $afterWorkbook.CellIsRule.Operator -eq 5 -and
        $afterWorkbook.CellIsRule.Formula1 -eq '=10' -and
        $afterWorkbook.CellIsRule.FillColor -eq
            (Convert-TestHexToOleColor '#ffc7ce') -and
        $afterWorkbook.CellIsRule.FontColor -eq
            (Convert-TestHexToOleColor '#9c0006') -and
        $afterWorkbook.CellIsRule.Bold -eq $true
    )
    Test-Condition 'Contains-text rule applied' (
        $afterWorkbook.ContainsTextRule.Count -eq 1 -and
        $afterWorkbook.ContainsTextRule.Type -eq 9 -and
        $afterWorkbook.ContainsTextRule.Text -eq 'Budget' -and
        $afterWorkbook.ContainsTextRule.TextOperator -eq 0 -and
        $afterWorkbook.ContainsTextRule.FillColor -eq
            (Convert-TestHexToOleColor '#ffc7ce') -and
        $afterWorkbook.ContainsTextRule.FontColor -eq
            (Convert-TestHexToOleColor '#9c0006') -and
        $afterWorkbook.ContainsTextRule.Bold -eq $true
    )
    Test-Condition 'Color-scale rule applied' (
        $afterWorkbook.ColorScaleRule.Count -eq 1 -and
        $afterWorkbook.ColorScaleRule.Type -eq 3 -and
        (Compare-Object $afterWorkbook.ColorScaleRule.CriteriaTypes @(1, 5, 2)).
            Count -eq 0 -and
        $afterWorkbook.ColorScaleRule.MiddleValue -eq 50 -and
        (Compare-Object $afterWorkbook.ColorScaleRule.Colors @(
            (Convert-TestHexToOleColor '#f8696b'),
            (Convert-TestHexToOleColor '#ffeb84'),
            (Convert-TestHexToOleColor '#63be7b')
        )).Count -eq 0
    )
    Test-Condition 'Data-bar rule applied' (
        $afterWorkbook.DataBarRule.Count -eq 1 -and
        $afterWorkbook.DataBarRule.Type -eq 4 -and
        $afterWorkbook.DataBarRule.MinType -eq 1 -and
        $afterWorkbook.DataBarRule.MaxType -eq 2 -and
        $afterWorkbook.DataBarRule.Color -eq
            (Convert-TestHexToOleColor '#5b9bd5')
    )
    Test-Condition 'Icon-set rule applied' (
        $afterWorkbook.IconSetRule.Count -eq 1 -and
        $afterWorkbook.IconSetRule.Type -eq 6 -and
        $afterWorkbook.IconSetRule.IconSetId -eq 4 -and
        (Compare-Object $afterWorkbook.IconSetRule.Thresholds @(33, 67)).
            Count -eq 0 -and
        (Compare-Object $afterWorkbook.IconSetRule.ThresholdTypes @(3, 3)).
            Count -eq 0 -and
        (Compare-Object $afterWorkbook.IconSetRule.ThresholdOperators @(7, 7)).
            Count -eq 0
    )

    $hashBeforeClear = (Get-FileHash -LiteralPath $workingPath -Algorithm SHA256).
        Hash.ToLowerInvariant()
    $thirdResult = Invoke-NativeEdit $workingPath @(
        @{
            kind = 'clearConditionalFormatting'
            sheetName = 'Donnees'
        }
    )
    [void]$backupPaths.Add([string]$thirdResult.backupPath)
    $clearedWorkbook = Read-WorkbookState $workingPath
    $clearedZip = Get-ZipState $workingPath

    Test-Condition 'Conditional formatting clear-all' (
        $clearedWorkbook.TotalConditionalFormattingCount -eq 0
    ) ([string]$clearedWorkbook.TotalConditionalFormattingCount)
    Test-Condition 'Dimensions survive conditional clear' (
        [Math]::Abs($clearedWorkbook.ColumnTWidth - 20) -lt 0.01 -and
        [Math]::Abs($clearedWorkbook.Row10Height - 36) -lt 0.01
    )
    Test-Condition 'Form button preserved' (
        $clearedWorkbook.ButtonCount -eq $beforeWorkbook.ButtonCount
    )
    Test-Condition 'OOXML entry set preserved' (
        (Compare-Object $beforeZip.Entries $clearedZip.Entries).Count -eq 0
    )

    $opaqueDifferences = @(
        foreach ($name in $beforeZip.OpaqueHashes.Keys) {
            if (
                $beforeZip.OpaqueHashes[$name] -ne
                $clearedZip.OpaqueHashes[$name]
            ) {
                $name
            }
        }
    )
    Test-Condition 'Opaque OOXML parts preserved' (
        $opaqueDifferences.Count -eq 0
    ) ($opaqueDifferences -join ', ')
    Test-Condition 'Persistent backups created' (
        $backupPaths.Count -eq 3 -and
        @(
            $backupPaths |
                Where-Object { Test-Path -LiteralPath $_ -PathType Leaf }
        ).Count -eq 3
    )
    Test-Condition 'Second backup matches pre-edit workbook' (
        (Get-FileHash -LiteralPath $secondResult.backupPath -Algorithm SHA256).
            Hash.ToLowerInvariant() -ceq $hashBeforeSecondEdit
    )
    Test-Condition 'Clear backup matches pre-clear workbook' (
        (Get-FileHash -LiteralPath $thirdResult.backupPath -Algorithm SHA256).
            Hash.ToLowerInvariant() -ceq $hashBeforeClear
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
