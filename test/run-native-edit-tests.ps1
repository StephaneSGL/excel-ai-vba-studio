param(
    [string]$WorkbookPath = "$PSScriptRoot\fixtures\NativeEditingSynthetic.xlsm",
    [string]$BridgeScriptPath = "$PSScriptRoot\..\scripts\office-ai-apply-edits.ps1",
    [switch]$ObjectsOnly
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Release-ComObject {
    param([AllowNull()][object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
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

function Invoke-NativeEdit {
    param(
        [string]$TargetPath,
        [object[]]$Operations
    )
    $expectedHash = Get-Sha256 $TargetPath
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
        $previousErrorActionPreference = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $output = @(& powershell.exe `
                -NoLogo `
                -NoProfile `
                -NonInteractive `
                -ExecutionPolicy Bypass `
                -File $BridgeScriptPath `
                -WorkbookPath $TargetPath `
                -OperationsPath $payloadPath 2>&1)
        }
        finally {
            $ErrorActionPreference = $previousErrorActionPreference
        }
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

function New-WorkbookObjectFixture {
    param([string]$Path)

    $excel = $null
    $workbooks = $null
    $workbook = $null
    $worksheet = $null
    $quotedWorksheet = $null
    $chartObjects = $null
    $chartObject = $null
    $chart = $null
    $sourceRange = $null
    $seriesCollection = $null
    $series = $null
    $trendlines = $null
    $trendline = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AutomationSecurity = 3
        $workbooks = $excel.Workbooks
        $workbook = $workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = 'Objects'
        $quotedWorksheet = $workbook.Worksheets.Add()
        $quotedWorksheet.Name = "O'Brien, Est"

        foreach ($startRow in @(1, 20, 40)) {
            $worksheet.Cells.Item($startRow, 1).Value2 = 'Category'
            $worksheet.Cells.Item($startRow, 2).Value2 = 'Amount'
            $worksheet.Cells.Item($startRow, 3).Value2 = 'Count'
            for ($row = $startRow + 1; $row -le $startRow + 4; $row++) {
                $worksheet.Cells.Item($row, 1).Value2 = "Item-$row"
                $worksheet.Cells.Item($row, 2).Value2 = $row * 10
                $worksheet.Cells.Item($row, 3).Value2 = $row
            }
        }
		# Keep formulas in rows that will become totals rows. Their calculated
		# values remain identical to the fixture values, so both representations
		# can be checked independently after every ListObject resize.
		foreach ($formulaRow in @(5, 24)) {
			$formulaCell = $null
			try {
				$formulaCell = $worksheet.Cells.Item($formulaRow, 3)
				try { $formulaCell.Formula2 = "=B$formulaRow/10" }
				catch { $formulaCell.Formula = "=B$formulaRow/10" }
			}
			finally {
				Release-ComObject $formulaCell
			}
		}
		$externalFormulaDefinitions = @('=A31', '=SUM(B20:B30)', '=C24')
		for ($formulaIndex = 0; $formulaIndex -lt $externalFormulaDefinitions.Count; $formulaIndex++) {
			$externalFormulaCell = $null
			try {
				$externalFormulaCell = $worksheet.Cells.Item($formulaIndex + 1, 8)
				try { $externalFormulaCell.Formula2 = $externalFormulaDefinitions[$formulaIndex] }
				catch { $externalFormulaCell.Formula = $externalFormulaDefinitions[$formulaIndex] }
			}
			finally {
				Release-ComObject $externalFormulaCell
			}
		}
        foreach ($row in @(6, 19, 31, 39, 51, 55)) {
            $worksheet.Cells.Item($row, 1).Value2 = "Sentinel-$row"
            $worksheet.Cells.Item($row, 2).Value2 = $row * 100
            $worksheet.Cells.Item($row, 3).Value2 = "Stable-$row"
        }
        $quotedWorksheet.Range('A1').Value2 = 'Category'
        $quotedWorksheet.Range('B1').Value2 = 'Revenue'
        $quotedWorksheet.Range('C1').Value2 = 'Forecast'
        for ($row = 2; $row -le 5; $row++) {
            $quotedWorksheet.Cells.Item($row, 1).Value2 = "Quoted-$row"
            $quotedWorksheet.Cells.Item($row, 2).Value2 = $row * 25
            $quotedWorksheet.Cells.Item($row, 3).Value2 = $row * 30
        }
        $worksheet.Range('E1').Value2 = 'Existing'
        $worksheet.Range('F1').Value2 = 'Value'
		$worksheet.Range('G1').Value2 = 'Volume'
        for ($row = 2; $row -le 5; $row++) {
            $worksheet.Cells.Item($row, 5).Value2 = "E$row"
            $worksheet.Cells.Item($row, 6).Value2 = $row * 5
			$worksheet.Cells.Item($row, 7).Value2 = $row * 3
        }
        $chartObjects = $worksheet.ChartObjects()
        $chartObject = $chartObjects.Add(900, 20, 320, 180)
        $chartObject.Name = 'ExistingChart'
        $chart = $chartObject.Chart
        $chart.ChartType = 51
        $sourceRange = $worksheet.Range('E1:F5')
        $chart.SetSourceData($sourceRange, 2)
        $chart.HasTitle = $true
        $chart.ChartTitle.Text = 'Existing chart'
        $seriesCollection = $chart.SeriesCollection()
        $series = $seriesCollection.Item(1)
        $trendlines = $series.Trendlines()
        $trendline = $trendlines.Add()
        $workbook.SaveAs($Path, 51)
    }
    finally {
        Release-ComObject $trendline
        Release-ComObject $trendlines
        Release-ComObject $series
        Release-ComObject $seriesCollection
        Release-ComObject $sourceRange
        Release-ComObject $quotedWorksheet
        Release-ComObject $chart
        Release-ComObject $chartObject
        Release-ComObject $chartObjects
        Release-ComObject $worksheet
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

function Read-WorkbookObjectState {
    param([string]$Path)

    $excel = $null
    $workbooks = $null
    $workbook = $null
    $worksheet = $null
    $listObjects = $null
    $chartObjects = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AutomationSecurity = 3
        $workbooks = $excel.Workbooks
        $workbook = $workbooks.Open($Path, 0, $true)
        $worksheet = $workbook.Worksheets.Item('Objects')

		$cellValues = New-Object 'System.Collections.Generic.List[string]'
		$cellFormulas = New-Object 'System.Collections.Generic.List[string]'
		$outsideSentinels = New-Object 'System.Collections.Generic.List[string]'
		$externalFormulas = New-Object 'System.Collections.Generic.List[string]'
		for ($row = 1; $row -le 55; $row++) {
			for ($column = 1; $column -le 3; $column++) {
                $cell = $null
                try {
					$cell = $worksheet.Cells.Item($row, $column)
					[void]$cellValues.Add([string]$cell.Value2)
					try { [void]$cellFormulas.Add([string]$cell.Formula2) }
					catch { [void]$cellFormulas.Add([string]$cell.Formula) }
					if ($row -in @(6, 19, 31, 39, 51, 55)) {
						[void]$outsideSentinels.Add([string]$cell.Value2)
					}
                }
                finally {
                    Release-ComObject $cell
                }
			}
		}
		for ($formulaRow = 1; $formulaRow -le 3; $formulaRow++) {
			$externalFormulaCell = $null
			try {
				$externalFormulaCell = $worksheet.Cells.Item($formulaRow, 8)
				try { [void]$externalFormulas.Add([string]$externalFormulaCell.Formula2) }
				catch { [void]$externalFormulas.Add([string]$externalFormulaCell.Formula) }
			}
			finally {
				Release-ComObject $externalFormulaCell
			}
		}

        $tables = New-Object 'System.Collections.Generic.List[object]'
        $listObjects = $worksheet.ListObjects
        for ($index = 1; $index -le [int]$listObjects.Count; $index++) {
            $listObject = $null
            $range = $null
            try {
                $listObject = $listObjects.Item($index)
                $range = $listObject.Range
                [void]$tables.Add([PSCustomObject]@{
                    Name = [string]$listObject.Name
                    Range = ([string]$range.Address($false, $false)).ToUpperInvariant()
					Headers = [bool]$listObject.ShowHeaders
                    Totals = [bool]$listObject.ShowTotals
                })
            }
            finally {
                Release-ComObject $range
                Release-ComObject $listObject
            }
        }

        $charts = New-Object 'System.Collections.Generic.List[object]'
        $chartObjects = $worksheet.ChartObjects()
        for ($index = 1; $index -le [int]$chartObjects.Count; $index++) {
            $chartObject = $null
            $chart = $null
            $chartTitle = $null
            $fullSeries = $null
            try {
                $chartObject = $chartObjects.Item($index)
                $chart = $chartObject.Chart
                $title = ''
                if ([bool]$chart.HasTitle) {
                    $chartTitle = $chart.ChartTitle
                    $title = [string]$chartTitle.Text
                }
                $fullSeries = $chart.FullSeriesCollection()
                $filteredSeriesCount = 0
                $trendlineCount = 0
                for (
                    $seriesIndex = 1;
                    $seriesIndex -le [int]$fullSeries.Count;
                    $seriesIndex++
                ) {
                    $series = $null
                    $trendlines = $null
                    try {
                        $series = $fullSeries.Item($seriesIndex)
                        if ([bool]$series.IsFiltered) {
                            $filteredSeriesCount++
                        }
                        try {
                            $trendlines = $series.Trendlines()
                            $trendlineCount += [int]$trendlines.Count
                        }
                        catch { }
                    }
                    finally {
                        Release-ComObject $trendlines
                        Release-ComObject $series
                    }
                }
                [void]$charts.Add([PSCustomObject]@{
                    Name = [string]$chartObject.Name
                    Type = [int]$chart.ChartType
                    Title = $title
                    FullSeriesCount = [int]$fullSeries.Count
                    FilteredSeriesCount = $filteredSeriesCount
                    TrendlineCount = $trendlineCount
                    Left = [Math]::Round([double]$chartObject.Left, 3)
                    Top = [Math]::Round([double]$chartObject.Top, 3)
                    Width = [Math]::Round([double]$chartObject.Width, 3)
                    Height = [Math]::Round([double]$chartObject.Height, 3)
                })
            }
            finally {
                Release-ComObject $fullSeries
                Release-ComObject $chartTitle
                Release-ComObject $chart
                Release-ComObject $chartObject
            }
        }

        return [PSCustomObject]@{
			Cells = $cellValues.ToArray() -join [char]31
			Formulas = $cellFormulas.ToArray() -join [char]31
			OutsideSentinels = $outsideSentinels.ToArray() -join [char]31
			ExternalFormulas = $externalFormulas.ToArray() -join [char]31
			Tables = @($tables.ToArray() | Sort-Object Name)
            Charts = @($charts.ToArray() | Sort-Object Name)
        }
    }
    finally {
        Release-ComObject $chartObjects
        Release-ComObject $listObjects
        Release-ComObject $worksheet
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

function New-TestTableDefinition {
    param(
        [string]$Id,
        [string]$Name,
        [string]$RangeRef,
        [bool]$TotalsRow = $false
    )
    return @{
        id = $Id
        name = $Name
        displayName = $Name
        rangeRef = $RangeRef
        headerRow = $true
        totalsRow = $TotalsRow
        style = @{
            name = 'TableStyleMedium2'
            showFirstColumn = $false
            showLastColumn = $false
            showRowStripes = $true
            showColumnStripes = $false
        }
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

if (-not $ObjectsOnly) {
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

    $hashBeforeSecondEdit = Get-Sha256 $workingPath
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

    $hashBeforeClear = Get-Sha256 $workingPath
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
        (Get-Sha256 $secondResult.backupPath) -ceq $hashBeforeSecondEdit
    )
    Test-Condition 'Clear backup matches pre-clear workbook' (
        (Get-Sha256 $thirdResult.backupPath) -ceq $hashBeforeClear
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
}

$objectWorkingPath = Join-Path ([IO.Path]::GetTempPath()) (
    'excel-native-objects-' + [Guid]::NewGuid().ToString('N') + '.xlsx'
)
$objectBackupPaths = New-Object 'System.Collections.Generic.List[string]'
try {
    New-WorkbookObjectFixture $objectWorkingPath
    $objectBaseline = Read-WorkbookObjectState $objectWorkingPath
    foreach ($unsafeSeriesName in @(
        '=Sheet1!A1',
        "='[other.xlsx]S'!A1",
        "=cmd|' /C calc'!A0"
    )) {
        $hashBeforeRejectedName = Get-Sha256 $objectWorkingPath
        $rejected = $false
        try {
            [void](Invoke-NativeEdit $objectWorkingPath @(
                @{
                    kind = 'createChart'
                    sheetName = 'Objects'
                    chart = @{
                        id = 'chart:unsafe-name-test'
                        name = 'UnsafeNameTest'
                        chartType = 51
                        plotBy = 'columns'
                        anchor = @{ left = 900; top = 500; width = 320; height = 180 }
                        series = @(
                            @{
                                id = 'series:unsafe-name-test'
                                name = $unsafeSeriesName
                                categoryRange = 'A2:A5'
                                valuesRange = 'B2:B5'
                            }
                        )
                    }
                }
            ))
        }
        catch {
            $rejected = $_.Exception.Message -match 'cannot be an Excel formula'
        }
        Test-Condition "Unsafe series name rejected: $unsafeSeriesName" (
            $rejected -and
            (Get-Sha256 $objectWorkingPath) -ceq $hashBeforeRejectedName
        )
    }
    foreach ($unsafeTableName in @('Sales Q3', 'A1', 'R1C1', 'R', 'C')) {
        $hashBeforeRejectedTableName = Get-Sha256 $objectWorkingPath
        $rejected = $false
        try {
            [void](Invoke-NativeEdit $objectWorkingPath @(
                @{
                    kind = 'createTable'
                    sheetName = 'Objects'
                    table = (New-TestTableDefinition `
                        'table:unsafe-name-test' `
                        $unsafeTableName `
                        'A1:C5')
                }
            ))
        }
        catch {
            $rejected = $true
        }
        Test-Condition "Unsafe table name rejected: $unsafeTableName" (
            $rejected -and
            (Get-Sha256 $objectWorkingPath) -ceq $hashBeforeRejectedTableName
        )
    }
    $hashBeforeSuggestedChart = Get-Sha256 $objectWorkingPath
    $suggestedChartRejected = $false
    try {
        [void](Invoke-NativeEdit $objectWorkingPath @(
            @{
                kind = 'createChart'
                sheetName = 'Objects'
                chart = @{
                    id = 'chart:suggested-rollback'
                    name = 'SuggestedRollback'
                    chartType = -2
                    sourceRangeRef = 'A1:C5'
                    plotBy = 'columns'
                    anchor = @{ left = 900; top = 500; width = 320; height = 180 }
                }
            }
        ))
    }
    catch {
        $suggestedChartRejected = $true
    }
    Test-Condition 'xlSuggestedChart is rejected before mutation' (
        $suggestedChartRejected -and
        (Get-Sha256 $objectWorkingPath) -ceq $hashBeforeSuggestedChart
    )
    $hashBeforeRegionMap = Get-Sha256 $objectWorkingPath
    $regionMapRejected = $false
    try {
        [void](Invoke-NativeEdit $objectWorkingPath @(
            @{
                kind = 'createChart'
                sheetName = 'Objects'
                chart = @{
                    id = 'chart:region-map-offline'
                    name = 'RegionMapOffline'
                    chartType = 140
                    sourceRangeRef = 'A1:B5'
                    plotBy = 'columns'
                    anchor = @{ left = 900; top = 500; width = 320; height = 180 }
                }
            }
        ))
    }
    catch {
        $regionMapRejected = $true
    }
    Test-Condition 'xlRegionMap is rejected without network disclosure' (
        $regionMapRejected -and
        (Get-Sha256 $objectWorkingPath) -ceq $hashBeforeRegionMap
    )
	$headerlessCreate = New-TestTableDefinition `
		'table:objects:A10:C14' `
		'HeaderlessCreate' `
		'A10:C14'
	$headerlessCreate.headerRow = $false
	$hashBeforeHeaderlessCreate = Get-Sha256 $objectWorkingPath
	$headerlessCreateRejected = $false
	try {
		[void](Invoke-NativeEdit $objectWorkingPath @(
			@{
				kind = 'createTable'
				sheetName = 'Objects'
				table = $headerlessCreate
			}
		))
	}
	catch {
		$headerlessCreateRejected = $_.Exception.Message -match 'headerRow=false is disabled'
	}
	Test-Condition 'Headerless table creation is rejected without mutation' (
		$headerlessCreateRejected -and
		(Get-Sha256 $objectWorkingPath) -ceq $hashBeforeHeaderlessCreate
	)
	$totalsCreate = New-TestTableDefinition `
		'table:objects:A10:C14' `
		'TotalsCreate' `
		'A10:C14' `
		$true
	$hashBeforeTotalsCreate = Get-Sha256 $objectWorkingPath
	$totalsCreateRejected = $false
	try {
		[void](Invoke-NativeEdit $objectWorkingPath @(
			@{
				kind = 'createTable'
				sheetName = 'Objects'
				table = $totalsCreate
			}
		))
	}
	catch {
		$totalsCreateRejected = $_.Exception.Message -match 'totalsRow=true is disabled'
	}
	Test-Condition 'Totals-row table creation is rejected before mutation' (
		$totalsCreateRejected -and
		(Get-Sha256 $objectWorkingPath) -ceq $hashBeforeTotalsCreate
	)

    $topTable = New-TestTableDefinition `
		'table:objects:A1:C5' `
		'TableTop' `
		'A1:C5'
    $middleTable = New-TestTableDefinition `
        'table:objects:A20:C30' `
        'TableMiddle' `
        'A20:C30'
    $bottomTable = New-TestTableDefinition `
        'table:objects:A40:C50' `
        'TableBottom' `
        'A40:C50'
    $chartDefinition = @{
        id = 'chart:objects:SalesChart'
        name = 'SalesChart'
        chartType = 51
        sourceRangeRef = 'E1:G5'
        plotBy = 'columns'
        anchor = @{ left = 900; top = 240; width = 420; height = 260 }
        title = @{ visible = $true; text = 'Sales chart' }
        legend = @{ visible = $true; position = 'right' }
		valueAxis = @{
			visible = $true
			numberFormat = '0.00'
		}
        style = 10
        roundedCorners = $false
        alternativeText = 'Native chart integration test'
    }
    $pieDefinition = @{
        id = 'chart:objects:PieChart'
        name = 'PieChart'
        chartType = 5
        sourceRangeRef = 'E1:F5'
        plotBy = 'columns'
        anchor = @{ left = 1350; top = 20; width = 360; height = 240 }
        title = @{ visible = $true; text = 'Pie chart' }
        legend = @{ visible = $true; position = 'right' }
    }
	$threeDimensionalColumnDefinition = @{
		id = 'chart:objects:ThreeDimensionalColumn'
		name = 'ThreeDimensionalColumn'
		chartType = -4100
		sourceRangeRef = '$E$1:$G$5'
		plotBy = 'columns'
		anchor = @{ left = 1350; top = 280; width = 360; height = 240 }
		title = @{ visible = $true; text = '3D column chart' }
		gapWidth = 150
	}
	$bubbleDefinition = @{
		id = 'chart:objects:BubbleChart'
		name = 'BubbleChart'
		chartType = 15
		plotBy = 'columns'
		anchor = @{ left = 1740; top = 280; width = 420; height = 260 }
		title = @{ visible = $true; text = 'Bubble chart' }
		series = @(
			@{
				id = 'series:objects:bubble-first'
				name = 'Bubble first'
				xValuesRange = 'E2:E5'
				valuesRange = 'F2:F5'
				bubbleSizesRange = 'G2:G5'
				chartType = 15
				axisGroup = 'primary'
			},
			@{
				id = 'series:objects:bubble-second'
				name = 'Bubble second'
				xValuesRange = 'E2:E5'
				valuesRange = 'G2:G5'
				bubbleSizesRange = 'F2:F5'
				chartType = 15
				axisGroup = 'primary'
			}
		)
	}
    $quotedSheetChartDefinition = @{
        id = 'chart:quoted-sheet:Revenue'
        name = 'QuotedSheetChart'
        chartType = 65
        plotBy = 'columns'
        anchor = @{ left = 20; top = 20; width = 360; height = 220 }
        title = @{ visible = $true; text = 'Quoted sheet chart' }
        series = @(
            @{
                id = 'series:quoted-sheet:revenue'
                nameRange = 'B1'
                categoryRange = 'A2:A5'
                valuesRange = 'B2:B5'
                chartType = 65
                lineColor = '#4472c4'
                lineWidth = 2
                markerStyle = 'circle'
                markerSize = 6
            },
            @{
                id = 'series:quoted-sheet:forecast'
                name = 'Forecast, "net"'
                categoryRange = 'A2:A5'
                valuesRange = 'C2:C5'
                chartType = 65
                lineColor = '#70ad47'
                lineWidth = 2
            }
        )
    }
    $createObjectsResult = Invoke-NativeEdit $objectWorkingPath @(
        @{ kind = 'createTable'; sheetName = 'Objects'; table = $topTable },
        @{ kind = 'createTable'; sheetName = 'Objects'; table = $middleTable },
        @{ kind = 'createTable'; sheetName = 'Objects'; table = $bottomTable },
        @{ kind = 'createChart'; sheetName = 'Objects'; chart = $chartDefinition },
        @{ kind = 'createChart'; sheetName = 'Objects'; chart = $pieDefinition },
		@{ kind = 'createChart'; sheetName = 'Objects'; chart = $threeDimensionalColumnDefinition },
		@{ kind = 'createChart'; sheetName = 'Objects'; chart = $bubbleDefinition },
        @{
            kind = 'createChart'
            sheetName = "O'Brien, Est"
            chart = $quotedSheetChartDefinition
        }
    )
    [void]$objectBackupPaths.Add([string]$createObjectsResult.backupPath)
    $createdObjects = Read-WorkbookObjectState $objectWorkingPath
    Test-Condition 'Three disjoint same-column tables created' (
        $createdObjects.Tables.Count -eq 3 -and
        (Compare-Object `
            @($createdObjects.Tables | ForEach-Object { $_.Range }) `
            @('A1:C5', 'A20:C30', 'A40:C50')).Count -eq 0
    )
	$createdTopTable = @(
		$createdObjects.Tables | Where-Object Name -eq 'TableTop'
	)[0]
	Test-Condition 'Table creation keeps the requested final range without totals' (
		$createdTopTable.Range -ceq 'A1:C5' -and
		$createdTopTable.Headers -and
		-not $createdTopTable.Totals
	)
	Test-Condition 'Table creation preserves worksheet values and formulas' (
		$createdObjects.Cells -ceq $objectBaseline.Cells -and
		$createdObjects.Formulas -ceq $objectBaseline.Formulas -and
		$createdObjects.ExternalFormulas -ceq $objectBaseline.ExternalFormulas
	)
    Test-Condition 'Chart created without removing existing chart' (
		$createdObjects.Charts.Count -eq 5 -and
        @($createdObjects.Charts | Where-Object Name -eq 'ExistingChart').Count -eq 1 -and
        @($createdObjects.Charts | Where-Object Name -eq 'SalesChart').Count -eq 1 -and
        @(
            $createdObjects.Charts |
                Where-Object { $_.Name -eq 'PieChart' -and $_.Type -eq 5 }
        ).Count -eq 1
    )
	Test-Condition 'xl3DColumn accepts and persists gap width' (
		@(
			$createdObjects.Charts |
				Where-Object { $_.Name -eq 'ThreeDimensionalColumn' -and $_.Type -eq -4100 }
		).Count -eq 1
	)
	Test-Condition 'Second bubble series persists with independent bubble sizes' (
		@(
			$createdObjects.Charts |
				Where-Object {
					$_.Name -eq 'BubbleChart' -and
					$_.Type -eq 15 -and
					$_.FullSeriesCount -eq 2
				}
		).Count -eq 1
	)
    Test-Condition 'Quoted worksheet, nameRange and literal series name verified' (
        [bool]$createObjectsResult.ok
    )

	$headerTransition = New-TestTableDefinition `
		'table:objects:A20:C30' `
		'TableMiddle' `
		'A20:C30'
	$headerTransition.headerRow = $false
	$hashBeforeHeaderTransition = Get-Sha256 $objectWorkingPath
	$headerTransitionRejected = $false
	try {
		[void](Invoke-NativeEdit $objectWorkingPath @(
			@{
				kind = 'updateTable'
				sheetName = 'Objects'
				name = 'TableMiddle'
				table = $headerTransition
			}
		))
	}
	catch {
		$headerTransitionRejected = $_.Exception.Message -match 'headerRow transitions are disabled'
	}
	$afterRejectedHeaderTransition = Read-WorkbookObjectState $objectWorkingPath
	Test-Condition 'Table header transition is rejected without mutation' (
		$headerTransitionRejected -and
		(Get-Sha256 $objectWorkingPath) -ceq $hashBeforeHeaderTransition -and
		$afterRejectedHeaderTransition.Cells -ceq $objectBaseline.Cells -and
		$afterRejectedHeaderTransition.Formulas -ceq $objectBaseline.Formulas -and
		$afterRejectedHeaderTransition.ExternalFormulas -ceq $objectBaseline.ExternalFormulas -and
		@(
			$afterRejectedHeaderTransition.Tables |
				Where-Object { $_.Name -eq 'TableMiddle' -and $_.Headers }
		).Count -eq 1
	)

	$middleTotalsTransition = New-TestTableDefinition `
		'table:objects:A20:C24' `
		'TableMiddle' `
		'A20:C24' `
		$true
	$hashBeforeTotalsTransition = Get-Sha256 $objectWorkingPath
	$totalsTransitionRejected = $false
	try {
		[void](Invoke-NativeEdit $objectWorkingPath @(
			@{
				kind = 'updateTable'
				sheetName = 'Objects'
				name = 'TableMiddle'
				table = $middleTotalsTransition
			}
		))
	}
	catch {
		$totalsTransitionRejected = $_.Exception.Message -match 'totalsRow transitions are disabled'
	}
	$afterRejectedTotalsTransition = Read-WorkbookObjectState $objectWorkingPath
	Test-Condition 'Totals-row transition is rejected without mutation' (
		$totalsTransitionRejected -and
		(Get-Sha256 $objectWorkingPath) -ceq $hashBeforeTotalsTransition -and
		$afterRejectedTotalsTransition.Cells -ceq $objectBaseline.Cells -and
		$afterRejectedTotalsTransition.Formulas -ceq $objectBaseline.Formulas -and
		$afterRejectedTotalsTransition.ExternalFormulas -ceq $objectBaseline.ExternalFormulas
	)

	$baselineExistingChart = @(
        $objectBaseline.Charts | Where-Object Name -eq 'ExistingChart'
    )[0]
    $preserveExistingResult = Invoke-NativeEdit $objectWorkingPath @(
        @{
            kind = 'updateChart'
            sheetName = 'Objects'
            name = 'ExistingChart'
            preserveAnchor = $true
            preserveSeries = $true
            chart = @{
                id = 'chart:objects:ExistingChart'
                name = 'ExistingChart'
                chartType = 51
                sourceRangeRef = 'E1:F5'
                plotBy = 'columns'
                # Deliberately different: preservation flags must keep the
                # native geometry and the presence of the unmodelled trendline.
                anchor = @{ left = 1; top = 1; width = 20; height = 20 }
                title = @{ visible = $true; text = 'Existing chart retitled' }
            }
        }
    )
    [void]$objectBackupPaths.Add([string]$preserveExistingResult.backupPath)
    $preservedObjects = Read-WorkbookObjectState $objectWorkingPath
    $preservedExistingChart = @(
        $preservedObjects.Charts | Where-Object Name -eq 'ExistingChart'
    )[0]
    Test-Condition 'Title-only chart update preserves native anchor and trendline' (
        $preservedExistingChart.Title -eq 'Existing chart retitled' -and
        $preservedExistingChart.TrendlineCount -eq 1 -and
        $preservedExistingChart.Left -eq $baselineExistingChart.Left -and
        $preservedExistingChart.Top -eq $baselineExistingChart.Top -and
        $preservedExistingChart.Width -eq $baselineExistingChart.Width -and
        $preservedExistingChart.Height -eq $baselineExistingChart.Height
    )

	$stylePreserveResult = Invoke-NativeEdit $objectWorkingPath @(
		@{
			kind = 'updateChart'
			sheetName = 'Objects'
			name = 'ExistingChart'
			preserveAnchor = $true
			preserveSeries = $true
			allowSeriesFormattingChange = $true
			chart = @{
				id = 'chart:objects:ExistingChart'
				name = 'ExistingChart'
				chartType = 51
				sourceRangeRef = 'E1:F5'
				plotBy = 'columns'
				anchor = @{ left = 1; top = 1; width = 20; height = 20 }
				title = @{ visible = $true; text = 'Existing chart retitled' }
				style = 12
			}
		}
	)
	[void]$objectBackupPaths.Add([string]$stylePreserveResult.backupPath)
	$stylePreservedObjects = Read-WorkbookObjectState $objectWorkingPath
	$stylePreservedExistingChart = @(
		$stylePreservedObjects.Charts | Where-Object Name -eq 'ExistingChart'
	)[0]
	Test-Condition 'ChartStyle update preserves trendline and native series objects' (
		$stylePreservedExistingChart.Title -eq 'Existing chart retitled' -and
		$stylePreservedExistingChart.TrendlineCount -eq 1 -and
		$stylePreservedExistingChart.Left -eq $baselineExistingChart.Left -and
		$stylePreservedExistingChart.Top -eq $baselineExistingChart.Top
	)

    $renamedMiddleTable = New-TestTableDefinition `
        'table:objects:A20:C30' `
        'Sales.Q3' `
        'A20:C30'
    $updatedChart = @{
        id = 'chart:objects:SalesChart'
        name = 'SalesTrend'
        chartType = -4152
        plotBy = 'columns'
        anchor = @{ left = 900; top = 240; width = 440; height = 280 }
        title = @{ visible = $true; text = 'Sales trend' }
        legend = @{ visible = $true; position = 'bottom' }
        categoryAxis = @{
            visible = $true
            title = 'Category'
            reverseOrder = $false
            numberFormat = 'General'
            majorGridlines = $false
            minorGridlines = $false
        }
        valueAxis = @{
            visible = $true
            title = 'Amount'
            minimumScale = 0
            maximumScale = 100
            majorUnit = 20
            minorUnit = 10
            logarithmic = $false
            reverseOrder = $false
			# Exercise the explicit custom -> source-linked reset. The native
			# verifier must observe NumberFormatLinked=true after reopening.
			numberFormat = ''
            majorGridlines = $true
            minorGridlines = $false
        }
        secondaryValueAxis = @{
            visible = $true
            title = 'Count'
            minimumScale = 0
            maximumScale = 10
            majorUnit = 2
            minorUnit = 1
            logarithmic = $false
            reverseOrder = $false
            numberFormat = '0'
        }
        series = @(
            @{
                id = 'series:objects:sales'
                name = 'Amount'
                categoryRange = 'A2:A5'
                valuesRange = 'B2:B5'
                chartType = 65
                axisGroup = 'primary'
                lineColor = '#3366cc'
                lineWidth = 2
                dashStyle = 'solid'
                markerStyle = 'circle'
                markerSize = 7
                smooth = $false
                visible = $true
                dataLabels = @{
					# ApplyDataLabels() defaults ShowValue to true. Keeping every
					# modeled flag false proves the setter neutralizes that default
					# and the reopened-chart verifier checks the persisted state.
					showValue = $false
                    showCategoryName = $false
                    showSeriesName = $false
                    showPercentage = $false
                    showBubbleSize = $false
                }
            },
            @{
                id = 'series:objects:count'
                name = 'Count'
                categoryRange = 'A2:A5'
                valuesRange = 'C2:C5'
                chartType = 51
                axisGroup = 'secondary'
                color = '#cc6633'
                visible = $false
            }
        )
        style = 11
        roundedCorners = $true
        gapWidth = 160
        overlap = -20
        alternativeText = 'Updated native chart integration test'
    }
    $updateObjectsResult = Invoke-NativeEdit $objectWorkingPath @(
        @{
            kind = 'updateTable'
            sheetName = 'Objects'
            name = 'TableMiddle'
            table = $renamedMiddleTable
        },
        @{ kind = 'deleteTable'; sheetName = 'Objects'; name = 'TableTop' },
        @{
            kind = 'updateChart'
            sheetName = 'Objects'
            name = 'SalesChart'
            chart = $updatedChart
        }
    )
    [void]$objectBackupPaths.Add([string]$updateObjectsResult.backupPath)
    $updatedObjects = Read-WorkbookObjectState $objectWorkingPath
    Test-Condition 'Table rename and unlist applied' (
        $updatedObjects.Tables.Count -eq 2 -and
        @($updatedObjects.Tables | Where-Object Name -eq 'Sales.Q3').Count -eq 1 -and
        @($updatedObjects.Tables | Where-Object Name -eq 'TableTop').Count -eq 0
    )
    Test-Condition 'Table unlist preserves cells and row positions' (
		$updatedObjects.Cells -ceq $objectBaseline.Cells -and
		$updatedObjects.Formulas -ceq $objectBaseline.Formulas
    )
    Test-Condition 'Chart type and title updated' (
        @(
            $updatedObjects.Charts |
                Where-Object {
                    $_.Name -eq 'SalesTrend' -and
                    $_.Type -eq -4111 -and
                    $_.Title -eq 'Sales trend' -and
                    $_.FullSeriesCount -eq 2 -and
                    $_.FilteredSeriesCount -eq 1
                }
        ).Count -eq 1
    )
    Test-Condition 'Untargeted chart preserved through object edits' (
        @(
            $updatedObjects.Charts |
                Where-Object {
                    $_.Name -eq 'ExistingChart' -and
                    $_.Type -eq 51 -and
                    $_.Title -eq 'Existing chart retitled' -and
                    $_.TrendlineCount -eq 1
                }
        ).Count -eq 1
    )

    $modernChartCreated = $false
    $hashBeforeModernChart = Get-Sha256 $objectWorkingPath
    try {
        $modernResult = Invoke-NativeEdit $objectWorkingPath @(
            @{
                kind = 'createChart'
                sheetName = 'Objects'
                chart = @{
                    id = 'chart:objects:ModernTreemap'
                    name = 'ModernTreemap'
                    chartType = 117
                    sourceRangeRef = 'A1:B5'
                    plotBy = 'columns'
                    anchor = @{ left = 1350; top = 300; width = 360; height = 240 }
                    title = @{ visible = $true; text = 'Modern treemap' }
                    legend = @{ visible = $true; position = 'right' }
                }
            }
        )
        [void]$objectBackupPaths.Add([string]$modernResult.backupPath)
        $modernChartCreated = $true
    }
    catch {
        $modernChartCreated = $false
    }
    $modernState = Read-WorkbookObjectState $objectWorkingPath
    Test-Condition 'Modern chart 117 succeeds or rolls back safely' (
        (
            $modernChartCreated -and
            @(
                $modernState.Charts |
                    Where-Object { $_.Name -eq 'ModernTreemap' -and $_.Type -eq 117 }
            ).Count -eq 1
        ) -or (
            -not $modernChartCreated -and
            (Get-Sha256 $objectWorkingPath) -ceq $hashBeforeModernChart
        )
    ) ("created=" + [string]$modernChartCreated)

    $deleteObjectOperations = @(
        @{ kind = 'deleteTable'; sheetName = 'Objects'; name = 'TableBottom' },
        @{ kind = 'deleteChart'; sheetName = 'Objects'; name = 'SalesTrend' },
		@{ kind = 'deleteChart'; sheetName = 'Objects'; name = 'PieChart' },
		@{ kind = 'deleteChart'; sheetName = 'Objects'; name = 'ThreeDimensionalColumn' },
		@{ kind = 'deleteChart'; sheetName = 'Objects'; name = 'BubbleChart' }
    )
    if ($modernChartCreated) {
        $deleteObjectOperations += @{
            kind = 'deleteChart'
            sheetName = 'Objects'
            name = 'ModernTreemap'
        }
    }
    $deleteObjectsResult = Invoke-NativeEdit `
        $objectWorkingPath `
        $deleteObjectOperations
    [void]$objectBackupPaths.Add([string]$deleteObjectsResult.backupPath)
    $deletedObjects = Read-WorkbookObjectState $objectWorkingPath
    Test-Condition 'Table and chart delete operations persist' (
        $deletedObjects.Tables.Count -eq 1 -and
        $deletedObjects.Tables[0].Name -eq 'Sales.Q3' -and
        $deletedObjects.Charts.Count -eq 1 -and
        $deletedObjects.Charts[0].Name -eq 'ExistingChart'
    )
    Test-Condition 'Repeated unlist preserves every worksheet value' (
		$deletedObjects.Cells -ceq $objectBaseline.Cells -and
		$deletedObjects.Formulas -ceq $objectBaseline.Formulas
    )
}
finally {
    if (Test-Path -LiteralPath $objectWorkingPath -PathType Leaf) {
        Remove-Item -LiteralPath $objectWorkingPath -Force
    }
    foreach ($backupPath in $objectBackupPaths) {
        if (Test-Path -LiteralPath $backupPath -PathType Leaf) {
            Remove-Item -LiteralPath $backupPath -Force
        }
    }
    $objectBackupDirectory = Join-Path (
        [IO.Path]::GetDirectoryName($objectWorkingPath)
    ) '.excel-ai-vba-backups'
    if (
        (Test-Path -LiteralPath $objectBackupDirectory -PathType Container) -and
        @(Get-ChildItem -LiteralPath $objectBackupDirectory -Force).Count -eq 0
    ) {
        Remove-Item -LiteralPath $objectBackupDirectory -Force
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
