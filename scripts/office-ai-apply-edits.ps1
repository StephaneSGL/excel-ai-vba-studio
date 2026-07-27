param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$OperationsPath
)

$ErrorActionPreference = 'Stop'
$MaxOperations = 10000
$MaxPayloadBytes = 4MB
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Release-ComObject {
    param([AllowNull()][object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Get-ExcelProcessId {
    param([object]$ExcelApplication)

    if (-not ('ExcelAiVbaStudio.NativeProcess' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace ExcelAiVbaStudio
{
    public static class NativeProcess
    {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(
            IntPtr hWnd,
            out uint processId
        );
    }
}
'@
    }

    $windowHandle = [IntPtr]([int64]$ExcelApplication.Hwnd)
    [uint32]$processId = 0
    [void][ExcelAiVbaStudio.NativeProcess]::GetWindowThreadProcessId(
        $windowHandle,
        [ref]$processId
    )
    if ($processId -eq 0) {
        throw 'Excel process ID could not be determined.'
    }
    return [int]$processId
}

function Get-RequiredProperty {
    param(
        [object]$Value,
        [string]$Name
    )
    $property = $Value.PSObject.Properties[$Name]
    if ($null -eq $property) {
        throw "Missing required property: $Name"
    }
    return $property.Value
}

function Has-Property {
    param(
        [object]$Value,
        [string]$Name
    )
    return $null -ne $Value.PSObject.Properties[$Name]
}

function Assert-LocalPath {
    param([string]$Path)

    $fullPath = [IO.Path]::GetFullPath($Path)
    if (
        $fullPath.StartsWith('\\') -or
        $fullPath.StartsWith('\\?\UNC\', [StringComparison]::OrdinalIgnoreCase)
    ) {
        throw 'Network and UNC paths are not supported.'
    }

    $root = [IO.Path]::GetPathRoot($fullPath)
    if ([string]::IsNullOrWhiteSpace($root)) {
        throw "Path has no local drive root: $fullPath"
    }
    $drive = New-Object IO.DriveInfo($root)
    if (
        $drive.DriveType -eq [IO.DriveType]::Network -or
        $drive.DriveType -eq [IO.DriveType]::Unknown -or
        $drive.DriveType -eq [IO.DriveType]::NoRootDirectory
    ) {
        throw "Path is not on a verified local drive: $fullPath"
    }
    return $fullPath
}

function Assert-NoReparsePointChain {
    param([string]$Path)

    $fullPath = [IO.Path]::GetFullPath($Path)
    $root = [IO.Path]::GetPathRoot($fullPath)
    if ([string]::IsNullOrWhiteSpace($root)) {
        throw "Path has no root: $fullPath"
    }

    $current = $root
    if (Test-Path -LiteralPath $current) {
        $rootItem = Get-Item -LiteralPath $current -Force
        if (($rootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Reparse point detected in path: $current"
        }
    }

    $relative = $fullPath.Substring($root.Length)
    $parts = $relative.Split(
        [char[]]@(
            [IO.Path]::DirectorySeparatorChar,
            [IO.Path]::AltDirectorySeparatorChar
        ),
        [StringSplitOptions]::RemoveEmptyEntries
    )
    foreach ($part in $parts) {
        $current = [IO.Path]::Combine($current, $part)
        if (-not (Test-Path -LiteralPath $current)) {
            break
        }
        $item = Get-Item -LiteralPath $current -Force
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Reparse point detected in path: $current"
        }
    }
    return $fullPath
}

function Get-FileSha256Hex {
    param([string]$Path)

    $stream = [IO.File]::OpenRead($Path)
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        return [BitConverter]::ToString(
            $sha256.ComputeHash($stream)
        ).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
        $stream.Dispose()
    }
}

function Test-MacroWorkbookPackage {
    param([string]$Path)

    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        foreach ($entryName in @(
            '[Content_Types].xml',
            'xl/workbook.xml'
        )) {
            if ($null -eq $archive.GetEntry($entryName)) {
                throw "Required OOXML entry is missing: $entryName"
            }
        }
        return $null -ne $archive.GetEntry('xl/vbaProject.bin')
    }
    finally {
        $archive.Dispose()
    }
}

function Convert-HexToOleColor {
    param([string]$Hex)
    if ($Hex -notmatch '^#(?<r>[0-9a-fA-F]{2})(?<g>[0-9a-fA-F]{2})(?<b>[0-9a-fA-F]{2})$') {
        throw "Invalid color: $Hex"
    }
    $red = [Convert]::ToInt32($Matches.r, 16)
    $green = [Convert]::ToInt32($Matches.g, 16)
    $blue = [Convert]::ToInt32($Matches.b, 16)
    return $red + ($green * 256) + ($blue * 65536)
}

function Get-NumberFormat {
    param(
        [AllowNull()][string]$Format,
        [object]$ExcelApplication
    )

    $dateSeparator = [string]$ExcelApplication.International(17)
    $timeSeparator = [string]$ExcelApplication.International(18)
    $yearCode = [string]$ExcelApplication.International(19)
    $monthCode = [string]$ExcelApplication.International(20)
    $dayCode = [string]$ExcelApplication.International(21)
    $hourCode = [string]$ExcelApplication.International(22)
    $minuteCode = [string]$ExcelApplication.International(23)
    $secondCode = [string]$ExcelApplication.International(24)
    $generalFormat = [string]$ExcelApplication.International(26)
    $dateOrder = [int]$ExcelApplication.International(32)
    $decimalSeparator = [string]$ExcelApplication.International(3)
    $thousandsSeparator = [string]$ExcelApplication.International(4)

    $dateFmt = switch ($dateOrder) {
        0 { "$monthCode$dateSeparator$dayCode$dateSeparator$yearCode" }
        2 { "$yearCode$dateSeparator$monthCode$dateSeparator$dayCode" }
        default { "$dayCode$dateSeparator$monthCode$dateSeparator$yearCode" }
    }
    $timeFmt = "$hourCode$timeSeparator$minuteCode$timeSeparator$secondCode"
    $numberFmt = "#$thousandsSeparator##0$decimalSeparator" + '00'
    $currencySymbol = [char]0x20AC
    $yenSymbol = [char]0x00A5

    switch ($Format) {
        { [string]::IsNullOrWhiteSpace($_) } { return $generalFormat }
        'normal' { return $generalFormat }
        'text' { return '@' }
        'percent' { return "0$decimalSeparator" + '00%' }
        'rmb' { return "$yenSymbol$numberFmt" }
        'usd' { return "`$$numberFmt" }
        'eur' { return "$currencySymbol$numberFmt" }
        'date' { return $dateFmt }
        'time' { return $timeFmt }
        'datetime' { return "$dateFmt $timeFmt" }
        'duration' { return "[$hourCode]$timeSeparator$minuteCode$timeSeparator$secondCode" }
        'number' { return $numberFmt }
        'number_plain' { return "0$decimalSeparator" + '00' }
        default { return [string]$Format }
    }
}

function Get-LineStyle {
    param([AllowNull()][string]$Style)
    switch ($Style) {
        'dashed' { return -4115 }
        'dotted' { return -4118 }
        'double' { return -4119 }
        'none' { return -4142 }
        default { return 1 }
    }
}

function Get-BorderWeight {
    param([AllowNull()][string]$Style)
    switch ($Style) {
        'medium' { return -4138 }
        'thick' { return 4 }
        default { return 2 }
    }
}

function Apply-BorderPatch {
    param(
        [object]$Range,
        [string]$Side,
        [AllowNull()][object]$Value
    )
    $index = switch ($Side) {
        'left' { 7 }
        'top' { 8 }
        'bottom' { 9 }
        'right' { 10 }
    }
    $border = $null
    try {
        $border = $Range.Borders.Item($index)
        if ($null -eq $Value) {
            $border.LineStyle = -4142
            return
        }
        $style = if ($Value.Count -gt 0) { [string]$Value[0] } else { 'thin' }
        $color = if ($Value.Count -gt 1) { [string]$Value[1] } else { '#000000' }
        $border.LineStyle = Get-LineStyle $style
        $border.Weight = Get-BorderWeight $style
        $border.Color = Convert-HexToOleColor $color
    }
    finally {
        Release-ComObject $border
    }
}

function Apply-CellOperation {
    param(
        [object]$Workbook,
        [object]$ExcelApplication,
        [object]$Operation
    )

    $sheetName = [string](Get-RequiredProperty $Operation 'sheetName')
    $row = [int](Get-RequiredProperty $Operation 'row')
    $column = [int](Get-RequiredProperty $Operation 'column')
    if ($row -lt 1 -or $row -gt 1048576 -or $column -lt 1 -or $column -gt 16384) {
        throw "Cell target outside Excel limits: $sheetName!R${row}C${column}"
    }

    $worksheet = $null
    $cell = $null
    $font = $null
    $interior = $null
    $normalStyle = $null
    $normalFont = $null
    try {
        $worksheet = $Workbook.Worksheets.Item($sheetName)
        $cell = $worksheet.Cells.Item($row, $column)

        if (Has-Property $Operation 'value') {
            $value = $Operation.value
            $kind = [string](Get-RequiredProperty $value 'kind')
            switch ($kind) {
                'blank' { $cell.ClearContents() }
                'formula' { $cell.Formula = [string](Get-RequiredProperty $value 'value') }
                'number' { $cell.Value2 = [double](Get-RequiredProperty $value 'value') }
                'text' { $cell.Value2 = [string](Get-RequiredProperty $value 'value') }
                default { throw "Unsupported value kind: $kind" }
            }
        }

        if (Has-Property $Operation 'style') {
            $style = $Operation.style
            if (Has-Property $style 'align') {
                $alignment = if ($null -eq $style.align) {
                    ''
                } else {
                    [string]$style.align
                }
                $cell.HorizontalAlignment = switch ($alignment) {
                    'left' { -4131 }
                    'center' { -4108 }
                    'right' { -4152 }
                    '' { 1 }
                    default { 1 }
                }
            }
            if (Has-Property $style 'valign') {
                $verticalAlignment = if ($null -eq $style.valign) {
                    ''
                } else {
                    [string]$style.valign
                }
                $cell.VerticalAlignment = switch ($verticalAlignment) {
                    'top' { -4160 }
                    'middle' { -4108 }
                    'bottom' { -4107 }
                    '' { -4107 }
                    default { -4107 }
                }
            }
            if (Has-Property $style 'format') {
                $cell.NumberFormatLocal = Get-NumberFormat $style.format $ExcelApplication
            }
            if (Has-Property $style 'textwrap') {
                $cell.WrapText = [bool]$style.textwrap
            }

            if (
                (Has-Property $style 'color') -or
                (Has-Property $style 'underline') -or
                (Has-Property $style 'strike') -or
                (Has-Property $style 'font')
            ) {
                $font = $cell.Font
                if (Has-Property $style 'color') {
                    if ($null -eq $style.color) {
                        $font.ColorIndex = -4105
                    }
                    else {
                        $font.Color = Convert-HexToOleColor ([string]$style.color)
                    }
                }
                if (Has-Property $style 'underline') {
                    $font.Underline = if ([bool]$style.underline) { 2 } else { -4142 }
                }
                if (Has-Property $style 'strike') {
                    $font.Strikethrough = [bool]$style.strike
                }
                if (Has-Property $style 'font') {
                    $fontPatch = $style.font
                    $needsNormalFont =
                        ((Has-Property $fontPatch 'name') -and $null -eq $fontPatch.name) -or
                        ((Has-Property $fontPatch 'size') -and $null -eq $fontPatch.size) -or
                        ((Has-Property $fontPatch 'bold') -and $null -eq $fontPatch.bold) -or
                        ((Has-Property $fontPatch 'italic') -and $null -eq $fontPatch.italic)
                    if ($needsNormalFont) {
                        $normalStyle = $Workbook.Styles.Item(1)
                        $normalFont = $normalStyle.Font
                    }
                    if (Has-Property $fontPatch 'name') {
                        $font.Name = if ($null -eq $fontPatch.name) {
                            [string]$normalFont.Name
                        } else {
                            [string]$fontPatch.name
                        }
                    }
                    if (Has-Property $fontPatch 'size') {
                        $font.Size = if ($null -eq $fontPatch.size) {
                            [double]$normalFont.Size
                        } else {
                            [double]$fontPatch.size
                        }
                    }
                    if (Has-Property $fontPatch 'bold') {
                        $font.Bold = if ($null -eq $fontPatch.bold) {
                            [bool]$normalFont.Bold
                        } else {
                            [bool]$fontPatch.bold
                        }
                    }
                    if (Has-Property $fontPatch 'italic') {
                        $font.Italic = if ($null -eq $fontPatch.italic) {
                            [bool]$normalFont.Italic
                        } else {
                            [bool]$fontPatch.italic
                        }
                    }
                }
            }

            if (Has-Property $style 'bgcolor') {
                $interior = $cell.Interior
                if ($null -eq $style.bgcolor -or [string]$style.bgcolor -eq '#ffffff') {
                    $interior.Pattern = -4142
                }
                else {
                    $interior.Pattern = 1
                    $interior.Color = Convert-HexToOleColor ([string]$style.bgcolor)
                }
            }

            if (Has-Property $style 'border') {
                foreach ($side in @('top', 'right', 'bottom', 'left')) {
                    if (Has-Property $style.border $side) {
                        Apply-BorderPatch $cell $side $style.border.$side
                    }
                }
            }
        }
    }
    finally {
        Release-ComObject $normalFont
        Release-ComObject $normalStyle
        Release-ComObject $interior
        Release-ComObject $font
        Release-ComObject $cell
        Release-ComObject $worksheet
    }
}

$workbookFullPath = Assert-LocalPath $WorkbookPath
$operationsFullPath = Assert-LocalPath $OperationsPath
if ([IO.Path]::GetExtension($workbookFullPath) -ine '.xlsm') {
    throw 'Native editing currently accepts only .xlsm files.'
}
Assert-NoReparsePointChain $workbookFullPath | Out-Null
Assert-NoReparsePointChain $operationsFullPath | Out-Null
if (-not (Test-Path -LiteralPath $workbookFullPath -PathType Leaf)) {
    throw 'Workbook does not exist.'
}
if (-not (Test-Path -LiteralPath $operationsFullPath -PathType Leaf)) {
    throw 'Operations payload does not exist.'
}
$workbookItem = Get-Item -LiteralPath $workbookFullPath -Force
if (($workbookItem.Attributes -band [IO.FileAttributes]::ReadOnly) -ne 0) {
    throw 'Read-only workbooks cannot be edited.'
}
$operationsItem = Get-Item -LiteralPath $operationsFullPath -Force
if ($operationsItem.Length -gt $MaxPayloadBytes) {
    throw 'Operations payload exceeds 4 MiB.'
}

$payload = Get-Content -Raw -LiteralPath $operationsFullPath -Encoding UTF8 | ConvertFrom-Json
if ([int](Get-RequiredProperty $payload 'version') -ne 2) {
    throw 'Unsupported native edit protocol version. Only version 2 is accepted.'
}
$transactionId = [string](Get-RequiredProperty $payload 'transactionId')
$parsedTransactionId = [Guid]::Empty
if (
    $transactionId -notmatch '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$' -or
    -not [Guid]::TryParseExact($transactionId, 'D', [ref]$parsedTransactionId) -or
    $parsedTransactionId.ToString('D') -cne $transactionId
) {
    throw 'Transaction ID must be a canonical lowercase UUID.'
}
$expectedWorkbookSha256 = [string](
    Get-RequiredProperty $payload 'expectedWorkbookSha256'
)
if ($expectedWorkbookSha256 -notmatch '^[0-9a-f]{64}$') {
    throw 'Expected workbook SHA-256 must be 64 lowercase hexadecimal characters.'
}
$operations = @(Get-RequiredProperty $payload 'operations')
if ($operations.Count -lt 1 -or $operations.Count -gt $MaxOperations) {
    throw "Operation count must be between 1 and $MaxOperations."
}

$initialWorkbookSha256 = Get-FileSha256Hex $workbookFullPath
if ($initialWorkbookSha256 -cne $expectedWorkbookSha256) {
    throw 'Workbook changed before the native edit transaction started.'
}
$sourceHasVbaProject = Test-MacroWorkbookPackage $workbookFullPath

$workbookDirectory = [IO.Path]::GetDirectoryName($workbookFullPath)
$workbookBaseName = [IO.Path]::GetFileNameWithoutExtension($workbookFullPath)
$workPath = Join-Path $workbookDirectory (
    '.' + $workbookBaseName +
    '.excel-ai-native-edit.' + $transactionId + '.xlsm'
)
Assert-NoReparsePointChain $workPath | Out-Null
if (Test-Path -LiteralPath $workPath) {
    throw "Native edit work file already exists: $workPath"
}

$excel = $null
$workbooks = $null
$workbook = $null
$blankWorkbook = $null

try {
    Copy-Item -LiteralPath $workbookFullPath -Destination $workPath
    Assert-NoReparsePointChain $workPath | Out-Null
    if ((Get-FileSha256Hex $workPath) -cne $expectedWorkbookSha256) {
        throw 'Native edit work copy does not match the source workbook.'
    }

    try {
        $excel = New-Object -ComObject Excel.Application
        $ownedExcelProcessId = Get-ExcelProcessId $excel
        [Console]::Out.WriteLine("OWNED_EXCEL_PID|$ownedExcelProcessId")
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.ScreenUpdating = $false
        $excel.AskToUpdateLinks = $false
        $excel.AutomationSecurity = 3
        $workbooks = $excel.Workbooks
        $blankWorkbook = $workbooks.Add()
        $excel.Calculation = -4135
        $excel.CalculateBeforeSave = $false
        $workbook = $workbooks.Open($workPath, 0, $false)

        foreach ($operation in $operations) {
            Apply-CellOperation $workbook $excel $operation
        }

        $workbook.Save()
    }
    finally {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch { }
        }
        Release-ComObject $workbook
        $workbook = $null
        if ($null -ne $blankWorkbook) {
            try { $blankWorkbook.Close($false) } catch { }
        }
        Release-ComObject $blankWorkbook
        $blankWorkbook = $null
        Release-ComObject $workbooks
        $workbooks = $null
        if ($null -ne $excel) {
            try { $excel.Quit() } catch { }
        }
        Release-ComObject $excel
        $excel = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    if ((Test-MacroWorkbookPackage $workPath) -ne $sourceHasVbaProject) {
        throw 'The VBA project presence changed in the native edit work copy.'
    }
    if ((Get-FileSha256Hex $workbookFullPath) -cne $expectedWorkbookSha256) {
        throw 'Workbook changed while the native edit transaction was running.'
    }

    $backupDirectory = Join-Path $workbookDirectory '.excel-ai-vba-backups'
    Assert-NoReparsePointChain $backupDirectory | Out-Null
    if (-not (Test-Path -LiteralPath $backupDirectory)) {
        New-Item -ItemType Directory -Path $backupDirectory | Out-Null
    }
    Assert-NoReparsePointChain $backupDirectory | Out-Null
    $backupDirectoryItem = Get-Item -LiteralPath $backupDirectory -Force
    if (-not $backupDirectoryItem.PSIsContainer) {
        throw "Backup path is not a directory: $backupDirectory"
    }

    $timestamp = [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ')
    $backupPath = Join-Path $backupDirectory (
        $workbookBaseName + '.before-native-edit.' +
        $timestamp + '.' + $transactionId + '.xlsm'
    )
    Assert-NoReparsePointChain $backupPath | Out-Null
    if (Test-Path -LiteralPath $backupPath) {
        throw "Backup path already exists: $backupPath"
    }

    [IO.File]::Replace(
        $workPath,
        $workbookFullPath,
        $backupPath,
        $true
    )

    Assert-NoReparsePointChain $backupPath | Out-Null
    Assert-NoReparsePointChain $workbookFullPath | Out-Null
    if ((Get-FileSha256Hex $backupPath) -cne $expectedWorkbookSha256) {
        throw 'Persistent backup does not match the pre-edit workbook.'
    }
    if ((Test-MacroWorkbookPackage $backupPath) -ne $sourceHasVbaProject) {
        throw 'The persistent backup has an unexpected VBA project state.'
    }
    if ((Test-MacroWorkbookPackage $workbookFullPath) -ne $sourceHasVbaProject) {
        throw 'The VBA project presence changed in the saved workbook.'
    }
    $workbookSha256 = Get-FileSha256Hex $workbookFullPath
    @{
        ok = $true
        backupPath = $backupPath
        workbookSha256 = $workbookSha256
    } |
        ConvertTo-Json -Compress |
        ForEach-Object { [Console]::Out.WriteLine($_) }
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch { }
    }
    Release-ComObject $workbook
    if ($null -ne $blankWorkbook) {
        try { $blankWorkbook.Close($false) } catch { }
    }
    Release-ComObject $blankWorkbook
    Release-ComObject $workbooks
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
    }
    Release-ComObject $excel

    if (Test-Path -LiteralPath $workPath -PathType Leaf) {
        Remove-Item -LiteralPath $workPath -Force -ErrorAction SilentlyContinue
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
