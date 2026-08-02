param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$OperationsPath
)

$ErrorActionPreference = 'Stop'
$MaxOperations = 10000
$MaxConditionalFormattingAddsPerSheet = 64
$MaxPayloadBytes = 4MB
$MaxWorkbookObjectOperationsPerSheet = 512
$MaxChartSeries = 255
$MaxWorkbookObjectRangeCells = 1000000
$MaxWorkbookObjectTransactionRangeCells = 5000000
$MaxNamedStreams = 64
$MaxNamedStreamBytes = 8MB
$MaxTotalNamedStreamBytes = 32MB
$MaxZoneIdentifierBytes = 64KB
$script:NativeWorkbookExtension = '.xlsm'
$script:ValidNativeChartTypes = @(
    -4169, -4152, -4151, -4120, -4102, -4101, -4100, -4098,
    1, 4, 5, 15, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62,
    63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77,
    78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92,
    93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105,
    106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117,
    118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128, 129,
    130, 131, 132, 133, 134, 135, 136, 137, 138, 139
)
$script:NativeSecondaryAxisChartTypes = @(
    1, 76, 77, 135, 136, 137,
    57, 58, 59, 132, 133, 134,
    51, 52, 53, 124, 125, 126,
    4, 63, 64, 65, 66, 67, 127, 128, 129,
    -4169, 72, 73, 74, 75, 138,
    15, 87, 139,
    -4152, 113, 114, 115, 116
)
$script:NativeCategoryScaleChartTypes = @(
    -4169, 72, 73, 74, 75, 138,
    15, 87, 139
)
$script:NativeBubbleChartTypes = @(15, 87, 139)
$script:NativeSmoothSeriesChartTypes = @(
    -4169, -4101,
    4, 63, 64, 65, 66, 67, 72, 73, 74, 75,
    127, 128, 129, 138
)
$script:NativePercentageDataLabelChartTypes = @(
    -4120, -4102,
    5, 68, 69, 70, 71, 80, 130, 131
)
$script:NativeGapWidthChartTypes = @(
	-4100,
    51, 52, 53, 54, 55, 56, 124, 125, 126,
    57, 58, 59, 60, 61, 62, 132, 133, 134,
    -4152, 113, 114, 115, 116,
    68, 71, 118, 122
)
$script:NativeOverlapChartTypes = @(
    51, 52, 53, 57, 58, 59,
    124, 125, 126, 132, 133, 134,
    -4152, 113, 114, 115, 116
)
Add-Type -AssemblyName System.IO.Compression.FileSystem
. (Join-Path $PSScriptRoot 'ooxml-package-signature.ps1')

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

function Get-ExcelProcessIdentity {
    param([object]$ExcelApplication)

    $processId = Get-ExcelProcessId $ExcelApplication
    $process = $null
    try {
        $process = [Diagnostics.Process]::GetProcessById($processId)
        if (
            -not [string]::Equals(
                $process.ProcessName,
                'EXCEL',
                [StringComparison]::OrdinalIgnoreCase
            )
        ) {
            throw 'The owned automation process is not EXCEL.EXE.'
        }
        return [PSCustomObject]@{
            ProcessId = $processId
            StartTimeUtcTicks = $process.StartTime.ToUniversalTime().Ticks.ToString(
                [Globalization.CultureInfo]::InvariantCulture
            )
        }
    }
    finally {
        if ($null -ne $process) { $process.Dispose() }
    }
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

function Get-StreamSha256Hex {
    param([IO.Stream]$Stream)

    if (-not $Stream.CanSeek) {
        throw 'Workbook lock stream must support seeking.'
    }
    $originalPosition = $Stream.Position
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        $Stream.Position = 0
        return [BitConverter]::ToString(
            $sha256.ComputeHash($Stream)
        ).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $Stream.Position = $originalPosition
        $sha256.Dispose()
    }
}

function Get-BoundedNamedStreamBytes {
    param(
        [string]$Path,
        [string]$StreamName,
        [long]$MaximumBytes,
        [string]$Label
    )

    $memory = $null
    try {
        $memory = New-Object IO.MemoryStream
        Get-Content `
            -LiteralPath $Path `
            -Stream $StreamName `
            -Encoding Byte `
            -ReadCount 8192 `
            -ErrorAction Stop |
            ForEach-Object {
            $chunk = [byte[]]$_
            if ($memory.Length + $chunk.Length -gt $MaximumBytes) {
                throw "$Label exceeds the $MaximumBytes-byte safety limit."
            }
            $memory.Write($chunk, 0, $chunk.Length)
            }
        return $memory.ToArray()
    }
    finally {
        if ($null -ne $memory) { $memory.Dispose() }
    }
}

function Get-NamedStreamState {
    param([string]$Path)

    $streamItems = @(Get-Item -LiteralPath $Path -Stream * -Force -ErrorAction Stop)
    if ($streamItems.Count -gt $MaxNamedStreams) {
        throw "Workbook has more than $MaxNamedStreams alternate data streams."
    }

    [long]$totalBytes = 0
    $seenNames = New-Object 'System.Collections.Generic.HashSet[string]' `
        ([StringComparer]::OrdinalIgnoreCase)
    $result = New-Object 'System.Collections.Generic.List[object]'
    foreach ($streamItem in $streamItems) {
        $streamName = [string]$streamItem.Stream
        if ($streamName -in @('$DATA', ':$DATA')) {
            continue
        }
        if ($streamName.EndsWith(':$DATA', [StringComparison]::OrdinalIgnoreCase)) {
            $streamName = $streamName.Substring(0, $streamName.Length - 6)
        }
        if (
            [string]::IsNullOrWhiteSpace($streamName) -or
            $streamName.Length -gt 255 -or
            $streamName -match '[:\\/*?"<>|\x00-\x1f\x7f]' -or
            -not $seenNames.Add($streamName)
        ) {
            throw 'Workbook contains an invalid or ambiguous alternate data stream name.'
        }
        [long]$length = $streamItem.Length
        if ($length -lt 0 -or $length -gt $MaxNamedStreamBytes) {
            throw (
                "Alternate data stream $streamName exceeds the " +
                "$MaxNamedStreamBytes-byte safety limit."
            )
        }
        $totalBytes += $length
        if ($totalBytes -gt $MaxTotalNamedStreamBytes) {
            throw (
                'Workbook alternate data streams exceed the ' +
                "$MaxTotalNamedStreamBytes-byte aggregate safety limit."
            )
        }
        $streamBytes = [byte[]]@(Get-BoundedNamedStreamBytes `
            $Path `
            $streamName `
            $MaxNamedStreamBytes `
            "Alternate data stream $streamName")
        $sha256 = [Security.Cryptography.SHA256]::Create()
        try {
            $streamSha256 = [BitConverter]::ToString(
                $sha256.ComputeHash($streamBytes)
            ).Replace('-', '').ToLowerInvariant()
        }
        finally {
            $sha256.Dispose()
        }
        $verifiedLength = [long](
            Get-Item -LiteralPath $Path -Stream $streamName -Force -ErrorAction Stop
        ).Length
        if ($verifiedLength -ne $length) {
            throw "Alternate data stream changed during inspection: $streamName"
        }
        [void]$result.Add([PSCustomObject]@{
            Name = $streamName
            Length = $length
            Sha256 = $streamSha256
        })
    }
    return @($result.ToArray() | Sort-Object Name)
}

function Test-NamedStreamStateEqual {
    param(
        [object[]]$Left,
        [object[]]$Right
    )

    $leftItems = @($Left | Sort-Object Name)
    $rightItems = @($Right | Sort-Object Name)
    if ($leftItems.Count -ne $rightItems.Count) { return $false }
    for ($index = 0; $index -lt $leftItems.Count; $index++) {
        if (
            -not [string]::Equals(
                [string]$leftItems[$index].Name,
                [string]$rightItems[$index].Name,
                [StringComparison]::OrdinalIgnoreCase
            ) -or
            [long]$leftItems[$index].Length -ne [long]$rightItems[$index].Length -or
            [string]$leftItems[$index].Sha256 -cne [string]$rightItems[$index].Sha256
        ) {
            return $false
        }
    }
    return $true
}

function Assert-SafeZoneIdentifierState {
    param(
        [string]$Path,
        [object[]]$NamedStreamState
    )

    $zoneEntries = @(
        $NamedStreamState | Where-Object {
            [string]::Equals(
                [string]$_.Name,
                'Zone.Identifier',
                [StringComparison]::OrdinalIgnoreCase
            )
        }
    )
    if ($zoneEntries.Count -eq 0) { return }
    if ($zoneEntries.Count -ne 1) {
        throw 'Native Excel automation refused: Zone.Identifier is ambiguous.'
    }
    if ([long]$zoneEntries[0].Length -gt $MaxZoneIdentifierBytes) {
        throw (
            'Native Excel automation refused: Zone.Identifier exceeds the ' +
            "$MaxZoneIdentifierBytes-byte safety limit."
        )
    }

    $bytes = @(Get-BoundedNamedStreamBytes `
        $Path `
        ([string]$zoneEntries[0].Name) `
        $MaxZoneIdentifierBytes `
        'Zone.Identifier')
    $byteArray = [byte[]]$bytes
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        $actualSha256 = [BitConverter]::ToString(
            $sha256.ComputeHash($byteArray)
        ).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }
    if ($actualSha256 -cne [string]$zoneEntries[0].Sha256) {
        throw 'Native Excel automation refused: Zone.Identifier changed during inspection.'
    }

    $utf8 = New-Object Text.UTF8Encoding($false, $true)
    try {
        $text = $utf8.GetString($byteArray)
    }
    catch {
        throw 'Native Excel automation refused: Zone.Identifier is not valid UTF-8.'
    }
    if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) {
        $text = $text.Substring(1)
    }
    if (
        [string]::IsNullOrEmpty($text) -or
        $text.Contains([char]0) -or
        [regex]::IsMatch($text, '\r(?!\n)')
    ) {
        throw 'Native Excel automation refused: Zone.Identifier is malformed.'
    }

    $section = ''
    $zoneTransferSections = 0
    $zoneIds = New-Object 'System.Collections.Generic.List[int]'
    foreach ($rawLine in [regex]::Split($text, '\r?\n')) {
        $line = $rawLine.Trim()
        if (
            [string]::IsNullOrEmpty($line) -or
            $line.StartsWith(';') -or
            $line.StartsWith('#')
        ) {
            continue
        }
        $sectionMatch = [regex]::Match($line, '^\[([^\[\]\r\n]+)\]$')
        if ($sectionMatch.Success) {
            $section = $sectionMatch.Groups[1].Value.Trim().ToLowerInvariant()
            if ([string]::IsNullOrEmpty($section)) {
                throw 'Native Excel automation refused: Zone.Identifier is malformed.'
            }
            if ($section -ceq 'zonetransfer') { $zoneTransferSections++ }
            continue
        }
        $assignment = [regex]::Match(
            $line,
            '^([A-Za-z][A-Za-z0-9._-]*)\s*=\s*(.*)$'
        )
        if (-not $assignment.Success -or [string]::IsNullOrEmpty($section)) {
            throw 'Native Excel automation refused: Zone.Identifier is malformed.'
        }
        if (
            $section -ceq 'zonetransfer' -and
            $assignment.Groups[1].Value -ieq 'ZoneId'
        ) {
            $zoneIdText = $assignment.Groups[2].Value.Trim()
            if ($zoneIdText -cnotmatch '^[0-4]$') {
                throw 'Native Excel automation refused: Zone.Identifier has an invalid ZoneId.'
            }
            [void]$zoneIds.Add([int]$zoneIdText)
        }
    }
    if ($zoneTransferSections -ne 1 -or $zoneIds.Count -ne 1) {
        throw 'Native Excel automation refused: Zone.Identifier is missing or ambiguous.'
    }
    if ($zoneIds[0] -in @(3, 4)) {
        throw (
            'Native Excel automation refused: the workbook is marked as ' +
            "Internet or Restricted Zone (ZoneId=$($zoneIds[0])). " +
            'Trust and unblock it explicitly before editing.'
        )
    }
}

function Copy-NamedStreamsFromSource {
    param(
        [string]$SourcePath,
        [string]$TargetPath,
        [object[]]$ExpectedState
    )

    foreach ($existing in @(Get-NamedStreamState $TargetPath)) {
        Remove-Item `
            -LiteralPath $TargetPath `
            -Stream ([string]$existing.Name) `
            -Force `
            -ErrorAction Stop
    }
    foreach ($entry in @($ExpectedState)) {
        $streamBytes = [byte[]]@(Get-BoundedNamedStreamBytes `
            $SourcePath `
            ([string]$entry.Name) `
            $MaxNamedStreamBytes `
            "Alternate data stream $($entry.Name)")
        Set-Content `
            -LiteralPath $TargetPath `
            -Stream ([string]$entry.Name) `
            -Encoding Byte `
            -Value $streamBytes `
            -ErrorAction Stop
    }
    $actualState = @(Get-NamedStreamState $TargetPath)
    if (-not (Test-NamedStreamStateEqual $ExpectedState $actualState)) {
        throw 'Alternate data streams could not be preserved on the native edit work file.'
    }
}

function Test-ProtectedPackageEntry {
    param([string]$EntryName)

    $normalized = $EntryName.Replace('\', '/')
    if ($normalized -ieq 'xl/vbaProject.bin') {
        return $false
    }
    return $normalized -imatch (
        '^(?:xl/(?:activeX|ctrlProps|embeddings|media|printerSettings)/.+|' +
        '.+\.(?:bin|png|jpe?g|gif|emf|wmf|vml))$'
    )
}

function Get-PackagePreservationState {
    param([string]$Path)

    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entryNames = @(
            $archive.Entries |
                ForEach-Object { $_.FullName } |
                Sort-Object
        )
        $protectedHashes = New-Object (
            'System.Collections.Generic.Dictionary[string,string]'
        ) ([StringComparer]::Ordinal)

        foreach ($entry in $archive.Entries) {
            if (-not (Test-ProtectedPackageEntry $entry.FullName)) {
                continue
            }
            $stream = $null
            $sha256 = $null
            try {
                $stream = $entry.Open()
                $sha256 = [Security.Cryptography.SHA256]::Create()
                $protectedHashes.Add(
                    $entry.FullName,
                    [BitConverter]::ToString(
                        $sha256.ComputeHash($stream)
                    ).Replace('-', '').ToLowerInvariant()
                )
            }
            finally {
                if ($null -ne $sha256) {
                    $sha256.Dispose()
                }
                if ($null -ne $stream) {
                    $stream.Dispose()
                }
            }
        }

        return [PSCustomObject]@{
            Entries = $entryNames
            ProtectedHashes = $protectedHashes
        }
    }
    finally {
        $archive.Dispose()
    }
}

function Compare-PackagePreservationState {
    param(
        [object]$Before,
        [object]$After,
        [bool]$AllowWorkbookObjectChanges = $false
    )

    $differences = New-Object 'System.Collections.Generic.List[string]'
    $entryDelta = @(
        @($Before.Entries) + @($After.Entries) |
            Group-Object -CaseSensitive |
            Where-Object { $_.Count -eq 1 } |
            ForEach-Object { [string]$_.Name }
    )
    if ($entryDelta.Count -gt 0) {
        $unexpectedEntries = @(
            if ($AllowWorkbookObjectChanges) {
                $entryDelta | Where-Object {
                [string]$_ -notmatch (
                    '^(?:\[Content_Types\]\.xml|' +
                    'xl/worksheets/(?:sheet\d+\.xml|_rels/sheet\d+\.xml\.rels)|' +
                    'xl/tables/table\d+\.xml|' +
                    'xl/charts/(?:chart|chartEx|style|colors)\d+\.xml|' +
                    'xl/charts/_rels/(?:chart|chartEx)\d+\.xml\.rels|' +
                    'xl/drawings/(?:drawing\d+\.xml|_rels/drawing\d+\.xml\.rels))$'
                )
                }
            } else {
                $entryDelta
            }
        )
        if ($unexpectedEntries.Count -gt 0) {
            [void]$differences.Add(
                'OOXML entry set: ' + ($unexpectedEntries -join ', ')
            )
        }
    }

    foreach ($entryName in $Before.ProtectedHashes.Keys) {
        if (-not $After.ProtectedHashes.ContainsKey($entryName)) {
            [void]$differences.Add("$entryName (missing)")
            continue
        }
        if (
            $Before.ProtectedHashes[$entryName] -cne
            $After.ProtectedHashes[$entryName]
        ) {
            [void]$differences.Add("$entryName (content changed)")
        }
    }
    foreach ($entryName in $After.ProtectedHashes.Keys) {
        if (-not $Before.ProtectedHashes.ContainsKey($entryName)) {
            [void]$differences.Add("$entryName (added)")
        }
    }
    return $differences.ToArray()
}

function Get-VbaProjectFingerprint {
    param(
        [string]$WorkbookPath,
        [string]$HelperPath,
        [string]$RequestPath
    )

    Assert-NoReparsePointChain $WorkbookPath | Out-Null
    Assert-NoReparsePointChain $HelperPath | Out-Null
    Assert-NoReparsePointChain $RequestPath | Out-Null
    if (-not (Test-Path -LiteralPath $HelperPath -PathType Leaf)) {
        throw "VBA fingerprint helper is missing: $HelperPath"
    }
    if (Test-Path -LiteralPath $RequestPath) {
        throw "VBA fingerprint request already exists: $RequestPath"
    }

    try {
        $requestJson = @{
            schemaVersion = 1
            operation = 'fingerprint'
            workbookPath = $WorkbookPath
        } | ConvertTo-Json -Compress
        [IO.File]::WriteAllText(
            $RequestPath,
            $requestJson,
            (New-Object Text.UTF8Encoding($false))
        )
        Assert-NoReparsePointChain $RequestPath | Out-Null

        $output = & $HelperPath $RequestPath 2>&1
        $exitCode = $LASTEXITCODE
        $resultLine = @(
            $output | Where-Object { [string]$_ -match '^\s*\{' }
        ) | Select-Object -Last 1
        if ($exitCode -ne 0 -or [string]::IsNullOrWhiteSpace($resultLine)) {
            throw "VBA fingerprint failed: $output"
        }
        $result = [string]$resultLine | ConvertFrom-Json
        if (
            $result.ok -ne $true -or
            $result.operation -cne 'fingerprint' -or
            [string]$result.projectFingerprintSha256 -notmatch '^[0-9a-f]{64}$' -or
            [int]$result.projectStreamCount -lt 1
        ) {
            throw "VBA fingerprint result is invalid: $resultLine"
        }
        return [PSCustomObject]@{
            ProjectName = [string]$result.projectName
            FingerprintSha256 = [string]$result.projectFingerprintSha256
            StreamCount = [int]$result.projectStreamCount
            StorageCount = [int]$result.projectStorageCount
            Protected = [bool]$result.protected
            Signed = [bool]$result.signed
        }
    }
    finally {
        if (Test-Path -LiteralPath $RequestPath -PathType Leaf) {
            Remove-Item -LiteralPath $RequestPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Save-DisplacedCandidate {
    param(
        [string]$CandidatePath,
        [string]$CandidateSha256,
        [string]$PrimaryDisplacedPath,
        [string]$WorkbookBaseName,
        [string]$BackupDirectory,
        [string]$TransactionId
    )

    if ($CandidatePath -ieq $PrimaryDisplacedPath) {
        return $PrimaryDisplacedPath
    }
    $conflictBackupPath = Join-Path $BackupDirectory (
        $WorkbookBaseName + '.concurrent-edit.' +
        [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ') + '.' +
        $TransactionId + '.' + [Guid]::NewGuid().ToString('N') +
        $script:NativeWorkbookExtension
    )
    Assert-NoReparsePointChain $conflictBackupPath | Out-Null
    if (Test-Path -LiteralPath $conflictBackupPath) {
        throw "Conflict backup path already exists: $conflictBackupPath"
    }
    Copy-Item -LiteralPath $CandidatePath -Destination $conflictBackupPath
    if ((Get-FileSha256Hex $conflictBackupPath) -cne $CandidateSha256) {
        throw 'Displaced concurrent workbook version could not be preserved.'
    }
    return $conflictBackupPath
}

function Restore-MissingWorkbook {
    param(
        [string]$WorkbookPath,
        [string]$CandidatePath,
        [string]$WorkbookDirectory,
        [string]$WorkbookBaseName,
        [string]$PrimaryDisplacedPath,
        [string]$BackupDirectory,
        [string]$TransactionId
    )

    $candidateSha256 = Get-FileSha256Hex $CandidatePath
    $preservedPath = Save-DisplacedCandidate `
        $CandidatePath `
        $candidateSha256 `
        $PrimaryDisplacedPath `
        $WorkbookBaseName `
        $BackupDirectory `
        $TransactionId
    $cleanupPaths = New-Object 'System.Collections.Generic.List[string]'
    try {
        for ($attempt = 0; $attempt -lt 8; $attempt++) {
            if (Test-Path -LiteralPath $WorkbookPath -PathType Leaf) {
                return [PSCustomObject]@{
                    Restored = $false
                    ExternalVersionAlreadyCurrent = $true
                    PreservedPath = $preservedPath
                }
            }
            $restorePath = Join-Path $WorkbookDirectory (
                '.' + $WorkbookBaseName + '.excel-ai-missing-restore.' +
                $TransactionId + '.' + $attempt + $script:NativeWorkbookExtension
            )
            Assert-NoReparsePointChain $restorePath | Out-Null
            if (Test-Path -LiteralPath $restorePath) {
                throw "Missing-workbook recovery path already exists: $restorePath"
            }
            [void]$cleanupPaths.Add($restorePath)
            Copy-Item -LiteralPath $CandidatePath -Destination $restorePath
            if ((Get-FileSha256Hex $restorePath) -cne $candidateSha256) {
                throw 'Missing-workbook recovery copy could not be verified.'
            }
            try {
                [IO.File]::Move($restorePath, $WorkbookPath)
            }
            catch {
                if (Test-Path -LiteralPath $WorkbookPath -PathType Leaf) {
                    return [PSCustomObject]@{
                        Restored = $false
                        ExternalVersionAlreadyCurrent = $true
                        PreservedPath = $preservedPath
                    }
                }
                continue
            }
            return [PSCustomObject]@{
                Restored = $true
                ExternalVersionAlreadyCurrent = $false
                PreservedPath = $preservedPath
            }
        }
        throw (
            'The workbook path remained unavailable during recovery. ' +
            "Displaced version preserved at: $preservedPath"
        )
    }
    finally {
        foreach ($path in $cleanupPaths) {
            if (Test-Path -LiteralPath $path -PathType Leaf) {
                Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

function Restore-DisplacedWorkbook {
    param(
        [string]$WorkbookPath,
        [string]$DisplacedPath,
        [string]$InstalledSha256,
        [string]$WorkbookDirectory,
        [string]$WorkbookBaseName,
        [string]$BackupDirectory,
        [string]$TransactionId
    )

    $candidatePath = $DisplacedPath
    $candidateSha256 = Get-FileSha256Hex $candidatePath
    $cleanupPaths = New-Object 'System.Collections.Generic.List[string]'
    $candidatePreservedPath = $DisplacedPath
    try {
        for ($attempt = 0; $attempt -lt 8; $attempt++) {
            $candidatePreservedPath = Save-DisplacedCandidate `
                $candidatePath `
                $candidateSha256 `
                $DisplacedPath `
                $WorkbookBaseName `
                $BackupDirectory `
                $TransactionId
            if (
                $candidatePath -ine $DisplacedPath -and
                -not $cleanupPaths.Contains($candidatePath)
            ) {
                [void]$cleanupPaths.Add($candidatePath)
            }
            $currentSha256 = Get-FileSha256Hex $WorkbookPath
            if ($currentSha256 -cne $InstalledSha256) {
                return [PSCustomObject]@{
                    Restored = $false
                    ExternalVersionAlreadyCurrent = $true
                    WorkbookSha256 = $currentSha256
                    PreservedPath = $candidatePreservedPath
                }
            }

            $restorePath = Join-Path $WorkbookDirectory (
                '.' + $WorkbookBaseName + '.excel-ai-conflict-restore.' +
                $TransactionId + '.' + $attempt + $script:NativeWorkbookExtension
            )
            $capturedPath = Join-Path $WorkbookDirectory (
                '.' + $WorkbookBaseName + '.excel-ai-conflict-captured.' +
                $TransactionId + '.' + $attempt + $script:NativeWorkbookExtension
            )
            foreach ($path in @($restorePath, $capturedPath)) {
                Assert-NoReparsePointChain $path | Out-Null
                if (Test-Path -LiteralPath $path) {
                    throw "Conflict recovery path already exists: $path"
                }
            }

            Copy-Item -LiteralPath $candidatePath -Destination $restorePath
            [void]$cleanupPaths.Add($restorePath)
            if ((Get-FileSha256Hex $restorePath) -cne $candidateSha256) {
                throw 'Conflict recovery copy could not be verified.'
            }
            try {
                [IO.File]::Replace(
                    $restorePath,
                    $WorkbookPath,
                    $capturedPath,
                    $false
                )
            }
            catch {
                $replaceError = $_.Exception.Message
                $recoveryCandidate = if (
                    Test-Path -LiteralPath $capturedPath -PathType Leaf
                ) {
                    $capturedPath
                }
                else {
                    $candidatePath
                }
                if (Test-Path -LiteralPath $WorkbookPath -PathType Leaf) {
                    $recoveryCandidateSha256 = Get-FileSha256Hex $recoveryCandidate
                    $preservedPath = Save-DisplacedCandidate `
                        $recoveryCandidate `
                        $recoveryCandidateSha256 `
                        $DisplacedPath `
                        $WorkbookBaseName `
                        $BackupDirectory `
                        $TransactionId
                }
                else {
                    $missingRecovery = Restore-MissingWorkbook `
                        $WorkbookPath `
                        $recoveryCandidate `
                        $WorkbookDirectory `
                        $WorkbookBaseName `
                        $DisplacedPath `
                        $BackupDirectory `
                        $TransactionId
                    $preservedPath = $missingRecovery.PreservedPath
                }
                if (
                    $recoveryCandidate -ine $DisplacedPath -and
                    -not $cleanupPaths.Contains($recoveryCandidate)
                ) {
                    [void]$cleanupPaths.Add($recoveryCandidate)
                }
                throw (
                    'Conflict recovery swap failed after preserving or restoring ' +
                    "the displaced workbook at ${preservedPath}: $replaceError"
                )
            }
            $capturedSha256 = Get-FileSha256Hex $capturedPath
            if ($capturedSha256 -ceq $InstalledSha256) {
                [void]$cleanupPaths.Add($capturedPath)
                if ((Get-FileSha256Hex $WorkbookPath) -cne $candidateSha256) {
                    throw 'Conflict recovery target does not match the displaced version.'
                }
                return [PSCustomObject]@{
                    Restored = $true
                    ExternalVersionAlreadyCurrent = $false
                    WorkbookSha256 = $candidateSha256
                    PreservedPath = $candidatePreservedPath
                }
            }

            # Another writer won the interval between the hash and atomic swap.
            # The just-displaced version is newer, so install it on the next pass.
            $InstalledSha256 = $candidateSha256
            $candidatePath = $capturedPath
            $candidateSha256 = $capturedSha256
        }

        $candidatePreservedPath = Save-DisplacedCandidate `
            $candidatePath `
            $candidateSha256 `
            $DisplacedPath `
            $WorkbookBaseName `
            $BackupDirectory `
            $TransactionId
        if ($candidatePath -ine $DisplacedPath) {
            [void]$cleanupPaths.Add($candidatePath)
        }
        throw (
            'Concurrent workbook replacements did not stop. ' +
            "Latest displaced version preserved at: $candidatePreservedPath"
        )
    }
    finally {
        foreach ($path in $cleanupPaths) {
            if (Test-Path -LiteralPath $path -PathType Leaf) {
                Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
            }
        }
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

    $sheetName = Get-OperationSheetName $Operation
    $row = Get-NativeInteger `
        (Get-RequiredProperty $Operation 'row') `
        'Native cell row' `
        1 `
        1048576
    $column = Get-NativeInteger `
        (Get-RequiredProperty $Operation 'column') `
        'Native cell column' `
        1 `
        16384

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

function Assert-AllowedProperties {
    param(
        [object]$Value,
        [string[]]$Allowed,
        [string]$Label
    )
    foreach ($property in $Value.PSObject.Properties) {
        if ($Allowed -notcontains $property.Name) {
            throw "$Label contains an unknown property: $($property.Name)"
        }
    }
}

function Get-OperationSheetName {
    param([object]$Operation)

    $value = Get-RequiredProperty $Operation 'sheetName'
    if (
        $value -isnot [string] -or
        [string]::IsNullOrWhiteSpace([string]$value) -or
        ([string]$value).Length -gt 31
    ) {
        throw 'Native edit operation contains an invalid worksheet name.'
    }
    return [string]$value
}

function Assert-NativeColor {
    param(
        [AllowNull()][object]$Value,
        [string]$Label
    )
    if ($Value -isnot [string] -or [string]$Value -cnotmatch '^#[0-9a-f]{6}$') {
        throw "$Label must be a lowercase #rrggbb color."
    }
}

function Assert-NativeShortText {
    param(
        [AllowNull()][object]$Value,
        [string]$Label
    )
    if (
        $Value -isnot [string] -or
        [string]::IsNullOrEmpty([string]$Value) -or
        ([string]$Value).Length -gt 255 -or
        ([string]$Value).Contains([char]0)
    ) {
        throw "$Label must contain 1-255 characters without NUL."
    }
}

function Test-NativeNumber {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return $false
    }
    $typeCode = [Type]::GetTypeCode($Value.GetType())
    return @(
        [TypeCode]::Byte,
        [TypeCode]::Decimal,
        [TypeCode]::Double,
        [TypeCode]::Int16,
        [TypeCode]::Int32,
        [TypeCode]::Int64,
        [TypeCode]::SByte,
        [TypeCode]::Single,
        [TypeCode]::UInt16,
        [TypeCode]::UInt32,
        [TypeCode]::UInt64
    ) -contains $typeCode
}

function Get-NativeInteger {
    param(
        [AllowNull()][object]$Value,
        [string]$Label,
        [int]$Minimum,
        [int]$Maximum
    )

    if (-not (Test-NativeNumber $Value)) {
        throw "$Label must be an integer."
    }
    $number = [double]$Value
    if (
        [double]::IsNaN($number) -or
        [double]::IsInfinity($number) -or
        [Math]::Truncate($number) -ne $number -or
        $number -lt $Minimum -or
        $number -gt $Maximum
    ) {
        throw "$Label is outside Excel limits."
    }
    return [int]$number
}

function Convert-ExcelColumnLettersToNumber {
    param([string]$Letters)

    $number = 0
    foreach ($character in $Letters.ToCharArray()) {
        $number = ($number * 26) + ([int]$character - [int][char]'A') + 1
    }
    return $number
}

function Get-NormalizedRangeRef {
    param([AllowNull()][object]$Value)

    if ($Value -isnot [string]) {
        throw 'Native workbook range is invalid.'
    }
    $normalized = ([string]$Value).Trim().Replace('$', '').ToUpperInvariant()
    $match = [regex]::Match(
        $normalized,
        '^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$'
    )
    if (-not $match.Success) {
        throw 'Native workbook range is invalid.'
    }
    $columns = @($match.Groups[1].Value)
    $rows = @([int]$match.Groups[2].Value)
    if ($match.Groups[3].Success) {
        $columns += $match.Groups[3].Value
        $rows += [int]$match.Groups[4].Value
    }
    foreach ($letters in $columns) {
        if ((Convert-ExcelColumnLettersToNumber $letters) -gt 16384) {
            throw 'Native workbook range exceeds Excel limits.'
        }
    }
    foreach ($row in $rows) {
        if ($row -gt 1048576) {
            throw 'Native workbook range exceeds Excel limits.'
        }
    }
    $startColumn = Convert-ExcelColumnLettersToNumber $match.Groups[1].Value
    $endColumn = if ($match.Groups[3].Success) {
        Convert-ExcelColumnLettersToNumber $match.Groups[3].Value
    }
    else {
        $startColumn
    }
    $startRow = [int]$match.Groups[2].Value
    $endRow = if ($match.Groups[4].Success) {
        [int]$match.Groups[4].Value
    }
    else {
        $startRow
    }
    if ($startColumn -gt $endColumn -or $startRow -gt $endRow) {
        throw 'Native workbook range must run from top-left to bottom-right.'
    }
    return $normalized
}

function Get-BoundedWorkbookObjectRangeCellCount {
    param(
        [AllowNull()][object]$Value,
        [string]$Label
    )

    $normalized = Get-NormalizedRangeRef $Value
    $match = [regex]::Match(
        $normalized,
        '^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$'
    )
    [long]$startColumn = Convert-ExcelColumnLettersToNumber $match.Groups[1].Value
    [long]$endColumn = if ($match.Groups[3].Success) {
        Convert-ExcelColumnLettersToNumber $match.Groups[3].Value
    }
    else {
        $startColumn
    }
    [long]$startRow = [long]$match.Groups[2].Value
    [long]$endRow = if ($match.Groups[4].Success) {
        [long]$match.Groups[4].Value
    }
    else {
        $startRow
    }
    [long]$cellCount = ($endColumn - $startColumn + 1) * ($endRow - $startRow + 1)
    if ($cellCount -gt $MaxWorkbookObjectRangeCells) {
        throw (
            "$Label exceeds the $MaxWorkbookObjectRangeCells-cell " +
            'native Excel safety limit.'
        )
    }
    return $cellCount
}

function Test-NativeWorkbookRangesOverlap {
    param(
        [string]$First,
        [string]$Second
    )

    $firstNormalized = Get-NormalizedRangeRef $First
    $secondNormalized = Get-NormalizedRangeRef $Second
    $pattern = '^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$'
    $firstMatch = [regex]::Match($firstNormalized, $pattern)
    $secondMatch = [regex]::Match($secondNormalized, $pattern)
    $firstStartColumn = Convert-ExcelColumnLettersToNumber $firstMatch.Groups[1].Value
    $firstEndColumn = if ($firstMatch.Groups[3].Success) {
        Convert-ExcelColumnLettersToNumber $firstMatch.Groups[3].Value
    } else { $firstStartColumn }
    $secondStartColumn = Convert-ExcelColumnLettersToNumber $secondMatch.Groups[1].Value
    $secondEndColumn = if ($secondMatch.Groups[3].Success) {
        Convert-ExcelColumnLettersToNumber $secondMatch.Groups[3].Value
    } else { $secondStartColumn }
    $firstStartRow = [int]$firstMatch.Groups[2].Value
    $firstEndRow = if ($firstMatch.Groups[4].Success) {
        [int]$firstMatch.Groups[4].Value
    } else { $firstStartRow }
    $secondStartRow = [int]$secondMatch.Groups[2].Value
    $secondEndRow = if ($secondMatch.Groups[4].Success) {
        [int]$secondMatch.Groups[4].Value
    } else { $secondStartRow }
    return $firstStartRow -le $secondEndRow -and
        $secondStartRow -le $firstEndRow -and
        $firstStartColumn -le $secondEndColumn -and
        $secondStartColumn -le $firstEndColumn
}

function Get-NativeTableRangeCellCost {
    param(
        [object]$Definition,
        [string]$Label
    )

    return [long](Get-BoundedWorkbookObjectRangeCellCount `
        (Get-RequiredProperty $Definition 'rangeRef') `
        "$Label.rangeRef")
}

function Get-NativeChartRangeCellCost {
    param(
        [object]$Definition,
        [string]$Label
    )

    [long]$total = 0
    if (Has-Property $Definition 'sourceRangeRef') {
        $total += Get-BoundedWorkbookObjectRangeCellCount `
            $Definition.sourceRangeRef `
            "$Label.sourceRangeRef"
    }
    if (Has-Property $Definition 'series') {
        $seriesIndex = 0
        foreach ($seriesDefinition in @($Definition.series)) {
            foreach ($propertyName in @(
                'nameRange',
                'categoryRange',
                'valuesRange',
                'xValuesRange',
                'bubbleSizesRange'
            )) {
                if (Has-Property $seriesDefinition $propertyName) {
                    $total += Get-BoundedWorkbookObjectRangeCellCount `
                        $seriesDefinition.$propertyName `
                        "$Label.series[$seriesIndex].$propertyName"
                }
            }
            $seriesIndex++
        }
    }
    return $total
}

function Apply-ColumnWidthOperation {
    param(
        [object]$Workbook,
        [object]$Operation
    )

    Assert-AllowedProperties `
        $Operation `
        @('kind', 'sheetName', 'column', 'widthPx') `
        'Native column-width operation'
    $sheetName = Get-OperationSheetName $Operation
    $column = Get-NativeInteger `
        (Get-RequiredProperty $Operation 'column') `
        'Native column-width column' `
        1 `
        16384
    $widthValue = Get-RequiredProperty $Operation 'widthPx'
    if (-not (Test-NativeNumber $widthValue)) {
        throw 'Native column width must be numeric.'
    }
    $widthPx = [double]$widthValue
    if (
        [double]::IsNaN($widthPx) -or
        [double]::IsInfinity($widthPx) -or
        $widthPx -le 5 -or
        $widthPx -gt 1790
    ) {
        throw 'Native column-width operation is outside Excel limits.'
    }

    $worksheet = $null
    $columns = $null
    $targetColumn = $null
    try {
        $worksheet = $Workbook.Worksheets.Item($sheetName)
        $columns = $worksheet.Columns
        $targetColumn = $columns.Item($column)
        $targetColumn.ColumnWidth = [Math]::Max(($widthPx - 5) / 7, 0)
    }
    finally {
        Release-ComObject $targetColumn
        Release-ComObject $columns
        Release-ComObject $worksheet
    }
}

function Apply-RowHeightOperation {
    param(
        [object]$Workbook,
        [object]$Operation
    )

    Assert-AllowedProperties `
        $Operation `
        @('kind', 'sheetName', 'row', 'heightPx') `
        'Native row-height operation'
    $sheetName = Get-OperationSheetName $Operation
    $row = Get-NativeInteger `
        (Get-RequiredProperty $Operation 'row') `
        'Native row-height row' `
        1 `
        1048576
    $heightValue = Get-RequiredProperty $Operation 'heightPx'
    if (-not (Test-NativeNumber $heightValue)) {
        throw 'Native row height must be numeric.'
    }
    $heightPx = [double]$heightValue
    if (
        [double]::IsNaN($heightPx) -or
        [double]::IsInfinity($heightPx) -or
        $heightPx -le 0 -or
        $heightPx -gt 546
    ) {
        throw 'Native row-height operation is outside Excel limits.'
    }

    $worksheet = $null
    $rows = $null
    $targetRow = $null
    try {
        $worksheet = $Workbook.Worksheets.Item($sheetName)
        $rows = $worksheet.Rows
        $targetRow = $rows.Item($row)
        $targetRow.RowHeight = $heightPx * 72 / 96
    }
    finally {
        Release-ComObject $targetRow
        Release-ComObject $rows
        Release-ComObject $worksheet
    }
}

function Apply-ClearConditionalFormattingOperation {
    param(
        [object]$Workbook,
        [object]$Operation
    )

    Assert-AllowedProperties `
        $Operation `
        @('kind', 'sheetName') `
        'Native clear-conditional-formatting operation'
    $sheetName = Get-OperationSheetName $Operation

    $worksheet = $null
    $cells = $null
    $conditions = $null
    try {
        $worksheet = $Workbook.Worksheets.Item($sheetName)
        $cells = $worksheet.Cells
        $conditions = $cells.FormatConditions
        $conditions.Delete()
    }
    finally {
        Release-ComObject $conditions
        Release-ComObject $cells
        Release-ComObject $worksheet
    }
}

function Apply-ConditionalFormattingOperation {
    param(
        [object]$Workbook,
        [object]$Operation
    )

    Assert-AllowedProperties `
        $Operation `
        @('kind', 'sheetName', 'rangeRef', 'rule') `
        'Native conditional-formatting operation'
    $sheetName = Get-OperationSheetName $Operation
    $rangeRef = Get-NormalizedRangeRef (
        Get-RequiredProperty $Operation 'rangeRef'
    )
    $rule = Get-RequiredProperty $Operation 'rule'
    $ruleType = [string](Get-RequiredProperty $rule 'type')

    $worksheet = $null
    $range = $null
    $conditions = $null
    $formatCondition = $null
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
    $iconSets = $null
    $iconSet = $null
    try {
        $worksheet = $Workbook.Worksheets.Item($sheetName)
        $range = $worksheet.Range($rangeRef)
        $conditions = $range.FormatConditions

        switch ($ruleType) {
            'cellIs' {
                Assert-AllowedProperties `
                    $rule `
                    @(
                        'type',
                        'operator',
                        'operand',
                        'fillColor',
                        'fontColor',
                        'bold'
                    ) `
                    'Native cellIs rule'
                $operatorName = [string](
                    Get-RequiredProperty $rule 'operator'
                )
                $operator = switch ($operatorName) {
                    'equal' { 3 }
                    'greaterThan' { 5 }
                    'lessThan' { 6 }
                    default { throw 'Native cellIs operator is invalid.' }
                }
                $operand = Get-RequiredProperty $rule 'operand'
                if ($operand -is [string]) {
                    Assert-NativeShortText $operand 'Native cellIs operand'
                    if ([string]$operand -match '^=') {
                        throw 'Native cellIs operand cannot be an Excel formula.'
                    }
                }
                elseif (-not (Test-NativeNumber $operand)) {
                    throw 'Native cellIs operand is invalid.'
                }
                elseif (
                    [double]::IsNaN([double]$operand) -or
                    [double]::IsInfinity([double]$operand)
                ) {
                    throw 'Native cellIs operand must be finite.'
                }
                Assert-NativeColor `
                    (Get-RequiredProperty $rule 'fillColor') `
                    'Native cellIs fill'
                Assert-NativeColor `
                    (Get-RequiredProperty $rule 'fontColor') `
                    'Native cellIs font'
                $bold = Get-RequiredProperty $rule 'bold'
                if ($bold -isnot [bool] -or $bold -ne $true) {
                    throw 'Native cellIs bold style is invalid.'
                }
                $formatCondition = $conditions.Add(1, $operator, $operand)
                $interior = $formatCondition.Interior
                $interior.Color = Convert-HexToOleColor ([string]$rule.fillColor)
                $font = $formatCondition.Font
                $font.Color = Convert-HexToOleColor ([string]$rule.fontColor)
                $font.Bold = $true
            }
            'containsText' {
                Assert-AllowedProperties `
                    $rule `
                    @('type', 'text', 'fillColor', 'fontColor', 'bold') `
                    'Native containsText rule'
                Assert-NativeShortText `
                    (Get-RequiredProperty $rule 'text') `
                    'Native containsText value'
                Assert-NativeColor `
                    (Get-RequiredProperty $rule 'fillColor') `
                    'Native containsText fill'
                Assert-NativeColor `
                    (Get-RequiredProperty $rule 'fontColor') `
                    'Native containsText font'
                $bold = Get-RequiredProperty $rule 'bold'
                if ($bold -isnot [bool] -or $bold -ne $true) {
                    throw 'Native containsText bold style is invalid.'
                }
                $missing = [Type]::Missing
                $formatCondition = $conditions.Add(
                    9,
                    $missing,
                    $missing,
                    $missing,
                    [string]$rule.text,
                    0,
                    $missing,
                    $missing
                )
                $interior = $formatCondition.Interior
                $interior.Color = Convert-HexToOleColor ([string]$rule.fillColor)
                $font = $formatCondition.Font
                $font.Color = Convert-HexToOleColor ([string]$rule.fontColor)
                $font.Bold = $true
            }
            'colorScale' {
                Assert-AllowedProperties `
                    $rule `
                    @('type', 'colors') `
                    'Native colorScale rule'
                $colors = @(Get-RequiredProperty $rule 'colors')
                if ($colors.Count -ne 3) {
                    throw 'Native colorScale rule requires three colors.'
                }
                foreach ($color in $colors) {
                    Assert-NativeColor $color 'Native colorScale color'
                }
                $formatCondition = $conditions.AddColorScale(3)
                $criteria = $formatCondition.ColorScaleCriteria
                $criterion1 = $criteria.Item(1)
                $criterion2 = $criteria.Item(2)
                $criterion3 = $criteria.Item(3)
                $criterion1.Type = 1
                $criterion2.Type = 5
                $criterion2.Value = 50
                $criterion3.Type = 2
                $formatColor1 = $criterion1.FormatColor
                $formatColor2 = $criterion2.FormatColor
                $formatColor3 = $criterion3.FormatColor
                $formatColor1.Color = Convert-HexToOleColor ([string]$colors[0])
                $formatColor2.Color = Convert-HexToOleColor ([string]$colors[1])
                $formatColor3.Color = Convert-HexToOleColor ([string]$colors[2])
            }
            'dataBar' {
                Assert-AllowedProperties `
                    $rule `
                    @('type', 'color') `
                    'Native dataBar rule'
                Assert-NativeColor `
                    (Get-RequiredProperty $rule 'color') `
                    'Native dataBar color'
                $formatCondition = $conditions.AddDatabar()
                $minPoint = $formatCondition.MinPoint
                $maxPoint = $formatCondition.MaxPoint
                $minPoint.Modify(1)
                $maxPoint.Modify(2)
                $barColor = $formatCondition.BarColor
                $barColor.Color = Convert-HexToOleColor ([string]$rule.color)
            }
            'iconSet' {
                Assert-AllowedProperties `
                    $rule `
                    @('type', 'iconSet', 'thresholds') `
                    'Native iconSet rule'
                if ([string](Get-RequiredProperty $rule 'iconSet') -cne '3TrafficLights1') {
                    throw 'Native iconSet name is invalid.'
                }
                $thresholds = @(Get-RequiredProperty $rule 'thresholds')
                if (
                    $thresholds.Count -ne 2 -or
                    -not (Test-NativeNumber $thresholds[0]) -or
                    -not (Test-NativeNumber $thresholds[1]) -or
                    [double]$thresholds[0] -ne 33 -or
                    [double]$thresholds[1] -ne 67
                ) {
                    throw 'Native iconSet thresholds must be 33 and 67.'
                }
                $formatCondition = $conditions.AddIconSetCondition()
                $iconSets = $Workbook.IconSets
                $iconSet = $iconSets.Item(4)
                $formatCondition.IconSet = $iconSet
                $criteria = $formatCondition.IconCriteria
                $criterion2 = $criteria.Item(2)
                $criterion3 = $criteria.Item(3)
                $criterion2.Type = 3
                $criterion2.Value = 33
                $criterion2.Operator = 7
                $criterion3.Type = 3
                $criterion3.Value = 67
                $criterion3.Operator = 7
            }
            default {
                throw "Unsupported native conditional-formatting rule: $ruleType"
            }
        }
        $formatCondition.SetLastPriority()
    }
    finally {
        Release-ComObject $iconSet
        Release-ComObject $iconSets
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
        Release-ComObject $formatCondition
        Release-ComObject $conditions
        Release-ComObject $range
        Release-ComObject $worksheet
    }
}

function Assert-NativeBoolean {
    param(
        [AllowNull()][object]$Value,
        [string]$Label
    )
    if ($Value -isnot [bool]) {
        throw "$Label must be a boolean."
    }
}

function Get-NativeFiniteNumber {
    param(
        [AllowNull()][object]$Value,
        [string]$Label,
        [double]$Minimum,
        [double]$Maximum,
        [bool]$Integer = $false
    )
    if (-not (Test-NativeNumber $Value)) {
        throw "$Label must be numeric."
    }
    $number = [double]$Value
    if (
        [double]::IsNaN($number) -or
        [double]::IsInfinity($number) -or
        $number -lt $Minimum -or
        $number -gt $Maximum -or
        ($Integer -and [Math]::Truncate($number) -ne $number)
    ) {
        throw "$Label is outside the supported range."
    }
    return $number
}

function Assert-NativeObjectId {
    param(
        [AllowNull()][object]$Value,
        [string]$Label
    )
    if (
        $Value -isnot [string] -or
        ([string]$Value).Length -lt 1 -or
        ([string]$Value).Length -gt 512 -or
        [string]$Value -match '[\x00-\x1f\x7f]'
    ) {
        throw "$Label must contain 1-512 characters without controls."
    }
}

function Test-NativeExcelTableName {
    param([string]$Value)

    if ($Value -cnotmatch '^[\p{L}_\\][\p{L}\p{N}._]{0,254}$') {
        return $false
    }
    if ($Value -match '^[RC]$') {
        return $false
    }
    $a1Match = [regex]::Match(
        $Value,
        '^([A-Za-z]{1,3})([1-9][0-9]{0,6})$',
        [Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    if (
        $a1Match.Success -and
        (Convert-ExcelColumnLettersToNumber $a1Match.Groups[1].Value.ToUpperInvariant()) -le 16384 -and
        [int]$a1Match.Groups[2].Value -le 1048576
    ) {
        return $false
    }
    $r1c1Match = [regex]::Match(
        $Value,
        '^R([1-9][0-9]{0,6})C([1-9][0-9]{0,4})$',
        [Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    if (
        $r1c1Match.Success -and
        [int]$r1c1Match.Groups[1].Value -le 1048576 -and
        [int]$r1c1Match.Groups[2].Value -le 16384
    ) {
        return $false
    }
    return $true
}

function Assert-NativeObjectName {
    param(
        [AllowNull()][object]$Value,
        [string]$Label,
        [bool]$IdentifierOnly = $false
    )
    if (
        $Value -isnot [string] -or
        ([string]$Value).Length -lt 1 -or
        ([string]$Value).Length -gt 255 -or
        [string]$Value -match '[\x00-\x1f\x7f]' -or
        ($IdentifierOnly -and -not (Test-NativeExcelTableName ([string]$Value)))
    ) {
        throw "$Label is invalid."
    }
}

function Assert-NativeOptionalText {
    param(
        [AllowNull()][object]$Value,
        [string]$Label,
        [int]$MaximumLength
    )
    if (
        $Value -isnot [string] -or
        ([string]$Value).Length -gt $MaximumLength -or
        ([string]$Value).Contains([char]0)
    ) {
        throw "$Label is invalid."
    }
}

function Assert-NativeTableDefinition {
    param([object]$Definition)

    Assert-AllowedProperties `
        $Definition `
        @('id', 'name', 'displayName', 'rangeRef', 'headerRow', 'totalsRow', 'style') `
        'Native table definition'
    Assert-NativeObjectId (Get-RequiredProperty $Definition 'id') 'Native table id'
    Assert-NativeObjectName `
        (Get-RequiredProperty $Definition 'name') `
        'Native table name' `
        $true
    Assert-NativeObjectName `
        (Get-RequiredProperty $Definition 'displayName') `
        'Native table display name' `
        $true
    $normalizedRange = Get-NormalizedRangeRef (Get-RequiredProperty $Definition 'rangeRef')
    Assert-NativeBoolean `
        (Get-RequiredProperty $Definition 'headerRow') `
        'Native table headerRow'
    Assert-NativeBoolean `
        (Get-RequiredProperty $Definition 'totalsRow') `
        'Native table totalsRow'
	$rangeMatch = [regex]::Match(
		$normalizedRange,
		'^[A-Z]{1,3}(?<startRow>[1-9][0-9]{0,6})(?::[A-Z]{1,3}(?<endRow>[1-9][0-9]{0,6}))?$'
	)
	$startRow = [int]$rangeMatch.Groups['startRow'].Value
	$endRow = if ($rangeMatch.Groups['endRow'].Success) {
		[int]$rangeMatch.Groups['endRow'].Value
	} else { $startRow }
	$minimumRows = if ([bool]$Definition.totalsRow) { 3 } else { 2 }
	if (($endRow - $startRow + 1) -lt $minimumRows) {
		throw 'Native table range does not contain enough rows for its header, data and optional totals row.'
	}

    $style = Get-RequiredProperty $Definition 'style'
    Assert-AllowedProperties `
        $style `
        @('name', 'showFirstColumn', 'showLastColumn', 'showRowStripes', 'showColumnStripes') `
        'Native table style'
    $styleName = Get-RequiredProperty $style 'name'
    if (
        $styleName -isnot [string] -or
        [string]$styleName -cnotmatch (
            '^TableStyle(?:Light(?:[1-9]|1[0-9]|2[01])|' +
            'Medium(?:[1-9]|1[0-9]|2[0-8])|Dark(?:[1-9]|1[01]))$'
        )
    ) {
        throw 'Native table style name is invalid.'
    }
    foreach ($propertyName in @(
        'showFirstColumn',
        'showLastColumn',
        'showRowStripes',
        'showColumnStripes'
    )) {
        Assert-NativeBoolean `
            (Get-RequiredProperty $style $propertyName) `
            "Native table style $propertyName"
    }
}

function Assert-NativeChartAxisDefinition {
    param(
        [object]$Definition,
        [string]$Label
    )
    Assert-AllowedProperties `
        $Definition `
        @(
            'visible', 'title', 'minimumScale', 'maximumScale', 'majorUnit',
            'minorUnit', 'logarithmic', 'reverseOrder', 'numberFormat',
            'majorGridlines', 'minorGridlines'
        ) `
        $Label
    foreach ($propertyName in @(
        'visible', 'logarithmic', 'reverseOrder',
        'majorGridlines', 'minorGridlines'
    )) {
        if (Has-Property $Definition $propertyName) {
            Assert-NativeBoolean $Definition.$propertyName "$Label.$propertyName"
        }
    }
    if (Has-Property $Definition 'title') {
        Assert-NativeOptionalText $Definition.title "$Label.title" 1000
    }
    if (Has-Property $Definition 'numberFormat') {
        Assert-NativeOptionalText $Definition.numberFormat "$Label.numberFormat" 255
    }
    foreach ($propertyName in @('minimumScale', 'maximumScale')) {
        if (Has-Property $Definition $propertyName) {
            $value = $Definition.$propertyName
            if ($null -ne $value) {
                Get-NativeFiniteNumber $value "$Label.$propertyName" -1e307 1e307 | Out-Null
            }
        }
    }
    foreach ($propertyName in @('majorUnit', 'minorUnit')) {
        if (Has-Property $Definition $propertyName) {
            $value = $Definition.$propertyName
            if ($null -ne $value) {
                $number = Get-NativeFiniteNumber $value "$Label.$propertyName" 0 1e307
                if ($number -le 0) {
                    throw "$Label.$propertyName must be greater than zero."
                }
            }
        }
    }
    if (
        (Has-Property $Definition 'minimumScale') -and
        (Has-Property $Definition 'maximumScale') -and
        $null -ne $Definition.minimumScale -and
        $null -ne $Definition.maximumScale -and
        [double]$Definition.minimumScale -ge [double]$Definition.maximumScale
    ) {
        throw "$Label minimumScale must be below maximumScale."
    }
    if (
        (Has-Property $Definition 'logarithmic') -and
        [bool]$Definition.logarithmic -and
        (
            ((Has-Property $Definition 'minimumScale') -and $null -ne $Definition.minimumScale -and [double]$Definition.minimumScale -le 0) -or
            ((Has-Property $Definition 'maximumScale') -and $null -ne $Definition.maximumScale -and [double]$Definition.maximumScale -le 0)
        )
    ) {
        throw "$Label logarithmic scale cannot use a non-positive bound."
    }
}

function Assert-NativeChartSeriesDefinition {
    param(
        [object]$Definition,
        [int]$Index
    )
    $label = "Native chart series $Index"
    Assert-AllowedProperties `
        $Definition `
        @(
            'id', 'name', 'nameRange', 'categoryRange', 'valuesRange',
            'xValuesRange', 'bubbleSizesRange', 'chartType', 'axisGroup',
            'color', 'lineColor', 'lineWidth', 'dashStyle', 'markerStyle',
            'markerSize', 'smooth', 'visible', 'dataLabels'
        ) `
        $label
    Assert-NativeObjectId (Get-RequiredProperty $Definition 'id') "$label.id"
    if (
        (Has-Property $Definition 'name') -and
        (Has-Property $Definition 'nameRange')
    ) {
        throw "$label cannot define both name and nameRange."
    }
    if (
        (Has-Property $Definition 'categoryRange') -and
        (Has-Property $Definition 'xValuesRange')
    ) {
        throw "$label cannot define both categoryRange and xValuesRange."
    }
    if (Has-Property $Definition 'name') {
        Assert-NativeObjectName $Definition.name "$label.name"
        if (([string]$Definition.name).Trim() -match '^[=+\-@]') {
            throw "$label.name cannot be an Excel formula."
        }
    }
    foreach ($propertyName in @(
        'nameRange', 'categoryRange', 'valuesRange',
        'xValuesRange', 'bubbleSizesRange'
    )) {
        if (
            $propertyName -ceq 'valuesRange' -or
            (Has-Property $Definition $propertyName)
        ) {
            Get-NormalizedRangeRef `
                (Get-RequiredProperty $Definition $propertyName) | Out-Null
        }
    }
    if (
        (Has-Property $Definition 'nameRange') -and
        (Get-BoundedWorkbookObjectRangeCellCount `
            $Definition.nameRange `
            "$label.nameRange") -ne 1
    ) {
        throw "$label.nameRange must identify exactly one cell."
    }
    if (Has-Property $Definition 'chartType') {
        $chartType = [int](Get-NativeFiniteNumber `
            $Definition.chartType "$label.chartType" -10000 10000 $true)
        if ($script:ValidNativeChartTypes -notcontains $chartType) {
            throw "$label.chartType is unknown."
        }
    }
    if (
        (Has-Property $Definition 'axisGroup') -and
        [string]$Definition.axisGroup -notin @('primary', 'secondary')
    ) {
        throw "$label.axisGroup is invalid."
    }
    foreach ($propertyName in @('color', 'lineColor')) {
        if (Has-Property $Definition $propertyName) {
            Assert-NativeColor $Definition.$propertyName "$label.$propertyName"
        }
    }
    if (Has-Property $Definition 'lineWidth') {
        Get-NativeFiniteNumber $Definition.lineWidth "$label.lineWidth" 0.1 20 | Out-Null
    }
    if (Has-Property $Definition 'markerSize') {
        Get-NativeFiniteNumber $Definition.markerSize "$label.markerSize" 2 72 $true | Out-Null
    }
    if (
        (Has-Property $Definition 'dashStyle') -and
        [string]$Definition.dashStyle -notin @('solid', 'dash', 'dot', 'dashDot')
    ) {
        throw "$label.dashStyle is invalid."
    }
    if (
        (Has-Property $Definition 'markerStyle') -and
        [string]$Definition.markerStyle -notin @(
            'automatic', 'circle', 'dash', 'diamond', 'dot', 'none',
            'picture', 'plus', 'square', 'star', 'triangle', 'x'
        )
    ) {
        throw "$label.markerStyle is invalid."
    }
    foreach ($propertyName in @('smooth', 'visible')) {
        if (Has-Property $Definition $propertyName) {
            Assert-NativeBoolean $Definition.$propertyName "$label.$propertyName"
        }
    }
	if (Has-Property $Definition 'dataLabels') {
		$dataLabels = $Definition.dataLabels
        Assert-AllowedProperties `
            $dataLabels `
            @(
                'showValue', 'showCategoryName', 'showSeriesName',
                'showPercentage', 'showBubbleSize', 'position'
            ) `
            "$label.dataLabels"
		$hasExplicitShowOption = $false
		$hasEnabledShowOption = $false
		foreach ($propertyName in @(
			'showValue', 'showCategoryName', 'showSeriesName',
			'showPercentage', 'showBubbleSize'
		)) {
			if (Has-Property $dataLabels $propertyName) {
				$hasExplicitShowOption = $true
				Assert-NativeBoolean `
					$dataLabels.$propertyName `
					"$label.dataLabels.$propertyName"
				if ([bool]$dataLabels.$propertyName) {
					$hasEnabledShowOption = $true
				}
			}
		}
		if (-not $hasExplicitShowOption) {
			throw "$label.dataLabels must explicitly define at least one show option."
		}
		if ((Has-Property $dataLabels 'position') -and -not $hasEnabledShowOption) {
			throw "$label.dataLabels.position requires at least one enabled show option."
		}
        if (
            (Has-Property $dataLabels 'position') -and
            [string]$dataLabels.position -notin @(
                'above', 'below', 'bestFit', 'center', 'insideBase',
                'insideEnd', 'left', 'outsideEnd', 'right'
            )
        ) {
            throw "$label.dataLabels.position is invalid."
        }
    }
}

function Assert-NativeChartDefinition {
    param([object]$Definition)

    Assert-AllowedProperties `
        $Definition `
        @(
            'id', 'name', 'chartType', 'sourceRangeRef', 'plotBy', 'anchor',
            'title', 'legend', 'categoryAxis', 'valueAxis',
            'secondaryCategoryAxis', 'secondaryValueAxis', 'series', 'style',
            'roundedCorners', 'gapWidth', 'overlap', 'alternativeText'
        ) `
        'Native chart definition'
    Assert-NativeObjectId (Get-RequiredProperty $Definition 'id') 'Native chart id'
    Assert-NativeObjectName (Get-RequiredProperty $Definition 'name') 'Native chart name'
    $chartType = [int](Get-NativeFiniteNumber `
        (Get-RequiredProperty $Definition 'chartType') `
        'Native chart type' `
        -10000 `
        10000 `
        $true)
    if ($script:ValidNativeChartTypes -notcontains $chartType) {
        throw 'Native chart type is unknown.'
    }
    $plotBy = [string](Get-RequiredProperty $Definition 'plotBy')
    if ($plotBy -notin @('columns', 'rows')) {
        throw 'Native chart plotBy is invalid.'
    }
    if (Has-Property $Definition 'sourceRangeRef') {
        Get-NormalizedRangeRef $Definition.sourceRangeRef | Out-Null
    }
    if (
        (Has-Property $Definition 'sourceRangeRef') -and
        (Has-Property $Definition 'series')
    ) {
        throw 'Native chart cannot define both sourceRangeRef and series.'
    }
    $anchor = Get-RequiredProperty $Definition 'anchor'
    Assert-AllowedProperties `
        $anchor `
        @('left', 'top', 'width', 'height') `
        'Native chart anchor'
    Get-NativeFiniteNumber `
        (Get-RequiredProperty $anchor 'left') `
        'Native chart left' 0 10000000 | Out-Null
    Get-NativeFiniteNumber `
        (Get-RequiredProperty $anchor 'top') `
        'Native chart top' 0 10000000 | Out-Null
    Get-NativeFiniteNumber `
        (Get-RequiredProperty $anchor 'width') `
        'Native chart width' 20 1000000 | Out-Null
    Get-NativeFiniteNumber `
        (Get-RequiredProperty $anchor 'height') `
        'Native chart height' 20 1000000 | Out-Null

    if (Has-Property $Definition 'title') {
        $title = $Definition.title
        Assert-AllowedProperties $title @('visible', 'text') 'Native chart title'
        Assert-NativeBoolean `
            (Get-RequiredProperty $title 'visible') `
            'Native chart title.visible'
        Assert-NativeOptionalText `
            (Get-RequiredProperty $title 'text') `
            'Native chart title.text' `
            1000
    }
    if (Has-Property $Definition 'legend') {
        $legend = $Definition.legend
        Assert-AllowedProperties $legend @('visible', 'position') 'Native chart legend'
        Assert-NativeBoolean `
            (Get-RequiredProperty $legend 'visible') `
            'Native chart legend.visible'
        if (
            [string](Get-RequiredProperty $legend 'position') -notin @(
                'bottom', 'corner', 'custom', 'left', 'right', 'top'
            )
        ) {
            throw 'Native chart legend.position is invalid.'
        }
    }
    foreach ($propertyName in @(
        'categoryAxis', 'valueAxis',
        'secondaryCategoryAxis', 'secondaryValueAxis'
    )) {
        if (Has-Property $Definition $propertyName) {
            Assert-NativeChartAxisDefinition `
                $Definition.$propertyName `
                "Native chart $propertyName"
        }
    }
    $seriesCount = 0
    $seriesDefinitions = @()
    if (Has-Property $Definition 'series') {
        $seriesDefinitions = @($Definition.series)
        $seriesCount = $seriesDefinitions.Count
        if ($seriesCount -gt $MaxChartSeries) {
            throw "Native chart series must contain at most $MaxChartSeries entries."
        }
        $seriesIds = New-Object 'System.Collections.Generic.HashSet[string]' `
            ([StringComparer]::Ordinal)
        for ($index = 0; $index -lt $seriesDefinitions.Count; $index++) {
            Assert-NativeChartSeriesDefinition $seriesDefinitions[$index] ($index + 1)
            $seriesId = [string]$seriesDefinitions[$index].id
            if (-not $seriesIds.Add($seriesId)) {
                throw 'Native chart series ids must be unique.'
            }
        }
    }
    if (-not (Has-Property $Definition 'sourceRangeRef') -and $seriesCount -eq 0) {
        throw 'Native chart requires a source range or at least one series.'
    }
	$effectiveSeriesTypes = New-Object 'System.Collections.Generic.HashSet[int]'
	foreach ($seriesDefinition in $seriesDefinitions) {
		$effectiveSeriesType = if (Has-Property $seriesDefinition 'chartType') {
			[int]$seriesDefinition.chartType
		} else { $chartType }
		[void]$effectiveSeriesTypes.Add($effectiveSeriesType)
		if (
			(Has-Property $seriesDefinition 'bubbleSizesRange') -and
			-not ($script:NativeBubbleChartTypes -contains $effectiveSeriesType)
		) {
			throw 'Native chart bubbleSizesRange requires a bubble series type.'
		}
		if (
			(Has-Property $seriesDefinition 'smooth') -and
			-not ($script:NativeSmoothSeriesChartTypes -contains $effectiveSeriesType)
		) {
			throw 'Native chart smooth requires a line or scatter series type.'
		}
		if (Has-Property $seriesDefinition 'dataLabels') {
			$seriesLabels = $seriesDefinition.dataLabels
			if (
				(Has-Property $seriesLabels 'showBubbleSize') -and
				[bool]$seriesLabels.showBubbleSize -and
				-not ($script:NativeBubbleChartTypes -contains $effectiveSeriesType)
			) {
				throw 'Native chart showBubbleSize requires a bubble series type.'
			}
			if (
				(Has-Property $seriesLabels 'showPercentage') -and
				[bool]$seriesLabels.showPercentage -and
				-not ($script:NativePercentageDataLabelChartTypes -contains $effectiveSeriesType)
			) {
				throw 'Native chart showPercentage requires a pie or doughnut series type.'
			}
			if (
				(Has-Property $seriesLabels 'position') -and
				-not (Test-NativeDataLabelPositionSupported `
					$effectiveSeriesType `
					([string]$seriesLabels.position))
			) {
				throw 'Native chart data label position is not supported by this chart type.'
			}
		}
	}
	$bubbleSeriesTypeCount = @(
		$effectiveSeriesTypes |
			Where-Object { $script:NativeBubbleChartTypes -contains [int]$_ }
	).Count
	if (
		$bubbleSeriesTypeCount -gt 0 -and
		$bubbleSeriesTypeCount -lt $effectiveSeriesTypes.Count
	) {
		throw 'Native charts cannot mix bubble and non-bubble series because Excel silently promotes every series to bubble.'
	}
	if ($seriesCount -gt 0 -and $effectiveSeriesTypes.Count -eq 1) {
		$singleEffectiveSeriesType = [int]@($effectiveSeriesTypes)[0]
		if ($chartType -ne $singleEffectiveSeriesType) {
			throw 'Native homogeneous explicit series require their concrete type as the top-level chart type.'
		}
	}
	if ($chartType -ne -4152 -and $effectiveSeriesTypes.Count -gt 1) {
		throw 'Native heterogeneous explicit series require xlCombo as the top-level chart type.'
	}
    $axisProperties = @(
        'categoryAxis', 'valueAxis',
        'secondaryCategoryAxis', 'secondaryValueAxis'
    )
    if (-not (Test-NativeChartTypeSupportsAxes $chartType)) {
        foreach ($axisProperty in $axisProperties) {
            if (Has-Property $Definition $axisProperty) {
                throw 'Native chart axes are not supported by this chart type.'
            }
        }
    }
	foreach ($axisSpec in @(
		@{ Name = 'categoryAxis'; Group = 'primary' },
		@{ Name = 'secondaryCategoryAxis'; Group = 'secondary' }
	)) {
		$axisProperty = [string]$axisSpec.Name
		if (
			(Has-Property $Definition $axisProperty) -and
			-not (Test-NativeChartAxisGroupSupportsCategoryScale $Definition ([string]$axisSpec.Group))
		) {
			$axisDefinition = $Definition.$axisProperty
			foreach ($scaleProperty in @(
				'minimumScale', 'maximumScale', 'majorUnit',
				'minorUnit', 'logarithmic'
			)) {
				if (Has-Property $axisDefinition $scaleProperty) {
					throw "Native chart $axisProperty numeric scale settings require a scatter or bubble series on the same axis group."
				}
			}
		}
	}
    foreach ($axisProperty in @('secondaryCategoryAxis', 'secondaryValueAxis')) {
        if (Has-Property $Definition $axisProperty) {
            $axisDefinition = $Definition.$axisProperty
            if (
                (Has-Property $axisDefinition 'majorGridlines') -or
                (Has-Property $axisDefinition 'minorGridlines')
            ) {
                throw "Native chart $axisProperty cannot define gridlines."
            }
        }
    }
    $supportsSecondaryAxes = Test-NativeChartDefinitionSupportsAnyType `
        $Definition `
        $script:NativeSecondaryAxisChartTypes
    if (-not $supportsSecondaryAxes -and (
            (Has-Property $Definition 'secondaryCategoryAxis') -or
            (Has-Property $Definition 'secondaryValueAxis')
        )) {
        throw 'Native chart secondary axes are not supported by this chart type.'
    }
    foreach ($seriesDefinition in $seriesDefinitions) {
        if (
            (Has-Property $seriesDefinition 'axisGroup') -and
            [string]$seriesDefinition.axisGroup -ceq 'secondary'
        ) {
            $effectiveSeriesType = if (Has-Property $seriesDefinition 'chartType') {
                [int]$seriesDefinition.chartType
            }
            else {
                $chartType
            }
            if (-not ($script:NativeSecondaryAxisChartTypes -contains $effectiveSeriesType)) {
                throw 'Native chart series cannot use a secondary axis with this chart type.'
            }
        }
    }
    if ($chartType -eq -4152) {
        if ($seriesCount -eq 0) {
            throw 'Native custom combo charts require explicit series.'
        }
        $comboTypes = New-Object 'System.Collections.Generic.HashSet[int]'
        foreach ($seriesDefinition in $seriesDefinitions) {
            if (
                -not (Has-Property $seriesDefinition 'chartType') -or
                [int]$seriesDefinition.chartType -eq -4152
            ) {
                throw 'Native custom combo charts require a concrete chartType on every series.'
            }
            [void]$comboTypes.Add([int]$seriesDefinition.chartType)
        }
        if ($comboTypes.Count -lt 2) {
            throw 'Native custom combo charts require at least two distinct concrete series chart types.'
        }
    }
    if (Has-Property $Definition 'style') {
        Get-NativeFiniteNumber $Definition.style 'Native chart style' 1 48 $true | Out-Null
    }
    if (Has-Property $Definition 'roundedCorners') {
        Assert-NativeBoolean $Definition.roundedCorners 'Native chart roundedCorners'
    }
    if (Has-Property $Definition 'gapWidth') {
        Get-NativeFiniteNumber $Definition.gapWidth 'Native chart gapWidth' 0 500 $true | Out-Null
        if (-not (Test-NativeChartDefinitionSupportsAnyType `
            $Definition `
            $script:NativeGapWidthChartTypes)) {
            throw 'Native chart gapWidth is not supported by this chart type.'
        }
    }
    if (Has-Property $Definition 'overlap') {
        Get-NativeFiniteNumber $Definition.overlap 'Native chart overlap' -100 100 $true | Out-Null
        if (-not (Test-NativeChartDefinitionSupportsAnyType `
            $Definition `
            $script:NativeOverlapChartTypes)) {
            throw 'Native chart overlap is not supported by this chart type.'
        }
    }
    if (Has-Property $Definition 'alternativeText') {
        Assert-NativeOptionalText `
            $Definition.alternativeText `
            'Native chart alternativeText' `
            1000
    }
}

function Set-NativeTableDefinition {
    param(
        [object]$ListObject,
        [object]$Definition,
        [object]$TargetRange,
		[bool]$Resize
    )

	$currentTableRange = $null
	try {
		$currentHasTotals = [bool]$ListObject.ShowTotals
		if ($currentHasTotals -ne [bool]$Definition.totalsRow) {
			throw 'Native table totalsRow transitions are disabled because Excel moves worksheet cells and rewrites formula references.'
		}
		if ($currentHasTotals) {
			$currentTableRange = $ListObject.Range
			$currentRangeRef = ([string]$currentTableRange.Address(
				$false, $false
			)).ToUpperInvariant()
			if ($currentRangeRef -cne (Get-NormalizedRangeRef $Definition.rangeRef)) {
				throw 'Native tables with an existing totals row cannot be resized because Excel moves worksheet cells and rewrites formula references.'
			}
		}
		elseif ($Resize) {
			$ListObject.Resize($TargetRange)
		}
    }
    finally {
		Release-ComObject $currentTableRange
	}
    $ListObject.Name = [string]$Definition.name
    $ListObject.DisplayName = [string]$Definition.displayName
    $ListObject.ShowHeaders = [bool]$Definition.headerRow
    $ListObject.TableStyle = [string]$Definition.style.name
    $ListObject.ShowTableStyleFirstColumn = [bool]$Definition.style.showFirstColumn
    $ListObject.ShowTableStyleLastColumn = [bool]$Definition.style.showLastColumn
    $ListObject.ShowTableStyleRowStripes = [bool]$Definition.style.showRowStripes
    $ListObject.ShowTableStyleColumnStripes = [bool]$Definition.style.showColumnStripes
}

function Apply-CreateTableOperation {
    param(
        [object]$Workbook,
        [object]$Operation
    )

    Assert-AllowedProperties `
        $Operation `
        @('kind', 'sheetName', 'table') `
        'Native createTable operation'
    $sheetName = Get-OperationSheetName $Operation
    $definition = Get-RequiredProperty $Operation 'table'
	Assert-NativeTableDefinition $definition
	if (-not [bool]$definition.headerRow) {
		throw 'Native table creation with headerRow=false is disabled because Excel can move worksheet cells.'
	}
	if ([bool]$definition.totalsRow) {
		throw 'Native table creation with totalsRow=true is disabled because Excel moves worksheet cells and rewrites formula references.'
	}
    $rangeRef = Get-NormalizedRangeRef $definition.rangeRef

    $worksheet = $null
    $range = $null
    $listObjects = $null
    $listObject = $null
    try {
        $worksheet = $Workbook.Worksheets.Item($sheetName)
        $range = $worksheet.Range($rangeRef)
        $listObjects = $worksheet.ListObjects
        $missing = [Type]::Missing
		# Always tell Excel that the first source row is the internal table
		# header. xlNo inserts/shifts cells and violates the no-corruption rule.
		$listObject = $listObjects.Add(1, $range, $missing, 1)
		Set-NativeTableDefinition $listObject $definition $range $false
    }
    finally {
        Release-ComObject $listObject
        Release-ComObject $listObjects
        Release-ComObject $range
        Release-ComObject $worksheet
    }
}

function Apply-UpdateTableOperation {
    param(
        [object]$Workbook,
        [object]$Operation
    )

    Assert-AllowedProperties `
        $Operation `
        @('kind', 'sheetName', 'name', 'table') `
        'Native updateTable operation'
    $sheetName = Get-OperationSheetName $Operation
    $currentName = Get-RequiredProperty $Operation 'name'
    Assert-NativeObjectName $currentName 'Native current table name' $true
    $definition = Get-RequiredProperty $Operation 'table'
    Assert-NativeTableDefinition $definition
    $rangeRef = Get-NormalizedRangeRef $definition.rangeRef

    $worksheet = $null
    $range = $null
    $listObjects = $null
    $listObject = $null
    try {
        $worksheet = $Workbook.Worksheets.Item($sheetName)
        $range = $worksheet.Range($rangeRef)
        $listObjects = $worksheet.ListObjects
        $listObject = $listObjects.Item([string]$currentName)
		if ([bool]$listObject.ShowHeaders -ne [bool]$definition.headerRow) {
			throw 'Native table headerRow transitions are disabled because Excel can move worksheet cells.'
		}
		Set-NativeTableDefinition $listObject $definition $range $true
    }
    finally {
        Release-ComObject $listObject
        Release-ComObject $listObjects
        Release-ComObject $range
        Release-ComObject $worksheet
    }
}

function Apply-DeleteTableOperation {
    param(
        [object]$Workbook,
        [object]$Operation
    )

    Assert-AllowedProperties `
        $Operation `
        @('kind', 'sheetName', 'name') `
        'Native deleteTable operation'
    $sheetName = Get-OperationSheetName $Operation
    $currentName = Get-RequiredProperty $Operation 'name'
    Assert-NativeObjectName $currentName 'Native current table name' $true

    $worksheet = $null
    $listObjects = $null
    $listObject = $null
    try {
        $worksheet = $Workbook.Worksheets.Item($sheetName)
        $listObjects = $worksheet.ListObjects
        $listObject = $listObjects.Item([string]$currentName)
        # Removing a table must not delete or shift its worksheet cells.
        $listObject.Unlist()
    }
    finally {
        Release-ComObject $listObject
        Release-ComObject $listObjects
        Release-ComObject $worksheet
    }
}

function Get-NativeMarkerStyle {
    param([string]$Value)
    switch ($Value) {
        'automatic' { return -4105 }
        'circle' { return 8 }
        'dash' { return -4115 }
        'diamond' { return 2 }
        'dot' { return -4118 }
        'none' { return -4142 }
        'picture' { return -4147 }
        'plus' { return 9 }
        'square' { return 1 }
        'star' { return 5 }
        'triangle' { return 3 }
        'x' { return -4168 }
        default { throw "Unsupported native marker style: $Value" }
    }
}

function Get-NativeDashStyle {
    param([string]$Value)
    switch ($Value) {
        'solid' { return 1 }
        'dash' { return 4 }
        'dot' { return 2 }
        'dashDot' { return 5 }
        default { throw "Unsupported native dash style: $Value" }
    }
}

function Get-NativeDataLabelPosition {
    param([string]$Value)
    switch ($Value) {
        'above' { return 0 }
        'below' { return 1 }
        'bestFit' { return 5 }
        'center' { return -4108 }
        'insideBase' { return 4 }
        'insideEnd' { return 3 }
        'left' { return -4131 }
        'outsideEnd' { return 2 }
        'right' { return -4152 }
        default { throw "Unsupported native data-label position: $Value" }
    }
}

function Test-NativeDataLabelPositionSupported {
    param(
        [int]$ChartType,
        [string]$Position
    )

    $supportedPositions = switch ($ChartType) {
        { $_ -in @(51, 57) } {
            @('center', 'insideBase', 'insideEnd', 'outsideEnd')
            break
        }
        { $_ -in @(52, 53, 58, 59) } {
            @('center', 'insideBase', 'insideEnd')
            break
        }
        { $_ -in @(4, 63, 64, 65, 66, 67, -4169, 72, 73, 74, 75, 15, 87) } {
            @('above', 'below', 'center', 'left', 'right')
            break
        }
        { $_ -in @(5, 68, 69, 70, 71, -4102) } {
            @('bestFit', 'center', 'insideEnd', 'outsideEnd')
            break
        }
        default { @() }
    }
    return $supportedPositions -contains $Position
}

function Test-NativeChartTypeSupportsAxes {
    param([int]$ChartType)

    return $ChartType -notin @(
        -4120, -4102,
        5, 68, 69, 70, 71, 80,
        117, 120, 123, 130, 131, 140
    )
}

function Get-NativeChartDefinitionEffectiveTypes {
    param([object]$Definition)

    $types = New-Object 'System.Collections.Generic.List[int]'
    $topLevelType = [int]$Definition.chartType
    if (Has-Property $Definition 'series') {
        foreach ($seriesDefinition in @($Definition.series)) {
            if (Has-Property $seriesDefinition 'chartType') {
                [void]$types.Add([int]$seriesDefinition.chartType)
            }
            else {
                [void]$types.Add($topLevelType)
            }
        }
    }
    if ($types.Count -eq 0) {
        [void]$types.Add($topLevelType)
    }
    return @($types)
}

function Test-NativeChartDefinitionSupportsAnyType {
    param(
        [object]$Definition,
        [int[]]$SupportedTypes
    )

    foreach ($chartType in @(Get-NativeChartDefinitionEffectiveTypes $Definition)) {
        if ($SupportedTypes -contains [int]$chartType) {
            return $true
        }
    }
    return $false
}

function Test-NativeChartAxisGroupSupportsCategoryScale {
    param(
        [object]$Definition,
        [string]$AxisGroup
    )

    $topLevelType = [int]$Definition.chartType
    if (-not (Has-Property $Definition 'series')) {
        return $AxisGroup -ceq 'primary' -and
            $script:NativeCategoryScaleChartTypes -contains $topLevelType
    }
    foreach ($seriesDefinition in @($Definition.series)) {
        $seriesAxisGroup = if (Has-Property $seriesDefinition 'axisGroup') {
            [string]$seriesDefinition.axisGroup
        }
        else {
            'primary'
        }
        if ($seriesAxisGroup -cne $AxisGroup) {
            continue
        }
        $seriesType = if (Has-Property $seriesDefinition 'chartType') {
            [int]$seriesDefinition.chartType
        }
        else {
            $topLevelType
        }
        if ($script:NativeCategoryScaleChartTypes -contains $seriesType) {
            return $true
        }
    }
    return $false
}

function Test-NativeChartGroupSupportsOption {
    param(
        [object]$ChartGroup,
        [int[]]$SupportedSeriesTypes,
        [int[]]$CompatibleGroupTypes
    )

    if ($CompatibleGroupTypes -contains [int]$ChartGroup.Type) {
        return $true
    }
    $seriesCollection = $null
    try {
        $seriesCollection = $ChartGroup.SeriesCollection()
        for ($index = 1; $index -le [int]$seriesCollection.Count; $index++) {
            $series = $null
            try {
                $series = $seriesCollection.Item($index)
                if ($SupportedSeriesTypes -contains [int]$series.ChartType) {
                    return $true
                }
            }
            finally {
                Release-ComObject $series
            }
        }
        return $false
    }
    finally {
        Release-ComObject $seriesCollection
    }
}

function Set-NativeChartAxisDefinition {
    param(
        [object]$Chart,
        [int]$AxisType,
        [int]$AxisGroup,
        [object]$Definition
    )

    if (Has-Property $Definition 'visible') {
        $Chart.HasAxis($AxisType, $AxisGroup) = [bool]$Definition.visible
        if (-not [bool]$Definition.visible) {
            return
        }
    }
    else {
        $Chart.HasAxis($AxisType, $AxisGroup) = $true
    }

    $axis = $null
    $axisTitle = $null
    $tickLabels = $null
    try {
        $axis = $Chart.Axes($AxisType, $AxisGroup)
        if (Has-Property $Definition 'title') {
            $axis.HasTitle = ([string]$Definition.title).Length -gt 0
            if ($axis.HasTitle) {
                $axisTitle = $axis.AxisTitle
                $axisTitle.Text = [string]$Definition.title
            }
        }
        if (Has-Property $Definition 'minimumScale') {
            if ($null -eq $Definition.minimumScale) {
                $axis.MinimumScaleIsAuto = $true
            }
            else {
                $axis.MinimumScaleIsAuto = $false
                $axis.MinimumScale = [double]$Definition.minimumScale
            }
        }
        if (Has-Property $Definition 'maximumScale') {
            if ($null -eq $Definition.maximumScale) {
                $axis.MaximumScaleIsAuto = $true
            }
            else {
                $axis.MaximumScaleIsAuto = $false
                $axis.MaximumScale = [double]$Definition.maximumScale
            }
        }
        if (Has-Property $Definition 'majorUnit') {
            if ($null -eq $Definition.majorUnit) {
                $axis.MajorUnitIsAuto = $true
            }
            else {
                $axis.MajorUnitIsAuto = $false
                $axis.MajorUnit = [double]$Definition.majorUnit
            }
        }
        if (Has-Property $Definition 'minorUnit') {
            if ($null -eq $Definition.minorUnit) {
                $axis.MinorUnitIsAuto = $true
            }
            else {
                $axis.MinorUnitIsAuto = $false
                $axis.MinorUnit = [double]$Definition.minorUnit
            }
        }
        if (Has-Property $Definition 'logarithmic') {
            $axis.ScaleType = if ([bool]$Definition.logarithmic) { -4133 } else { -4132 }
        }
        if (Has-Property $Definition 'reverseOrder') {
            $axis.ReversePlotOrder = [bool]$Definition.reverseOrder
        }
		if (Has-Property $Definition 'numberFormat') {
			$tickLabels = $axis.TickLabels
			$numberFormat = [string]$Definition.numberFormat
			if ($numberFormat.Length -eq 0) {
				$tickLabels.NumberFormatLinked = $true
			}
			else {
				$tickLabels.NumberFormat = $numberFormat
				$tickLabels.NumberFormatLinked = $false
				# Excel localizes some custom formats (for example 0.00 -> 0,00
				# on a French installation). Keep Excel's canonical value as the
				# expected definition so the post-reopen check remains strict
				# without being tied to the caller's locale.
				$normalizedNumberFormat = [string]$tickLabels.NumberFormat
				if ([string]::IsNullOrEmpty($normalizedNumberFormat)) {
					throw 'Excel returned an empty custom axis number format.'
				}
				$Definition.numberFormat = $normalizedNumberFormat
			}
        }
        if (Has-Property $Definition 'majorGridlines') {
            $axis.HasMajorGridlines = [bool]$Definition.majorGridlines
        }
        if (Has-Property $Definition 'minorGridlines') {
            $axis.HasMinorGridlines = [bool]$Definition.minorGridlines
        }
    }
    finally {
        Release-ComObject $tickLabels
        Release-ComObject $axisTitle
        Release-ComObject $axis
    }
}

function Clear-NativeChartSeries {
    param([object]$Chart)

    $seriesCollection = $null
    try {
        $seriesCollection = $Chart.FullSeriesCollection()
        while ([int]$seriesCollection.Count -gt 0) {
            $series = $null
            try {
                $series = $seriesCollection.Item(1)
                $series.Delete()
            }
            finally {
                Release-ComObject $series
            }
        }
    }
    finally {
        Release-ComObject $seriesCollection
    }
}

function Set-NativeChartSeriesDefinitions {
    param(
        [object]$Chart,
        [object]$Worksheet,
        [object[]]$Definitions
    )

    Clear-NativeChartSeries $Chart
    $seriesCollection = $null
    try {
        $seriesCollection = $Chart.SeriesCollection()
        foreach ($definition in $Definitions) {
            $series = $null
            $categoryRange = $null
            $xValuesRange = $null
            $valuesRange = $null
            $bubbleSizesRange = $null
            $format = $null
            $fill = $null
            $fillColor = $null
            $line = $null
            $lineColor = $null
            $dataLabels = $null
            try {
                $series = $seriesCollection.NewSeries()
				$effectiveSeriesType = if (Has-Property $definition 'chartType') {
					[int]$definition.chartType
				} else { [int]$Chart.ChartType }
                if (Has-Property $definition 'nameRange') {
                    $normalizedNameRange = Get-NormalizedRangeRef `
                        $definition.nameRange
                    $escapedWorksheetName = ([string]$Worksheet.Name).Replace("'", "''")
                    $series.Name = "='$escapedWorksheetName'!$normalizedNameRange"
                }
                elseif (Has-Property $definition 'name') {
                    $series.Name = [string]$definition.name
                }
                if (Has-Property $definition 'categoryRange') {
                    $categoryRange = $Worksheet.Range(
                        (Get-NormalizedRangeRef $definition.categoryRange)
                    )
                    $series.XValues = $categoryRange
                }
                if (Has-Property $definition 'xValuesRange') {
                    $xValuesRange = $Worksheet.Range(
                        (Get-NormalizedRangeRef $definition.xValuesRange)
                    )
                    $series.XValues = $xValuesRange
                }
                $valuesRange = $Worksheet.Range(
                    (Get-NormalizedRangeRef $definition.valuesRange)
                )
                $series.Values = $valuesRange
				if (Has-Property $definition 'chartType') {
					$series.ChartType = [int]$definition.chartType
				}
                if (Has-Property $definition 'bubbleSizesRange') {
                    $bubbleSizesRange = $Worksheet.Range(
                        (Get-NormalizedRangeRef $definition.bubbleSizesRange)
                    )
                    $series.BubbleSizes = $bubbleSizesRange
                }
                if (Has-Property $definition 'axisGroup') {
                    $series.AxisGroup = if (
                        [string]$definition.axisGroup -ceq 'secondary'
                    ) { 2 } else { 1 }
                }
                if (Has-Property $definition 'markerStyle') {
                    $series.MarkerStyle = Get-NativeMarkerStyle (
                        [string]$definition.markerStyle
                    )
                }
                if (Has-Property $definition 'markerSize') {
                    $series.MarkerSize = [int]$definition.markerSize
                }
                if (Has-Property $definition 'smooth') {
                    $series.Smooth = [bool]$definition.smooth
                }

                if (
                    (Has-Property $definition 'color') -or
                    (Has-Property $definition 'lineColor') -or
                    (Has-Property $definition 'lineWidth') -or
                    (Has-Property $definition 'dashStyle')
                ) {
                    $format = $series.Format
                    $fill = $format.Fill
                    $line = $format.Line
                    if (Has-Property $definition 'color') {
                        $fill.Visible = -1
                        [void]$fill.Solid()
                        $fillColor = $fill.ForeColor
                        $fillColor.RGB = Convert-HexToOleColor ([string]$definition.color)
                    }
                    if (Has-Property $definition 'lineColor') {
                        $line.Visible = -1
                        $lineColor = $line.ForeColor
                        $lineColor.RGB = Convert-HexToOleColor ([string]$definition.lineColor)
                    }
                    if (Has-Property $definition 'lineWidth') {
                        $line.Weight = [double]$definition.lineWidth
                    }
                    if (Has-Property $definition 'dashStyle') {
                        $line.DashStyle = Get-NativeDashStyle ([string]$definition.dashStyle)
                    }
                }

				if (Has-Property $definition 'dataLabels') {
					$labels = $definition.dataLabels
					$labelShowProperties = @(
						'showValue', 'showCategoryName', 'showSeriesName',
						'showPercentage', 'showBubbleSize'
					)
					$hasEnabledShowOption = @(
						$labelShowProperties | Where-Object {
							(Has-Property $labels $_) -and [bool]$labels.$_
						}
					).Count -gt 0
					if (-not $hasEnabledShowOption) {
						$series.HasDataLabels = $false
					}
					else {
						[void]$series.ApplyDataLabels()
						$dataLabels = $series.DataLabels()
						foreach ($propertyName in $labelShowProperties) {
						if (
							$propertyName -ceq 'showPercentage' -and
							-not ($script:NativePercentageDataLabelChartTypes -contains $effectiveSeriesType)
						) { continue }
						if (
							$propertyName -ceq 'showBubbleSize' -and
							-not ($script:NativeBubbleChartTypes -contains $effectiveSeriesType)
						) { continue }
							$dataLabels.$propertyName = (
							(Has-Property $labels $propertyName) -and
							[bool]$labels.$propertyName
						)
						}
						if (Has-Property $labels 'position') {
							$dataLabels.Position = Get-NativeDataLabelPosition (
								[string]$labels.position
							)
						}
					}
                }
                if (Has-Property $definition 'visible') {
                    [void]($series.IsFiltered = -not [bool]$definition.visible)
                }
            }
            finally {
                Release-ComObject $dataLabels
                Release-ComObject $lineColor
                Release-ComObject $line
                Release-ComObject $fillColor
                Release-ComObject $fill
                Release-ComObject $format
                Release-ComObject $bubbleSizesRange
                Release-ComObject $valuesRange
                Release-ComObject $xValuesRange
                Release-ComObject $categoryRange
                Release-ComObject $series
            }
        }
    }
    finally {
        Release-ComObject $seriesCollection
    }
}

function Set-NativeChartDefinition {
    param(
        [object]$ChartObject,
        [object]$Chart,
        [object]$Worksheet,
        [object]$Definition,
        [bool]$PreserveAnchor = $false,
		[bool]$PreserveSeries = $false,
		[bool]$AllowSeriesFormattingChange = $false
    )

    $ChartObject.Name = [string]$Definition.name
    if (-not $PreserveAnchor) {
        $ChartObject.Left = [double]$Definition.anchor.left
        $ChartObject.Top = [double]$Definition.anchor.top
        $ChartObject.Width = [double]$Definition.anchor.width
        $ChartObject.Height = [double]$Definition.anchor.height
    }
    $initialChartType = [int]$Definition.chartType
    $definitionSeries = @()
    if (Has-Property $Definition 'series') {
        $definitionSeries = @($Definition.series)
    }
    if (
        $initialChartType -eq -4152 -and
        $definitionSeries.Count -gt 0 -and
        (Has-Property $definitionSeries[0] 'chartType')
    ) {
        # Excel reports a heterogeneous per-series chart as xlCombo only after
        # its series types exist; xlCombo cannot reliably seed an empty chart.
        $initialChartType = [int]$definitionSeries[0].chartType
    }
    $plotBy = if ([string]$Definition.plotBy -ceq 'rows') { 1 } else { 2 }
    if ($PreserveSeries) {
        $actualChartType = [int]$Chart.ChartType
        $chartTypeMatches =
            $actualChartType -eq [int]$Definition.chartType -or
            (
                [int]$Definition.chartType -eq -4152 -and
                $actualChartType -eq -4111
            )
        if (
            -not $chartTypeMatches -or
            (
                (Has-Property $Definition 'sourceRangeRef') -and
                [int]$Chart.PlotBy -ne $plotBy
            )
        ) {
            throw (
                'preserveSeries cannot be used while changing chartType or plotBy: ' +
                [string]$Definition.name
            )
        }
		if (
			(Has-Property $Definition 'style') -and
			[int]$Chart.ChartStyle -ne [int]$Definition.style -and
			-not $AllowSeriesFormattingChange
		) {
            throw (
                'preserveSeries cannot be used while changing chart style: ' +
                [string]$Definition.name
            )
        }
    }
    else {
        $Chart.ChartType = $initialChartType
    }

    $sourceRange = $null
    $chartTitle = $null
    $shapeRange = $null
    $chartGroups = $null
    try {
        # ChartStyle can reset explicit series formatting. Apply it before
        # rebuilding or formatting any series, and never reapply an unchanged
        # style when the caller asks to preserve existing series without rebuilding them.
        if (
            (Has-Property $Definition 'style') -and
			(-not $PreserveSeries -or $AllowSeriesFormattingChange)
        ) {
            $Chart.ChartStyle = [int]$Definition.style
        }
        if (
            -not $PreserveSeries -and
            (Has-Property $Definition 'sourceRangeRef')
        ) {
            $sourceRange = $Worksheet.Range(
                (Get-NormalizedRangeRef $Definition.sourceRangeRef)
            )
            $Chart.SetSourceData($sourceRange, $plotBy)
        }
        if (-not $PreserveSeries -and (Has-Property $Definition 'series')) {
            Set-NativeChartSeriesDefinitions $Chart $Worksheet @($Definition.series)
        }
        if (
            -not $PreserveSeries -and
            (Has-Property $Definition 'sourceRangeRef')
        ) {
            $Chart.PlotBy = $plotBy
        }
        if (Has-Property $Definition 'title') {
            $Chart.HasTitle = [bool]$Definition.title.visible
            if ($Chart.HasTitle) {
                $chartTitle = $Chart.ChartTitle
                $chartTitle.Text = [string]$Definition.title.text
            }
        }
        if (Has-Property $Definition 'legend') {
            $Chart.HasLegend = [bool]$Definition.legend.visible
            if (
                $Chart.HasLegend -and
                [string]$Definition.legend.position -cne 'custom'
            ) {
                $Chart.Legend.Position = switch ([string]$Definition.legend.position) {
                    'bottom' { -4107 }
                    'corner' { 2 }
                    'left' { -4131 }
                    'right' { -4152 }
                    'top' { -4160 }
                }
            }
        }
        if (Test-NativeChartTypeSupportsAxes ([int]$Definition.chartType)) {
            if (Has-Property $Definition 'categoryAxis') {
                Set-NativeChartAxisDefinition $Chart 1 1 $Definition.categoryAxis
            }
            if (Has-Property $Definition 'valueAxis') {
                Set-NativeChartAxisDefinition $Chart 2 1 $Definition.valueAxis
            }
            if (Has-Property $Definition 'secondaryCategoryAxis') {
                Set-NativeChartAxisDefinition $Chart 1 2 $Definition.secondaryCategoryAxis
            }
            if (Has-Property $Definition 'secondaryValueAxis') {
                Set-NativeChartAxisDefinition $Chart 2 2 $Definition.secondaryValueAxis
            }
        }
        if (Has-Property $Definition 'roundedCorners') {
            $ChartObject.RoundedCorners = [bool]$Definition.roundedCorners
        }
        if (
            (Has-Property $Definition 'gapWidth') -or
            (Has-Property $Definition 'overlap')
        ) {
            $chartGroups = $Chart.ChartGroups()
            $gapWidthCompatibleGroups = 0
            $overlapCompatibleGroups = 0
            for ($index = 1; $index -le [int]$chartGroups.Count; $index++) {
                $chartGroup = $null
                try {
                    $chartGroup = $chartGroups.Item($index)
                    if (
                        (Has-Property $Definition 'gapWidth') -and
                        (Test-NativeChartGroupSupportsOption `
                            $chartGroup `
                            $script:NativeGapWidthChartTypes `
                            @(2, 3, 5))
                    ) {
                        $gapWidthCompatibleGroups++
                        $chartGroup.GapWidth = [int]$Definition.gapWidth
                        if ([int]$chartGroup.GapWidth -ne [int]$Definition.gapWidth) {
                            throw "Native chart gapWidth failed for chart group ${index}."
                        }
                    }
                    if (
                        (Has-Property $Definition 'overlap') -and
                        (Test-NativeChartGroupSupportsOption `
                            $chartGroup `
                            $script:NativeOverlapChartTypes `
                            @(2, 3))
                    ) {
                        $overlapCompatibleGroups++
                        $chartGroup.Overlap = [int]$Definition.overlap
                        if ([int]$chartGroup.Overlap -ne [int]$Definition.overlap) {
                            throw "Native chart overlap failed for chart group ${index}."
                        }
                    }
                }
                finally {
                    Release-ComObject $chartGroup
                }
            }
            if (
                (Has-Property $Definition 'gapWidth') -and
                $gapWidthCompatibleGroups -eq 0
            ) {
                throw 'Native chart gapWidth is unsupported by every chart group.'
            }
            if (
                (Has-Property $Definition 'overlap') -and
                $overlapCompatibleGroups -eq 0
            ) {
                throw 'Native chart overlap is unsupported by every chart group.'
            }
        }
        if (Has-Property $Definition 'alternativeText') {
            $shapeRange = $ChartObject.ShapeRange
            $shapeRange.AlternativeText = [string]$Definition.alternativeText
        }
    }
    finally {
        Release-ComObject $chartGroups
        Release-ComObject $shapeRange
        Release-ComObject $chartTitle
        Release-ComObject $sourceRange
    }
}

function Apply-CreateChartOperation {
    param(
        [object]$Workbook,
        [object]$Operation
    )

    Assert-AllowedProperties `
        $Operation `
        @('kind', 'sheetName', 'chart') `
        'Native createChart operation'
    $sheetName = Get-OperationSheetName $Operation
    $definition = Get-RequiredProperty $Operation 'chart'
    Assert-NativeChartDefinition $definition
	if (
		(Has-Property $definition 'legend') -and
		[string]$definition.legend.position -ceq 'custom'
	) {
		throw 'Native createChart cannot create a custom legend layout; custom only preserves an existing manual Excel layout.'
	}

    $worksheet = $null
    $chartObjects = $null
    $chartObject = $null
    $chart = $null
    try {
        $worksheet = $Workbook.Worksheets.Item($sheetName)
        $chartObjects = $worksheet.ChartObjects()
        $chartObject = $chartObjects.Add(
            [double]$definition.anchor.left,
            [double]$definition.anchor.top,
            [double]$definition.anchor.width,
            [double]$definition.anchor.height
        )
        $chart = $chartObject.Chart
        Set-NativeChartDefinition $chartObject $chart $worksheet $definition
    }
    finally {
        Release-ComObject $chart
        Release-ComObject $chartObject
        Release-ComObject $chartObjects
        Release-ComObject $worksheet
    }
}

function Apply-UpdateChartOperation {
    param(
        [object]$Workbook,
        [object]$Operation
    )

    Assert-AllowedProperties `
        $Operation `
        @(
            'kind', 'sheetName', 'name', 'chart',
			'preserveAnchor', 'preserveSeries', 'allowSeriesFormattingChange'
        ) `
        'Native updateChart operation'
    $sheetName = Get-OperationSheetName $Operation
    $currentName = Get-RequiredProperty $Operation 'name'
    Assert-NativeObjectName $currentName 'Native current chart name'
    $definition = Get-RequiredProperty $Operation 'chart'
    Assert-NativeChartDefinition $definition
	foreach ($propertyName in @('preserveAnchor', 'preserveSeries', 'allowSeriesFormattingChange')) {
        if (Has-Property $Operation $propertyName) {
            Assert-NativeBoolean `
                $Operation.$propertyName `
                "Native updateChart $propertyName"
        }
    }

    $worksheet = $null
    $chartObjects = $null
    $chartObject = $null
    $chart = $null
    try {
        $worksheet = $Workbook.Worksheets.Item($sheetName)
        $chartObjects = $worksheet.ChartObjects()
        $chartObject = $chartObjects.Item([string]$currentName)
        $chart = $chartObject.Chart
		if (
			(Has-Property $definition 'legend') -and
			[bool]$definition.legend.visible -and
			[string]$definition.legend.position -ceq 'custom'
		) {
			if (-not [bool]$chart.HasLegend) {
				throw 'Native updateChart custom legend position requires an existing manual Excel legend layout.'
			}
			$currentLegend = $null
			try {
				$currentLegend = $chart.Legend
				if ([int]$currentLegend.Position -ne -4161) {
					throw 'Native updateChart custom legend position requires an existing manual Excel legend layout.'
				}
			}
			finally {
				Release-ComObject $currentLegend
			}
		}
        $preserveAnchor =
            (Has-Property $Operation 'preserveAnchor') -and
            [bool]$Operation.preserveAnchor
        $preserveSeries =
            (Has-Property $Operation 'preserveSeries') -and
            [bool]$Operation.preserveSeries
		$allowSeriesFormattingChange =
			(Has-Property $Operation 'allowSeriesFormattingChange') -and
			[bool]$Operation.allowSeriesFormattingChange
		if ($allowSeriesFormattingChange -and -not $preserveSeries) {
			throw 'Native updateChart allowSeriesFormattingChange requires preserveSeries=true.'
		}
        Set-NativeChartDefinition `
            $chartObject `
            $chart `
            $worksheet `
            $definition `
            $preserveAnchor `
			$preserveSeries `
			$allowSeriesFormattingChange
    }
    finally {
        Release-ComObject $chart
        Release-ComObject $chartObject
        Release-ComObject $chartObjects
        Release-ComObject $worksheet
    }
}

function Apply-DeleteChartOperation {
    param(
        [object]$Workbook,
        [object]$Operation
    )

    Assert-AllowedProperties `
        $Operation `
        @('kind', 'sheetName', 'name') `
        'Native deleteChart operation'
    $sheetName = Get-OperationSheetName $Operation
    $currentName = Get-RequiredProperty $Operation 'name'
    Assert-NativeObjectName $currentName 'Native current chart name'

    $worksheet = $null
    $chartObjects = $null
    $chartObject = $null
    try {
        $worksheet = $Workbook.Worksheets.Item($sheetName)
        $chartObjects = $worksheet.ChartObjects()
        $chartObject = $chartObjects.Item([string]$currentName)
        $chartObject.Delete()
    }
    finally {
        Release-ComObject $chartObject
        Release-ComObject $chartObjects
        Release-ComObject $worksheet
    }
}

function Get-NativeWorkbookObjectKey {
    param(
        [string]$SheetName,
        [string]$ObjectName
    )
    return $SheetName + [char]31 + $ObjectName
}

function Get-NativeTargetedWorkbookObjectSets {
    param([object[]]$Operations)

    $tables = New-Object 'System.Collections.Generic.HashSet[string]' `
        ([StringComparer]::OrdinalIgnoreCase)
    $charts = New-Object 'System.Collections.Generic.HashSet[string]' `
        ([StringComparer]::OrdinalIgnoreCase)
    foreach ($operation in $Operations) {
        $kind = if (Has-Property $operation 'kind') {
            [string]$operation.kind
        }
        else {
            'cell'
        }
        if ($kind -notin @(
            'createTable', 'updateTable', 'deleteTable',
            'createChart', 'updateChart', 'deleteChart'
        )) {
            continue
        }
        $sheetName = Get-OperationSheetName $operation
        $isTableOperation = $kind -like '*Table'
        if (Has-Property $operation 'name') {
            $key = Get-NativeWorkbookObjectKey $sheetName ([string]$operation.name)
            if ($isTableOperation) {
                [void]$tables.Add($key)
            }
            else {
                [void]$charts.Add($key)
            }
        }
        $definitionProperty = if ($isTableOperation) { 'table' } else { 'chart' }
        if (Has-Property $operation $definitionProperty) {
            $definition = $operation.$definitionProperty
            $key = Get-NativeWorkbookObjectKey $sheetName ([string]$definition.name)
            if ($isTableOperation) {
                [void]$tables.Add($key)
            }
            else {
                [void]$charts.Add($key)
            }
        }
    }
    return [PSCustomObject]@{
        Tables = $tables
        Charts = $charts
    }
}

function Get-NativeUntargetedWorkbookObjectSnapshot {
    param(
        [object]$Workbook,
        [object]$TargetedObjects
    )

    $items = New-Object 'System.Collections.Generic.List[object]'
    $worksheets = $null
    try {
        $worksheets = $Workbook.Worksheets
        for ($sheetIndex = 1; $sheetIndex -le [int]$worksheets.Count; $sheetIndex++) {
            $worksheet = $null
            $listObjects = $null
            $chartObjects = $null
            try {
                $worksheet = $worksheets.Item($sheetIndex)
                $sheetName = [string]$worksheet.Name
                $listObjects = $worksheet.ListObjects
                for ($index = 1; $index -le [int]$listObjects.Count; $index++) {
                    $listObject = $null
                    $range = $null
                    $tableStyle = $null
                    try {
                        $listObject = $listObjects.Item($index)
                        $key = Get-NativeWorkbookObjectKey $sheetName ([string]$listObject.Name)
                        if ($TargetedObjects.Tables.Contains($key)) {
                            continue
                        }
                        $range = $listObject.Range
                        $tableStyle = $listObject.TableStyle
                        [void]$items.Add([PSCustomObject]@{
                            type = 'table'
                            sheet = $sheetName
                            name = [string]$listObject.Name
                            displayName = [string]$listObject.DisplayName
                            rangeRef = ([string]$range.Address($false, $false)).ToUpperInvariant()
                            headerRow = [bool]$listObject.ShowHeaders
                            totalsRow = [bool]$listObject.ShowTotals
                            style = [string]$tableStyle.Name
                            firstColumn = [bool]$listObject.ShowTableStyleFirstColumn
                            lastColumn = [bool]$listObject.ShowTableStyleLastColumn
                            rowStripes = [bool]$listObject.ShowTableStyleRowStripes
                            columnStripes = [bool]$listObject.ShowTableStyleColumnStripes
                        })
                    }
                    finally {
                        Release-ComObject $tableStyle
                        Release-ComObject $range
                        Release-ComObject $listObject
                    }
                }

                $chartObjects = $worksheet.ChartObjects()
                for ($index = 1; $index -le [int]$chartObjects.Count; $index++) {
                    $chartObject = $null
                    $chart = $null
                    $chartTitle = $null
                    $legend = $null
                    $seriesCollection = $null
                    try {
                        $chartObject = $chartObjects.Item($index)
                        $key = Get-NativeWorkbookObjectKey $sheetName ([string]$chartObject.Name)
                        if ($TargetedObjects.Charts.Contains($key)) {
                            continue
                        }
                        $chart = $chartObject.Chart
                        $seriesFormulas = New-Object 'System.Collections.Generic.List[string]'
                        $seriesCollection = $chart.FullSeriesCollection()
                        for ($seriesIndex = 1; $seriesIndex -le [int]$seriesCollection.Count; $seriesIndex++) {
                            $series = $null
                            try {
                                $series = $seriesCollection.Item($seriesIndex)
                                [void]$seriesFormulas.Add([string]$series.Formula)
                            }
                            finally {
                                Release-ComObject $series
                            }
                        }
                        $titleText = ''
                        if ([bool]$chart.HasTitle) {
                            $chartTitle = $chart.ChartTitle
                            $titleText = [string]$chartTitle.Text
                        }
                        $legendPosition = 0
                        if ([bool]$chart.HasLegend) {
                            $legend = $chart.Legend
                            $legendPosition = [int]$legend.Position
                        }
                        [void]$items.Add([PSCustomObject]@{
                            type = 'chart'
                            sheet = $sheetName
                            name = [string]$chartObject.Name
                            chartType = [int]$chart.ChartType
                            plotBy = [int]$chart.PlotBy
                            left = [Math]::Round([double]$chartObject.Left, 3)
                            top = [Math]::Round([double]$chartObject.Top, 3)
                            width = [Math]::Round([double]$chartObject.Width, 3)
                            height = [Math]::Round([double]$chartObject.Height, 3)
                            hasTitle = [bool]$chart.HasTitle
                            title = $titleText
                            hasLegend = [bool]$chart.HasLegend
                            legendPosition = $legendPosition
                            roundedCorners = [bool]$chartObject.RoundedCorners
                            series = $seriesFormulas.ToArray()
                        })
                    }
                    finally {
                        Release-ComObject $seriesCollection
                        Release-ComObject $legend
                        Release-ComObject $chartTitle
                        Release-ComObject $chart
                        Release-ComObject $chartObject
                    }
                }
            }
            finally {
                Release-ComObject $chartObjects
                Release-ComObject $listObjects
                Release-ComObject $worksheet
            }
        }
    }
    finally {
        Release-ComObject $worksheets
    }
    return @(
        $items.ToArray() |
            Sort-Object type, sheet, name |
            ConvertTo-Json -Compress -Depth 8
    ) -join ''
}

function Compare-NativeWorkbookObjectSnapshots {
    param(
        [string]$BeforeJson,
        [string]$AfterJson
    )

    $beforeItems = if ([string]::IsNullOrWhiteSpace($BeforeJson)) {
        @()
    }
    else {
        @($BeforeJson | ConvertFrom-Json)
    }
    $afterItems = if ([string]::IsNullOrWhiteSpace($AfterJson)) {
        @()
    }
    else {
        @($AfterJson | ConvertFrom-Json)
    }
    $beforeByKey = @{}
    $afterByKey = @{}
    foreach ($item in $beforeItems) {
        $beforeByKey["$($item.type):$($item.sheet):$($item.name)"] = $item
    }
    foreach ($item in $afterItems) {
        $afterByKey["$($item.type):$($item.sheet):$($item.name)"] = $item
    }
    $differences = New-Object 'System.Collections.Generic.List[string]'
    $keys = @($beforeByKey.Keys + $afterByKey.Keys | Sort-Object -Unique)
    foreach ($key in $keys) {
        if (-not $beforeByKey.ContainsKey($key)) {
            [void]$differences.Add("$key added")
            continue
        }
        if (-not $afterByKey.ContainsKey($key)) {
            [void]$differences.Add("$key missing")
            continue
        }
        $beforeItem = $beforeByKey[$key]
        $afterItem = $afterByKey[$key]
        $propertyNames = @(
            @($beforeItem.PSObject.Properties.Name) +
            @($afterItem.PSObject.Properties.Name) |
                Sort-Object -Unique
        )
        $changedProperties = @(
            foreach ($propertyName in $propertyNames) {
                $beforeValue = $beforeItem.PSObject.Properties[$propertyName].Value
                $afterValue = $afterItem.PSObject.Properties[$propertyName].Value
                if (
                    [string]$beforeItem.type -ceq 'chart' -and
                    $propertyName -in @('left', 'top', 'width', 'height') -and
                    [Math]::Abs([double]$beforeValue - [double]$afterValue) -le 1
                ) {
                    continue
                }
                if (
                    ($beforeValue | ConvertTo-Json -Compress -Depth 8) -cne
                    ($afterValue | ConvertTo-Json -Compress -Depth 8)
                ) {
					if (
						[string]$beforeItem.type -ceq 'chart' -and
						$propertyName -in @('left', 'top', 'width', 'height')
					) {
						"$propertyName ($beforeValue -> $afterValue)"
					}
					else {
						$propertyName
					}
                }
            }
        )
        if ($changedProperties.Count -gt 0) {
            [void]$differences.Add(
                "$key changed: " + ($changedProperties -join ', ')
            )
        }
    }
    return $differences.ToArray()
}

function Restore-NativeUntargetedChartAnchors {
	param(
		[object]$Workbook,
		[object]$TargetedObjects,
		[string]$SnapshotJson
	)

	if ([string]::IsNullOrWhiteSpace($SnapshotJson)) { return }
	$snapshotItems = @($SnapshotJson | ConvertFrom-Json)
	foreach ($item in $snapshotItems) {
		if ([string]$item.type -cne 'chart') { continue }
		$key = Get-NativeWorkbookObjectKey ([string]$item.sheet) ([string]$item.name)
		if ($TargetedObjects.Charts.Contains($key)) { continue }

		$worksheet = $null
		$chartObjects = $null
		$chartObject = $null
		try {
			$worksheet = $Workbook.Worksheets.Item([string]$item.sheet)
			$chartObjects = $worksheet.ChartObjects()
			$chartObject = $chartObjects.Item([string]$item.name)
			# Creating or resizing a ListObject can move charts whose placement is
			# tied to cells. Reapply the captured point geometry before saving so a
			# mixed table/chart transaction remains non-destructive.
			$chartObject.Width = [double]$item.width
			$chartObject.Height = [double]$item.height
			$chartObject.Left = [double]$item.left
			$chartObject.Top = [double]$item.top
		}
		finally {
			Release-ComObject $chartObject
			Release-ComObject $chartObjects
			Release-ComObject $worksheet
		}
	}
}

function Get-NativePreservedChartSnapshot {
    param(
        [object]$Workbook,
        [object[]]$Operations,
        [bool]$AfterReopen
    )

    $items = New-Object 'System.Collections.Generic.List[object]'
    for ($operationIndex = 0; $operationIndex -lt $Operations.Count; $operationIndex++) {
        $operation = $Operations[$operationIndex]
        if (
            -not (Has-Property $operation 'kind') -or
            [string]$operation.kind -cne 'updateChart'
        ) {
            continue
        }
        $preserveAnchor =
            (Has-Property $operation 'preserveAnchor') -and
            [bool]$operation.preserveAnchor
        $preserveSeries =
            (Has-Property $operation 'preserveSeries') -and
            [bool]$operation.preserveSeries
		$allowSeriesFormattingChange =
			(Has-Property $operation 'allowSeriesFormattingChange') -and
			[bool]$operation.allowSeriesFormattingChange
        if (-not $preserveAnchor -and -not $preserveSeries) {
            continue
        }

        $worksheet = $null
        $chartObjects = $null
        $chartObject = $null
        $chart = $null
        $seriesCollection = $null
        try {
            $worksheet = $Workbook.Worksheets.Item([string]$operation.sheetName)
            $chartObjects = $worksheet.ChartObjects()
            $chartName = if ($AfterReopen) {
                [string]$operation.chart.name
            }
            else {
                [string]$operation.name
            }
            $chartObject = $chartObjects.Item($chartName)
            $chart = $chartObject.Chart
            $item = [ordered]@{
                operationIndex = $operationIndex
                preserveAnchor = $preserveAnchor
                preserveSeries = $preserveSeries
				allowSeriesFormattingChange = $allowSeriesFormattingChange
            }
            if ($preserveAnchor) {
                $item.left = [Math]::Round([double]$chartObject.Left, 3)
                $item.top = [Math]::Round([double]$chartObject.Top, 3)
                $item.width = [Math]::Round([double]$chartObject.Width, 3)
                $item.height = [Math]::Round([double]$chartObject.Height, 3)
            }
            if ($preserveSeries) {
                $seriesItems = New-Object 'System.Collections.Generic.List[object]'
                $seriesCollection = $chart.FullSeriesCollection()
                for ($index = 1; $index -le [int]$seriesCollection.Count; $index++) {
                    $series = $null
                    $trendlines = $null
                    $format = $null
                    $fill = $null
                    $fillColor = $null
                    $line = $null
                    $lineColor = $null
                    $labels = $null
                    try {
                        $series = $seriesCollection.Item($index)
                        $trendlineCount = try {
                            $trendlines = $series.Trendlines()
                            [int]$trendlines.Count
                        }
                        catch { -1 }
                        $hasErrorBars = try { [bool]$series.HasErrorBars } catch { $null }
                        $fillRgb = $null
                        $fillVisible = $null
                        $lineRgb = $null
                        $lineVisible = $null
                        $lineWeight = $null
                        $lineDashStyle = $null
                        try {
                            $format = $series.Format
                            $fill = $format.Fill
                            $fillVisible = [int]$fill.Visible
                            $fillColor = $fill.ForeColor
                            $fillRgb = [int]$fillColor.RGB
                            $line = $format.Line
                            $lineVisible = [int]$line.Visible
                            $lineColor = $line.ForeColor
                            $lineRgb = [int]$lineColor.RGB
                            $lineWeight = [Math]::Round([double]$line.Weight, 3)
                            $lineDashStyle = [int]$line.DashStyle
                        }
                        catch { }
                        $labelState = $null
                        $hasDataLabels = try { [bool]$series.HasDataLabels } catch { $false }
                        if ($hasDataLabels) {
                            try {
                                $labels = $series.DataLabels()
                                $labelState = [ordered]@{
                                    showValue = [bool]$labels.ShowValue
                                    showCategoryName = [bool]$labels.ShowCategoryName
                                    showSeriesName = [bool]$labels.ShowSeriesName
                                    showPercentage = [bool]$labels.ShowPercentage
                                    showBubbleSize = [bool]$labels.ShowBubbleSize
                                    position = [int]$labels.Position
                                }
                            }
                            catch { $labelState = '__unsupported__' }
                        }
                        $seriesChartType = try { [int]$series.ChartType } catch { $null }
                        $seriesAxisGroup = try { [int]$series.AxisGroup } catch { $null }
                        $seriesFiltered = try { [bool]$series.IsFiltered } catch { $null }
                        $seriesMarkerStyle = try { [int]$series.MarkerStyle } catch { $null }
                        $seriesMarkerSize = try { [int]$series.MarkerSize } catch { $null }
                        $seriesSmooth = try { [bool]$series.Smooth } catch { $null }
						$seriesItem = [ordered]@{
                            formula = [string]$series.Formula
                            chartType = $seriesChartType
                            axisGroup = $seriesAxisGroup
                            filtered = $seriesFiltered
                            smooth = $seriesSmooth
                            trendlineCount = $trendlineCount
                            hasErrorBars = $hasErrorBars
                            dataLabels = $labelState
						}
						if (-not $allowSeriesFormattingChange) {
							$seriesItem.markerStyle = $seriesMarkerStyle
							$seriesItem.markerSize = $seriesMarkerSize
							$seriesItem.fillVisible = $fillVisible
							$seriesItem.fillRgb = $fillRgb
							$seriesItem.lineVisible = $lineVisible
							$seriesItem.lineRgb = $lineRgb
							$seriesItem.lineWeight = $lineWeight
							$seriesItem.lineDashStyle = $lineDashStyle
						}
						[void]$seriesItems.Add([PSCustomObject]$seriesItem)
                    }
                    finally {
                        Release-ComObject $labels
                        Release-ComObject $lineColor
                        Release-ComObject $line
                        Release-ComObject $fillColor
                        Release-ComObject $fill
                        Release-ComObject $format
                        Release-ComObject $trendlines
                        Release-ComObject $series
                    }
                }
                $item.series = $seriesItems.ToArray()
            }
            [void]$items.Add([PSCustomObject]$item)
        }
        finally {
            Release-ComObject $seriesCollection
            Release-ComObject $chart
            Release-ComObject $chartObject
            Release-ComObject $chartObjects
            Release-ComObject $worksheet
        }
    }
    return @($items.ToArray() | ConvertTo-Json -Compress -Depth 10) -join ''
}

function Assert-NativeTableMatches {
    param(
        [object]$Worksheet,
        [object]$Definition
    )

    $listObjects = $null
    $listObject = $null
    $range = $null
    $tableStyle = $null
    try {
        $listObjects = $Worksheet.ListObjects
        $listObject = $listObjects.Item([string]$Definition.name)
        $range = $listObject.Range
        $tableStyle = $listObject.TableStyle
        $actualRange = ([string]$range.Address($false, $false)).ToUpperInvariant()
        if (
            [string]$listObject.Name -cne [string]$Definition.name -or
            [string]$listObject.DisplayName -cne [string]$Definition.displayName -or
            $actualRange -cne (Get-NormalizedRangeRef $Definition.rangeRef) -or
            [bool]$listObject.ShowHeaders -ne [bool]$Definition.headerRow -or
            [bool]$listObject.ShowTotals -ne [bool]$Definition.totalsRow -or
            [string]$tableStyle.Name -cne [string]$Definition.style.name -or
            [bool]$listObject.ShowTableStyleFirstColumn -ne [bool]$Definition.style.showFirstColumn -or
            [bool]$listObject.ShowTableStyleLastColumn -ne [bool]$Definition.style.showLastColumn -or
            [bool]$listObject.ShowTableStyleRowStripes -ne [bool]$Definition.style.showRowStripes -or
            [bool]$listObject.ShowTableStyleColumnStripes -ne [bool]$Definition.style.showColumnStripes
        ) {
			throw (
				"Table verification failed after reopen: $($Definition.name); " +
				"range=$actualRange expected=$(Get-NormalizedRangeRef $Definition.rangeRef); " +
				"headers=$([bool]$listObject.ShowHeaders) expectedHeaders=$([bool]$Definition.headerRow); " +
				"totals=$([bool]$listObject.ShowTotals) expectedTotals=$([bool]$Definition.totalsRow); " +
				"name=$([string]$listObject.Name); displayName=$([string]$listObject.DisplayName); " +
				"style=$([string]$tableStyle.Name)"
			)
        }
    }
    finally {
        Release-ComObject $tableStyle
        Release-ComObject $range
        Release-ComObject $listObject
        Release-ComObject $listObjects
    }
}

function ConvertFrom-NativeSeriesFormula {
    param([string]$Formula)

    $match = [regex]::Match(
        $Formula,
        '^\s*=\s*SERIES\s*\((?<arguments>.*)\)\s*$',
        [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
            [Text.RegularExpressions.RegexOptions]::Singleline
    )
    if (-not $match.Success) {
        throw 'Native chart series formula is invalid.'
    }

    $source = $match.Groups['arguments'].Value
    $arguments = New-Object 'System.Collections.Generic.List[string]'
    $current = New-Object System.Text.StringBuilder
    $inString = $false
    $inSheetQuote = $false
    $bracketDepth = 0
	$braceDepth = 0
    $parenthesisDepth = 0
    $separator = [char]0

    for ($offset = 0; $offset -lt $source.Length; $offset++) {
        $character = $source[$offset]
        if ($inString) {
            [void]$current.Append($character)
            if ($character -eq [char]34) {
                if (
                    $offset + 1 -lt $source.Length -and
                    $source[$offset + 1] -eq [char]34
                ) {
                    [void]$current.Append($source[$offset + 1])
                    $offset++
                }
                else {
                    $inString = $false
                }
            }
            continue
        }
        if ($inSheetQuote) {
            [void]$current.Append($character)
            if ($character -eq [char]39) {
                if (
                    $offset + 1 -lt $source.Length -and
                    $source[$offset + 1] -eq [char]39
                ) {
                    [void]$current.Append($source[$offset + 1])
                    $offset++
                }
                else {
                    $inSheetQuote = $false
                }
            }
            continue
        }
        if ($character -eq [char]34) {
            $inString = $true
            [void]$current.Append($character)
            continue
        }
        if ($character -eq [char]39) {
            $inSheetQuote = $true
            [void]$current.Append($character)
            continue
        }
        if ($character -eq [char]91) {
            $bracketDepth++
            [void]$current.Append($character)
            continue
        }
        if ($character -eq [char]93) {
            if ($bracketDepth -le 0) {
                throw 'Native chart series formula contains an unmatched bracket.'
            }
            $bracketDepth--
            [void]$current.Append($character)
            continue
        }
		if ($character -eq [char]123) {
			$braceDepth++
			[void]$current.Append($character)
			continue
		}
		if ($character -eq [char]125) {
			if ($braceDepth -le 0) {
				throw 'Native chart series formula contains an unmatched brace.'
			}
			$braceDepth--
			[void]$current.Append($character)
			continue
		}
        if ($character -eq [char]40) {
            $parenthesisDepth++
            [void]$current.Append($character)
            continue
        }
        if ($character -eq [char]41) {
            if ($parenthesisDepth -le 0) {
                throw 'Native chart series formula contains an unmatched parenthesis.'
            }
            $parenthesisDepth--
            [void]$current.Append($character)
            continue
        }

        $isSeparator = $character -eq [char]44 -or $character -eq [char]59
        if (
            $isSeparator -and
            $bracketDepth -eq 0 -and
			$braceDepth -eq 0 -and
            $parenthesisDepth -eq 0
        ) {
            if ($separator -eq [char]0) {
                $separator = $character
            }
            elseif ($separator -ne $character) {
                throw 'Native chart series formula mixes argument separators.'
            }
            [void]$arguments.Add($current.ToString().Trim())
            [void]$current.Clear()
            continue
        }
        [void]$current.Append($character)
    }

    if (
        $inString -or
        $inSheetQuote -or
        $bracketDepth -ne 0 -or
		$braceDepth -ne 0 -or
        $parenthesisDepth -ne 0
    ) {
        throw 'Native chart series formula is not balanced.'
    }
    [void]$arguments.Add($current.ToString().Trim())
    if ($arguments.Count -lt 4 -or $arguments.Count -gt 5) {
		throw (
			'Native chart SERIES formula must contain four or five arguments; ' +
			"received $($arguments.Count)."
		)
    }
    if ([string]::IsNullOrWhiteSpace($arguments[2])) {
        throw 'Native chart SERIES values argument is missing.'
    }

    $plotOrder = 0
    if (
        -not [int]::TryParse(
            $arguments[3],
            [Globalization.NumberStyles]::Integer,
            [Globalization.CultureInfo]::InvariantCulture,
            [ref]$plotOrder
        ) -or
        $plotOrder -lt 1
    ) {
        throw 'Native chart SERIES plot order is invalid.'
    }

    return [PSCustomObject]@{
        NameArgument = $arguments[0]
        XValuesArgument = $arguments[1]
        ValuesArgument = $arguments[2]
        PlotOrder = $plotOrder
        BubbleSizesArgument = if ($arguments.Count -eq 5) {
            $arguments[4]
        }
        else {
            $null
        }
    }
}

function Get-NativeSeriesFormulaLiteral {
    param([string]$Argument)

    $value = $Argument.Trim()
    if ($value -cnotmatch '^"(?:[^"]|"")*"$') {
        throw 'Native chart SERIES name is not a literal string.'
    }
    return $value.Substring(1, $value.Length - 2).Replace('""', '"')
}

function Get-NormalizedNativeSeriesFormulaRange {
    param(
        [string]$Argument,
        [string]$WorksheetName
    )

    $candidate = $Argument.Trim()
    if ($candidate.StartsWith('=')) {
        $candidate = $candidate.Substring(1).Trim()
    }
    $bangIndex = $candidate.LastIndexOf([char]33)
    if ($bangIndex -ge 0) {
        if ($bangIndex -eq 0 -or $bangIndex -eq $candidate.Length - 1) {
            throw 'Native chart SERIES range qualification is invalid.'
        }
        $sheetToken = $candidate.Substring(0, $bangIndex).Trim()
        $candidate = $candidate.Substring($bangIndex + 1).Trim()
        if ($sheetToken[0] -eq [char]39) {
            if (
                $sheetToken.Length -lt 2 -or
                $sheetToken[$sheetToken.Length - 1] -ne [char]39
            ) {
                throw 'Native chart SERIES sheet quoting is invalid.'
            }
            $quoted = $sheetToken.Substring(1, $sheetToken.Length - 2)
            $decoded = New-Object System.Text.StringBuilder
            for ($offset = 0; $offset -lt $quoted.Length; $offset++) {
                if ($quoted[$offset] -eq [char]39) {
                    if (
                        $offset + 1 -ge $quoted.Length -or
                        $quoted[$offset + 1] -ne [char]39
                    ) {
                        throw 'Native chart SERIES sheet apostrophe is invalid.'
                    }
                    [void]$decoded.Append([char]39)
                    $offset++
                }
                else {
                    [void]$decoded.Append($quoted[$offset])
                }
            }
            $sheetToken = $decoded.ToString()
        }
        elseif ($sheetToken.IndexOf([char]39) -ge 0) {
            throw 'Native chart SERIES sheet quoting is invalid.'
        }
        if (
            $sheetToken.IndexOf([char]91) -ge 0 -or
            $sheetToken.IndexOf([char]93) -ge 0
        ) {
            throw 'Native chart SERIES external workbook reference is forbidden.'
        }
        if (
            -not [StringComparer]::OrdinalIgnoreCase.Equals(
                $sheetToken,
                $WorksheetName
            )
        ) {
            throw 'Native chart SERIES range targets another worksheet.'
        }
    }
    return Get-NormalizedRangeRef ($candidate.Replace('$', ''))
}

function Assert-NativeSeriesFormulaRangeArgument {
    param(
        [string]$ActualArgument,
        [string]$ExpectedRange,
        [string]$WorksheetName,
        [string]$Label
    )

    $actual = Get-NormalizedNativeSeriesFormulaRange `
        $ActualArgument `
        $WorksheetName
    $expected = Get-NormalizedRangeRef $ExpectedRange
    if ($actual -cne $expected) {
        throw "$Label verification failed: actual=$actual expected=$expected"
    }
}

function Get-NativeRangeBounds {
    param([string]$RangeRef)

    $normalized = Get-NormalizedRangeRef $RangeRef
    $match = [regex]::Match(
        $normalized,
        '^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$'
    )
    return [PSCustomObject]@{
        StartColumn = Convert-ExcelColumnLettersToNumber $match.Groups[1].Value
        StartRow = [int]$match.Groups[2].Value
        EndColumn = if ($match.Groups[3].Success) {
            Convert-ExcelColumnLettersToNumber $match.Groups[3].Value
        }
        else {
            Convert-ExcelColumnLettersToNumber $match.Groups[1].Value
        }
        EndRow = if ($match.Groups[4].Success) {
            [int]$match.Groups[4].Value
        }
        else {
            [int]$match.Groups[2].Value
        }
    }
}

function Add-NativeChartSourceCoverageRange {
    param(
        [string]$RangeRef,
        [object]$ExpectedBounds,
        [bool[]]$CoveredCells
    )

    $bounds = Get-NativeRangeBounds $RangeRef
    if (
        $bounds.StartColumn -lt $ExpectedBounds.StartColumn -or
        $bounds.StartRow -lt $ExpectedBounds.StartRow -or
        $bounds.EndColumn -gt $ExpectedBounds.EndColumn -or
        $bounds.EndRow -gt $ExpectedBounds.EndRow
    ) {
        throw 'Chart source formula references a cell outside the requested source range.'
    }
    $width = $ExpectedBounds.EndColumn - $ExpectedBounds.StartColumn + 1
    for ($row = $bounds.StartRow; $row -le $bounds.EndRow; $row++) {
        for ($column = $bounds.StartColumn; $column -le $bounds.EndColumn; $column++) {
            $offset =
                (($row - $ExpectedBounds.StartRow) * $width) +
                ($column - $ExpectedBounds.StartColumn)
            $CoveredCells[$offset] = $true
        }
    }
}

function Test-NativeSeriesFormulaConstantArgument {
    param([string]$Argument)

    $candidate = $Argument.Trim()
    return (
        $candidate -match '^"(?:[^"]|"")*"$' -or
        $candidate -match '^=?\{.*\}$'
    )
}

function Assert-NativeChartSourceCoverage {
    param(
        [object]$SeriesCollection,
        [string]$ExpectedRange,
        [string]$WorksheetName
    )

    if ([int]$SeriesCollection.Count -lt 1) {
        throw 'Chart source-range verification found no series.'
    }
    $expectedBounds = Get-NativeRangeBounds $ExpectedRange
    $width = $expectedBounds.EndColumn - $expectedBounds.StartColumn + 1
    $height = $expectedBounds.EndRow - $expectedBounds.StartRow + 1
    [long]$cellCount = [long]$width * [long]$height
    if ($cellCount -gt $MaxWorkbookObjectRangeCells) {
        throw 'Chart source-range verification exceeds the native safety limit.'
    }
    $coveredCells = New-Object 'bool[]' ([int]$cellCount)
    # Excel does not include the corner shared by row and column headers in any
    # SERIES argument. It is nevertheless part of SetSourceData's source range.
    $coveredCells[0] = $true

    for ($index = 1; $index -le [int]$SeriesCollection.Count; $index++) {
        $series = $null
        try {
            $series = $SeriesCollection.Item($index)
			try {
				$formula = ConvertFrom-NativeSeriesFormula ([string]$series.Formula)
			}
			catch {
				throw "Chart source series $index formula validation failed. $($_.Exception.Message)"
			}
            if ($formula.PlotOrder -ne $index) {
                throw 'Chart source-range plot order verification failed.'
            }
            foreach ($argumentSpec in @(
                @{ Value = $formula.NameArgument; ConstantsAllowed = $true; Required = $false },
                @{ Value = $formula.XValuesArgument; ConstantsAllowed = $true; Required = $false },
                @{ Value = $formula.ValuesArgument; ConstantsAllowed = $false; Required = $true },
                @{ Value = $formula.BubbleSizesArgument; ConstantsAllowed = $true; Required = $false }
            )) {
                $argument = [string]$argumentSpec.Value
                if ([string]::IsNullOrWhiteSpace($argument)) {
                    if ([bool]$argumentSpec.Required) {
                        throw 'Chart source-range formula is missing a required range.'
                    }
                    continue
                }
                if (
                    [bool]$argumentSpec.ConstantsAllowed -and
                    (Test-NativeSeriesFormulaConstantArgument $argument)
                ) {
                    continue
                }
                $normalizedRange = Get-NormalizedNativeSeriesFormulaRange `
                    $argument `
                    $WorksheetName
                Add-NativeChartSourceCoverageRange `
                    $normalizedRange `
                    $expectedBounds `
                    $coveredCells
            }
        }
        finally {
            Release-ComObject $series
        }
    }
    for ($offset = 0; $offset -lt $coveredCells.Length; $offset++) {
        if (-not $coveredCells[$offset]) {
            throw 'Chart source-range formulas do not cover the requested source range.'
        }
    }
}

function Assert-NativeChartAxisMatches {
    param(
        [object]$Chart,
        [int]$AxisType,
        [int]$AxisGroup,
        [object]$Definition,
        [string]$Label
    )

    $expectedVisible = if (Has-Property $Definition 'visible') {
        [bool]$Definition.visible
    }
    else {
        $true
    }
    $actualVisible = [bool]$Chart.HasAxis($AxisType, $AxisGroup)
    if ($actualVisible -ne $expectedVisible) {
        throw "$Label visibility verification failed."
    }
    if (-not $expectedVisible) {
        return
    }

    $axis = $null
    $axisTitle = $null
    $tickLabels = $null
    try {
        $axis = $Chart.Axes($AxisType, $AxisGroup)
        if (Has-Property $Definition 'title') {
            $expectedHasTitle = ([string]$Definition.title).Length -gt 0
            if ([bool]$axis.HasTitle -ne $expectedHasTitle) {
                throw "$Label title visibility verification failed."
            }
            if ($expectedHasTitle) {
                $axisTitle = $axis.AxisTitle
                if ([string]$axisTitle.Text -cne [string]$Definition.title) {
                    throw "$Label title verification failed."
                }
            }
        }
        foreach ($scaleProperty in @(
            @{ Name = 'minimumScale'; Auto = 'MinimumScaleIsAuto'; Value = 'MinimumScale' },
            @{ Name = 'maximumScale'; Auto = 'MaximumScaleIsAuto'; Value = 'MaximumScale' },
            @{ Name = 'majorUnit'; Auto = 'MajorUnitIsAuto'; Value = 'MajorUnit' },
            @{ Name = 'minorUnit'; Auto = 'MinorUnitIsAuto'; Value = 'MinorUnit' }
        )) {
            if (-not (Has-Property $Definition $scaleProperty.Name)) {
                continue
            }
            $expectedValue = $Definition.($scaleProperty.Name)
            $actualIsAuto = [bool]$axis.($scaleProperty.Auto)
            if ($null -eq $expectedValue) {
                if (-not $actualIsAuto) {
                    throw "$Label $($scaleProperty.Name) auto verification failed."
                }
                continue
            }
            if ($actualIsAuto) {
                throw "$Label $($scaleProperty.Name) unexpectedly remained automatic."
            }
            $actualValue = [double]$axis.($scaleProperty.Value)
            $tolerance = [Math]::Max(1e-9, [Math]::Abs([double]$expectedValue) * 1e-9)
            if ([Math]::Abs($actualValue - [double]$expectedValue) -gt $tolerance) {
                throw "$Label $($scaleProperty.Name) verification failed."
            }
        }
        if (
            (Has-Property $Definition 'logarithmic') -and
            ([int]$axis.ScaleType -eq -4133) -ne [bool]$Definition.logarithmic
        ) {
            throw "$Label logarithmic-scale verification failed."
        }
        if (
            (Has-Property $Definition 'reverseOrder') -and
            [bool]$axis.ReversePlotOrder -ne [bool]$Definition.reverseOrder
        ) {
            throw "$Label reverse-order verification failed."
        }
		if (Has-Property $Definition 'numberFormat') {
			$tickLabels = $axis.TickLabels
			$expectedNumberFormat = [string]$Definition.numberFormat
			$expectedNumberFormatLinked = $expectedNumberFormat.Length -eq 0
			if ([bool]$tickLabels.NumberFormatLinked -ne $expectedNumberFormatLinked) {
				throw "$Label number-format linked-state verification failed."
			}
			if (-not $expectedNumberFormatLinked) {
				$persistedNumberFormat = [string]$tickLabels.NumberFormat
				if ($persistedNumberFormat -cne $expectedNumberFormat) {
					# Excel can serialize an equivalent custom format using a
					# different locale-specific spelling. Canonicalize the expected
					# value through the same reopened, read-only chart and accept it
					# only when Excel resolves both spellings to the same exact value.
					$tickLabels.NumberFormat = $expectedNumberFormat
					$tickLabels.NumberFormatLinked = $false
					$canonicalExpectedNumberFormat = [string]$tickLabels.NumberFormat
					if ($persistedNumberFormat -cne $canonicalExpectedNumberFormat) {
						$expectedJson = ConvertTo-Json $expectedNumberFormat -Compress
						$actualJson = ConvertTo-Json $persistedNumberFormat -Compress
						$canonicalJson = ConvertTo-Json `
							$canonicalExpectedNumberFormat `
							-Compress
						throw (
							"$Label number-format verification failed; " +
							"expected=$expectedJson, actual=$actualJson, " +
							"canonicalExpected=$canonicalJson."
						)
					}
				}
            }
        }
        if (
            (Has-Property $Definition 'majorGridlines') -and
            [bool]$axis.HasMajorGridlines -ne [bool]$Definition.majorGridlines
        ) {
            throw "$Label major-gridlines verification failed."
        }
        if (
            (Has-Property $Definition 'minorGridlines') -and
            [bool]$axis.HasMinorGridlines -ne [bool]$Definition.minorGridlines
        ) {
            throw "$Label minor-gridlines verification failed."
        }
    }
    finally {
        Release-ComObject $tickLabels
        Release-ComObject $axisTitle
        Release-ComObject $axis
    }
}

function Assert-NativeChartMatches {
    param(
        [object]$Worksheet,
        [object]$Operation
    )

    $Definition = $Operation.chart
    $preserveAnchor =
        (Has-Property $Operation 'preserveAnchor') -and
        [bool]$Operation.preserveAnchor

    $chartObjects = $null
    $chartObject = $null
    $chart = $null
    $chartTitle = $null
    $legend = $null
    $seriesCollection = $null
    $shapeRange = $null
    $chartGroups = $null
    try {
        $chartObjects = $Worksheet.ChartObjects()
        $chartObject = $chartObjects.Item([string]$Definition.name)
        $chart = $chartObject.Chart
        $expectedPlotBy = if ([string]$Definition.plotBy -ceq 'rows') { 1 } else { 2 }
        $actualChartType = [int]$chart.ChartType
        $chartTypeMatches =
            $actualChartType -eq [int]$Definition.chartType -or
            (
                [int]$Definition.chartType -eq -4152 -and
                $actualChartType -eq -4111
            )
        if (
            [string]$chartObject.Name -cne [string]$Definition.name -or
            -not $chartTypeMatches -or
            (
                (Has-Property $Definition 'sourceRangeRef') -and
                [int]$chart.PlotBy -ne $expectedPlotBy
            ) -or
            (
                -not $preserveAnchor -and
                (
                    [Math]::Abs([double]$chartObject.Left - [double]$Definition.anchor.left) -gt 1 -or
                    [Math]::Abs([double]$chartObject.Top - [double]$Definition.anchor.top) -gt 1 -or
                    [Math]::Abs([double]$chartObject.Width - [double]$Definition.anchor.width) -gt 1 -or
                    [Math]::Abs([double]$chartObject.Height - [double]$Definition.anchor.height) -gt 1
                )
            )
        ) {
            throw (
                "Chart verification failed after reopen: $($Definition.name); " +
                "type=${actualChartType}/$([int]$Definition.chartType), " +
                "plotBy=$([int]$chart.PlotBy)/$expectedPlotBy, " +
                "anchor=$([double]$chartObject.Left),$([double]$chartObject.Top)," +
                "$([double]$chartObject.Width),$([double]$chartObject.Height)"
            )
        }
        if (Has-Property $Definition 'title') {
            if ([bool]$chart.HasTitle -ne [bool]$Definition.title.visible) {
                throw "Chart title visibility verification failed: $($Definition.name)"
            }
            if ([bool]$Definition.title.visible) {
                $chartTitle = $chart.ChartTitle
                if ([string]$chartTitle.Text -cne [string]$Definition.title.text) {
                    throw "Chart title verification failed: $($Definition.name)"
                }
            }
        }
        if (Has-Property $Definition 'legend') {
            if ([bool]$chart.HasLegend -ne [bool]$Definition.legend.visible) {
                throw "Chart legend visibility verification failed: $($Definition.name)"
            }
            if (
                [bool]$Definition.legend.visible -and
                [string]$Definition.legend.position -cne 'custom'
            ) {
                $legend = $chart.Legend
                $expectedPosition = switch ([string]$Definition.legend.position) {
                    'bottom' { -4107 }
                    'corner' { 2 }
                    'left' { -4131 }
                    'right' { -4152 }
                    'top' { -4160 }
                }
                if ([int]$legend.Position -ne $expectedPosition) {
                    throw "Chart legend position verification failed: $($Definition.name)"
                }
            }
        }
        if (
            (Has-Property $Definition 'style') -and
            [int]$chart.ChartStyle -ne [int]$Definition.style
        ) {
            throw "Chart style verification failed: $($Definition.name)"
        }
        if (
            (Has-Property $Definition 'roundedCorners') -and
            [bool]$chartObject.RoundedCorners -ne [bool]$Definition.roundedCorners
        ) {
            throw "Chart rounded-corners verification failed: $($Definition.name)"
        }
        if (Has-Property $Definition 'alternativeText') {
            $shapeRange = $chartObject.ShapeRange
            if ([string]$shapeRange.AlternativeText -cne [string]$Definition.alternativeText) {
                throw "Chart alternative-text verification failed: $($Definition.name)"
            }
        }
        if (Test-NativeChartTypeSupportsAxes ([int]$Definition.chartType)) {
            foreach ($axisSpec in @(
                @{ Name = 'categoryAxis'; Type = 1; Group = 1 },
                @{ Name = 'valueAxis'; Type = 2; Group = 1 },
                @{ Name = 'secondaryCategoryAxis'; Type = 1; Group = 2 },
                @{ Name = 'secondaryValueAxis'; Type = 2; Group = 2 }
            )) {
                if (Has-Property $Definition $axisSpec.Name) {
                    Assert-NativeChartAxisMatches `
                        $chart `
                        $axisSpec.Type `
                        $axisSpec.Group `
                        $Definition.($axisSpec.Name) `
                        "Chart $($axisSpec.Name): $($Definition.name)"
                }
            }
        }

        $seriesCollection = $chart.FullSeriesCollection()
        if (Has-Property $Definition 'series') {
            $seriesDefinitions = @($Definition.series)
            if ([int]$seriesCollection.Count -ne $seriesDefinitions.Count) {
                throw "Chart series-count verification failed: $($Definition.name)"
            }
            for ($index = 0; $index -lt $seriesDefinitions.Count; $index++) {
                $series = $null
                $format = $null
                $fill = $null
                $fillColor = $null
                $line = $null
                $lineColor = $null
                $dataLabels = $null
                try {
                    $series = $seriesCollection.Item($index + 1)
                    $seriesDefinition = $seriesDefinitions[$index]
					$effectiveSeriesType = if (Has-Property $seriesDefinition 'chartType') {
						[int]$seriesDefinition.chartType
					} else { [int]$Definition.chartType }
					try {
						$formulaParts = ConvertFrom-NativeSeriesFormula ([string]$series.Formula)
					}
					catch {
						throw (
							"Chart series $($index + 1) formula validation failed: " +
							"$($Definition.name). $($_.Exception.Message)"
						)
					}
                    $seriesLabel = "Chart series $($index + 1)"
                    $worksheetName = [string]$Worksheet.Name
                    if (Has-Property $seriesDefinition 'nameRange') {
                        Assert-NativeSeriesFormulaRangeArgument `
                            $formulaParts.NameArgument `
                            ([string]$seriesDefinition.nameRange) `
                            $worksheetName `
                            "$seriesLabel name range"
                    }
                    elseif (Has-Property $seriesDefinition 'name') {
                        $formulaName = Get-NativeSeriesFormulaLiteral `
                            $formulaParts.NameArgument
                        if (
                            $formulaName -cne [string]$seriesDefinition.name -or
                            [string]$series.Name -cne [string]$seriesDefinition.name
                        ) {
                            throw "$seriesLabel name verification failed: $($Definition.name)"
                        }
                    }
                    if (Has-Property $seriesDefinition 'categoryRange') {
                        Assert-NativeSeriesFormulaRangeArgument `
                            $formulaParts.XValuesArgument `
                            ([string]$seriesDefinition.categoryRange) `
                            $worksheetName `
                            "$seriesLabel category range"
                    }
                    elseif (Has-Property $seriesDefinition 'xValuesRange') {
                        Assert-NativeSeriesFormulaRangeArgument `
                            $formulaParts.XValuesArgument `
                            ([string]$seriesDefinition.xValuesRange) `
                            $worksheetName `
                            "$seriesLabel X range"
                    }
                    Assert-NativeSeriesFormulaRangeArgument `
                        $formulaParts.ValuesArgument `
                        ([string]$seriesDefinition.valuesRange) `
                        $worksheetName `
                        "$seriesLabel values range"
                    if (Has-Property $seriesDefinition 'bubbleSizesRange') {
                        if ($null -eq $formulaParts.BubbleSizesArgument) {
                            throw "$seriesLabel bubble range is missing: $($Definition.name)"
                        }
                        Assert-NativeSeriesFormulaRangeArgument `
                            $formulaParts.BubbleSizesArgument `
                            ([string]$seriesDefinition.bubbleSizesRange) `
                            $worksheetName `
                            "$seriesLabel bubble range"
                    }
                    if ($formulaParts.PlotOrder -ne $index + 1) {
                        throw "$seriesLabel plot-order verification failed: $($Definition.name)"
                    }
                    if (
                        (Has-Property $seriesDefinition 'chartType') -and
                        [int]$series.ChartType -ne [int]$seriesDefinition.chartType
                    ) {
						throw (
							"Chart series type verification failed: $($Definition.name); " +
							"series=$($index + 1), actual=$([int]$series.ChartType), " +
							"expected=$([int]$seriesDefinition.chartType)"
						)
                    }
                    if (Has-Property $seriesDefinition 'axisGroup') {
                        $expectedAxisGroup = if (
                            [string]$seriesDefinition.axisGroup -ceq 'secondary'
                        ) { 2 } else { 1 }
                        if ([int]$series.AxisGroup -ne $expectedAxisGroup) {
                            throw "Chart series axis verification failed: $($Definition.name)"
                        }
                    }
                    if (
                        (Has-Property $seriesDefinition 'markerStyle') -and
                        [int]$series.MarkerStyle -ne
                            (Get-NativeMarkerStyle ([string]$seriesDefinition.markerStyle))
                    ) {
                        throw "Chart marker verification failed: $($Definition.name)"
                    }
                    if (
                        (Has-Property $seriesDefinition 'markerSize') -and
                        [int]$series.MarkerSize -ne [int]$seriesDefinition.markerSize
                    ) {
                        throw "Chart marker-size verification failed: $($Definition.name)"
                    }
                    if (
                        (Has-Property $seriesDefinition 'smooth') -and
						($script:NativeSmoothSeriesChartTypes -contains $effectiveSeriesType) -and
                        [bool]$series.Smooth -ne [bool]$seriesDefinition.smooth
                    ) {
                        throw "Chart smooth-line verification failed: $($Definition.name)"
                    }
                    if (
                        (Has-Property $seriesDefinition 'visible') -and
                        [bool]$series.IsFiltered -ne (-not [bool]$seriesDefinition.visible)
                    ) {
                        throw "Chart series visibility verification failed: $($Definition.name)"
                    }
                    if (
                        (Has-Property $seriesDefinition 'color') -or
                        (Has-Property $seriesDefinition 'lineColor') -or
                        (Has-Property $seriesDefinition 'lineWidth') -or
                        (Has-Property $seriesDefinition 'dashStyle')
                    ) {
                        $format = $series.Format
                    }
                    if (Has-Property $seriesDefinition 'color') {
                        $fill = $format.Fill
                        $fillColor = $fill.ForeColor
                        if (
                            [int]$fill.Visible -eq 0 -or
                            [int]$fillColor.RGB -ne
                                (Convert-HexToOleColor ([string]$seriesDefinition.color))
                        ) {
                            throw "Chart series fill verification failed: $($Definition.name)"
                        }
                    }
                    if (
                        (Has-Property $seriesDefinition 'lineColor') -or
                        (Has-Property $seriesDefinition 'lineWidth') -or
                        (Has-Property $seriesDefinition 'dashStyle')
                    ) {
                        $line = $format.Line
                    }
                    if (Has-Property $seriesDefinition 'lineColor') {
                        $lineColor = $line.ForeColor
                        if (
                            [int]$line.Visible -eq 0 -or
                            [int]$lineColor.RGB -ne
                                (Convert-HexToOleColor ([string]$seriesDefinition.lineColor))
                        ) {
                            throw "Chart series line-color verification failed: $($Definition.name)"
                        }
                    }
                    if (
                        (Has-Property $seriesDefinition 'lineWidth') -and
                        [Math]::Abs(
                            [double]$line.Weight - [double]$seriesDefinition.lineWidth
                        ) -gt 0.05
                    ) {
                        throw "Chart series line-width verification failed: $($Definition.name)"
                    }
                    if (
                        (Has-Property $seriesDefinition 'dashStyle') -and
                        [int]$line.DashStyle -ne
                            (Get-NativeDashStyle ([string]$seriesDefinition.dashStyle))
                    ) {
                        throw "Chart series dash-style verification failed: $($Definition.name)"
                    }
					if (Has-Property $seriesDefinition 'dataLabels') {
						$definedLabels = $seriesDefinition.dataLabels
						$labelShowProperties = @(
							'showValue', 'showCategoryName', 'showSeriesName',
							'showPercentage', 'showBubbleSize'
						)
						$hasEnabledShowOption = @(
							$labelShowProperties | Where-Object {
								(Has-Property $definedLabels $_) -and [bool]$definedLabels.$_
							}
						).Count -gt 0
						if (-not $hasEnabledShowOption) {
							if ([bool]$series.HasDataLabels) {
								throw "Chart data-label removal verification failed: $($Definition.name)"
							}
						}
						else {
							if (-not [bool]$series.HasDataLabels) {
								throw "Chart data-label creation verification failed: $($Definition.name)"
							}
							$dataLabels = $series.DataLabels()
							foreach ($labelProperty in $labelShowProperties) {
							if (
								$labelProperty -ceq 'showPercentage' -and
								-not ($script:NativePercentageDataLabelChartTypes -contains $effectiveSeriesType)
							) { continue }
							if (
								$labelProperty -ceq 'showBubbleSize' -and
								-not ($script:NativeBubbleChartTypes -contains $effectiveSeriesType)
							) { continue }
								$expectedLabelProperty = (
								(Has-Property $definedLabels $labelProperty) -and
								[bool]$definedLabels.$labelProperty
							)
							if ([bool]$dataLabels.$labelProperty -ne $expectedLabelProperty) {
                                throw (
                                    "Chart data-label $labelProperty verification failed: " +
                                    [string]$Definition.name
                                )
                            }
                        }
							if (
							(Has-Property $definedLabels 'position') -and
							[int]$dataLabels.Position -ne
								(Get-NativeDataLabelPosition (
									[string]$definedLabels.position
								))
                        ) {
                            throw "Chart data-label position verification failed: $($Definition.name)"
							}
						}
                    }
                }
                finally {
                    Release-ComObject $dataLabels
                    Release-ComObject $lineColor
                    Release-ComObject $line
                    Release-ComObject $fillColor
                    Release-ComObject $fill
                    Release-ComObject $format
                    Release-ComObject $series
                }
            }
        }
        elseif (Has-Property $Definition 'sourceRangeRef') {
            Assert-NativeChartSourceCoverage `
                $seriesCollection `
                ([string]$Definition.sourceRangeRef) `
                ([string]$Worksheet.Name)
        }
        if (
            (Has-Property $Definition 'gapWidth') -or
            (Has-Property $Definition 'overlap')
        ) {
            $chartGroups = $chart.ChartGroups()
            $gapWidthCompatibleGroups = 0
            $overlapCompatibleGroups = 0
            for ($index = 1; $index -le [int]$chartGroups.Count; $index++) {
                $chartGroup = $null
                try {
                    $chartGroup = $chartGroups.Item($index)
                    if (
                        (Has-Property $Definition 'gapWidth') -and
                        (Test-NativeChartGroupSupportsOption `
                            $chartGroup `
                            $script:NativeGapWidthChartTypes `
                            @(2, 3, 5))
                    ) {
                        $actualGapWidth = [int]$chartGroup.GapWidth
                        $gapWidthCompatibleGroups++
                        if ($actualGapWidth -ne [int]$Definition.gapWidth) {
                            throw (
                                "Chart gap-width verification failed for group ${index}: " +
                                "$actualGapWidth/$([int]$Definition.gapWidth)"
                            )
                        }
                    }
                    if (
                        (Has-Property $Definition 'overlap') -and
                        (Test-NativeChartGroupSupportsOption `
                            $chartGroup `
                            $script:NativeOverlapChartTypes `
                            @(2, 3))
                    ) {
                        $actualOverlap = [int]$chartGroup.Overlap
                        $overlapCompatibleGroups++
                        if ($actualOverlap -ne [int]$Definition.overlap) {
                            throw (
                                "Chart overlap verification failed for group ${index}: " +
                                "$actualOverlap/$([int]$Definition.overlap)"
                            )
                        }
                    }
                }
                finally {
                    Release-ComObject $chartGroup
                }
            }
            if (
                (Has-Property $Definition 'gapWidth') -and
                $gapWidthCompatibleGroups -eq 0
            ) {
                throw "Chart gap-width verification failed: $($Definition.name)"
            }
            if (
                (Has-Property $Definition 'overlap') -and
                $overlapCompatibleGroups -eq 0
            ) {
                throw "Chart overlap verification failed: $($Definition.name)"
            }
        }
    }
    finally {
        Release-ComObject $chartGroups
        Release-ComObject $shapeRange
        Release-ComObject $seriesCollection
        Release-ComObject $legend
        Release-ComObject $chartTitle
        Release-ComObject $chart
        Release-ComObject $chartObject
        Release-ComObject $chartObjects
    }
}

function Assert-NativeWorkbookObjectOperations {
    param(
        [object]$Workbook,
        [object[]]$Operations
    )

    $desiredTableNames = New-Object 'System.Collections.Generic.HashSet[string]' `
        ([StringComparer]::OrdinalIgnoreCase)
    $desiredChartNames = New-Object 'System.Collections.Generic.HashSet[string]' `
        ([StringComparer]::OrdinalIgnoreCase)
    foreach ($operation in $Operations) {
        $kind = if (Has-Property $operation 'kind') { [string]$operation.kind } else { 'cell' }
        if ($kind -in @('createTable', 'updateTable')) {
            [void]$desiredTableNames.Add((Get-NativeWorkbookObjectKey `
                ([string]$operation.sheetName) `
                ([string]$operation.table.name)))
        }
        if ($kind -in @('createChart', 'updateChart')) {
            [void]$desiredChartNames.Add((Get-NativeWorkbookObjectKey `
                ([string]$operation.sheetName) `
                ([string]$operation.chart.name)))
        }
    }

    foreach ($operation in $Operations) {
        $kind = if (Has-Property $operation 'kind') { [string]$operation.kind } else { 'cell' }
        if ($kind -notin @(
            'createTable', 'updateTable', 'deleteTable',
            'createChart', 'updateChart', 'deleteChart'
        )) {
            continue
        }
        $worksheet = $null
        try {
            $worksheet = $Workbook.Worksheets.Item([string]$operation.sheetName)
            if ($kind -in @('createTable', 'updateTable')) {
                Assert-NativeTableMatches $worksheet $operation.table
            }
            elseif ($kind -in @('createChart', 'updateChart')) {
                Assert-NativeChartMatches $worksheet $operation
            }
            elseif ($kind -ceq 'deleteTable') {
                $key = Get-NativeWorkbookObjectKey `
                    ([string]$operation.sheetName) `
                    ([string]$operation.name)
                if (-not $desiredTableNames.Contains($key)) {
                    $listObjects = $null
                    $deletedObject = $null
                    try {
                        $listObjects = $worksheet.ListObjects
                        try {
                            $deletedObject = $listObjects.Item([string]$operation.name)
                        }
                        catch {
                            $deletedObject = $null
                        }
                        if ($null -ne $deletedObject) {
                            throw "Deleted table is still present: $($operation.name)"
                        }
                    }
                    finally {
                        Release-ComObject $deletedObject
                        Release-ComObject $listObjects
                    }
                }
            }
            elseif ($kind -ceq 'deleteChart') {
                $key = Get-NativeWorkbookObjectKey `
                    ([string]$operation.sheetName) `
                    ([string]$operation.name)
                if (-not $desiredChartNames.Contains($key)) {
                    $chartObjects = $null
                    $deletedObject = $null
                    try {
                        $chartObjects = $worksheet.ChartObjects()
                        try {
                            $deletedObject = $chartObjects.Item([string]$operation.name)
                        }
                        catch {
                            $deletedObject = $null
                        }
                        if ($null -ne $deletedObject) {
                            throw "Deleted chart is still present: $($operation.name)"
                        }
                    }
                    finally {
                        Release-ComObject $deletedObject
                        Release-ComObject $chartObjects
                    }
                }
            }
        }
        finally {
            Release-ComObject $worksheet
        }
    }
}

function Apply-NativeOperation {
    param(
        [object]$Workbook,
        [object]$ExcelApplication,
        [object]$Operation
    )

    $kind = if (Has-Property $Operation 'kind') {
        [string]$Operation.kind
    } else {
        'cell'
    }
    switch ($kind) {
        'cell' {
            Apply-CellOperation $Workbook $ExcelApplication $Operation
        }
        'columnWidth' {
            Apply-ColumnWidthOperation $Workbook $Operation
        }
        'rowHeight' {
            Apply-RowHeightOperation $Workbook $Operation
        }
        'addConditionalFormatting' {
            Apply-ConditionalFormattingOperation $Workbook $Operation
        }
        'clearConditionalFormatting' {
            Apply-ClearConditionalFormattingOperation $Workbook $Operation
        }
        'createTable' {
			Apply-CreateTableOperation $Workbook $Operation
        }
        'updateTable' {
			Apply-UpdateTableOperation $Workbook $Operation
        }
        'deleteTable' {
            Apply-DeleteTableOperation $Workbook $Operation
        }
        'createChart' {
            Apply-CreateChartOperation $Workbook $Operation
        }
        'updateChart' {
            Apply-UpdateChartOperation $Workbook $Operation
        }
        'deleteChart' {
            Apply-DeleteChartOperation $Workbook $Operation
        }
        default {
            throw "Unsupported native edit operation: $kind"
        }
    }
}

$workbookFullPath = Assert-LocalPath $WorkbookPath
$operationsFullPath = Assert-LocalPath $OperationsPath
$script:NativeWorkbookExtension = [IO.Path]::GetExtension($workbookFullPath).ToLowerInvariant()
if ($script:NativeWorkbookExtension -notin @('.xlsx', '.xlsm')) {
    throw 'Native editing accepts only .xlsx and .xlsm files.'
}
Assert-NoReparsePointChain $workbookFullPath | Out-Null
Assert-NoReparsePointChain $operationsFullPath | Out-Null
if (-not (Test-Path -LiteralPath $workbookFullPath -PathType Leaf)) {
    throw 'Workbook does not exist.'
}
if (-not (Test-Path -LiteralPath $operationsFullPath -PathType Leaf)) {
    throw 'Operations payload does not exist.'
}
$preflightNamedStreamState = @(Get-NamedStreamState $workbookFullPath)
Assert-SafeZoneIdentifierState $workbookFullPath $preflightNamedStreamState
Assert-OoxmlPackageUnsigned $workbookFullPath
Assert-OoxmlPackageHasNoXlmMacroSheets $workbookFullPath
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
$conditionalAddsBySheet = @{}
$workbookObjectOperationsBySheet = @{}
$requestedTableRangesBySheet = @{}
[long]$workbookObjectRangeCells = 0
foreach ($operation in $operations) {
    $kind = if (Has-Property $operation 'kind') {
        [string]$operation.kind
    } else {
        'cell'
    }
    if ($kind -ceq 'addConditionalFormatting') {
        $sheetName = Get-OperationSheetName $operation
        $currentCount = if ($conditionalAddsBySheet.ContainsKey($sheetName)) {
            [int]$conditionalAddsBySheet[$sheetName]
        } else {
            0
        }
        $currentCount++
        if ($currentCount -gt $MaxConditionalFormattingAddsPerSheet) {
            throw (
                'Native edit cannot append more than ' +
                "$MaxConditionalFormattingAddsPerSheet conditional-formatting " +
                'rules per worksheet.'
            )
        }
        $conditionalAddsBySheet[$sheetName] = $currentCount
    }

    if ($kind -in @(
        'createTable', 'updateTable', 'deleteTable',
        'createChart', 'updateChart', 'deleteChart'
    )) {
        $sheetName = Get-OperationSheetName $operation
        $objectCount = if ($workbookObjectOperationsBySheet.ContainsKey($sheetName)) {
            [int]$workbookObjectOperationsBySheet[$sheetName]
        }
        else {
            0
        }
        $objectCount++
        if ($objectCount -gt $MaxWorkbookObjectOperationsPerSheet) {
            throw (
                'Native edit cannot apply more than ' +
                "$MaxWorkbookObjectOperationsPerSheet workbook-object " +
                'operations per worksheet.'
            )
        }
        $workbookObjectOperationsBySheet[$sheetName] = $objectCount
    }

    switch ($kind) {
        'createTable' {
            Assert-AllowedProperties `
                $operation `
                @('kind', 'sheetName', 'table') `
                'Native createTable operation'
            $tableDefinition = Get-RequiredProperty $operation 'table'
            Assert-NativeTableDefinition $tableDefinition
			if (-not [bool]$tableDefinition.headerRow) {
				throw 'Native table creation with headerRow=false is disabled because Excel can move worksheet cells.'
			}
			if ([bool]$tableDefinition.totalsRow) {
				throw 'Native table creation with totalsRow=true is disabled because Excel moves worksheet cells and rewrites formula references.'
			}
            $sheetName = Get-OperationSheetName $operation
            $requestedRanges = if ($requestedTableRangesBySheet.ContainsKey($sheetName)) {
                @($requestedTableRangesBySheet[$sheetName])
            } else { @() }
            foreach ($requestedRange in $requestedRanges) {
                if (Test-NativeWorkbookRangesOverlap $requestedRange.RangeRef $tableDefinition.rangeRef) {
                    throw "Native tables $($requestedRange.Name) and $($tableDefinition.name) overlap on worksheet $sheetName."
                }
            }
            $requestedTableRangesBySheet[$sheetName] = @($requestedRanges) + [pscustomobject]@{
                Name = [string]$tableDefinition.name
                RangeRef = Get-NormalizedRangeRef $tableDefinition.rangeRef
            }
            $workbookObjectRangeCells += Get-NativeTableRangeCellCost `
                $tableDefinition `
                'Native createTable operation.table'
        }
        'updateTable' {
            Assert-AllowedProperties `
                $operation `
                @('kind', 'sheetName', 'name', 'table') `
                'Native updateTable operation'
            Assert-NativeObjectName `
                (Get-RequiredProperty $operation 'name') `
                'Native current table name' `
                $true
            $tableDefinition = Get-RequiredProperty $operation 'table'
            Assert-NativeTableDefinition $tableDefinition
            $sheetName = Get-OperationSheetName $operation
            $requestedRanges = if ($requestedTableRangesBySheet.ContainsKey($sheetName)) {
                @($requestedTableRangesBySheet[$sheetName])
            } else { @() }
            foreach ($requestedRange in $requestedRanges) {
                if (Test-NativeWorkbookRangesOverlap $requestedRange.RangeRef $tableDefinition.rangeRef) {
                    throw "Native tables $($requestedRange.Name) and $($tableDefinition.name) overlap on worksheet $sheetName."
                }
            }
            $requestedTableRangesBySheet[$sheetName] = @($requestedRanges) + [pscustomobject]@{
                Name = [string]$tableDefinition.name
                RangeRef = Get-NormalizedRangeRef $tableDefinition.rangeRef
            }
            $workbookObjectRangeCells += Get-NativeTableRangeCellCost `
                $tableDefinition `
                'Native updateTable operation.table'
        }
        'deleteTable' {
            Assert-AllowedProperties `
                $operation `
                @('kind', 'sheetName', 'name') `
                'Native deleteTable operation'
            Assert-NativeObjectName `
                (Get-RequiredProperty $operation 'name') `
                'Native current table name' `
                $true
        }
        'createChart' {
            Assert-AllowedProperties `
                $operation `
                @('kind', 'sheetName', 'chart') `
                'Native createChart operation'
            $chartDefinition = Get-RequiredProperty $operation 'chart'
            Assert-NativeChartDefinition $chartDefinition
			if (
				(Has-Property $chartDefinition 'legend') -and
				[string]$chartDefinition.legend.position -ceq 'custom'
			) {
				throw 'Native createChart cannot create a custom legend layout; custom only preserves an existing manual Excel layout.'
			}
            $workbookObjectRangeCells += Get-NativeChartRangeCellCost `
                $chartDefinition `
                'Native createChart operation.chart'
        }
        'updateChart' {
            Assert-AllowedProperties `
                $operation `
                @(
                    'kind', 'sheetName', 'name', 'chart',
					'preserveAnchor', 'preserveSeries', 'allowSeriesFormattingChange'
                ) `
                'Native updateChart operation'
            Assert-NativeObjectName `
                (Get-RequiredProperty $operation 'name') `
                'Native current chart name'
			foreach ($propertyName in @('preserveAnchor', 'preserveSeries', 'allowSeriesFormattingChange')) {
                if (Has-Property $operation $propertyName) {
                    Assert-NativeBoolean `
                        $operation.$propertyName `
                        "Native updateChart $propertyName"
                }
            }
			if (
				(Has-Property $operation 'allowSeriesFormattingChange') -and
				[bool]$operation.allowSeriesFormattingChange -and
				(
					-not (Has-Property $operation 'preserveSeries') -or
					-not [bool]$operation.preserveSeries
				)
			) {
				throw 'Native updateChart allowSeriesFormattingChange requires preserveSeries=true.'
			}
            $chartDefinition = Get-RequiredProperty $operation 'chart'
            Assert-NativeChartDefinition $chartDefinition
            $workbookObjectRangeCells += Get-NativeChartRangeCellCost `
                $chartDefinition `
                'Native updateChart operation.chart'
        }
        'deleteChart' {
            Assert-AllowedProperties `
                $operation `
                @('kind', 'sheetName', 'name') `
                'Native deleteChart operation'
            Assert-NativeObjectName `
                (Get-RequiredProperty $operation 'name') `
                'Native current chart name'
        }
    }
    if (
        $workbookObjectRangeCells -gt
        $MaxWorkbookObjectTransactionRangeCells
    ) {
        throw (
            'Native workbook-object ranges exceed the ' +
            "$MaxWorkbookObjectTransactionRangeCells-cell transaction safety budget."
        )
    }
}
$hasWorkbookObjectOperation = @(
    $operations | Where-Object {
        (Has-Property $_ 'kind') -and
        [string]$_.kind -in @(
            'createTable', 'updateTable', 'deleteTable',
            'createChart', 'updateChart', 'deleteChart'
        )
    }
).Count -gt 0

$workbookDirectory = [IO.Path]::GetDirectoryName($workbookFullPath)
$workbookBaseName = [IO.Path]::GetFileNameWithoutExtension($workbookFullPath)
$vbaHelperPath = Assert-LocalPath (
    Join-Path $PSScriptRoot '..\bin\win32-x64\excel-ai-vba-writeback.exe'
)
Assert-NoReparsePointChain $vbaHelperPath | Out-Null
$workPath = Join-Path $workbookDirectory (
    '.' + $workbookBaseName +
    '.excel-ai-native-edit.' + $transactionId + $script:NativeWorkbookExtension
)
Assert-NoReparsePointChain $workPath | Out-Null
if (Test-Path -LiteralPath $workPath) {
    throw "Native edit work file already exists: $workPath"
}

$excel = $null
$workbooks = $null
$workbook = $null
$blankWorkbook = $null
$sourceLock = $null
$backupPath = $null
$sourceNamedStreamState = @()

try {
    try {
        $sourceLock = [IO.File]::Open(
            $workbookFullPath,
            [IO.FileMode]::Open,
            [IO.FileAccess]::Read,
            [IO.FileShare]::Read -bor [IO.FileShare]::Delete
        )
    }
    catch {
        throw (
            'Workbook cannot be locked against concurrent writers: ' +
            $_.Exception.Message
        )
    }

    $initialWorkbookSha256 = Get-StreamSha256Hex $sourceLock
    if ($initialWorkbookSha256 -cne $expectedWorkbookSha256) {
        throw 'Workbook changed after it was loaded in the editor. Reload before saving.'
    }
    $sourceNamedStreamState = @(Get-NamedStreamState $workbookFullPath)
    Assert-SafeZoneIdentifierState $workbookFullPath $sourceNamedStreamState
    if (-not (Test-NamedStreamStateEqual `
        $preflightNamedStreamState `
        $sourceNamedStreamState)) {
        throw 'Workbook alternate data streams changed before the transaction lock.'
    }

    $workStream = [IO.File]::Open(
        $workPath,
        [IO.FileMode]::CreateNew,
        [IO.FileAccess]::Write,
        [IO.FileShare]::None
    )
    try {
        $sourceLock.Position = 0
        $sourceLock.CopyTo($workStream)
        $workStream.Flush($true)
    }
    finally {
        $workStream.Dispose()
    }
    Assert-NoReparsePointChain $workPath | Out-Null
    if ((Get-FileSha256Hex $workPath) -cne $expectedWorkbookSha256) {
        throw 'Native edit work copy does not match the locked source workbook.'
    }
    Copy-NamedStreamsFromSource `
        $workbookFullPath `
        $workPath `
        $sourceNamedStreamState
Assert-OoxmlPackageUnsigned $workPath
Assert-OoxmlPackageHasNoXlmMacroSheets $workPath

    $sourceHasVbaProject = Test-MacroWorkbookPackage $workPath
    $sourcePackageState = Get-PackagePreservationState $workPath
    if ($sourceHasVbaProject) {
        $sourceVbaFingerprint = Get-VbaProjectFingerprint `
            $workPath `
            $vbaHelperPath `
            (Join-Path $workbookDirectory (
                '.' + $workbookBaseName + '.excel-ai-vba-fingerprint.before.' +
                $transactionId + '.json'
            ))
    }

    try {
        $excel = New-Object -ComObject Excel.Application
        $ownedExcelProcess = Get-ExcelProcessIdentity $excel
        [Console]::Out.WriteLine(
            'OWNED_EXCEL_PID|' + $ownedExcelProcess.ProcessId + '|' +
            $ownedExcelProcess.StartTimeUtcTicks
        )
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

        $targetedWorkbookObjects = $null
        $untargetedWorkbookObjectSnapshot = $null
        $preservedChartSnapshot = $null
        if ($hasWorkbookObjectOperation) {
            $targetedWorkbookObjects = Get-NativeTargetedWorkbookObjectSets $operations
            $untargetedWorkbookObjectSnapshot = `
                Get-NativeUntargetedWorkbookObjectSnapshot `
                    $workbook `
                    $targetedWorkbookObjects
            $preservedChartSnapshot = Get-NativePreservedChartSnapshot `
                $workbook `
                $operations `
                $false
        }

        foreach ($operation in $operations) {
            Apply-NativeOperation $workbook $excel $operation
        }
		if ($hasWorkbookObjectOperation) {
			Restore-NativeUntargetedChartAnchors `
				$workbook `
				$targetedWorkbookObjects `
				$untargetedWorkbookObjectSnapshot
		}

        $workbook.Save()
        if ($hasWorkbookObjectOperation) {
            $workbook.Close($false)
            Release-ComObject $workbook
            $workbook = $null
            $workbook = $workbooks.Open($workPath, 0, $true)
            Assert-NativeWorkbookObjectOperations $workbook $operations
            $reopenedPreservedChartSnapshot = Get-NativePreservedChartSnapshot `
                $workbook `
                $operations `
                $true
            if ($reopenedPreservedChartSnapshot -cne $preservedChartSnapshot) {
                throw (
                    'Excel changed a chart anchor or series that the update ' +
                    'declared preserved; the native edit was refused.'
                )
            }
            $reopenedUntargetedSnapshot = `
                Get-NativeUntargetedWorkbookObjectSnapshot `
                    $workbook `
                    $targetedWorkbookObjects
            if ($reopenedUntargetedSnapshot -cne $untargetedWorkbookObjectSnapshot) {
                $snapshotDifferences = @(
                    Compare-NativeWorkbookObjectSnapshots `
                        $untargetedWorkbookObjectSnapshot `
                        $reopenedUntargetedSnapshot
                )
                if ($snapshotDifferences.Count -gt 0) {
                    throw (
                        'Excel changed a non-targeted table or chart; the native ' +
                        'edit was refused: ' + ($snapshotDifferences -join '; ')
                    )
                }
            }
        }
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

    $currentSourceNamedStreamState = @(Get-NamedStreamState $workbookFullPath)
    if (-not (Test-NamedStreamStateEqual `
        $sourceNamedStreamState `
        $currentSourceNamedStreamState)) {
        throw 'Workbook alternate data streams changed during native Excel automation.'
    }
    Copy-NamedStreamsFromSource `
        $workbookFullPath `
        $workPath `
        $sourceNamedStreamState

    if ((Test-MacroWorkbookPackage $workPath) -ne $sourceHasVbaProject) {
        throw 'The VBA project presence changed in the native edit work copy.'
    }
    if ($sourceHasVbaProject) {
        $workVbaFingerprint = Get-VbaProjectFingerprint `
            $workPath `
            $vbaHelperPath `
            (Join-Path $workbookDirectory (
                '.' + $workbookBaseName + '.excel-ai-vba-fingerprint.after.' +
                $transactionId + '.json'
            ))
        if (
            $sourceVbaFingerprint.ProjectName -cne $workVbaFingerprint.ProjectName -or
            $sourceVbaFingerprint.FingerprintSha256 -cne
                $workVbaFingerprint.FingerprintSha256 -or
            $sourceVbaFingerprint.StreamCount -ne $workVbaFingerprint.StreamCount -or
            $sourceVbaFingerprint.StorageCount -ne $workVbaFingerprint.StorageCount -or
            $sourceVbaFingerprint.Protected -ne $workVbaFingerprint.Protected -or
            $sourceVbaFingerprint.Signed -ne $workVbaFingerprint.Signed
        ) {
            throw (
                'Excel changed the logical VBA project streams; save was refused. ' +
                "before=$($sourceVbaFingerprint.FingerprintSha256) " +
                "after=$($workVbaFingerprint.FingerprintSha256)"
            )
        }
    }
    $workPackageState = Get-PackagePreservationState $workPath
    $packageDifferences = @(
        Compare-PackagePreservationState `
            $sourcePackageState `
            $workPackageState `
            $hasWorkbookObjectOperation
    )
    if ($packageDifferences.Count -gt 0) {
        $hasDimensionOperation = @(
            $operations | Where-Object {
                (Has-Property $_ 'kind') -and
                ([string]$_.kind -in @('columnWidth', 'rowHeight'))
            }
        ).Count -gt 0
        $hasVmlDifference = @(
            $packageDifferences | Where-Object {
                [string]$_ -match '(?i)\.vml\s+\('
            }
        ).Count -gt 0
        $dimensionGuidance = if (
            $hasDimensionOperation -and
            $hasVmlDifference
        ) {
            'The dimension change would move a protected worksheet control ' +
            'or drawing; resize outside its anchored rows/columns or use ' +
            'native Excel. '
        } else {
            ''
        }
        throw (
            $dimensionGuidance +
            'Excel changed protected workbook package content; save was refused: ' +
            ($packageDifferences -join ', ')
        )
    }
    if ((Get-StreamSha256Hex $sourceLock) -cne $expectedWorkbookSha256) {
        throw 'Workbook changed while the native edit transaction was running.'
    }
    if (-not (Test-NamedStreamStateEqual `
        $sourceNamedStreamState `
        @(Get-NamedStreamState $workbookFullPath))) {
        throw 'Workbook alternate data streams changed before the atomic commit.'
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
        $timestamp + '.' + $transactionId + $script:NativeWorkbookExtension
    )
    Assert-NoReparsePointChain $backupPath | Out-Null
    if (Test-Path -LiteralPath $backupPath) {
        throw "Backup path already exists: $backupPath"
    }

    $validatedWorkSha256 = Get-FileSha256Hex $workPath
    try {
        [IO.File]::Replace(
            $workPath,
            $workbookFullPath,
            $backupPath,
            $false
        )
    }
    catch {
        $replaceError = $_.Exception.Message
        if (Test-Path -LiteralPath $backupPath -PathType Leaf) {
            if (Test-Path -LiteralPath $workbookFullPath -PathType Leaf) {
                $currentAfterFailureSha256 = Get-FileSha256Hex $workbookFullPath
                if ($currentAfterFailureSha256 -ceq $validatedWorkSha256) {
                    $replaceRecovery = Restore-DisplacedWorkbook `
                        $workbookFullPath `
                        $backupPath `
                        $validatedWorkSha256 `
                        $workbookDirectory `
                        $workbookBaseName `
                        $backupDirectory `
                        $transactionId
                    $recoveryPath = $replaceRecovery.PreservedPath
                }
                else {
                    $recoveryPath = $backupPath
                }
            }
            else {
                $replaceRecovery = Restore-MissingWorkbook `
                    $workbookFullPath `
                    $backupPath `
                    $workbookDirectory `
                    $workbookBaseName `
                    $backupPath `
                    $backupDirectory `
                    $transactionId
                $recoveryPath = $replaceRecovery.PreservedPath
            }
            throw (
                'Atomic replacement failed after displacing the workbook. ' +
                "The displaced version was restored or preserved at " +
                "${recoveryPath}: $replaceError"
            )
        }
        if (-not (Test-Path -LiteralPath $workbookFullPath -PathType Leaf)) {
            $recoveryPath = Join-Path $backupDirectory (
                $workbookBaseName + '.validated-edit-recovery.' +
                [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ') + '.' +
                $transactionId + $script:NativeWorkbookExtension
            )
            Assert-NoReparsePointChain $recoveryPath | Out-Null
            Copy-Item -LiteralPath $workPath -Destination $recoveryPath
            if ((Get-FileSha256Hex $recoveryPath) -cne $validatedWorkSha256) {
                throw (
                    'Atomic replacement failed, the workbook path is missing, ' +
                    'and the validated edited copy could not be preserved.'
                )
            }
            throw (
                'Atomic replacement failed and the workbook path is missing. ' +
                "The validated edited copy is preserved at ${recoveryPath}: " +
                $replaceError
            )
        }
        throw "Atomic replacement failed without displacing the workbook: $replaceError"
    }
    $displacedWorkbookSha256 = Get-FileSha256Hex $backupPath
    if ($displacedWorkbookSha256 -cne $expectedWorkbookSha256) {
        $conflictRecovery = Restore-DisplacedWorkbook `
            $workbookFullPath `
            $backupPath `
            $validatedWorkSha256 `
            $workbookDirectory `
            $workbookBaseName `
            $backupDirectory `
            $transactionId
        $conflictOutcome = if ($conflictRecovery.Restored) {
            'The version displaced by the atomic swap was restored.'
        }
        else {
            'A newer external version was already current and was left untouched.'
        }
        throw (
            'Workbook changed during the atomic commit; the edit was refused. ' +
            $conflictOutcome + " Displaced version retained at: $backupPath"
        )
    }

    try {
        Assert-NoReparsePointChain $backupPath | Out-Null
        Assert-NoReparsePointChain $workbookFullPath | Out-Null
        if ((Get-FileSha256Hex $backupPath) -cne $expectedWorkbookSha256) {
            throw 'Persistent backup does not match the pre-edit workbook.'
        }
        if (-not (Test-NamedStreamStateEqual `
            $sourceNamedStreamState `
            @(Get-NamedStreamState $backupPath))) {
            throw 'Persistent backup alternate data streams do not match the pre-edit workbook.'
        }
        if (-not (Test-NamedStreamStateEqual `
            $sourceNamedStreamState `
            @(Get-NamedStreamState $workbookFullPath))) {
            throw 'Saved workbook alternate data streams were not preserved.'
        }
        if ((Test-MacroWorkbookPackage $backupPath) -ne $sourceHasVbaProject) {
            throw 'The persistent backup has an unexpected VBA project state.'
        }
        if ((Test-MacroWorkbookPackage $workbookFullPath) -ne $sourceHasVbaProject) {
            throw 'The VBA project presence changed in the saved workbook.'
        }
        $savedPackageState = Get-PackagePreservationState $workbookFullPath
        $savedDifferences = @(
            Compare-PackagePreservationState $workPackageState $savedPackageState
        )
        if ($savedDifferences.Count -gt 0) {
            throw (
                'The atomic replacement changed the validated workbook: ' +
                ($savedDifferences -join ', ')
            )
        }
        $workbookSha256 = Get-FileSha256Hex $workbookFullPath
        if ($workbookSha256 -cne $validatedWorkSha256) {
            throw 'Workbook changed immediately after the atomic replacement.'
        }
    }
    catch {
        $validationMessage = $_.Exception.Message
        $rollbackMessage = 'Automatic rollback was not possible.'
        try {
            if ((Get-FileSha256Hex $backupPath) -cne $expectedWorkbookSha256) {
                throw 'The persistent backup is not a verified baseline.'
            }
            if (-not (Test-NamedStreamStateEqual `
                $sourceNamedStreamState `
                @(Get-NamedStreamState $backupPath))) {
                throw (
                    'The persistent backup alternate data streams are not a ' +
                    'verified baseline; the current workbook was left untouched.'
                )
            }
            $rollbackResult = Restore-DisplacedWorkbook `
                $workbookFullPath `
                $backupPath `
                $validatedWorkSha256 `
                $workbookDirectory `
                $workbookBaseName `
                $backupDirectory `
                $transactionId
            $rollbackMessage = if ($rollbackResult.Restored) {
                'The original workbook was restored automatically. ' +
                "Verified backup: $backupPath"
            }
            else {
                'A newer external version was already current and was left untouched. ' +
                "Verified baseline backup: $backupPath"
            }
        }
        catch {
            $rollbackMessage = (
                'Automatic rollback failed: ' + $_.Exception.Message +
                ". Inspect the retained recovery artifacts manually: $backupPath"
            )
        }
        throw "Post-save validation failed: $validationMessage $rollbackMessage"
    }

    @{
        ok = $true
        backupPath = $backupPath
        workbookSha256 = $workbookSha256
    } |
        ConvertTo-Json -Compress |
        ForEach-Object { [Console]::Out.WriteLine($_) }
}
catch {
    $failureLine = [int]$_.InvocationInfo.ScriptLineNumber
    $failureMessage = $_.Exception.Message
    throw "Native Excel edit failed at line ${failureLine}: $failureMessage"
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

    if ($null -ne $sourceLock) {
        $sourceLock.Dispose()
        $sourceLock = $null
    }
    if (Test-Path -LiteralPath $workPath -PathType Leaf) {
        Remove-Item -LiteralPath $workPath -Force -ErrorAction SilentlyContinue
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
