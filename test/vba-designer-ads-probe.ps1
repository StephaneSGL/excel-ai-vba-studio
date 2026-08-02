[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DesignerScriptPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Assert-Probe {
    param(
        [bool]$Condition,
        [string]$Message
    )
    if (-not $Condition) { throw $Message }
}

function Set-ProbeStreamBytes {
    param(
        [string]$Path,
        [string]$StreamName,
        [byte[]]$Bytes
    )
    Set-Content `
        -LiteralPath $Path `
        -Stream $StreamName `
        -Encoding Byte `
        -Value $Bytes `
        -ErrorAction Stop
}

function Get-ProbeStreamBytes {
    param(
        [string]$Path,
        [string]$StreamName
    )
    $memory = New-Object IO.MemoryStream
    try {
        Get-Content `
            -LiteralPath $Path `
            -Stream $StreamName `
            -Encoding Byte `
            -ReadCount 8192 `
            -ErrorAction Stop |
            ForEach-Object {
                $chunk = [byte[]]$_
                $memory.Write($chunk, 0, $chunk.Length)
            }
        return ,$memory.ToArray()
    }
    finally {
        $memory.Dispose()
    }
}

$designerFullPath = [IO.Path]::GetFullPath($DesignerScriptPath)
if (-not (Test-Path -LiteralPath $designerFullPath -PathType Leaf)) {
    throw "Designer script not found: $designerFullPath"
}

$tokens = $null
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile(
    $designerFullPath,
    [ref]$tokens,
    [ref]$parseErrors
)
if ($parseErrors.Count -gt 0) {
    throw "Designer PowerShell parse failed: $($parseErrors[0])"
}

$requiredFunctions = @(
    'Get-BoundedNamedStreamBytes',
    'Get-NamedStreamState',
    'Test-NamedStreamStateEqual',
    'Assert-SafeZoneIdentifierState',
    'Copy-NamedStreamsFromSource'
)
foreach ($functionName in $requiredFunctions) {
    $definitions = @($ast.FindAll({
        param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
            [string]::Equals(
                $node.Name,
                $functionName,
                [StringComparison]::OrdinalIgnoreCase
            )
    }, $true))
    Assert-Probe ($definitions.Count -eq 1) "Expected one $functionName definition"
    Invoke-Expression $definitions[0].Extent.Text
}

$MaxNamedStreams = 64
$MaxNamedStreamBytes = 8MB
$MaxTotalNamedStreamBytes = 32MB
$MaxZoneIdentifierBytes = 64KB
$utf8 = New-Object Text.UTF8Encoding($false)
$probeRoot = [IO.Path]::Combine(
    [IO.Path]::GetTempPath(),
    'excel-vba-designer-ads-' + [Guid]::NewGuid().ToString('N')
)
[void][IO.Directory]::CreateDirectory($probeRoot)

try {
    $workbookPath = [IO.Path]::Combine($probeRoot, 'source.xlsm')
    $candidatePath = [IO.Path]::Combine($probeRoot, 'candidate.xlsm')
    $backupPath = [IO.Path]::Combine($probeRoot, 'backup.xlsm')
    $rollbackPath = [IO.Path]::Combine($probeRoot, 'rollback.xlsm')
    $failedPath = [IO.Path]::Combine($probeRoot, 'failed.xlsm')
    [IO.File]::WriteAllBytes($workbookPath, $utf8.GetBytes('original workbook'))
    [IO.File]::WriteAllBytes($candidatePath, $utf8.GetBytes('validated candidate'))

    $trustedZone = $utf8.GetBytes("[ZoneTransfer]`r`nZoneId=2`r`n")
    $customBytes = [byte[]](0, 1, 2, 3, 127, 128, 254, 255)
    Set-ProbeStreamBytes $workbookPath 'Zone.Identifier' $trustedZone
    Set-ProbeStreamBytes $workbookPath 'Workbook.Metadata' $customBytes
    Set-ProbeStreamBytes $workbookPath 'Empty.Metadata' ([byte[]]@())
    Set-ProbeStreamBytes $candidatePath 'Stale.Metadata' ([byte[]](9, 9, 9))

    $expectedState = @(Get-NamedStreamState $workbookPath)
    Assert-Probe ($expectedState.Count -eq 3) 'ADS inventory did not find all source streams'
    foreach ($entry in $expectedState) {
        Assert-Probe (-not [string]::IsNullOrWhiteSpace([string]$entry.Name)) 'ADS name missing'
        Assert-Probe ([long]$entry.Length -ge 0) 'ADS length missing'
        Assert-Probe ([string]$entry.Sha256 -cmatch '^[0-9a-f]{64}$') 'ADS SHA-256 missing'
    }
    Assert-SafeZoneIdentifierState $workbookPath $expectedState

    Copy-NamedStreamsFromSource $workbookPath $candidatePath $expectedState
    $candidateState = @(Get-NamedStreamState $candidatePath)
    Assert-Probe `
        (Test-NamedStreamStateEqual $expectedState $candidateState) `
        'Candidate ADS do not match the inventoried source state'
    Assert-Probe `
        ((Get-ProbeStreamBytes $candidatePath 'Zone.Identifier').Length -eq $trustedZone.Length) `
        'Zone.Identifier length changed on the candidate'
    Assert-Probe `
        ((Get-ProbeStreamBytes $candidatePath 'Workbook.Metadata').Length -eq $customBytes.Length) `
        'Custom ADS length changed on the candidate'
    Assert-Probe `
        ((Get-ProbeStreamBytes $candidatePath 'Empty.Metadata').Length -eq 0) `
        'Empty ADS was not preserved on the candidate'

    $existingWriter = [IO.File]::Open(
        $workbookPath,
        [IO.FileMode]::Open,
        [IO.FileAccess]::Write,
        [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
    )
    try {
        $guardRejected = $false
        $rejectedGuard = $null
        try {
            $rejectedGuard = [IO.File]::Open(
                $workbookPath,
                [IO.FileMode]::Open,
                [IO.FileAccess]::Read,
                ([IO.FileShare]::Read -bor [IO.FileShare]::Delete)
            )
        }
        catch [IO.IOException] { $guardRejected = $true }
        finally { if ($null -ne $rejectedGuard) { $rejectedGuard.Dispose() } }
        Assert-Probe `
            $guardRejected `
            'Commit guard was acquired while an existing writer was open'
    }
    finally { $existingWriter.Dispose() }

    $commitGuard = [IO.File]::Open(
        $workbookPath,
        [IO.FileMode]::Open,
        [IO.FileAccess]::Read,
        ([IO.FileShare]::Read -bor [IO.FileShare]::Delete)
    )
    try {
        $writerBlocked = $false
        $writer = $null
        try {
            $writer = [IO.File]::Open(
                $workbookPath,
                [IO.FileMode]::Open,
                [IO.FileAccess]::Write,
                [IO.FileShare]::ReadWrite
            )
        }
        catch [IO.IOException] { $writerBlocked = $true }
        finally { if ($null -ne $writer) { $writer.Dispose() } }
        Assert-Probe $writerBlocked 'Commit guard did not refuse a concurrent writer'

        [IO.File]::Replace($candidatePath, $workbookPath, $backupPath)
        Assert-Probe $commitGuard.CanRead 'Commit guard closed during File.Replace'
    }
    finally { $commitGuard.Dispose() }
    Assert-Probe `
        (Test-NamedStreamStateEqual $expectedState @(Get-NamedStreamState $workbookPath)) `
        'ADS were not preserved on the replaced workbook'
    Assert-Probe `
        (Test-NamedStreamStateEqual $expectedState @(Get-NamedStreamState $backupPath)) `
        'ADS were not preserved on the displaced backup'

    $postCommitWriter = $null
    try {
        $postCommitWriter = [IO.File]::Open(
            $workbookPath,
            [IO.FileMode]::Open,
            [IO.FileAccess]::Write,
            [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
        )
        Assert-Probe `
            $postCommitWriter.CanWrite `
            'Commit guard was not released before simulated rollback'
    }
    finally {
        if ($null -ne $postCommitWriter) { $postCommitWriter.Dispose() }
    }

    Copy-Item -LiteralPath $backupPath -Destination $rollbackPath
    Copy-NamedStreamsFromSource $backupPath $rollbackPath $expectedState
    [IO.File]::Replace($rollbackPath, $workbookPath, $failedPath)
    Assert-Probe `
        (Test-NamedStreamStateEqual $expectedState @(Get-NamedStreamState $workbookPath)) `
        'ADS were not preserved by rollback replacement'

    Set-ProbeStreamBytes `
        $workbookPath `
        'Zone.Identifier' `
        $utf8.GetBytes("[ZoneTransfer]`r`nZoneId=3`r`n")
    $zone3Message = ''
    try {
        Assert-SafeZoneIdentifierState `
            $workbookPath `
            @(Get-NamedStreamState $workbookPath)
    }
    catch { $zone3Message = $_.Exception.Message }
    Assert-Probe `
        ($zone3Message -match 'Internet or Restricted Zone \(ZoneId=3\)') `
        'ZoneId=3 was not refused'

    Set-ProbeStreamBytes `
        $workbookPath `
        'Zone.Identifier' `
        $utf8.GetBytes("[ZoneTransfer]`r`nZoneId=4`r`n")
    $zone4Message = ''
    try {
        Assert-SafeZoneIdentifierState `
            $workbookPath `
            @(Get-NamedStreamState $workbookPath)
    }
    catch { $zone4Message = $_.Exception.Message }
    Assert-Probe `
        ($zone4Message -match 'Internet or Restricted Zone \(ZoneId=4\)') `
        'ZoneId=4 was not refused'

    Set-ProbeStreamBytes `
        $workbookPath `
        'Zone.Identifier' `
        $utf8.GetBytes("[ZoneTransfer]`r`nZoneId=2`r`nZoneId=3`r`n")
    $ambiguousMessage = ''
    try {
        Assert-SafeZoneIdentifierState `
            $workbookPath `
            @(Get-NamedStreamState $workbookPath)
    }
    catch { $ambiguousMessage = $_.Exception.Message }
    Assert-Probe `
        ($ambiguousMessage -match 'missing or ambiguous') `
        'Ambiguous Zone.Identifier was not refused'

    Set-ProbeStreamBytes `
        $workbookPath `
        'Zone.Identifier' `
        (New-Object byte[] ($MaxZoneIdentifierBytes + 1))
    $oversizedZoneMessage = ''
    try {
        Assert-SafeZoneIdentifierState `
            $workbookPath `
            @(Get-NamedStreamState $workbookPath)
    }
    catch { $oversizedZoneMessage = $_.Exception.Message }
    Assert-Probe `
        ($oversizedZoneMessage -match '65536-byte safety limit') `
        'Oversized Zone.Identifier was not refused'

    Set-ProbeStreamBytes `
        $workbookPath `
        'Oversized.Metadata' `
        (New-Object byte[] ($MaxNamedStreamBytes + 1))
    $oversizedStreamMessage = ''
    try { [void]@(Get-NamedStreamState $workbookPath) }
    catch { $oversizedStreamMessage = $_.Exception.Message }
    Assert-Probe `
        ($oversizedStreamMessage -match '8388608-byte safety limit') `
        'Oversized alternate data stream was not refused'

    Write-Output 'VBA designer NTFS ADS probe passed.'
}
finally {
    if (Test-Path -LiteralPath $probeRoot) {
        Remove-Item -LiteralPath $probeRoot -Recurse -Force
    }
}
