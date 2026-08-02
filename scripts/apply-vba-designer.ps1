[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$RequestPathBase64,

    [Parameter(Mandatory = $true)]
    [string]$HelperPathBase64
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
Add-Type -AssemblyName System.IO.Compression.FileSystem
. (Join-Path $PSScriptRoot 'ooxml-package-signature.ps1')

$MaxNamedStreams = 64
$MaxNamedStreamBytes = 8MB
$MaxTotalNamedStreamBytes = 32MB
$MaxZoneIdentifierBytes = 64KB

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------
function Decode-Base64Utf8 {
    param([string]$Value)
    return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Value))
}

function Release-ComObject {
    param([AllowNull()][object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) }
        catch { }
    }
}

function Test-ObjectProperty {
    param([object]$Object, [string]$PropertyName)
    if ($null -eq $Object) { return $false }
    try {
        $prop = $Object.PSObject.Properties[$PropertyName]
        return $null -ne $prop -and $prop.IsGettable
    }
    catch { return $false }
}

function Assert-LocalFixedDrive {
    param([string]$Path)
    if ($Path -match '^(\\\\|//|\\\?\\)') {
        throw "UNC/device paths are not allowed: $Path"
    }
    $full = [IO.Path]::GetFullPath($Path)
    $root = [IO.Path]::GetPathRoot($full)
    try {
        $drive = New-Object IO.DriveInfo($root)
        if ($drive.DriveType -ne [IO.DriveType]::Fixed) {
            throw "Drive must be fixed local: $Path"
        }
    } catch { throw "Invalid drive root: $root" }
}

function Assert-NoManagedBackupComponent {
    param([string]$Path)
    $full = [IO.Path]::GetFullPath($Path)
    $parts = $full.Split([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    foreach ($p in $parts) {
        if ($p -ieq '.excel-ai-vba-backups') {
            throw "Path component '.excel-ai-vba-backups' is not allowed: $Path"
        }
    }
}

function Assert-NoReparsePointChain {
    param([string]$Path)
    $full = [IO.Path]::GetFullPath($Path)
    $root = [IO.Path]::GetPathRoot($full)
    $current = $root
    $relative = $full.Substring($root.Length)
    foreach ($part in $relative.Split(@([IO.Path]::DirectorySeparatorChar,[IO.Path]::AltDirectorySeparatorChar), [StringSplitOptions]::RemoveEmptyEntries)) {
        $current = [IO.Path]::Combine($current, $part)
        if (-not (Test-Path -LiteralPath $current)) { break }
        $item = Get-Item -LiteralPath $current -Force -ErrorAction SilentlyContinue
        if ($item -and ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw "ReparsePoint found in path: $current"
        }
    }
}

function Assert-IsValidIdentifier {
    param([string]$Name)
    if ($Name -notmatch '^[A-Za-z_][A-Za-z0-9_]{0,30}$') {
        throw "Invalid identifier: '$Name'. Must match ^[A-Za-z_][A-Za-z0-9_]{0,30}$"
    }
}

function Assert-IsValidMacroName {
    param([string]$Name)
    $parts = $Name.Split('.')
    if ($parts.Length -gt 2) { throw "Invalid macro name (too many dots): $Name" }
    foreach ($p in $parts) { Assert-IsValidIdentifier $p }
}

function ConvertTo-WorksheetButtonOnActionIdentity {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OnAction
    )

    $value = $OnAction.Trim()
    $separator = $value.LastIndexOf('!')
    if ($separator -le 0 -or $separator -ge ($value.Length - 1)) {
        throw "Invalid worksheet button OnAction: '$OnAction'"
    }

    $workbookTarget = $value.Substring(0, $separator).Trim()
    if (
        $workbookTarget.Length -ge 2 -and
        $workbookTarget[0] -eq [char]39 -and
        $workbookTarget[$workbookTarget.Length - 1] -eq [char]39
    ) {
        $workbookTarget = $workbookTarget.Substring(
            1,
            $workbookTarget.Length - 2
        ).Replace("''", "'")
    }

    $bracketedWorkbook = [regex]::Matches(
        $workbookTarget,
        '\[(?<name>[^\[\]]+)\]'
    )
    $workbookName = if ($bracketedWorkbook.Count -gt 0) {
        $bracketedWorkbook[$bracketedWorkbook.Count - 1].Groups['name'].Value
    } else {
        [IO.Path]::GetFileName($workbookTarget)
    }
    $macroName = $value.Substring($separator + 1).Trim()
    if (
        [string]::IsNullOrWhiteSpace($workbookName) -or
        [string]::IsNullOrWhiteSpace($macroName)
    ) {
        throw "Invalid worksheet button OnAction: '$OnAction'"
    }

    return [pscustomobject]@{
        workbookName = $workbookName
        macroName = $macroName
    }
}

function Test-WorksheetButtonOnActionEquivalent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Actual,

        [Parameter(Mandatory = $true)]
        [string]$Expected
    )

    $actualIdentity = ConvertTo-WorksheetButtonOnActionIdentity $Actual
    $expectedIdentity = ConvertTo-WorksheetButtonOnActionIdentity $Expected
    return (
        [StringComparer]::OrdinalIgnoreCase.Equals(
            [string]$actualIdentity.workbookName,
            [string]$expectedIdentity.workbookName
        ) -and
        [StringComparer]::OrdinalIgnoreCase.Equals(
            [string]$actualIdentity.macroName,
            [string]$expectedIdentity.macroName
        )
    )
}

function Assert-MacroProcedureExists {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Components,

        [Parameter(Mandatory = $true)]
        [string]$MacroName
    )

    $parts = $MacroName.Split('.')
    $requestedModule = if ($parts.Length -eq 2) { $parts[0] } else { $null }
    $requestedProcedure = $parts[$parts.Length - 1]
    $procedurePattern = "(?im)^\s*(?:Public\s+)?(?:Static\s+)?Sub\s+" +
        [Regex]::Escape($requestedProcedure) +
        "\s*(?:\(\s*\))?\s*(?:'.*)?$"

    for ($index = 1; $index -le $Components.Count; $index++) {
        $component = $null
        $codeModule = $null
        try {
            $component = $Components.Item($index)
            if ($component.Type -ne 1) {
                continue
            }
            if (
                $requestedModule -and
                -not [StringComparer]::OrdinalIgnoreCase.Equals(
                    [string]$component.Name,
                    $requestedModule
                )
            ) {
                continue
            }
            $codeModule = $component.CodeModule
            $lineCount = [int]$codeModule.CountOfLines
            if ($lineCount -le 0) {
                continue
            }
            $source = [string]$codeModule.Lines(1, $lineCount)
            if ($source -match $procedurePattern) {
                return
            }
        } finally {
            Release-ComObject $codeModule
            Release-ComObject $component
        }
    }

    throw "Public macro procedure '$MacroName' was not found in a standard module"
}

function Assert-NoNulString {
    param([string]$Value)
    if ($Value -and $Value.Contains("`0")) { throw "String contains NUL character" }
}

function Get-Sha256 {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    $stream = [IO.File]::OpenRead($Path)
    try {
        $hash = [Security.Cryptography.SHA256]::Create().ComputeHash($stream)
        return ($hash | ForEach-Object { $_.ToString('x2') }) -join ''
    } finally { $stream.Dispose() }
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
        throw 'VBA designer refused: Zone.Identifier is ambiguous.'
    }
    if ([long]$zoneEntries[0].Length -gt $MaxZoneIdentifierBytes) {
        throw (
            'VBA designer refused: Zone.Identifier exceeds the ' +
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
        throw 'VBA designer refused: Zone.Identifier changed during inspection.'
    }

    $utf8 = New-Object Text.UTF8Encoding($false, $true)
    try {
        $text = $utf8.GetString($byteArray)
    }
    catch {
        throw 'VBA designer refused: Zone.Identifier is not valid UTF-8.'
    }
    if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) {
        $text = $text.Substring(1)
    }
    if (
        [string]::IsNullOrEmpty($text) -or
        $text.Contains([char]0) -or
        [regex]::IsMatch($text, '\r(?!\n)')
    ) {
        throw 'VBA designer refused: Zone.Identifier is malformed.'
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
                throw 'VBA designer refused: Zone.Identifier is malformed.'
            }
            if ($section -ceq 'zonetransfer') { $zoneTransferSections++ }
            continue
        }
        $assignment = [regex]::Match(
            $line,
            '^([A-Za-z][A-Za-z0-9._-]*)\s*=\s*(.*)$'
        )
        if (-not $assignment.Success -or [string]::IsNullOrEmpty($section)) {
            throw 'VBA designer refused: Zone.Identifier is malformed.'
        }
        if (
            $section -ceq 'zonetransfer' -and
            $assignment.Groups[1].Value -ieq 'ZoneId'
        ) {
            $zoneIdText = $assignment.Groups[2].Value.Trim()
            if ($zoneIdText -cnotmatch '^[0-4]$') {
                throw 'VBA designer refused: Zone.Identifier has an invalid ZoneId.'
            }
            [void]$zoneIds.Add([int]$zoneIdText)
        }
    }
    if ($zoneTransferSections -ne 1 -or $zoneIds.Count -ne 1) {
        throw 'VBA designer refused: Zone.Identifier is missing or ambiguous.'
    }
    if ($zoneIds[0] -in @(3, 4)) {
        throw (
            'VBA designer refused: the workbook is marked as Internet or ' +
            "Restricted Zone (ZoneId=$($zoneIds[0])). Trust and unblock it " +
            'explicitly before editing.'
        )
    }
}

function Copy-NamedStreamsFromSource {
    param(
        [string]$SourcePath,
        [string]$TargetPath,
        [object[]]$ExpectedState
    )

    $existingState = @(Get-NamedStreamState $TargetPath)
    # Never erase a newly applied or malformed Mark-of-the-Web while bringing a
    # work file back to the inspected source state. Unsafe state fails closed.
    Assert-SafeZoneIdentifierState $TargetPath $existingState
    foreach ($existing in $existingState) {
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
        throw 'Alternate data streams could not be preserved on the VBA designer work file.'
    }
}

function Get-GuidString { [Guid]::NewGuid().ToString('N') }

function New-InspectRequest {
    param([string]$WorkbookDir, [string]$WorkbookPath)
    $guid = Get-GuidString
    $path = [IO.Path]::Combine($WorkbookDir, "inspect_$guid.json")
    if (Test-Path -LiteralPath $path) { throw "Inspect request path already exists: $path" }
    Assert-NoReparsePointChain $path
    $body = @{ schemaVersion = 1; operation = 'inspect'; workbookPath = $WorkbookPath } | ConvertTo-Json
    [IO.File]::WriteAllText($path, $body, [Text.UTF8Encoding]::new($false))
    return $path
}

function Invoke-Helper {
    param([string]$HelperExe, [string]$RequestPath)
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = $HelperExe
    $psi.Arguments = "`"$RequestPath`""
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $proc = [Diagnostics.Process]::Start($psi)
    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()
    $exitCode = $proc.ExitCode
    $proc.Dispose()
    if ($exitCode -ne 0 -or -not $stdout) {
        throw "Helper failed. ExitCode: $exitCode StdErr: $stderr"
    }
    $lines = $stdout -split "`r?`n"
    $lastNonEmpty = $lines | Where-Object { $_.TrimEnd() -ne '' } | Select-Object -Last 1
    if (-not $lastNonEmpty) { throw "No JSON output from helper" }
    try {
        return $lastNonEmpty | ConvertFrom-Json
    } catch { throw "Invalid JSON from helper: $_" }
}

function Get-ExcelProcessIdentity {
    param([object]$ExcelApp)
    if (-not ('BudgetArtifact.NativeProcess' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace BudgetArtifact {
    public static class NativeProcess {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }
}
'@
    }
    [uint32]$ownedProcessId = 0
    [void][BudgetArtifact.NativeProcess]::GetWindowThreadProcessId([IntPtr][int64]$ExcelApp.Hwnd, [ref]$ownedProcessId)
    if ($ownedProcessId -eq 0) {
        throw 'Excel process ID could not be determined.'
    }

    $ownedProcess = $null
    try {
        $ownedProcess = [Diagnostics.Process]::GetProcessById([int]$ownedProcessId)
        if (
            -not [string]::Equals(
                $ownedProcess.ProcessName,
                'EXCEL',
                [StringComparison]::OrdinalIgnoreCase
            )
        ) {
            throw 'The owned automation process is not EXCEL.EXE.'
        }
        return [PSCustomObject]@{
            ProcessId = [int]$ownedProcessId
            StartTimeUtcTicks = $ownedProcess.StartTime.ToUniversalTime().Ticks.ToString(
                [Globalization.CultureInfo]::InvariantCulture
            )
        }
    }
    finally {
        if ($null -ne $ownedProcess) { $ownedProcess.Dispose() }
    }
}

function Stop-OwnedExcelProcess {
    param(
        [AllowNull()][object]$Identity,
        [int]$GracePeriodMilliseconds = 5000
    )

    if (
        $null -eq $Identity -or
        $GracePeriodMilliseconds -lt 0 -or
        $null -eq $Identity.PSObject.Properties['ProcessId'] -or
        $null -eq $Identity.PSObject.Properties['StartTimeUtcTicks']
    ) {
        return
    }

    $ownedProcessId = [int]$Identity.ProcessId
    [long]$expectedStartTimeUtcTicks = 0
    if (
        $ownedProcessId -le 0 -or
        -not [long]::TryParse(
            [string]$Identity.StartTimeUtcTicks,
            [Globalization.NumberStyles]::Integer,
            [Globalization.CultureInfo]::InvariantCulture,
            [ref]$expectedStartTimeUtcTicks
        )
    ) {
        return
    }

    $ownedProcess = $null
    try {
        $ownedProcess = [Diagnostics.Process]::GetProcessById($ownedProcessId)
        if (
            -not [string]::Equals(
                $ownedProcess.ProcessName,
                'EXCEL',
                [StringComparison]::OrdinalIgnoreCase
            ) -or
            $ownedProcess.StartTime.ToUniversalTime().Ticks -ne
                $expectedStartTimeUtcTicks
        ) {
            return
        }
        if ($ownedProcess.WaitForExit($GracePeriodMilliseconds)) {
            return
        }

        # Kill through the already verified Process object. Resolving the PID a
        # second time here could target an unrelated process after PID reuse.
        $ownedProcess.Kill()
        [void]$ownedProcess.WaitForExit(5000)
    }
    catch {
        # Cleanup is best-effort. Identity mismatch or an already-exited process
        # must fail closed and must never fall back to a PID-only termination.
    }
    finally {
        if ($null -ne $ownedProcess) { $ownedProcess.Dispose() }
    }
}

function Ensure-ExcelSession {
    param()
    $excel = $null
    $ownedProcessIdentity = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.AutomationSecurity = 3
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AskToUpdateLinks = $false
        $excel.ScreenUpdating = $false
        $excel.Visible = $false
        $ownedProcessIdentity = Get-ExcelProcessIdentity $excel
        [Console]::Out.WriteLine(
            'OWNED_EXCEL_PID|' + $ownedProcessIdentity.ProcessId + '|' +
            $ownedProcessIdentity.StartTimeUtcTicks
        )
        return $excel, $ownedProcessIdentity
    }
    catch {
        if ($null -ne $excel) {
            try { $excel.Quit() } catch { }
            Release-ComObject $excel
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Stop-OwnedExcelProcess $ownedProcessIdentity 0
        throw
    }
}

function Cleanup-Excel {
    param([object]$Excel, [AllowNull()][object]$OwnedProcessIdentity)
    if ($null -ne $Excel) {
        try { $Excel.Quit() } catch { }
        Release-ComObject $Excel
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Stop-OwnedExcelProcess $OwnedProcessIdentity
}

function Open-WorkbookReadOnly {
    param([object]$Excel, [string]$Path)
    $workbooks = $null; $wb = $null
    try {
        $workbooks = $Excel.Workbooks
        $wb = $workbooks.Open($Path, 0, $true)  # ReadOnly
        return $wb
    } finally {
        Release-ComObject $workbooks
    }
}

function Release-ComObjectSafe {
    param($Value)
    if ($null -ne $Value) {
        try { Release-ComObject $Value } catch { }
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$requestPath = [IO.Path]::GetFullPath((Decode-Base64Utf8 $RequestPathBase64))
$helperPath  = [IO.Path]::GetFullPath((Decode-Base64Utf8 $HelperPathBase64))

# Validate helper and request exist
if (-not (Test-Path -LiteralPath $helperPath -PathType Leaf)) { throw "Helper not found: $helperPath" }
if ([IO.Path]::GetExtension($helperPath) -ine '.exe') { throw "Helper must be .exe" }
if (-not (Test-Path -LiteralPath $requestPath -PathType Leaf)) { throw "Request file not found: $requestPath" }
$requestSize = (Get-Item -LiteralPath $requestPath).Length
if ($requestSize -gt 1MB) { throw "Request file exceeds 1 MiB" }

# Validate paths
Assert-LocalFixedDrive $requestPath
Assert-LocalFixedDrive $helperPath
Assert-NoManagedBackupComponent $requestPath
Assert-NoManagedBackupComponent $helperPath
Assert-NoReparsePointChain $requestPath
Assert-NoReparsePointChain $helperPath

# Read and parse request
$requestJson = [IO.File]::ReadAllText($requestPath, [Text.UTF8Encoding]::new($false))
$request = $requestJson | ConvertFrom-Json
if (-not $request -or ([int]$request.schemaVersion -notin @(1, 2))) { throw "Invalid request JSON or schemaVersion" }

# Validate required properties
if (-not (Test-ObjectProperty $request 'workbookPath') -or -not $request.workbookPath) { throw "Missing workbookPath in request" }
if (-not (Test-ObjectProperty $request 'expectedWorkbookSha256') -or -not $request.expectedWorkbookSha256) { throw "Missing expectedWorkbookSha256 in request" }
if (-not (Test-ObjectProperty $request 'operations')) { throw "Missing operations in request" }

$workbookPath = [IO.Path]::GetFullPath($request.workbookPath)
$expectedSha256 = $request.expectedWorkbookSha256
$operations = @($request.operations)  # Ensure array even if empty
if ($operations.Count -lt 1 -or $operations.Count -gt 100) {
    throw "Operations must contain between 1 and 100 items"
}

# Validate workbook
if (-not (Test-Path -LiteralPath $workbookPath -PathType Leaf)) { throw "Workbook not found: $workbookPath" }
if ([IO.Path]::GetExtension($workbookPath) -ine '.xlsm') { throw "Workbook must be .xlsm" }
Assert-LocalFixedDrive $workbookPath
Assert-NoManagedBackupComponent $workbookPath
Assert-NoReparsePointChain $workbookPath

# Compute original hash
$originalHash = Get-Sha256 $workbookPath
if (-not $originalHash) { throw "Cannot compute SHA256 of original workbook" }
if ($originalHash -cne $expectedSha256) { throw "Workbook SHA256 does not match expected: $originalHash vs $expectedSha256" }
$sourceNamedStreamState = @(Get-NamedStreamState $workbookPath)
Assert-SafeZoneIdentifierState $workbookPath $sourceNamedStreamState
Assert-OoxmlPackageUnsigned $workbookPath
Assert-OoxmlPackageHasNoXlmMacroSheets $workbookPath

# Pre-validate operations and collect validation sets
$seenFormNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$seenControlKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$seenUpdatedControlKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$seenEventHandlerKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$seenButtonKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$seenButtonAssignmentKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$seenActiveXKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$seenActiveXBindingKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

# Helper to validate a control object (for nested or standalone)
$controlTypeMap = @{
    'label'        = 'Forms.Label.1'
    'textBox'      = 'Forms.TextBox.1'
    'commandButton'= 'Forms.CommandButton.1'
    'comboBox'     = 'Forms.ComboBox.1'
    'listBox'      = 'Forms.ListBox.1'
    'checkBox'     = 'Forms.CheckBox.1'
    'optionButton' = 'Forms.OptionButton.1'
    'toggleButton' = 'Forms.ToggleButton.1'
    'frame'        = 'Forms.Frame.1'
    'image'        = 'Forms.Image.1'
    'spinButton'   = 'Forms.SpinButton.1'
    'scrollBar'    = 'Forms.ScrollBar.1'
}

$allowedCustomActiveXProgIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
if (Test-ObjectProperty $request 'allowedCustomActiveXProgIds') {
    $customProgIds = @($request.allowedCustomActiveXProgIds)
    if ($customProgIds.Count -gt 32) {
        throw "allowedCustomActiveXProgIds cannot contain more than 32 items"
    }
    foreach ($customProgIdValue in $customProgIds) {
        if ($customProgIdValue -isnot [string]) { throw "Custom ActiveX ProgID must be a string" }
        $customProgId = [string]$customProgIdValue
        Assert-NoNulString $customProgId
        if ($customProgId -notmatch '^[A-Za-z][A-Za-z0-9_.-]{1,127}$') {
            throw "Invalid custom ActiveX ProgID: $customProgId"
        }
        if (-not $allowedCustomActiveXProgIds.Add($customProgId)) {
            throw "Duplicate custom ActiveX ProgID: $customProgId"
        }
    }
}

function Resolve-ControlProgId {
    param([Parameter(Mandatory = $true)][object]$Control)

    $typeStr = [string]$Control.type
    if ($typeStr -ieq 'customActiveX') {
        if (-not (Test-ObjectProperty $Control 'progId') -or -not $Control.progId) {
            throw "customActiveX control requires progId"
        }
        $progId = [string]$Control.progId
        Assert-NoNulString $progId
        if ($progId -notmatch '^[A-Za-z][A-Za-z0-9_.-]{1,127}$') {
            throw "Invalid custom ActiveX ProgID: $progId"
        }
        if (-not $allowedCustomActiveXProgIds.Contains($progId)) {
            throw "Custom ActiveX ProgID is not allowlisted: $progId"
        }
        return $progId
    }
    if (Test-ObjectProperty $Control 'progId') {
        throw "progId is accepted only for customActiveX controls"
    }
    if (-not $controlTypeMap.ContainsKey($typeStr)) {
        throw "Unknown control type: $typeStr"
    }
    return [string]$controlTypeMap[$typeStr]
}

function Validate-Control {
    param(
        [object]$Control,
        [System.Collections.Generic.HashSet[string]]$ControlKeySet,
        [string]$FormName,
        [ref]$ErrorRef
    )
    if (-not (Test-ObjectProperty $Control 'name') -or -not $Control.name) { throw "Control missing name" }
    $ctrlName = [string]$Control.name
    Assert-IsValidIdentifier $ctrlName
    $key = "$FormName.$ctrlName"
    if (-not $ControlKeySet.Add($key)) { throw "Duplicate control key in request: $key" }

    if (-not (Test-ObjectProperty $Control 'type') -or -not $Control.type) { throw "Control $ctrlName missing type" }
    $typeStr = [string]$Control.type
    [void](Resolve-ControlProgId $Control)

    $left = if (Test-ObjectProperty $Control 'left') { [double]($Control.left) } else { throw "Control $ctrlName missing left" }
    $top = if (Test-ObjectProperty $Control 'top') { [double]($Control.top) } else { throw "Control $ctrlName missing top" }
    $width = if (Test-ObjectProperty $Control 'width') { [double]($Control.width) } else { throw "Control $ctrlName missing width" }
    $height = if (Test-ObjectProperty $Control 'height') { [double]($Control.height) } else { throw "Control $ctrlName missing height" }

    if ([double]::IsNaN($left) -or [double]::IsInfinity($left) -or $left -lt 0 -or $left -gt 10000) { throw "Invalid left for ${ctrlName}: $left" }
    if ([double]::IsNaN($top) -or [double]::IsInfinity($top) -or $top -lt 0 -or $top -gt 10000) { throw "Invalid top for ${ctrlName}: $top" }
    if ([double]::IsNaN($width) -or [double]::IsInfinity($width) -or $width -le 0 -or $width -gt 10000) { throw "Invalid width for ${ctrlName}: $width (must be >0)" }
    if ([double]::IsNaN($height) -or [double]::IsInfinity($height) -or $height -le 0 -or $height -gt 10000) { throw "Invalid height for ${ctrlName}: $height (must be >0)" }

    $caption = if (Test-ObjectProperty $Control 'caption') { [string]$Control.caption } else { '' }
    Assert-NoNulString $caption
    if ($caption.Length -gt 1000) { throw "caption too long for $ctrlName" }

    $enabled = if (Test-ObjectProperty $Control 'enabled') {
        $val = $Control.enabled
        if ($val -isnot [bool]) { throw "enabled must be boolean for $ctrlName" }
        $val
    } else { $true }

    $visible = if (Test-ObjectProperty $Control 'visible') {
        $val = $Control.visible
        if ($val -isnot [bool]) { throw "visible must be boolean for $ctrlName" }
        $val
    } else { $true }

    $tabIndex = if (Test-ObjectProperty $Control 'tabIndex') {
        $ti = $Control.tabIndex
        if ($ti -isnot [int] -and $ti -isnot [double]) { throw "tabIndex must be numeric for $ctrlName" }
        $tiInt = [int]$ti
        if ($tiInt -lt 0 -or $tiInt -gt 32767) { throw "tabIndex out of range (0-32767) for $ctrlName" }
        $tiInt
    } else { $null }

    $tip = if (Test-ObjectProperty $Control 'controlTipText') { [string]$Control.controlTipText } else { '' }
    Assert-NoNulString $tip
    if ($tip.Length -gt 1000) { throw "controlTipText too long for $ctrlName" }
}

function Validate-ControlChanges {
    param([object]$Changes, [string]$ControlName)

    if ($null -eq $Changes) { throw "changes missing for $ControlName" }
    $allowedProperties = @(
        'left', 'top', 'width', 'height', 'caption',
        'enabled', 'visible', 'tabIndex', 'controlTipText'
    )
    $specified = 0
    foreach ($property in $Changes.PSObject.Properties) {
        if ($allowedProperties -notcontains $property.Name) {
            throw "Unknown control change property '$($property.Name)' for $ControlName"
        }
        $specified++
    }
    if ($specified -lt 1) { throw "changes must contain at least one property for $ControlName" }

    foreach ($propertyName in @('left', 'top', 'width', 'height')) {
        if (-not (Test-ObjectProperty $Changes $propertyName)) { continue }
        $value = [double]$Changes.$propertyName
        if (
            [double]::IsNaN($value) -or
            [double]::IsInfinity($value) -or
            (($propertyName -in @('left', 'top')) -and $value -lt 0) -or
            (($propertyName -in @('width', 'height')) -and $value -le 0) -or
            $value -gt 10000
        ) {
            throw "Invalid $propertyName for ${ControlName}: $value"
        }
    }
    foreach ($propertyName in @('caption', 'controlTipText')) {
        if (-not (Test-ObjectProperty $Changes $propertyName)) { continue }
        $value = [string]$Changes.$propertyName
        Assert-NoNulString $value
        if ($value.Length -gt 1000) { throw "$propertyName too long for $ControlName" }
    }
    foreach ($propertyName in @('enabled', 'visible')) {
        if (
            (Test-ObjectProperty $Changes $propertyName) -and
            $Changes.$propertyName -isnot [bool]
        ) {
            throw "$propertyName must be boolean for $ControlName"
        }
    }
    if (Test-ObjectProperty $Changes 'tabIndex') {
        $tabIndex = $Changes.tabIndex
        if ($tabIndex -isnot [int] -and $tabIndex -isnot [double]) {
            throw "tabIndex must be numeric for $ControlName"
        }
        $tabIndexInt = [int]$tabIndex
        if ($tabIndexInt -lt 0 -or $tabIndexInt -gt 32767) {
            throw "tabIndex out of range (0-32767) for $ControlName"
        }
    }
}

function Assert-ValidUserFormEventProcedure {
    param(
        [string]$ProcedureSource,
        [string]$ObjectName,
        [string]$EventName
    )

    if (-not $ProcedureSource) { throw "procedureSource cannot be empty" }
    if ($ProcedureSource.Length -gt 200000) { throw "procedureSource exceeds 200000 characters" }
    Assert-NoNulString $ProcedureSource
    Assert-IsValidIdentifier $ObjectName
    Assert-IsValidIdentifier $EventName
    $procedureName = "${ObjectName}_${EventName}"
    $escapedProcedureName = [regex]::Escape($procedureName)
    $normalizedSource = Normalize-VbaProcedureSource $ProcedureSource
    $lines = @($normalizedSource -split "`n")
    $headerPattern = "(?i)^[ \t]*Private[ \t]+Sub[ \t]+${escapedProcedureName}[ \t]*\([^\r\n]*\)[ \t]*$"
    $invalidNestedProcedure = $false
    if ($lines.Count -gt 2) {
        for ($lineIndex = 1; $lineIndex -lt ($lines.Count - 1); $lineIndex++) {
            if (
                $lines[$lineIndex] -match '(?i)^[ \t]*(?:(?:Private|Public|Friend|Static)[ \t]+)?(?:Sub|Function|Property[ \t]+(?:Get|Let|Set))\b' -or
                $lines[$lineIndex] -match '(?i)^[ \t]*End[ \t]+Sub\b'
            ) {
                $invalidNestedProcedure = $true
                break
            }
        }
    }
    if (
        $lines.Count -lt 2 -or
        $lines[0] -notmatch $headerPattern -or
        $lines[$lines.Count - 1] -notmatch '(?i)^[ \t]*End[ \t]+Sub[ \t]*$' -or
        $invalidNestedProcedure
    ) {
        throw "procedureSource must contain only Private Sub ${procedureName}(...) ... End Sub"
    }
    foreach ($line in ($ProcedureSource -split "`r?`n")) {
        $trimmed = $line.TrimStart()
        if ($trimmed -match '(?i)^(VERSION(?:\s|$)|BEGIN(?:\s|$)|Attribute\s+VB_)') {
            throw "procedureSource contains forbidden line: $trimmed"
        }
    }
}

function Normalize-VbaProcedureSource {
    param([string]$Source)
    return (($Source -replace "`r`n", "`n") -replace "`r", "`n").Trim()
}

function Normalize-VbaProcedureHeader {
    param([string]$Header)
    return (($Header.Trim() -replace '[ \t]+', ' ').ToLowerInvariant())
}

function Get-VbaProcedureHeader {
    param(
        [object]$CodeModule,
        [string]$ProcedureName
    )

    $startLine = [int]$CodeModule.ProcStartLine($ProcedureName, 0)
    $lineCount = [int]$CodeModule.ProcCountLines($ProcedureName, 0)
    $escapedProcedureName = [regex]::Escape($ProcedureName)
    for ($lineNumber = $startLine; $lineNumber -lt ($startLine + $lineCount); $lineNumber++) {
        $line = [string]$CodeModule.Lines($lineNumber, 1)
        if ($line -match "(?i)^[ \t]*Private[ \t]+Sub[ \t]+${escapedProcedureName}[ \t]*\(") {
            return $line
        }
    }
    throw "Procedure declaration line not found for '$ProcedureName'"
}

# Prevalidate operations
foreach ($op in $operations) {
    if (-not (Test-ObjectProperty $op 'kind') -or -not $op.kind) { throw "Operation missing 'kind'" }
    $opKind = [string]$op.kind
    switch ($opKind) {
        'createUserForm' {
            if (-not (Test-ObjectProperty $op 'name') -or -not $op.name) { throw "createUserForm missing name" }
            $name = [string]$op.name
            Assert-IsValidIdentifier $name
            if (-not $seenFormNames.Add($name)) { throw "Duplicate form name in request: $name" }

            $caption = if (Test-ObjectProperty $op 'caption') { [string]$op.caption } else { '' }
            Assert-NoNulString $caption
            if ($caption.Length -gt 1000) { throw "form caption too long" }

            $width = if (Test-ObjectProperty $op 'width') { [double]($op.width) } else { 400.0 }
            $height = if (Test-ObjectProperty $op 'height') { [double]($op.height) } else { 300.0 }
            if ([double]::IsNaN($width) -or [double]::IsInfinity($width) -or $width -le 0 -or $width -gt 10000) { throw "Invalid width for form ${name}: $width (must be >0)" }
            if ([double]::IsNaN($height) -or [double]::IsInfinity($height) -or $height -le 0 -or $height -gt 10000) { throw "Invalid height for form ${name}: $height (must be >0)" }

            $source = if (Test-ObjectProperty $op 'source') { [string]$op.source } else { '' }
            if ($source) {
                if ($source.Length -gt 2000000) { throw "Source too long for $name" }
                Assert-NoNulString $source
                $lines = $source -split "`r?`n"
                foreach ($line in $lines) {
                    $trimmed = $line.TrimStart()
                    if ($trimmed -match '(?i)^(VERSION(?:\s|$)|BEGIN(?:\s|$)|Attribute\s+VB_)') {
                        throw "Source for $name contains forbidden line: $trimmed"
                    }
                }
            }

            # Validate nested controls if present
            if (Test-ObjectProperty $op 'controls') {
                $nestedControls = @($op.controls)
                foreach ($ctrl in $nestedControls) {
                    Validate-Control $ctrl $seenControlKeys $name
                }
            }
        }
        'addUserFormControl' {
            if (-not (Test-ObjectProperty $op 'formName') -or -not $op.formName) { throw "addUserFormControl missing formName" }
            $formName = [string]$op.formName
            Assert-IsValidIdentifier $formName
            # Note: form may exist in workbook or be previously created; we don't require it in seenFormNames here.
            if (-not (Test-ObjectProperty $op 'control')) { throw "addUserFormControl missing control" }
            $ctrl = $op.control
            Validate-Control $ctrl $seenControlKeys $formName
        }
        'updateUserFormControl' {
            if (-not (Test-ObjectProperty $op 'formName') -or -not $op.formName) { throw "updateUserFormControl missing formName" }
            $formName = [string]$op.formName
            Assert-IsValidIdentifier $formName
            if (-not (Test-ObjectProperty $op 'name') -or -not $op.name) { throw "updateUserFormControl missing name" }
            $name = [string]$op.name
            Assert-IsValidIdentifier $name
            if (-not (Test-ObjectProperty $op 'changes')) { throw "updateUserFormControl missing changes" }
            Validate-ControlChanges $op.changes "$formName.$name"
            $key = "$formName.$name"
            if (-not $seenUpdatedControlKeys.Add($key)) {
                throw "Duplicate control update in request: $key"
            }
        }
        'setUserFormEventHandler' {
            if (-not (Test-ObjectProperty $op 'formName') -or -not $op.formName) { throw "setUserFormEventHandler missing formName" }
            $formName = [string]$op.formName
            Assert-IsValidIdentifier $formName
            if (-not (Test-ObjectProperty $op 'objectName') -or -not $op.objectName) { throw "setUserFormEventHandler missing objectName" }
            $objectName = [string]$op.objectName
            Assert-IsValidIdentifier $objectName
            if (-not (Test-ObjectProperty $op 'eventName') -or -not $op.eventName) { throw "setUserFormEventHandler missing eventName" }
            $eventName = [string]$op.eventName
            Assert-IsValidIdentifier $eventName
            if (-not (Test-ObjectProperty $op 'procedureSource') -or -not $op.procedureSource) { throw "setUserFormEventHandler missing procedureSource" }
            $procedureSource = [string]$op.procedureSource
            Assert-ValidUserFormEventProcedure $procedureSource $objectName $eventName
            if (
                (Test-ObjectProperty $op 'replaceExisting') -and
                $op.replaceExisting -isnot [bool]
            ) {
                throw "replaceExisting must be boolean for $formName.${objectName}_${eventName}"
            }
            $key = "$formName.${objectName}_${eventName}"
            if (-not $seenEventHandlerKeys.Add($key)) {
                throw "Duplicate UserForm event handler in request: $key"
            }
        }
        'createWorksheetButton' {
            if (-not (Test-ObjectProperty $op 'sheetName') -or -not $op.sheetName) { throw "createWorksheetButton missing sheetName" }
            $sheetName = [string]$op.sheetName
            if ($sheetName.Length -gt 1000) { throw "sheetName too long" }
            Assert-NoNulString $sheetName

            if (-not (Test-ObjectProperty $op 'name') -or -not $op.name) { throw "createWorksheetButton missing name" }
            $name = [string]$op.name
            Assert-IsValidIdentifier $name

            if (-not (Test-ObjectProperty $op 'caption') -or -not $op.caption) { throw "createWorksheetButton missing caption" }
            $caption = [string]$op.caption
            Assert-NoNulString $caption
            if ($caption.Length -gt 1000) { throw "button caption too long" }

            if (-not (Test-ObjectProperty $op 'macroName') -or -not $op.macroName) { throw "createWorksheetButton missing macroName" }
            $macroName = [string]$op.macroName
            Assert-IsValidMacroName $macroName

            $left = if (Test-ObjectProperty $op 'left') { [double]($op.left) } else { throw "createWorksheetButton missing left" }
            $top = if (Test-ObjectProperty $op 'top') { [double]($op.top) } else { throw "createWorksheetButton missing top" }
            $width = if (Test-ObjectProperty $op 'width') { [double]($op.width) } else { throw "createWorksheetButton missing width" }
            $height = if (Test-ObjectProperty $op 'height') { [double]($op.height) } else { throw "createWorksheetButton missing height" }
            if ([double]::IsNaN($left) -or [double]::IsInfinity($left) -or $left -lt 0 -or $left -gt 10000) { throw "Invalid left for button ${name}: $left" }
            if ([double]::IsNaN($top) -or [double]::IsInfinity($top) -or $top -lt 0 -or $top -gt 10000) { throw "Invalid top for button ${name}: $top" }
            if ([double]::IsNaN($width) -or [double]::IsInfinity($width) -or $width -le 0 -or $width -gt 10000) { throw "Invalid width for button ${name}: $width (must be >0)" }
            if ([double]::IsNaN($height) -or [double]::IsInfinity($height) -or $height -le 0 -or $height -gt 10000) { throw "Invalid height for button ${name}: $height (must be >0)" }

            $key = "$sheetName.$name"
            if (-not $seenButtonKeys.Add($key)) { throw "Duplicate button key in request: $key" }
        }
        'assignWorksheetButtonMacro' {
            if (-not (Test-ObjectProperty $op 'sheetName') -or -not $op.sheetName) { throw "assignWorksheetButtonMacro missing sheetName" }
            $sheetName = [string]$op.sheetName
            Assert-NoNulString $sheetName
            if ($sheetName.Length -gt 1000) { throw "sheetName too long" }
            if (-not (Test-ObjectProperty $op 'name') -or -not $op.name) { throw "assignWorksheetButtonMacro missing name" }
            $name = [string]$op.name
            Assert-IsValidIdentifier $name
            if (-not (Test-ObjectProperty $op 'macroName') -or -not $op.macroName) { throw "assignWorksheetButtonMacro missing macroName" }
            $macroName = [string]$op.macroName
            Assert-IsValidMacroName $macroName
            $key = "$sheetName.$name"
            if (-not $seenButtonAssignmentKeys.Add($key)) { throw "Duplicate button assignment in request: $key" }
        }
        'createWorksheetActiveXControl' {
            if (-not (Test-ObjectProperty $op 'sheetName') -or -not $op.sheetName) { throw "createWorksheetActiveXControl missing sheetName" }
            $sheetName = [string]$op.sheetName
            Assert-NoNulString $sheetName
            if ($sheetName.Length -gt 1000) { throw "sheetName too long" }
            if (-not (Test-ObjectProperty $op 'control')) { throw "createWorksheetActiveXControl missing control" }
            Validate-Control $op.control $seenActiveXKeys $sheetName
        }
        'bindWorksheetActiveXMacro' {
            if (-not (Test-ObjectProperty $op 'sheetName') -or -not $op.sheetName) { throw "bindWorksheetActiveXMacro missing sheetName" }
            $sheetName = [string]$op.sheetName
            Assert-NoNulString $sheetName
            if ($sheetName.Length -gt 1000) { throw "sheetName too long" }
            if (-not (Test-ObjectProperty $op 'name') -or -not $op.name) { throw "bindWorksheetActiveXMacro missing name" }
            $name = [string]$op.name
            Assert-IsValidIdentifier $name
            if (-not (Test-ObjectProperty $op 'macroName') -or -not $op.macroName) { throw "bindWorksheetActiveXMacro missing macroName" }
            $macroName = [string]$op.macroName
            Assert-IsValidMacroName $macroName
            $key = "$sheetName.$name"
            if (-not $seenActiveXBindingKeys.Add($key)) { throw "Duplicate ActiveX binding in request: $key" }
        }
        default { throw "Unknown operation kind: $opKind" }
    }
}

# Pre-inspect original workbook
$workbookDir = [IO.Path]::GetDirectoryName($workbookPath)
$backupDir = [IO.Path]::Combine($workbookDir, '.excel-ai-vba-backups')

$inspectReqPath = New-InspectRequest $workbookDir $workbookPath
try {
    $preInspect = Invoke-Helper $helperPath $inspectReqPath
} finally {
    if (Test-Path -LiteralPath $inspectReqPath) { Remove-Item -LiteralPath $inspectReqPath -Force }
}
if ($preInspect.ok -ne $true) { throw "Pre-inspect did not return ok true" }
if ($preInspect.workbookSha256 -cne $expectedSha256) { throw "Pre-inspect hash mismatch" }
if ($preInspect.protected -ne $false) { throw "Workbook is VBA protected" }
if ($preInspect.signed -ne $false) { throw "Workbook is signed" }

# Gather existing module names with componentKind from preInspect for reference (not used for duplicate check yet)
$preModules = @($preInspect.modules)
foreach ($m in $preModules) {
    # Just validate structure - not needed immediately
}

# ---------------------------------------------------------------------------
# Transaction
# ---------------------------------------------------------------------------
$stagingGuid = Get-GuidString
$stagingPath = [IO.Path]::Combine($workbookDir, "staging_$stagingGuid.xlsm")
if (Test-Path -LiteralPath $stagingPath) { throw "Staging file already exists: $stagingPath" }
try {
    if ((Get-Sha256 $workbookPath) -cne $originalHash) {
        throw "Original workbook changed before staging copy"
    }
    $currentSourceNamedStreamState = @(Get-NamedStreamState $workbookPath)
    Assert-SafeZoneIdentifierState $workbookPath $currentSourceNamedStreamState
    if (-not (Test-NamedStreamStateEqual `
        $sourceNamedStreamState `
        $currentSourceNamedStreamState)) {
        throw "Workbook alternate data streams changed before staging copy"
    }
    Copy-Item -LiteralPath $workbookPath -Destination $stagingPath
    Assert-NoReparsePointChain $stagingPath
    if ((Get-Sha256 $stagingPath) -cne $originalHash) {
        throw "Staging copy does not match the original workbook"
    }
    Copy-NamedStreamsFromSource `
        $workbookPath `
        $stagingPath `
        $sourceNamedStreamState
    Assert-SafeZoneIdentifierState `
        $stagingPath `
        @(Get-NamedStreamState $stagingPath)
    Assert-OoxmlPackageUnsigned $stagingPath
    Assert-OoxmlPackageHasNoXlmMacroSheets $stagingPath
}
catch {
    if (Test-Path -LiteralPath $stagingPath) {
        Remove-Item -LiteralPath $stagingPath -Force -ErrorAction SilentlyContinue
    }
    throw "Failed to prepare staging workbook: $_"
}

$excel1 = $null; $excel1Identity = $null
$wbStaging = $null
$vbaProject = $null; $components = $null
$excel2 = $null; $excel2Identity = $null
$wbVerify = $null
$operationError = $null
$createdForms = [System.Collections.Generic.List[string]]::new()
$createdControls = [System.Collections.Generic.List[string]]::new()
$updatedControls = [System.Collections.Generic.List[string]]::new()
$updatedEventHandlers = [System.Collections.Generic.List[string]]::new()
$createdButtons = [System.Collections.Generic.List[string]]::new()
$assignedButtons = [System.Collections.Generic.List[string]]::new()
$createdActiveXControls = [System.Collections.Generic.List[string]]::new()
$boundActiveXControls = [System.Collections.Generic.List[string]]::new()
$backupPath = $null
$commitGuard = $null
$commitCompleted = $false
$rollbackCompleted = $false
$rollbackError = $null

# Keep structured verification data because worksheet names may contain dots.
$expectedButtons = [System.Collections.Generic.List[object]]::new()
$expectedUpdatedControls = [System.Collections.Generic.List[object]]::new()
$expectedEventHandlers = [System.Collections.Generic.List[object]]::new()
$expectedActiveXControls = [System.Collections.Generic.List[object]]::new()
$expectedActiveXBindings = [System.Collections.Generic.List[object]]::new()

try {
    # First Excel instance for modifications
    $excel1, $excel1Identity = Ensure-ExcelSession
    $workbooks1 = $null
    try {
        $workbooks1 = $excel1.Workbooks
        $wbStaging = $workbooks1.Open($stagingPath, 0, $false)
        if (-not $wbStaging) { throw "Failed to open staging workbook" }
        $vbaProject = $wbStaging.VBProject
        $components = $vbaProject.VBComponents
    } catch {
        $errMsg = $_.Exception.Message
        if ($errMsg -match 'VBProject|Access is denied|Cannot access') {
            throw "Please manually enable 'Trust access to the VBA project object model' in Excel Trust Center. Extension never changes this setting. Error: $errMsg"
        } else {
            throw "Failed to open staging workbook. Error: $errMsg"
        }
    }

    if ($vbaProject.Protection -ne 0) { throw "VBProject is protected" }

    # Enumerate existing forms and controls from workbook
    $existingFormNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $existControlKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $compCount = $components.Count
    for ($i = 0; $i -lt $compCount; $i++) {
        $comp = $null
        try {
            $comp = $components.Item($i+1)
            if ($comp.Type -eq 3) {
                $fname = $comp.Name
                $existingFormNames.Add($fname) > $null
                # Enumerate controls
                $designer = $null
                $controls = $null
                try {
                    $designer = $comp.Designer
                    $controls = $designer.Controls
                    $ctrlCount = $controls.Count
                    for ($j = 0; $j -lt $ctrlCount; $j++) {
                        $ctrl = $null
                        try {
                            $ctrl = $controls.Item($j)
                            $existControlKeys.Add("$fname.$($ctrl.Name)") > $null
                        } finally { Release-ComObject $ctrl }
                    }
                } finally {
                    Release-ComObject $controls
                    Release-ComObject $designer
                }
            }
        } finally { Release-ComObject $comp }
    }

    # Check for duplicate form names between request and existing (already checked prevalidation among request)
    foreach ($op in $operations) {
        if ($op.kind -eq 'createUserForm') {
            $name = [string]$op.name
            if ($existingFormNames.Contains($name)) { throw "Form name '$name' already exists in workbook" }
        }
    }

    # Enumerate existing worksheet buttons
    $existingButtonKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $existingActiveXKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $sheets = $wbStaging.Worksheets
    $sheetCount = $sheets.Count
    for ($i = 0; $i -lt $sheetCount; $i++) {
        $ws = $null
        try {
            $ws = $sheets.Item($i+1)
            $btns = $null
            try {
                $btns = $ws.Buttons()
                if ($btns) {
                    $btnCount = $btns.Count
                    for ($j = 0; $j -lt $btnCount; $j++) {
                        $btn = $null
                        try {
                            $btn = $btns.Item($j+1)
                            $existingButtonKeys.Add("$($ws.Name).$($btn.Name)") > $null
                        } finally { Release-ComObject $btn }
                    }
                }
            } finally { Release-ComObject $btns }
            $oleObjects = $null
            try {
                $oleObjects = $ws.OLEObjects()
                if ($oleObjects) {
                    $oleCount = [int]$oleObjects.Count
                    for ($j = 1; $j -le $oleCount; $j++) {
                        $oleObject = $null
                        try {
                            $oleObject = $oleObjects.Item($j)
                            $existingActiveXKeys.Add("$($ws.Name).$($oleObject.Name)") > $null
                        } finally { Release-ComObject $oleObject }
                    }
                }
            } finally { Release-ComObject $oleObjects }
        } finally { Release-ComObject $ws }
    }

    # Check duplicate buttons between request and existing
    foreach ($op in $operations) {
        if ($op.kind -eq 'createWorksheetButton') {
            $key = "$($op.sheetName).$($op.name)"
            if ($existingButtonKeys.Contains($key)) { throw "Button '$key' already exists in workbook" }
        }
        if ($op.kind -eq 'createWorksheetActiveXControl') {
            $key = "$($op.sheetName).$($op.control.name)"
            if ($existingActiveXKeys.Contains($key)) { throw "ActiveX control '$key' already exists in workbook" }
        }
    }

    # Apply operations
    foreach ($op in $operations) {
        $opKind = [string]$op.kind
        switch ($opKind) {
            'createUserForm' {
                $name = [string]$op.name
                $caption = if (Test-ObjectProperty $op 'caption') { [string]$op.caption } else { $name }
                $width = if (Test-ObjectProperty $op 'width') { [double]$op.width } else { 400.0 }
                $height = if (Test-ObjectProperty $op 'height') { [double]$op.height } else { 300.0 }
                $source = if (Test-ObjectProperty $op 'source') { [string]$op.source } else { '' }
                if (-not $source) { $source = "Option Explicit" }

                $comp = $null; $designer = $null; $props = $null; $widthProp = $null; $heightProp = $null; $codeMod = $null
                try {
                    $comp = $components.Add(3)
                    $comp.Name = $name
                    $designer = $comp.Designer
                    $designer.Caption = $caption
                    $props = $comp.Properties
                    $widthProp = $props.Item('Width')
                    $heightProp = $props.Item('Height')
                    $widthProp.Value = $width
                    $heightProp.Value = $height
                    if ($source) {
                        $codeMod = $comp.CodeModule
                        $codeMod.AddFromString($source)
                        if ([int]$codeMod.CountOfLines -le 0) {
                            throw "Source code resulted in empty module for $name"
                        }
                    }
                    $createdForms.Add($name)

                    # Process nested controls
                    if (Test-ObjectProperty $op 'controls') {
                        $nestedControls = @($op.controls)
                        foreach ($ctrl in $nestedControls) {
                            $ctrlName = $ctrl.name
                            $progId = Resolve-ControlProgId $ctrl
                            $left = [double]$ctrl.left
                            $top = [double]$ctrl.top
                            $widthCtrl = [double]$ctrl.width
                            $heightCtrl = [double]$ctrl.height
                            $captionCtrl = if (Test-ObjectProperty $ctrl 'caption') { [string]$ctrl.caption } else { '' }
                            $enabledSpecified = Test-ObjectProperty $ctrl 'enabled'
                            $enabledCtrl = if ($enabledSpecified) { [bool]$ctrl.enabled } else { $true }
                            $visibleSpecified = Test-ObjectProperty $ctrl 'visible'
                            $visibleCtrl = if ($visibleSpecified) { [bool]$ctrl.visible } else { $true }
                            $tabIndexCtrl = if (Test-ObjectProperty $ctrl 'tabIndex') { [int]$ctrl.tabIndex } else { $null }
                            $tipCtrl = if (Test-ObjectProperty $ctrl 'controlTipText') { [string]$ctrl.controlTipText } else { '' }

                            $ctrlObj = $null; $ctrlControls = $null
                            try {
                                $ctrlControls = $designer.Controls
                                $ctrlObj = $ctrlControls.Add($progId, $ctrlName, $true)
                                $ctrlObj.Left = $left
                                $ctrlObj.Top = $top
                                $ctrlObj.Width = $widthCtrl
                                $ctrlObj.Height = $heightCtrl
                                if ($captionCtrl) { $ctrlObj.Caption = $captionCtrl }
                                if ($ctrl.type -ine 'customActiveX' -or $enabledSpecified) { $ctrlObj.Enabled = $enabledCtrl }
                                if ($ctrl.type -ine 'customActiveX' -or $visibleSpecified) { $ctrlObj.Visible = $visibleCtrl }
                                if ($null -ne $tabIndexCtrl) { $ctrlObj.TabIndex = $tabIndexCtrl }
                                if ($tipCtrl) { $ctrlObj.ControlTipText = $tipCtrl }
                                $createdControls.Add("$name.$ctrlName")
                            } finally {
                                Release-ComObject $ctrlObj
                                Release-ComObject $ctrlControls
                            }
                        }
                    }
                } finally {
                    Release-ComObject $codeMod
                    Release-ComObject $heightProp
                    Release-ComObject $widthProp
                    Release-ComObject $props
                    Release-ComObject $designer
                    Release-ComObject $comp
                }
            }
            'addUserFormControl' {
                $formName = [string]$op.formName
                $ctrl = $op.control
                $ctrlName = $ctrl.name
                $progId = Resolve-ControlProgId $ctrl
                $left = [double]$ctrl.left
                $top = [double]$ctrl.top
                $widthCtrl = [double]$ctrl.width
                $heightCtrl = [double]$ctrl.height
                $captionCtrl = if (Test-ObjectProperty $ctrl 'caption') { [string]$ctrl.caption } else { '' }
                $enabledSpecified = Test-ObjectProperty $ctrl 'enabled'
                $enabledCtrl = if ($enabledSpecified) { [bool]$ctrl.enabled } else { $true }
                $visibleSpecified = Test-ObjectProperty $ctrl 'visible'
                $visibleCtrl = if ($visibleSpecified) { [bool]$ctrl.visible } else { $true }
                $tabIndexCtrl = if (Test-ObjectProperty $ctrl 'tabIndex') { [int]$ctrl.tabIndex } else { $null }
                $tipCtrl = if (Test-ObjectProperty $ctrl 'controlTipText') { [string]$ctrl.controlTipText } else { '' }

                $targetComp = $null
                try {
                    $targetComp = $components.Item($formName)
                    if ($targetComp.Type -ne 3) { throw "Component '$formName' is not a UserForm" }
                    $designer = $null; $ctrlControls = $null; $ctrlObj = $null
                    try {
                        $designer = $targetComp.Designer
                        $ctrlControls = $designer.Controls
                        $ctrlObj = $ctrlControls.Add($progId, $ctrlName, $true)
                        $ctrlObj.Left = $left
                        $ctrlObj.Top = $top
                        $ctrlObj.Width = $widthCtrl
                        $ctrlObj.Height = $heightCtrl
                        if ($captionCtrl) { $ctrlObj.Caption = $captionCtrl }
                        if ($ctrl.type -ine 'customActiveX' -or $enabledSpecified) { $ctrlObj.Enabled = $enabledCtrl }
                        if ($ctrl.type -ine 'customActiveX' -or $visibleSpecified) { $ctrlObj.Visible = $visibleCtrl }
                        if ($null -ne $tabIndexCtrl) { $ctrlObj.TabIndex = $tabIndexCtrl }
                        if ($tipCtrl) { $ctrlObj.ControlTipText = $tipCtrl }
                        $createdControls.Add("$formName.$ctrlName")
                    } finally {
                        Release-ComObject $ctrlObj
                        Release-ComObject $ctrlControls
                        Release-ComObject $designer
                    }
                } finally {
                    Release-ComObject $targetComp
                }
            }
            'updateUserFormControl' {
                $formName = [string]$op.formName
                $ctrlName = [string]$op.name
                $changes = $op.changes
                $targetComp = $null; $designer = $null; $ctrlControls = $null; $ctrlObj = $null
                try {
                    $targetComp = $components.Item($formName)
                    if ($targetComp.Type -ne 3) { throw "Component '$formName' is not a UserForm" }
                    $designer = $targetComp.Designer
                    $ctrlControls = $designer.Controls
                    $ctrlObj = $ctrlControls.Item($ctrlName)
                    if (-not $ctrlObj) { throw "Control '$ctrlName' not found on form '$formName'" }
                    if (Test-ObjectProperty $changes 'left') { $ctrlObj.Left = [double]$changes.left }
                    if (Test-ObjectProperty $changes 'top') { $ctrlObj.Top = [double]$changes.top }
                    if (Test-ObjectProperty $changes 'width') { $ctrlObj.Width = [double]$changes.width }
                    if (Test-ObjectProperty $changes 'height') { $ctrlObj.Height = [double]$changes.height }
                    if (Test-ObjectProperty $changes 'caption') { $ctrlObj.Caption = [string]$changes.caption }
                    if (Test-ObjectProperty $changes 'enabled') { $ctrlObj.Enabled = [bool]$changes.enabled }
                    if (Test-ObjectProperty $changes 'visible') { $ctrlObj.Visible = [bool]$changes.visible }
                    if (Test-ObjectProperty $changes 'tabIndex') { $ctrlObj.TabIndex = [int]$changes.tabIndex }
                    if (Test-ObjectProperty $changes 'controlTipText') { $ctrlObj.ControlTipText = [string]$changes.controlTipText }
                    $updatedControls.Add("$formName.$ctrlName")
                    $expectedUpdatedControls.Add([pscustomobject]@{
                        formName = $formName
                        name = $ctrlName
                        changes = $changes
                    })
                } finally {
                    Release-ComObject $ctrlObj
                    Release-ComObject $ctrlControls
                    Release-ComObject $designer
                    Release-ComObject $targetComp
                }
            }
            'setUserFormEventHandler' {
                $formName = [string]$op.formName
                $objectName = [string]$op.objectName
                $eventName = [string]$op.eventName
                $procedureName = "${objectName}_${eventName}"
                $procedureSource = [string]$op.procedureSource
                $replaceExisting = if (Test-ObjectProperty $op 'replaceExisting') {
                    [bool]$op.replaceExisting
                } else { $false }
                $targetComp = $null; $designer = $null; $ctrlControls = $null; $ctrlObj = $null; $codeModule = $null
                try {
                    $targetComp = $components.Item($formName)
                    if ($targetComp.Type -ne 3) { throw "Component '$formName' is not a UserForm" }
                    if ($objectName -ine 'UserForm') {
                        $designer = $targetComp.Designer
                        $ctrlControls = $designer.Controls
                        $ctrlObj = $ctrlControls.Item($objectName)
                        if (-not $ctrlObj) {
                            throw "Control '$objectName' not found on form '$formName'"
                        }
                    }
                    $codeModule = $targetComp.CodeModule
                    $startLine = 0
                    try {
                        $startLine = [int]$codeModule.ProcStartLine($procedureName, 0)
                    } catch {
                        $startLine = 0
                    }
                    if ($startLine -gt 0 -and -not $replaceExisting) {
                        throw "UserForm event handler '$formName.$procedureName' already exists; set replaceExisting=true to replace it"
                    }
                    $providedHeader = ((Normalize-VbaProcedureSource $procedureSource) -split "`n")[0]
                    if ($startLine -gt 0) {
                        $existingHeader = Get-VbaProcedureHeader $codeModule $procedureName
                        if (
                            (Normalize-VbaProcedureHeader $existingHeader) -cne
                            (Normalize-VbaProcedureHeader $providedHeader)
                        ) {
                            throw "UserForm event signature mismatch for '$formName.$procedureName'. Existing signature: $existingHeader"
                        }
                        $lineCount = [int]$codeModule.ProcCountLines($procedureName, 0)
                        $codeModule.DeleteLines($startLine, $lineCount)
                        $codeModule.InsertLines($startLine, $procedureSource)
                    } else {
                        try {
                            [void]$codeModule.CreateEventProc($eventName, $objectName)
                        } catch {
                            throw "Excel does not expose event '$eventName' for '$formName.$objectName': $($_.Exception.Message)"
                        }
                        $startLine = [int]$codeModule.ProcStartLine($procedureName, 0)
                        $generatedHeader = Get-VbaProcedureHeader $codeModule $procedureName
                        if (
                            (Normalize-VbaProcedureHeader $generatedHeader) -cne
                            (Normalize-VbaProcedureHeader $providedHeader)
                        ) {
                            throw "UserForm event signature mismatch for '$formName.$procedureName'. Excel expects: $generatedHeader"
                        }
                        $lineCount = [int]$codeModule.ProcCountLines($procedureName, 0)
                        $codeModule.DeleteLines($startLine, $lineCount)
                        $codeModule.InsertLines($startLine, $procedureSource)
                    }
                    $persistedStart = [int]$codeModule.ProcStartLine($procedureName, 0)
                    if ($persistedStart -le 0) {
                        throw "UserForm event handler '$formName.$procedureName' was not created"
                    }
                    $updatedEventHandlers.Add("$formName.$procedureName")
                    $expectedEventHandlers.Add([pscustomobject]@{
                        formName = $formName
                        procedureName = $procedureName
                        procedureSource = $procedureSource
                    })
                } finally {
                    Release-ComObject $codeModule
                    Release-ComObject $ctrlObj
                    Release-ComObject $ctrlControls
                    Release-ComObject $designer
                    Release-ComObject $targetComp
                }
            }
            'createWorksheetButton' {
                $sheetName = [string]$op.sheetName
                $name = [string]$op.name
                $caption = [string]$op.caption
                $macroName = [string]$op.macroName
                $left = [double]$op.left
                $top = [double]$op.top
                $width = [double]$op.width
                $height = [double]$op.height

                Assert-MacroProcedureExists $components $macroName

                $ws = $null; $btns = $null; $btn = $null
                try {
                    $ws = $wbStaging.Worksheets.Item($sheetName)
                    if (-not $ws) { throw "Worksheet '$sheetName' not found" }
                    $btns = $ws.Buttons()
                    $btn = $btns.Add($left, $top, $width, $height)
                    $btn.Name = $name
                    $btn.Caption = $caption
                    # OnAction should qualify original workbook filename (not staging)
                    $originalWorkbookName = [IO.Path]::GetFileNameWithoutExtension($workbookPath) + '.xlsm'
                    $btn.OnAction = "'$originalWorkbookName'!$macroName"
                    $createdButtons.Add("$sheetName.$name")
                    $expectedButtons.Add([pscustomobject]@{
                        sheetName = $sheetName
                        name = $name
                        caption = $caption
                        onAction = [string]$btn.OnAction
                    })
                } finally {
                    Release-ComObject $btn
                    Release-ComObject $btns
                    Release-ComObject $ws
                }
            }
            'assignWorksheetButtonMacro' {
                $sheetName = [string]$op.sheetName
                $name = [string]$op.name
                $macroName = [string]$op.macroName
                Assert-MacroProcedureExists $components $macroName

                $ws = $null; $btns = $null; $btn = $null
                try {
                    $ws = $wbStaging.Worksheets.Item($sheetName)
                    if (-not $ws) { throw "Worksheet '$sheetName' not found" }
                    $btns = $ws.Buttons()
                    $btn = $btns.Item($name)
                    if (-not $btn) { throw "Button '$name' not found on sheet '$sheetName'" }
                    $originalWorkbookName = [IO.Path]::GetFileNameWithoutExtension($workbookPath) + '.xlsm'
                    $btn.OnAction = "'$originalWorkbookName'!$macroName"
                    $assignedButtons.Add("$sheetName.$name")
                    $expectedButtons.Add([pscustomobject]@{
                        sheetName = $sheetName
                        name = $name
                        caption = $null
                        onAction = [string]$btn.OnAction
                    })
                } finally {
                    Release-ComObject $btn
                    Release-ComObject $btns
                    Release-ComObject $ws
                }
            }
            'createWorksheetActiveXControl' {
                $sheetName = [string]$op.sheetName
                $ctrl = $op.control
                $ctrlName = [string]$ctrl.name
                $progId = Resolve-ControlProgId $ctrl
                $left = [double]$ctrl.left
                $top = [double]$ctrl.top
                $widthCtrl = [double]$ctrl.width
                $heightCtrl = [double]$ctrl.height
                $captionCtrl = if (Test-ObjectProperty $ctrl 'caption') { [string]$ctrl.caption } else { '' }
                $enabledSpecified = Test-ObjectProperty $ctrl 'enabled'
                $enabledCtrl = if ($enabledSpecified) { [bool]$ctrl.enabled } else { $true }
                $visibleSpecified = Test-ObjectProperty $ctrl 'visible'
                $visibleCtrl = if ($visibleSpecified) { [bool]$ctrl.visible } else { $true }
                $tipCtrl = if (Test-ObjectProperty $ctrl 'controlTipText') { [string]$ctrl.controlTipText } else { '' }

                $ws = $null; $oleObjects = $null; $oleObject = $null; $embeddedControl = $null
                try {
                    $ws = $wbStaging.Worksheets.Item($sheetName)
                    if (-not $ws) { throw "Worksheet '$sheetName' not found" }
                    $oleObjects = $ws.OLEObjects()
                    $missing = [Type]::Missing
                    try {
                        $oleObject = $oleObjects.Add(
                            $progId,
                            $missing,
                            $false,
                            $false,
                            $missing,
                            $missing,
                            $missing,
                            $left,
                            $top,
                            $widthCtrl,
                            $heightCtrl
                        )
                    } catch {
                        throw "Excel refused ActiveX insertion for '$progId'. Enable ActiveX controls in Excel Trust Center only if your policy permits it. The extension never changes this setting. Error: $($_.Exception.Message)"
                    }
                    if (-not $oleObject) { throw "Excel did not create ActiveX control '$ctrlName'" }
                    $oleObject.Name = $ctrlName
                    $oleObject.Visible = $visibleCtrl
                    $embeddedControl = $oleObject.Object
                    if ($captionCtrl) {
                        try { $embeddedControl.Caption = $captionCtrl }
                        catch { throw "ActiveX control '$ctrlName' does not accept Caption: $($_.Exception.Message)" }
                    }
                    if ($ctrl.type -ine 'customActiveX' -or $enabledSpecified) {
                        try { $embeddedControl.Enabled = $enabledCtrl } catch {
                            throw "ActiveX control '$ctrlName' does not accept Enabled: $($_.Exception.Message)"
                        }
                    }
                    if ($tipCtrl) {
                        try { $embeddedControl.ControlTipText = $tipCtrl }
                        catch { throw "ActiveX control '$ctrlName' does not accept ControlTipText: $($_.Exception.Message)" }
                    }
                    $persistedProgId = [string]$oleObject.progID
                    if ($persistedProgId -ine $progId) {
                        throw "ActiveX control '$ctrlName' persisted unexpected ProgID '$persistedProgId'"
                    }
                    $createdActiveXControls.Add("$sheetName.$ctrlName")
                    $expectedActiveXControls.Add([pscustomobject]@{
                        sheetName = $sheetName
                        name = $ctrlName
                        progId = $progId
                    })
                } finally {
                    Release-ComObject $embeddedControl
                    Release-ComObject $oleObject
                    Release-ComObject $oleObjects
                    Release-ComObject $ws
                }
            }
            'bindWorksheetActiveXMacro' {
                $sheetName = [string]$op.sheetName
                $name = [string]$op.name
                $macroName = [string]$op.macroName
                Assert-MacroProcedureExists $components $macroName

                $ws = $null; $oleObjects = $null; $oleObject = $null
                $sheetComponent = $null; $codeModule = $null
                try {
                    $ws = $wbStaging.Worksheets.Item($sheetName)
                    if (-not $ws) { throw "Worksheet '$sheetName' not found" }
                    $oleObjects = $ws.OLEObjects()
                    $oleObject = $oleObjects.Item($name)
                    if (-not $oleObject) { throw "ActiveX control '$name' not found on sheet '$sheetName'" }
                    $progId = [string]$oleObject.progID
                    if ($progId -ine 'Forms.CommandButton.1' -and $progId -ine 'Forms.ToggleButton.1') {
                        throw "ActiveX binding supports only Forms.CommandButton.1 and Forms.ToggleButton.1, got '$progId'"
                    }

                    $sheetComponent = $components.Item([string]$ws.CodeName)
                    $codeModule = $sheetComponent.CodeModule
                    $existingSource = if ([int]$codeModule.CountOfLines -gt 0) {
                        [string]$codeModule.Lines(1, [int]$codeModule.CountOfLines)
                    } else { '' }
                    $escapedName = [regex]::Escape($name)
                    if ($existingSource -match "(?im)^\s*(?:Private|Public|Friend)?\s*Sub\s+${escapedName}_Click\s*\(") {
                        throw "ActiveX event handler '${name}_Click' already exists; refusing to overwrite it"
                    }
                    $handler = @"

Private Sub ${name}_Click()
    Call $macroName
End Sub
"@
                    $codeModule.AddFromString($handler)
                    $boundActiveXControls.Add("$sheetName.$name")
                    $expectedActiveXBindings.Add([pscustomobject]@{
                        sheetName = $sheetName
                        name = $name
                        macroName = $macroName
                    })
                } finally {
                    Release-ComObject $codeModule
                    Release-ComObject $sheetComponent
                    Release-ComObject $oleObject
                    Release-ComObject $oleObjects
                    Release-ComObject $ws
                }
            }
        }
    }

    # Save and close staging
    $wbStaging.Save()
    $wbStaging.Close($false)
    $wbStaging = $null

    # Release VBProject and components before workbook close
    Release-ComObject $components
    Release-ComObject $vbaProject

    # Cleanup first Excel
    $workbooks1 = $excel1.Workbooks
    Release-ComObject $workbooks1
    Cleanup-Excel $excel1 $excel1Identity
    $excel1 = $null; $excel1Identity = $null

    # Excel may rewrite or discard alternate data streams while saving. Restore
    # the exact inspected source state before any second COM open, then verify it.
    $currentSourceNamedStreamState = @(Get-NamedStreamState $workbookPath)
    Assert-SafeZoneIdentifierState $workbookPath $currentSourceNamedStreamState
    if (-not (Test-NamedStreamStateEqual `
        $sourceNamedStreamState `
        $currentSourceNamedStreamState)) {
        throw "Workbook alternate data streams changed during VBA designer automation"
    }
    Copy-NamedStreamsFromSource `
        $workbookPath `
        $stagingPath `
        $sourceNamedStreamState
    Assert-SafeZoneIdentifierState `
        $stagingPath `
        @(Get-NamedStreamState $stagingPath)

    # Second Excel instance for read-only verification
    $excel2, $excel2Identity = Ensure-ExcelSession
    $wbVerify = Open-WorkbookReadOnly $excel2 $stagingPath

    # Verify created forms
    $verifyComponents = $null
    try {
        $verifyVba = $wbVerify.VBProject
        $verifyComponents = $verifyVba.VBComponents
        $verifyCompCount = $verifyComponents.Count
        $verifyFormNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        for ($i = 0; $i -lt $verifyCompCount; $i++) {
            $comp = $null
            try {
                $comp = $verifyComponents.Item($i+1)
                if ($comp.Type -eq 3) {
                    $verifyFormNames.Add($comp.Name) > $null
                }
            } finally { Release-ComObject $comp }
        }
        foreach ($fn in $createdForms) {
            if (-not $verifyFormNames.Contains($fn)) { throw "Form '$fn' not found during verification" }
        }

        # Verify controls
        foreach ($ctrlKey in $createdControls) {
            $parts = $ctrlKey.Split('.')
            $formName = $parts[0]; $ctrlName = $parts[1]
            $comp = $null; $designer = $null; $controls = $null
            try {
                $comp = $verifyComponents.Item($formName)
                if ($comp.Type -ne 3) { throw "$formName is not a form" }
                $designer = $comp.Designer
                $controls = $designer.Controls
                # Check control exists
                $ctrlFound = $false
                $ctrlCount = $controls.Count
                for ($j = 0; $j -lt $ctrlCount; $j++) {
                    $c = $null
                    try {
                        $c = $controls.Item($j)
                        if ($c.Name -ieq $ctrlName) { $ctrlFound = $true; break }
                    } finally { Release-ComObject $c }
                }
                if (-not $ctrlFound) { throw "Control '$ctrlName' not found on form '$formName'" }
            } finally {
                Release-ComObject $controls
                Release-ComObject $designer
                Release-ComObject $comp
            }
        }

        # Verify updated UserForm control geometry and properties.
        foreach ($expectedControl in $expectedUpdatedControls) {
            $formName = [string]$expectedControl.formName
            $ctrlName = [string]$expectedControl.name
            $changes = $expectedControl.changes
            $comp = $null; $designer = $null; $controls = $null; $ctrl = $null
            try {
                $comp = $verifyComponents.Item($formName)
                if ($comp.Type -ne 3) { throw "$formName is not a form" }
                $designer = $comp.Designer
                $controls = $designer.Controls
                $ctrl = $controls.Item($ctrlName)
                if (-not $ctrl) { throw "Updated control '$ctrlName' not found on form '$formName'" }
                foreach ($propertyName in @('left', 'top', 'width', 'height')) {
                    if (-not (Test-ObjectProperty $changes $propertyName)) { continue }
                    $actual = [double]$ctrl.$propertyName
                    $expected = [double]$changes.$propertyName
                    if ([Math]::Abs($actual - $expected) -gt 0.1) {
                        throw "Updated control '$formName.$ctrlName' $propertyName mismatch: $actual vs $expected"
                    }
                }
                foreach ($propertyName in @('caption', 'controlTipText')) {
                    if (-not (Test-ObjectProperty $changes $propertyName)) { continue }
                    if ([string]$ctrl.$propertyName -cne [string]$changes.$propertyName) {
                        throw "Updated control '$formName.$ctrlName' $propertyName mismatch"
                    }
                }
                foreach ($propertyName in @('enabled', 'visible')) {
                    if (-not (Test-ObjectProperty $changes $propertyName)) { continue }
                    if ([bool]$ctrl.$propertyName -ne [bool]$changes.$propertyName) {
                        throw "Updated control '$formName.$ctrlName' $propertyName mismatch"
                    }
                }
                if (
                    (Test-ObjectProperty $changes 'tabIndex') -and
                    [int]$ctrl.TabIndex -ne [int]$changes.tabIndex
                ) {
                    throw "Updated control '$formName.$ctrlName' tabIndex mismatch"
                }
            } finally {
                Release-ComObject $ctrl
                Release-ComObject $controls
                Release-ComObject $designer
                Release-ComObject $comp
            }
        }

        # Verify exact event procedures without executing VBA.
        foreach ($expectedEvent in $expectedEventHandlers) {
            $formName = [string]$expectedEvent.formName
            $procedureName = [string]$expectedEvent.procedureName
            $comp = $null; $codeModule = $null
            try {
                $comp = $verifyComponents.Item($formName)
                if ($comp.Type -ne 3) { throw "$formName is not a form" }
                $codeModule = $comp.CodeModule
                $startLine = [int]$codeModule.ProcStartLine($procedureName, 0)
                $lineCount = [int]$codeModule.ProcCountLines($procedureName, 0)
                if ($startLine -le 0 -or $lineCount -le 0) {
                    throw "UserForm event handler '$formName.$procedureName' was not persisted"
                }
                $actualSource = [string]$codeModule.Lines($startLine, $lineCount)
                if (
                    (Normalize-VbaProcedureSource $actualSource) -cne
                    (Normalize-VbaProcedureSource ([string]$expectedEvent.procedureSource))
                ) {
                    throw "UserForm event handler '$formName.$procedureName' source mismatch"
                }
            } finally {
                Release-ComObject $codeModule
                Release-ComObject $comp
            }
        }

        # Verify buttons
        foreach ($expectedButton in $expectedButtons) {
            $sheetName = [string]$expectedButton.sheetName
            $btnName = [string]$expectedButton.name
            $btnKey = "$sheetName.$btnName"
            $ws = $null; $btns = $null; $btn = $null
            try {
                $ws = $wbVerify.Worksheets.Item($sheetName)
                $btns = $ws.Buttons()
                $btnFound = $false
                $btnCount = $btns.Count
                for ($j = 0; $j -lt $btnCount; $j++) {
                    $b = $null
                    try {
                        $b = $btns.Item($j+1)
                        if ($b.Name -ieq $btnName) {
                            # Verify Caption and OnAction
                            if ($null -ne $expectedButton.caption -and $b.Caption -cne $expectedButton.caption) { throw "Button '$btnKey' caption mismatch" }
                            if (
                                -not (Test-WorksheetButtonOnActionEquivalent `
                                    ([string]$b.OnAction) `
                                    ([string]$expectedButton.onAction))
                            ) {
                                throw "Button '$btnKey' OnAction mismatch"
                            }
                            $btnFound = $true; break
                        }
                    } finally { Release-ComObject $b }
                }
                if (-not $btnFound) { throw "Button '$btnName' not found on sheet '$sheetName'" }
            } finally {
                Release-ComObject $btn
                Release-ComObject $btns
                Release-ComObject $ws
            }
        }

        # Verify worksheet ActiveX controls and their persisted classes.
        foreach ($expectedControl in $expectedActiveXControls) {
            $sheetName = [string]$expectedControl.sheetName
            $controlName = [string]$expectedControl.name
            $ws = $null; $oleObjects = $null; $oleObject = $null
            try {
                $ws = $wbVerify.Worksheets.Item($sheetName)
                $oleObjects = $ws.OLEObjects()
                $oleObject = $oleObjects.Item($controlName)
                if (-not $oleObject) { throw "ActiveX control '$controlName' not found on sheet '$sheetName'" }
                if ([string]$oleObject.progID -ine [string]$expectedControl.progId) {
                    throw "ActiveX control '$sheetName.$controlName' ProgID mismatch"
                }
            } finally {
                Release-ComObject $oleObject
                Release-ComObject $oleObjects
                Release-ComObject $ws
            }
        }

        # Verify generated Click handlers without executing them.
        foreach ($expectedBinding in $expectedActiveXBindings) {
            $sheetName = [string]$expectedBinding.sheetName
            $controlName = [string]$expectedBinding.name
            $macroName = [string]$expectedBinding.macroName
            $ws = $null; $sheetComponent = $null; $codeModule = $null
            try {
                $ws = $wbVerify.Worksheets.Item($sheetName)
                $sheetComponent = $verifyComponents.Item([string]$ws.CodeName)
                $codeModule = $sheetComponent.CodeModule
                $source = if ([int]$codeModule.CountOfLines -gt 0) {
                    [string]$codeModule.Lines(1, [int]$codeModule.CountOfLines)
                } else { '' }
                $escapedControl = [regex]::Escape($controlName)
                $escapedMacro = [regex]::Escape($macroName)
                if ($source -notmatch "(?ims)^\s*Private\s+Sub\s+${escapedControl}_Click\s*\(\s*\).*?^\s*Call\s+${escapedMacro}\s*$.*?^\s*End\s+Sub\b") {
                    throw "ActiveX binding '$sheetName.$controlName' was not persisted as expected"
                }
            } finally {
                Release-ComObject $codeModule
                Release-ComObject $sheetComponent
                Release-ComObject $ws
            }
        }
    } finally {
        Release-ComObject $verifyComponents
        Release-ComObject $verifyVba
    }

    $wbVerify.Close($false)
    $wbVerify = $null
    # Release workbooks of excel2
    $workbooks2 = $excel2.Workbooks
    Release-ComObject $workbooks2
    # Close second Excel
    Cleanup-Excel $excel2 $excel2Identity
    $excel2 = $null; $excel2Identity = $null

    if (-not (Test-NamedStreamStateEqual `
        $sourceNamedStreamState `
        @(Get-NamedStreamState $stagingPath))) {
        throw "Staging alternate data streams changed during VBA designer verification"
    }

    # Post-verification with native helper on staging
    $inspectReqPath2 = New-InspectRequest $workbookDir $stagingPath
    try {
        $postInspect = Invoke-Helper $helperPath $inspectReqPath2
    } finally {
        if (Test-Path -LiteralPath $inspectReqPath2) { Remove-Item -LiteralPath $inspectReqPath2 -Force }
    }
    if ($postInspect.ok -ne $true) { throw "Post-inspect did not return ok true" }
    if ($postInspect.protected -ne $false) { throw "Staging workbook became protected" }
    if ($postInspect.signed -ne $false) { throw "Staging workbook became signed" }
    if (-not $postInspect.workbookSha256) { throw "Post-inspect missing workbookSha256" }

    # Check that each created form appears in modules with componentKind userform
    $postModules = @($postInspect.modules)
    $postFormNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($m in $postModules) {
        $name = if (Test-ObjectProperty $m 'name') { [string]$m.name } else { $null }
        $kind = if (Test-ObjectProperty $m 'componentKind') { [string]$m.componentKind } else { $null }
        if ($name -and $kind -ieq 'userform') {
            $postFormNames.Add($name) > $null
        }
    }
    foreach ($fn in $createdForms) {
        if (-not $postFormNames.Contains($fn)) { throw "Form '$fn' not found in post-inspect modules" }
    }

    # Check global designerStreamsSha256 has entries for each form name+slash
    $ds = $postInspect.designerStreamsSha256
    if (-not $ds) { throw "No designerStreamsSha256 in post-inspect" }
    $dsProps = $ds.PSObject.Properties
    foreach ($fn in $createdForms) {
        $prefix = $fn + '/'
        $found = $false
        foreach ($prop in $dsProps) {
            if ($prop.Name.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) -and $prop.Value -and [string]$prop.Value -ne '') {
                $found = $true
                break
            }
        }
        if (-not $found) { throw "Designer stream for '$fn' not found or empty in designerStreamsSha256" }
    }

    # Verify staging zip contains nonempty xl/vbaProject.bin
    $zip = $null
    try {
        $zip = [IO.Compression.ZipFile]::OpenRead($stagingPath)
        $entry = $zip.GetEntry('xl/vbaProject.bin')
        if (-not $entry -or $entry.Length -eq 0) { throw "xl/vbaProject.bin missing or empty in staging" }
    } finally { if ($zip) { $zip.Dispose() } }

    # Hold a read/delete-share guard across the final source check and atomic
    # replacement. It allows File.Replace's rename semantics but refuses any
    # existing or newly opened writer during the commit window.
    $commitGuard = [IO.File]::Open(
        $workbookPath,
        [IO.FileMode]::Open,
        [IO.FileAccess]::Read,
        ([IO.FileShare]::Read -bor [IO.FileShare]::Delete)
    )

    # Recompute original hash again before commit while the writer guard is held.
    $originalHash2 = Get-Sha256 $workbookPath
    if ($originalHash2 -cne $originalHash) { throw "Original workbook changed before commit" }
    if (-not (Test-Path -LiteralPath $stagingPath)) { throw "Staging file disappeared before commit" }
    $currentSourceNamedStreamState = @(Get-NamedStreamState $workbookPath)
    Assert-SafeZoneIdentifierState $workbookPath $currentSourceNamedStreamState
    if (-not (Test-NamedStreamStateEqual `
        $sourceNamedStreamState `
        $currentSourceNamedStreamState)) {
        throw "Workbook alternate data streams changed before commit"
    }
    $stagingNamedStreamState = @(Get-NamedStreamState $stagingPath)
    Assert-SafeZoneIdentifierState $stagingPath $stagingNamedStreamState
    if (-not (Test-NamedStreamStateEqual `
        $sourceNamedStreamState `
        $stagingNamedStreamState)) {
        throw "Staging alternate data streams do not match the original before commit"
    }

    # Create backup directory only now
    if (-not (Test-Path -LiteralPath $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir | Out-Null
        # Recheck that it's not a reparse point
        $backupItem = Get-Item -LiteralPath $backupDir -Force
        if ($backupItem.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Backup directory is a reparse point" }
    } else {
        $backupItem = Get-Item -LiteralPath $backupDir -Force
        if (-not ($backupItem.PSIsContainer)) { throw "Backup path exists but is not a directory: $backupDir" }
        if ($backupItem.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Backup directory cannot be a reparse point" }
    }
    $backupGuid = Get-GuidString
    $backupPath = [IO.Path]::Combine($backupDir, "backup_$backupGuid.xlsm")

    # Commit: replace original with staging, using backup
    [IO.File]::Replace($stagingPath, $workbookPath, $backupPath)
    $commitCompleted = $true

    # Verify both sides of the atomic replacement before reporting success.
    if (-not (Test-Path -LiteralPath $backupPath)) {
        throw "Atomic replacement did not create the recovery backup"
    }
    $backupHash = Get-Sha256 $backupPath
    if ($backupHash -cne $originalHash) {
        throw "Atomic replacement backup does not match the original workbook"
    }
    $backupNamedStreamState = @(Get-NamedStreamState $backupPath)
    Assert-SafeZoneIdentifierState $backupPath $backupNamedStreamState
    if (-not (Test-NamedStreamStateEqual `
        $sourceNamedStreamState `
        $backupNamedStreamState)) {
        throw "Atomic replacement backup alternate data streams do not match the original workbook"
    }
    $finalHash = Get-Sha256 $workbookPath
    if ($finalHash -cne $postInspect.workbookSha256) {
        throw "Committed workbook hash does not match the verified staging workbook"
    }
    $finalNamedStreamState = @(Get-NamedStreamState $workbookPath)
    Assert-SafeZoneIdentifierState $workbookPath $finalNamedStreamState
    if (-not (Test-NamedStreamStateEqual `
        $sourceNamedStreamState `
        $finalNamedStreamState)) {
        throw "Committed workbook alternate data streams were not preserved"
    }
    $commitGuard.Dispose()
    $commitGuard = $null

    # Build success output
    $result = [ordered]@{
        ok = $true
        targetWorkbookPath = $workbookPath
        sourceWorkbookPath = $workbookPath
        convertedToXlsm = $false
        changed = $true
        createdUserForms = @($createdForms)
        addedControls = @($createdControls)
        updatedControls = @($updatedControls)
        updatedEventHandlers = @($updatedEventHandlers)
        createdButtons = @($createdButtons)
        assignedButtons = @($assignedButtons)
        createdActiveXControls = @($createdActiveXControls)
        boundActiveXControls = @($boundActiveXControls)
        workbookSha256 = $finalHash
        backupPath = $backupPath
        macrosExecuted = $false
        accessVbomChanged = $false
        designerVerified = $true
    }
    Write-Output ($result | ConvertTo-Json -Compress)

} catch {
    $operationError = $_
} finally {
    # Cleanup COM objects (safely)
    if ($null -ne $wbStaging) {
        try { $wbStaging.Close($false) } catch { }
        Release-ComObject $wbStaging
    }
    if ($null -ne $vbaProject) { Release-ComObject $vbaProject }
    if ($null -ne $components) { Release-ComObject $components }
    if ($null -ne $excel1) {
        try { $excel1.Quit() } catch { }
        Release-ComObject $excel1
    }
    if ($null -ne $wbVerify) {
        try { $wbVerify.Close($false) } catch { }
        Release-ComObject $wbVerify
    }
    if ($null -ne $excel2) {
        try { $excel2.Quit() } catch { }
        Release-ComObject $excel2
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    Stop-OwnedExcelProcess $excel1Identity
    Stop-OwnedExcelProcess $excel2Identity

    if ($null -ne $commitGuard) {
        try { $commitGuard.Dispose() } catch { }
        $commitGuard = $null
    }

    # A failure after File.Replace must restore the verified displaced original.
    # Keep the persistent backup intact by replacing from a same-directory copy.
    if ($null -ne $operationError -and $commitCompleted) {
        $rollbackGuid = Get-GuidString
        $rollbackStagingPath = [IO.Path]::Combine(
            $workbookDir,
            "rollback_$rollbackGuid.xlsm"
        )
        $failedReplacementPath = [IO.Path]::Combine(
            $backupDir,
            "failed_$rollbackGuid.xlsm"
        )
        try {
            if (-not $backupPath -or -not (Test-Path -LiteralPath $backupPath)) {
                throw "Recovery backup is unavailable after commit failure"
            }
            if ((Get-Sha256 $backupPath) -cne $originalHash) {
                throw "Recovery backup hash changed after commit failure"
            }
            $recoveryNamedStreamState = @(Get-NamedStreamState $backupPath)
            Assert-SafeZoneIdentifierState $backupPath $recoveryNamedStreamState
            if (-not (Test-NamedStreamStateEqual `
                $sourceNamedStreamState `
                $recoveryNamedStreamState)) {
                throw "Recovery backup alternate data streams are not a verified baseline"
            }
            Copy-Item -LiteralPath $backupPath -Destination $rollbackStagingPath
            Assert-NoReparsePointChain $rollbackStagingPath
            Copy-NamedStreamsFromSource `
                $backupPath `
                $rollbackStagingPath `
                $sourceNamedStreamState
            [IO.File]::Replace(
                $rollbackStagingPath,
                $workbookPath,
                $failedReplacementPath
            )
            if ((Get-Sha256 $workbookPath) -cne $originalHash) {
                throw "Restored workbook does not match the original hash"
            }
            $restoredNamedStreamState = @(Get-NamedStreamState $workbookPath)
            Assert-SafeZoneIdentifierState $workbookPath $restoredNamedStreamState
            if (-not (Test-NamedStreamStateEqual `
                $sourceNamedStreamState `
                $restoredNamedStreamState)) {
                throw "Restored workbook alternate data streams do not match the original"
            }
            $rollbackCompleted = $true
            if (Test-Path -LiteralPath $failedReplacementPath) {
                Remove-Item -LiteralPath $failedReplacementPath -Force -ErrorAction SilentlyContinue
            }
        } catch {
            $rollbackError = $_
        } finally {
            if (Test-Path -LiteralPath $rollbackStagingPath) {
                Remove-Item -LiteralPath $rollbackStagingPath -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # Cleanup staging and backup on precommit failure
    if ($null -ne $operationError) {
        if (Test-Path -LiteralPath $stagingPath) { Remove-Item -LiteralPath $stagingPath -Force -ErrorAction SilentlyContinue }
        # Also remove any leftover inspect files already handled in try blocks
    }
}

if ($null -ne $operationError) {
    $errMsg = $operationError.Exception.Message
    $stackTrace = $operationError.ScriptStackTrace
    if ($null -ne $rollbackError) {
        $errMsg += "`nROLLBACK_FAILED: $($rollbackError.Exception.Message)`nRecovery backup: $backupPath"
    } elseif ($rollbackCompleted) {
        $errMsg += "`nROLLBACK_OK: the original workbook hash was restored."
    }
    Write-Error "$errMsg`n$stackTrace" -ErrorAction Continue
    exit 1
}
