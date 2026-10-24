# Shared policy for extension-owned Excel automation. Never use this against an
# interactive/user-owned Excel instance. Package/XLM/ADS checks remain at callers.
function Assert-ExcelAutomationSafety {
    param([Parameter(Mandatory = $true)][object]$Excel, [switch]$BeforeGuard)

    if ([int]$Excel.AutomationSecurity -ne 3) {
        throw 'Excel automation refused: macros are not force-disabled.'
    }
    $events = $Excel.EnableEvents
    $links = $Excel.AskToUpdateLinks
    if ($events -isnot [bool] -or $events) {
        throw 'Excel automation refused: events are not verifiably disabled.'
    }
    if ($links -isnot [bool] -or $links) {
        throw 'Excel automation refused: link-update prompts are not verifiably disabled.'
    }
    if (-not $BeforeGuard) {
        if ([int]$Excel.Calculation -ne -4135) {
            throw 'Excel automation refused: calculation is not manual.'
        }
        $calculateBeforeSave = $Excel.CalculateBeforeSave
        if ($calculateBeforeSave -isnot [bool] -or $calculateBeforeSave) {
            throw 'Excel automation refused: calculation before save is not verifiably disabled.'
        }
    }
}

function Set-ExcelAutomationSafety {
    param([Parameter(Mandatory = $true)][object]$Excel, [switch]$BeforeGuard)

    $Excel.AutomationSecurity = 3
    $Excel.EnableEvents = $false
    $Excel.AskToUpdateLinks = $false
    if (-not $BeforeGuard) {
        # Excel requires an open workbook before accepting manual calculation.
        $Excel.Calculation = -4135
        $Excel.CalculateBeforeSave = $false
    }
    Assert-ExcelAutomationSafety $Excel -BeforeGuard:$BeforeGuard
}

function Get-ExcelAutomationProcessIds {
    # Snapshot every existing process. Enumeration failure must not grant
    # ownership, and processes must not be left holding native handles.
    $processes = [Diagnostics.Process]::GetProcesses()
    try {
        foreach ($candidate in $processes) { [int]$candidate.Id }
    }
    finally {
        foreach ($candidate in $processes) { $candidate.Dispose() }
    }
}

function Get-ExcelAutomationProcess {
    param([int]$ProcessId)
    return [Diagnostics.Process]::GetProcessById($ProcessId)
}

function Get-ExcelAutomationProcessIdentity {
    param(
        [int]$ProcessId,
        [AllowEmptyCollection()][int[]]$PreexistingProcessIds,
        [long]$NotBeforeUtcTicks
    )

    if ($ProcessId -le 0 -or $PreexistingProcessIds -contains $ProcessId) {
        throw 'Excel COM did not create a new extension-owned process; automation was refused.'
    }
    $candidate = $null
    try {
        $candidate = Get-ExcelAutomationProcess $ProcessId
        $startTicks = $candidate.StartTime.ToUniversalTime().Ticks
        if (
            -not [string]::Equals($candidate.ProcessName, 'EXCEL', [StringComparison]::OrdinalIgnoreCase) -or
            $NotBeforeUtcTicks -le 0 -or $startTicks -lt $NotBeforeUtcTicks
        ) {
            throw 'Excel process ownership could not be verified; automation was refused.'
        }
        return [PSCustomObject]@{
            ProcessId = $ProcessId
            StartTimeUtcTicks = $startTicks.ToString([Globalization.CultureInfo]::InvariantCulture)
        }
    }
    finally {
        if ($null -ne $candidate) { $candidate.Dispose() }
    }
}

function Stop-ExcelAutomationProcess {
    param([AllowNull()][object]$Identity, [int]$GracePeriodMilliseconds = 5000)

    if ($null -eq $Identity -or $GracePeriodMilliseconds -lt 0) { return }
    $candidate = $null
    try {
        if (
            $null -eq $Identity.PSObject.Properties['ProcessId'] -or
            $null -eq $Identity.PSObject.Properties['StartTimeUtcTicks'] -or
            [int]$Identity.ProcessId -le 0 -or
            [string]$Identity.StartTimeUtcTicks -notmatch '^\d{15,19}$'
        ) { return }
        $candidate = Get-ExcelAutomationProcess ([int]$Identity.ProcessId)
        if (
            -not [string]::Equals($candidate.ProcessName, 'EXCEL', [StringComparison]::OrdinalIgnoreCase) -or
            $candidate.StartTime.ToUniversalTime().Ticks -ne [long]$Identity.StartTimeUtcTicks
        ) { return }
        if (-not $candidate.WaitForExit($GracePeriodMilliseconds)) {
            # Kill the verified object, never resolve the PID a second time.
            $candidate.Kill()
            [void]$candidate.WaitForExit(5000)
        }
    }
    catch { } # Already exited/unverifiable processes are never terminated.
    finally {
        if ($null -ne $candidate) { $candidate.Dispose() }
    }
}
