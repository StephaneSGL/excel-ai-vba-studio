[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$WorkbookPath,

    [switch]$ShowVbe
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

function Release-ComObject {
    param(
        [AllowNull()]
        [object]$ComObject
    )

    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        try {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
        }
        catch {
            # Excel remains user-owned. Releasing a proxy is best-effort only.
        }
    }
}

function Assert-LocalWorkbookPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    # Validate path syntax and drive type before any filesystem lookup so even
    # an unreachable UNC/mapped path is never touched.
    $fullPath = [IO.Path]::GetFullPath($Path)
    if (
        $fullPath.StartsWith('\\', [StringComparison]::Ordinal) -or
        $fullPath.StartsWith('//', [StringComparison]::Ordinal)
    ) {
        throw 'Le classeur doit etre sur un disque local; les chemins UNC et reseau sont refuses.'
    }

    $driveRoot = [IO.Path]::GetPathRoot($fullPath)
    try {
        $driveInfo = New-Object IO.DriveInfo($driveRoot)
        if ($driveInfo.DriveType -eq [IO.DriveType]::Network) {
            throw "Le classeur doit etre sur un disque local; le lecteur reseau '$driveRoot' est refuse."
        }
    }
    catch {
        if ($_.Exception.Message -like '*reseau*') {
            throw
        }
        throw "Impossible de verifier que le lecteur du classeur est local : $($_.Exception.Message)"
    }

    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        throw "Classeur introuvable : $fullPath"
    }
    return [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $fullPath).ProviderPath)
}

function Get-InstalledExcelPath {
    $registryPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe'
    )

    foreach ($registryPath in $registryPaths) {
        try {
            $candidate = (Get-Item -LiteralPath $registryPath -ErrorAction Stop).GetValue('')
            if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
                return [IO.Path]::GetFullPath([string]$candidate)
            }
        }
        catch {
            # Try the next discovery method.
        }
    }

    $programRoots = @(
        [Environment]::GetEnvironmentVariable('ProgramFiles'),
        [Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
    ) | Where-Object { $_ }

    $relativeCandidates = @(
        'Microsoft Office\root\Office16\EXCEL.EXE',
        'Microsoft Office\Office16\EXCEL.EXE',
        'Microsoft Office\Office15\EXCEL.EXE',
        'Microsoft Office\Office14\EXCEL.EXE'
    )

    foreach ($programRoot in $programRoots) {
        foreach ($relativeCandidate in $relativeCandidates) {
            $candidate = Join-Path -Path $programRoot -ChildPath $relativeCandidate
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                return [IO.Path]::GetFullPath($candidate)
            }
        }
    }

    $command = Get-Command -Name 'EXCEL.EXE' -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($command -and $command.Source) {
        return [IO.Path]::GetFullPath($command.Source)
    }

    return $null
}

function Get-RunningExcel {
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    }
    catch {
        return $null
    }
}

function Find-OpenWorkbook {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Excel,

        [Parameter(Mandatory = $true)]
        [string]$ResolvedWorkbookPath
    )

    for ($index = 1; $index -le $Excel.Workbooks.Count; $index++) {
        $candidate = $null
        try {
            $candidate = $Excel.Workbooks.Item($index)
            if (
                $candidate.FullName -and
                [StringComparer]::OrdinalIgnoreCase.Equals(
                    [IO.Path]::GetFullPath([string]$candidate.FullName),
                    $ResolvedWorkbookPath
                )
            ) {
                return $candidate
            }
        }
        catch {
            # A workbook may be closing while the collection is enumerated.
        }

        Release-ComObject -ComObject $candidate
    }

    return $null
}

function Initialize-WindowInterop {
    if ($null -ne ('OfficeAi.NativeWindow' -as [type])) {
        return
    }

    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace OfficeAi
{
    public static class NativeWindow
    {
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(
            IntPtr hWnd,
            out uint processId
        );
    }
}
'@
}

function Send-VbeShortcutToExcelWindow {
    param(
        [Parameter(Mandatory = $true)]
        [IntPtr]$ExcelWindowHandle,

        [Parameter(Mandatory = $true)]
        [uint32]$ExcelProcessId
    )

    $automationShell = $null
    try {
        Initialize-WindowInterop
        if ($ExcelWindowHandle -eq [IntPtr]::Zero) {
            throw "Excel n'a fourni aucun handle de fenetre."
        }
        if ($ExcelProcessId -eq 0) {
            throw 'Le processus proprietaire de la fenetre Excel est introuvable.'
        }

        # SW_RESTORE (9) handles a minimized workbook before AppActivate.
        [void][OfficeAi.NativeWindow]::ShowWindowAsync($ExcelWindowHandle, 9)
        [void][OfficeAi.NativeWindow]::SetForegroundWindow($ExcelWindowHandle)

        $automationShell = New-Object -ComObject WScript.Shell
        if (-not $automationShell.AppActivate([int]$ExcelProcessId)) {
            throw "Windows n'a pas autorise l'activation de la fenetre Excel."
        }

        Start-Sleep -Milliseconds 250
        $foregroundWindowHandle = [OfficeAi.NativeWindow]::GetForegroundWindow()
        if ($foregroundWindowHandle -ne $ExcelWindowHandle) {
            throw "La fenetre du classeur Excel cible n'est pas au premier plan; aucune touche n'a ete envoyee."
        }
        [uint32]$foregroundProcessId = 0
        [void][OfficeAi.NativeWindow]::GetWindowThreadProcessId(
            $foregroundWindowHandle,
            [ref]$foregroundProcessId
        )
        if ($foregroundProcessId -ne $ExcelProcessId) {
            throw "La fenetre Excel ciblee n'est pas au premier plan; aucune touche n'a ete envoyee."
        }

        [void]$automationShell.SendKeys('%{F11}')
        Start-Sleep -Milliseconds 300
        return $true
    }
    catch {
        [Console]::Out.WriteLine("VBE_UI_FALLBACK_FAILED|$($_.Exception.Message)")
        return $false
    }
    finally {
        Release-ComObject -ComObject $automationShell
    }
}

function Show-VbeWithKeyboardFallback {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Excel,

        [Parameter(Mandatory = $true)]
        [object]$Workbook
    )

    try {
        # Do not touch the VBA project or its security settings. Activate the
        # exact workbook first, then delegate to the verified window sender.
        [void]$Workbook.Activate()
        $Excel.Visible = $true
        Initialize-WindowInterop

        $excelWindowHandle = [IntPtr]([int64]$Excel.Hwnd)
        [uint32]$excelProcessId = 0
        [void][OfficeAi.NativeWindow]::GetWindowThreadProcessId(
            $excelWindowHandle,
            [ref]$excelProcessId
        )
        return Send-VbeShortcutToExcelWindow `
            -ExcelWindowHandle $excelWindowHandle `
            -ExcelProcessId $excelProcessId
    }
    catch {
        [Console]::Out.WriteLine("VBE_UI_FALLBACK_FAILED|$($_.Exception.Message)")
        return $false
    }
}

function Show-VbeForWorkbookPathWithKeyboardFallback {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ResolvedWorkbookPath
    )

    try {
        $workbookFileName = [IO.Path]::GetFileName($ResolvedWorkbookPath)
        $matchingProcesses = @(
            Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.Responding -and
                    $_.MainWindowHandle -ne 0 -and
                    $_.MainWindowTitle.IndexOf(
                        $workbookFileName,
                        [StringComparison]::OrdinalIgnoreCase
                    ) -ge 0
                }
        )

        if ($matchingProcesses.Count -ne 1) {
            throw (
                'Impossible d''identifier sans ambiguite la fenetre Excel du classeur "{0}" (candidats : {1}).' -f
                $workbookFileName,
                $matchingProcesses.Count
            )
        }

        $targetProcess = $matchingProcesses[0]
        if (
            $targetProcess.MainWindowTitle.IndexOf(
                'Visual Basic',
                [StringComparison]::OrdinalIgnoreCase
            ) -ge 0
        ) {
            return 'ALREADY_VISIBLE'
        }

        if (
            Send-VbeShortcutToExcelWindow `
                -ExcelWindowHandle ([IntPtr]$targetProcess.MainWindowHandle) `
                -ExcelProcessId ([uint32]$targetProcess.Id)
        ) {
            return 'ALT_F11_SENT'
        }
        return $null
    }
    catch {
        [Console]::Out.WriteLine("VBE_UI_FALLBACK_FAILED|$($_.Exception.Message)")
        return $null
    }
}

$excel = $null
$workbook = $null
$vbe = $null
$failureCode = 0
$failureMessage = $null
$resolvedWorkbookPath = $null

try {
    $resolvedWorkbookPath = Assert-LocalWorkbookPath -Path $WorkbookPath
    $excelPath = Get-InstalledExcelPath
    if (-not $excelPath) {
        $failureCode = 2
        throw "Microsoft Excel (EXCEL.EXE) est introuvable sur ce PC."
    }

    # This is deliberately a normal user-facing launch. It lets Excel and
    # Windows apply the user's Protected View, Trust Center, add-in and file
    # association policies. No security property is changed and this script
    # never calls Workbooks.Open(), Quit() or Stop-Process.
    $quotedWorkbookPath = '"' + $resolvedWorkbookPath + '"'
    [void](Start-Process -FilePath $excelPath -ArgumentList @($quotedWorkbookPath) -PassThru)
    Write-Output "EXCEL_LAUNCH_REQUESTED|$resolvedWorkbookPath"

    # A normal user-facing open needs no COM attachment or confirmation. Excel
    # receives the exact local path through its regular launch path, while VS
    # Code can release the progress notification immediately. VBE handoff keeps
    # the stricter exact-window verification below.
    if (-not $ShowVbe) {
        Write-Output "EXCEL_FAST_OPEN_REQUESTED|$resolvedWorkbookPath"
        exit 0
    }

    # COM is used only after the shell launch, and only to find/activate a
    # workbook Excel has already opened. A pre-existing Excel instance is never
    # asked to open the file through automation.
    for ($attempt = 0; $attempt -lt 80 -and $null -eq $workbook; $attempt++) {
        Start-Sleep -Milliseconds 250
        $candidateExcel = Get-RunningExcel
        if ($null -eq $candidateExcel) {
            continue
        }
        $candidateWorkbook = Find-OpenWorkbook `
            -Excel $candidateExcel `
            -ResolvedWorkbookPath $resolvedWorkbookPath
        if ($null -ne $candidateWorkbook) {
            $excel = $candidateExcel
            $workbook = $candidateWorkbook
            break
        }
        Release-ComObject -ComObject $candidateExcel
    }

    if ($null -ne $workbook) {
        $excel.Visible = $true
        [void]$workbook.Activate()
        try {
            $excel.WindowState = -4137 # xlMaximized
        }
        catch {
            # Maximizing is cosmetic and must not make opening fail.
        }
        Write-Output "EXCEL_OPENED|$resolvedWorkbookPath"

        if ($ShowVbe) {
            try {
                [void]$excel.CommandBars.ExecuteMso('TabDeveloper')
            }
            catch {
                Write-Output 'DEVELOPER_TAB_UNAVAILABLE'
            }

            # Alt+F11 is sent only after verifying the exact Excel foreground
            # process. This displays the VBE without touching AccessVBOM or any
            # Trust Center setting.
            if (Show-VbeWithKeyboardFallback -Excel $excel -Workbook $workbook) {
                Write-Output 'VBE_SHOWN_UI_FALLBACK|ALT_F11_SENT'
            }
            else {
                $failureCode = 3
                $failureMessage = 'VBE_NOT_AVAILABLE|La fenetre Excel cible n''a pas pu etre focalisee en toute securite.'
            }
        }
    }
    else {
        # Some Excel installations do not expose the launched instance through
        # the Running Object Table. Fall back to a uniquely identified visible
        # workbook window; never open the file again through COM.
        if ($ShowVbe) {
            $pathFallbackStatus = Show-VbeForWorkbookPathWithKeyboardFallback `
                -ResolvedWorkbookPath $resolvedWorkbookPath
            if ($pathFallbackStatus) {
                Write-Output "EXCEL_OPENED|$resolvedWorkbookPath"
                Write-Output "VBE_SHOWN_UI_FALLBACK|$pathFallbackStatus"
            }
            else {
                $failureCode = 3
                $failureMessage = "VBE_NOT_AVAILABLE|Excel a recu la demande d'ouverture, mais la fenetre cible n'a pas pu etre identifiee."
            }
        }
        else {
            $workbookFileName = [IO.Path]::GetFileName($resolvedWorkbookPath)
            $visibleMatches = @(
                Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
                    Where-Object {
                        $_.Responding -and
                        $_.MainWindowHandle -ne 0 -and
                        $_.MainWindowTitle.IndexOf(
                            $workbookFileName,
                            [StringComparison]::OrdinalIgnoreCase
                        ) -ge 0
                    }
            )
            if ($visibleMatches.Count -ge 1) {
                Write-Output "EXCEL_OPENED_UI_CONFIRMED|$resolvedWorkbookPath"
            }
            else {
                $failureCode = 1
                $failureMessage = "EXCEL_OPEN_UNCONFIRMED|Excel a recu la demande, mais l'ouverture du classeur n'a pas pu etre confirmee."
            }
        }
    }
}
catch {
    $automationError = $_.Exception.Message
    $pathFallbackStatus = $null
    if ($ShowVbe -and $resolvedWorkbookPath) {
        $pathFallbackStatus = Show-VbeForWorkbookPathWithKeyboardFallback `
            -ResolvedWorkbookPath $resolvedWorkbookPath
    }
    if ($pathFallbackStatus) {
        # Excel can reject all COM calls while its UI is busy. If the workbook
        # window can still be uniquely identified and safely focused, Alt+F11 is
        # a complete successful fallback and the extension must return success.
        $failureCode = 0
        $failureMessage = $null
        Write-Output "EXCEL_OPENED|$resolvedWorkbookPath"
        Write-Output "VBE_SHOWN_UI_FALLBACK|$pathFallbackStatus"
    }
    else {
        if ($failureCode -eq 0) {
            $failureCode = 1
        }
        $failureMessage = "EXCEL_OPEN_FAILED|$automationError"
    }
}
finally {
    # Excel remains user-owned: both new and pre-existing instances stay open.
    Release-ComObject -ComObject $workbook
    Release-ComObject -ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($failureCode -ne 0) {
    [Console]::Error.WriteLine($failureMessage)
    exit $failureCode
}
