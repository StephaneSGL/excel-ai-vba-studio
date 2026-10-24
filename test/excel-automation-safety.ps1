param([Parameter(Mandatory = $true)][string]$RepositoryRoot)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
. (Join-Path $RepositoryRoot 'scripts/excel-automation-safety.ps1')

function Assert-Test([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
function Assert-Refused([scriptblock]$Action, [string]$Message) {
    $refused = $false
    try { & $Action } catch { $refused = $true }
    Assert-Test $refused $Message
}

foreach ($file in @('excel-automation-safety.ps1', 'office-ai-export.ps1', 'prepare-macro-workbook.ps1', 'office-ai-apply-edits.ps1', 'apply-vba-designer.ps1')) {
    $fileTokens = $null; $fileErrors = $null
    [void][Management.Automation.Language.Parser]::ParseFile(
        (Join-Path $RepositoryRoot (Join-Path 'scripts' $file)), [ref]$fileTokens, [ref]$fileErrors
    )
    Assert-Test ($fileErrors.Count -eq 0) "$file must parse without errors."
}

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
namespace ExcelAutomationSafetyTest {
    public class Application {
        public string FaultProperty, FaultMode, DriftAfterOpen;
        public int Writes, OpenCount, QuitCount;
        public bool GuardOpen;
        public bool DisplayAlerts, ScreenUpdating, Visible;
        public object Macros = 1, Events = true, Links = true, Calc = -4105, BeforeSave = true;
        public Workbooks Workbooks;
        public Workbook LastOpened, Guard;
        public Application() { Workbooks = new Workbooks(this); }
        object Get(string name, object value) {
            if (FaultProperty == name && FaultMode == "get-throw") throw new Exception("blocked getter");
            if (FaultProperty == name && FaultMode == "get-null") return null;
            return value;
        }
        void Set(string name, ref object field, object value) {
            Writes++;
            if (FaultProperty == name && FaultMode == "set-throw") throw new Exception("blocked setter");
            if (FaultProperty == name && FaultMode == "set-ignore") return;
            if (name == "Calculation" && !GuardOpen) throw new Exception("calculation requires guard");
            field = value;
        }
        public object AutomationSecurity { get { return Get("AutomationSecurity", Macros); } set { Set("AutomationSecurity", ref Macros, value); } }
        public object EnableEvents { get { return Get("EnableEvents", Events); } set { Set("EnableEvents", ref Events, value); } }
        public object AskToUpdateLinks { get { return Get("AskToUpdateLinks", Links); } set { Set("AskToUpdateLinks", ref Links, value); } }
        public object Calculation { get { return Get("Calculation", Calc); } set { Set("Calculation", ref Calc, value); } }
        public object CalculateBeforeSave { get { return Get("CalculateBeforeSave", BeforeSave); } set { Set("CalculateBeforeSave", ref BeforeSave, value); } }
        public void Quit() { QuitCount++; }
        public void Drift(string name) {
            if (name == "AutomationSecurity") Macros = 1;
            if (name == "EnableEvents") Events = true;
            if (name == "AskToUpdateLinks") Links = true;
            if (name == "Calculation") Calc = -4105;
            if (name == "CalculateBeforeSave") BeforeSave = true;
        }
    }
    public class Workbooks {
        Application app;
        public Workbooks(Application value) { app = value; }
        public Workbook Add(int template) {
            if (template != -4167 || (int)app.AutomationSecurity != 3 || (bool)app.EnableEvents || (bool)app.AskToUpdateLinks)
                throw new Exception("unsafe guard creation");
            app.GuardOpen = true;
            app.Guard = new Workbook();
            return app.Guard;
        }
        public Workbook Open(string path, int updateLinks, bool readOnly) {
            if (updateLinks != 0 || !readOnly) throw new Exception("invalid open parameters");
            app.OpenCount++;
            app.LastOpened = new Workbook();
            app.Drift(app.DriftAfterOpen);
            return app.LastOpened;
        }
    }
    public class Workbook {
        public bool Closed;
        public void Close(bool save) { if (save) throw new Exception("unexpected save"); Closed = true; }
    }
    public class Process {
        public int Id = 4242, KillCount, DisposeCount;
        public string ProcessName = "EXCEL";
        public DateTime StartTime = DateTime.UtcNow.AddMinutes(1);
        public bool WaitForExit(int timeout) { return KillCount != 0; }
        public void Kill() { KillCount++; }
        public void Dispose() { DisposeCount++; }
    }
}
'@

# Load only these real function definitions, never the designer main program.
$tokens = $null; $parseErrors = $null
$designerAst = [Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $RepositoryRoot 'scripts/apply-vba-designer.ps1'), [ref]$tokens, [ref]$parseErrors
)
Assert-Test ($parseErrors.Count -eq 0) 'Designer PowerShell must parse.'
foreach ($name in @('Release-ComObject', 'Ensure-ExcelSession', 'Cleanup-Excel', 'Open-WorkbookReadOnly')) {
    $definition = $designerAst.Find({ param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    }, $true)
    Assert-Test ($null -ne $definition) "Missing designer function $name"
    Invoke-Expression $definition.Extent.Text
}

# All test objects below are synthetic; no COM object or workbook is opened.
$script:fakeProcess = [ExcelAutomationSafetyTest.Process]::new()
$script:fakeExcel = $null
$script:preexistingIds = @()
function Get-ExcelAutomationProcess { param([int]$ProcessId) return $script:fakeProcess }
function Get-ExcelAutomationProcessIds { return $script:preexistingIds }
function Get-ExcelProcessIdentity { param([object]$ExcelApp) return @{ ProcessId = 4242 } }
function Stop-OwnedExcelProcess { param($Identity, $GracePeriodMilliseconds = 0)
    Stop-ExcelAutomationProcess $Identity $GracePeriodMilliseconds
}
function New-Object { param([string]$ComObject)
    Assert-Test ($ComObject -eq 'Excel.Application') 'Only the synthetic Excel factory is permitted.'
    return $script:fakeExcel
}

$script:fakeExcel = [ExcelAutomationSafetyTest.Application]::new()
$script:preexistingIds = @(4242)
Assert-Refused { Ensure-ExcelSession } 'A preexisting Excel process must be refused.'
Assert-Test ($script:fakeExcel.Writes -eq 0 -and $script:fakeExcel.QuitCount -eq 0) 'Unowned Excel must not be changed or quit.'
Assert-Test ($script:fakeProcess.KillCount -eq 0) 'Unowned Excel must not be killed.'
$script:preexistingIds = @()

$script:fakeExcel = [ExcelAutomationSafetyTest.Application]::new()
$excel, $identity = Ensure-ExcelSession
Assert-Test ($identity.StartTimeUtcTicks -match '^\d{15,19}$') 'Owned identity must contain start ticks.'
Assert-Test ($excel.GuardOpen -and $excel.OpenCount -eq 0) 'Session must create a safe private guard before target open.'
$first = Open-WorkbookReadOnly $excel 'synthetic.xlsx'
$excel.Drift('EnableEvents')
$second = Open-WorkbookReadOnly $excel 'synthetic.xlsx'
Assert-Test ($excel.OpenCount -eq 2 -and -not $excel.EnableEvents) 'Every reopen must reapply the safety policy.'
$first.Close($false); $second.Close($false)
Cleanup-Excel $excel $identity
Assert-Test ($excel.Guard.Closed -and $excel.QuitCount -eq 1) 'Owned session cleanup must close the guard and quit.'

$properties = @('AutomationSecurity', 'EnableEvents', 'AskToUpdateLinks', 'Calculation', 'CalculateBeforeSave')
foreach ($property in $properties) {
    foreach ($mode in @('set-throw', 'set-ignore', 'get-throw', 'get-null')) {
        $excel = [ExcelAutomationSafetyTest.Application]::new()
        $excel.GuardOpen = $true
        $excel.FaultProperty = $property; $excel.FaultMode = $mode
        Assert-Refused { Open-WorkbookReadOnly $excel 'synthetic.xlsx' } "Unsafe $property/$mode must be refused."
        Assert-Test ($excel.OpenCount -eq 0) "Unsafe $property/$mode reached Open."
    }
    $excel = [ExcelAutomationSafetyTest.Application]::new()
    $excel.GuardOpen = $true; $excel.DriftAfterOpen = $property
    Assert-Refused { Open-WorkbookReadOnly $excel 'synthetic.xlsx' } "Post-open $property drift must be refused."
    Assert-Test ($excel.OpenCount -eq 1 -and $excel.LastOpened.Closed) 'A rejected post-open workbook must be closed.'
}

$script:fakeExcel = [ExcelAutomationSafetyTest.Application]::new()
$script:fakeExcel.FaultProperty = 'Calculation'; $script:fakeExcel.FaultMode = 'set-throw'
Assert-Refused { Ensure-ExcelSession } 'Guard initialization failure must be refused.'
Assert-Test ($script:fakeExcel.Guard.Closed -and $script:fakeExcel.QuitCount -eq 1) 'Guard initialization failure must clean up owned Excel.'

$script:fakeProcess = [ExcelAutomationSafetyTest.Process]::new()
$now = [DateTime]::UtcNow.Ticks
$identity = Get-ExcelAutomationProcessIdentity 4242 @() $now
Assert-Refused { Get-ExcelAutomationProcessIdentity 4242 @(4242) $now } 'Snapshot ownership rejection failed.'
$script:fakeProcess.ProcessName = 'OTHER'
Assert-Refused { Get-ExcelAutomationProcessIdentity 4242 @() $now } 'Non-Excel ownership rejection failed.'
Stop-ExcelAutomationProcess $identity 0
Assert-Test ($script:fakeProcess.KillCount -eq 0) 'Name mismatch must prevent termination.'
$script:fakeProcess.ProcessName = 'EXCEL'
$script:fakeProcess.StartTime = $script:fakeProcess.StartTime.AddTicks(1)
Stop-ExcelAutomationProcess $identity 0
Assert-Test ($script:fakeProcess.KillCount -eq 0) 'Reused PID must prevent termination.'
$script:fakeProcess.StartTime = [DateTime]::new([long]$identity.StartTimeUtcTicks, [DateTimeKind]::Utc)
Stop-ExcelAutomationProcess $identity 0
Assert-Test ($script:fakeProcess.KillCount -eq 1) 'Exact identity must terminate the owned process.'
$script:fakeProcess.StartTime = [DateTime]::UtcNow.AddMinutes(-1)
Assert-Refused { Get-ExcelAutomationProcessIdentity 4242 @() $now } 'An old process omitted from the snapshot must be refused.'

[Console]::Out.WriteLine('Synthetic Excel automation safety tests passed (no Office/COM/workbooks used).')
