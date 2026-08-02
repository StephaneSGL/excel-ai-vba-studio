param(
    [string]$BridgeScriptPath = "$PSScriptRoot\..\scripts\office-ai-apply-edits.ps1"
)

$ErrorActionPreference = 'Stop'
$tokens = $null
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile(
    (Resolve-Path -LiteralPath $BridgeScriptPath).Path,
    [ref]$tokens,
    [ref]$parseErrors
)
if ($parseErrors.Count -gt 0) {
    throw 'Native bridge PowerShell could not be parsed.'
}

$requiredFunctions = @(
    'Convert-ExcelColumnLettersToNumber',
    'Get-NormalizedRangeRef',
    'ConvertFrom-NativeSeriesFormula',
    'Get-NativeSeriesFormulaLiteral',
    'Get-NormalizedNativeSeriesFormulaRange',
    'Assert-NativeSeriesFormulaRangeArgument'
)
$functionAsts = @(
    $ast.FindAll(
        {
            param($node)
            $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
            $requiredFunctions -contains $node.Name
        },
        $true
    ) | Sort-Object { $_.Extent.StartOffset }
)
if ($functionAsts.Count -ne $requiredFunctions.Count) {
    throw 'Native SERIES parser test could not locate every required function.'
}
Invoke-Expression (($functionAsts | ForEach-Object { $_.Extent.Text }) -join "`n`n")

$formula = '=SERIES("Revenue, ""net""",''O''''Brien, Est''!$A$2:$A$5,''O''''Brien, Est''!$AB$2:$AB$5,1)'
$parts = ConvertFrom-NativeSeriesFormula $formula
if ((Get-NativeSeriesFormulaLiteral $parts.NameArgument) -cne 'Revenue, "net"') {
    throw 'SERIES literal-name parsing failed.'
}
Assert-NativeSeriesFormulaRangeArgument `
    $parts.XValuesArgument `
    'A2:A5' `
    "O'Brien, Est" `
    'Quoted category range'
Assert-NativeSeriesFormulaRangeArgument `
    $parts.ValuesArgument `
    'AB2:AB5' `
    "O'Brien, Est" `
    'Quoted values range'

$collisionRejected = $false
try {
    Assert-NativeSeriesFormulaRangeArgument `
        $parts.ValuesArgument `
        'B2:B5' `
        "O'Brien, Est" `
        'Substring collision'
}
catch {
    $collisionRejected = $_.Exception.Message -match 'actual=AB2:AB5 expected=B2:B5'
}
if (-not $collisionRejected) {
    throw 'SERIES range verification accepted AB2:AB5 as B2:B5.'
}

$externalRejected = $false
try {
    Get-NormalizedNativeSeriesFormulaRange `
        "'[other.xlsx]Sheet'!`$B`$2:`$B`$5" `
        'Sheet' | Out-Null
}
catch {
    $externalRejected = $_.Exception.Message -match 'external workbook reference'
}
if (-not $externalRejected) {
    throw 'SERIES range verification accepted an external workbook reference.'
}

$mixedSeparatorRejected = $false
try {
    ConvertFrom-NativeSeriesFormula '=SERIES("Name";A2:A5,B2:B5;1)' | Out-Null
}
catch {
    $mixedSeparatorRejected = $_.Exception.Message -match 'mixes argument separators'
}
if (-not $mixedSeparatorRejected) {
    throw 'SERIES parsing accepted mixed comma and semicolon separators.'
}

$arrayParts = ConvertFrom-NativeSeriesFormula `
    '=SERIES("Array series",{"A","B","C"},{1,2,3},1)'
if (
    $arrayParts.XValuesArgument -cne '{"A","B","C"}' -or
    $arrayParts.ValuesArgument -cne '{1,2,3}' -or
    $arrayParts.PlotOrder -ne 1
) {
    throw 'SERIES array-constant parsing failed.'
}

$unmatchedBraceRejected = $false
try {
    ConvertFrom-NativeSeriesFormula `
        '=SERIES("Array series",{"A","B",{1,2,3},1)' | Out-Null
}
catch {
    $unmatchedBraceRejected = $_.Exception.Message -match 'not balanced'
}
if (-not $unmatchedBraceRejected) {
    throw 'SERIES parsing accepted an unmatched array brace.'
}

Write-Output 'Native SERIES parser tests passed: exact arguments, arrays, quoted sheet, collision and external-reference rejection.'
