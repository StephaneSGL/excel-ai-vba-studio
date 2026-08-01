[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPathBase64,

    [Parameter(Mandatory = $true)]
    [string]$ComponentFileBase64,

    [Parameter(Mandatory = $true)]
    [string]$SourceBase64
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
Add-Type -AssemblyName System.IO.Compression.FileSystem
. (Join-Path $PSScriptRoot 'ooxml-package-signature.ps1')

function Decode-Base64Utf8 {
    param([string]$Value)

    try {
        return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Value))
    }
    catch {
        throw "Argument base64 invalide : $($_.Exception.Message)"
    }
}

function Assert-NoManagedBackupComponent {
    param([string]$Path)

    $fullPath = [IO.Path]::GetFullPath($Path)
    $root = [IO.Path]::GetPathRoot($fullPath)
    $relative = $fullPath.Substring($root.Length)
    $parts = $relative.Split(
        [char[]]@(
            [IO.Path]::DirectorySeparatorChar,
            [IO.Path]::AltDirectorySeparatorChar
        ),
        [StringSplitOptions]::RemoveEmptyEntries
    )
    foreach ($part in $parts) {
        if ($part -ieq '.excel-ai-vba-backups') {
            throw 'Les fichiers du dossier .excel-ai-vba-backups ne peuvent jamais etre une cible VBA.'
        }
    }
}

function Assert-LocalPathSyntax {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw 'Le chemin du classeur est vide.'
    }
    if (
        $Path.StartsWith('\\', [StringComparison]::Ordinal) -or
        $Path.StartsWith('//', [StringComparison]::Ordinal) -or
        $Path.StartsWith('\\?\', [StringComparison]::Ordinal) -or
        $Path.StartsWith('\\.\', [StringComparison]::Ordinal)
    ) {
        throw 'Les chemins UNC, reseau et device sont refuses.'
    }

    $fullPath = [IO.Path]::GetFullPath($Path)
    $root = [IO.Path]::GetPathRoot($fullPath)
    if ([string]::IsNullOrWhiteSpace($root)) {
        throw "Le chemin ne possede pas de racine locale : $fullPath"
    }
    $drive = New-Object IO.DriveInfo($root)
    if (
        $drive.DriveType -eq [IO.DriveType]::Network -or
        $drive.DriveType -eq [IO.DriveType]::Unknown -or
        $drive.DriveType -eq [IO.DriveType]::NoRootDirectory
    ) {
        throw "Le lecteur n'est pas un disque local verifie : $root"
    }
    return $fullPath
}

function Assert-NoReparsePointChain {
    param([string]$Path)

    $fullPath = [IO.Path]::GetFullPath($Path)
    $root = [IO.Path]::GetPathRoot($fullPath)
    if ([string]::IsNullOrWhiteSpace($root)) {
        throw "Le chemin ne possede pas de racine : $fullPath"
    }

    $current = $root
    if (Test-Path -LiteralPath $current) {
        $rootItem = Get-Item -LiteralPath $current -Force
        if (($rootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Point de reanalyse detecte : $current"
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
            throw "Point de reanalyse detecte : $current"
        }
    }
    return $fullPath
}

function Release-ComObject {
    param([AllowNull()][object]$Value)

    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
        }
        catch {
            # Cleanup remains best effort after Excel has been closed.
        }
    }
}

function Get-Sha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

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

function Remove-CreatedTarget {
    param(
        [string]$TargetPath,
        [string]$ExpectedTargetPath
    )

    if (
        [string]::IsNullOrWhiteSpace($TargetPath) -or
        [string]::IsNullOrWhiteSpace($ExpectedTargetPath)
    ) {
        return
    }
    $fullTarget = [IO.Path]::GetFullPath($TargetPath)
    $fullExpected = [IO.Path]::GetFullPath($ExpectedTargetPath)
    if (-not [StringComparer]::OrdinalIgnoreCase.Equals($fullTarget, $fullExpected)) {
        throw 'Le nettoyage a refuse un chemin cible inattendu.'
    }
    Assert-NoManagedBackupComponent $fullTarget
    Assert-NoReparsePointChain $fullTarget | Out-Null
    if (Test-Path -LiteralPath $fullTarget -PathType Leaf) {
        Remove-Item -LiteralPath $fullTarget -Force
    }
}

if ($null -eq ('ExcelAiVbaStudio.NativeProcess' -as [type])) {
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

$sourceInput = Decode-Base64Utf8 $WorkbookPathBase64
$componentFile = Decode-Base64Utf8 $ComponentFileBase64
$sourceCode = Decode-Base64Utf8 $SourceBase64

$sourceSyntaxPath = Assert-LocalPathSyntax $sourceInput
Assert-NoManagedBackupComponent $sourceSyntaxPath
Assert-NoReparsePointChain $sourceSyntaxPath | Out-Null
if (-not (Test-Path -LiteralPath $sourceSyntaxPath -PathType Leaf)) {
    throw "Classeur source introuvable : $sourceSyntaxPath"
}
$sourcePath = [IO.Path]::GetFullPath(
    (Resolve-Path -LiteralPath $sourceSyntaxPath).ProviderPath
)
Assert-NoManagedBackupComponent $sourcePath
Assert-NoReparsePointChain $sourcePath | Out-Null
if ([IO.Path]::GetExtension($sourcePath) -ine '.xlsx') {
    throw 'La preparation macro accepte uniquement un classeur source .xlsx.'
}
Assert-OoxmlPackageUnsigned $sourcePath

if (
    [IO.Path]::GetFileName($componentFile) -cne $componentFile -or
    [string]::IsNullOrWhiteSpace($componentFile)
) {
    throw 'componentFile doit etre un nom .bas ou .cls sans chemin.'
}
$componentExtension = [IO.Path]::GetExtension($componentFile)
if ($componentExtension -ine '.bas' -and $componentExtension -ine '.cls') {
    throw 'La preparation initiale accepte uniquement les modules .bas et .cls.'
}
if ($sourceCode.Length -gt 2000000) {
    throw 'Le code VBA depasse la limite de 2 000 000 de caracteres.'
}
if ($sourceCode.IndexOf([char]0) -ge 0) {
    throw 'Le code VBA contient un caractere NUL.'
}

$lines = @($sourceCode.Replace("`r`n", "`n").Replace("`r", "`n").Split("`n"))
$fileModuleName = [IO.Path]::GetFileNameWithoutExtension($componentFile)
$moduleName = $fileModuleName
$nameAttributeIndex = -1
for ($index = 0; $index -lt $lines.Count; $index++) {
    if ($lines[$index] -match '^[ \t]*Attribute[ \t]+VB_Name[ \t]*=[ \t]*"([^"]+)"[ \t]*$') {
        if (
            $nameAttributeIndex -ge 0 -and
            -not [StringComparer]::OrdinalIgnoreCase.Equals($moduleName, $matches[1])
        ) {
            throw 'Le source VBA contient plusieurs attributs VB_Name incompatibles.'
        }
        $moduleName = $matches[1]
        $nameAttributeIndex = $index
    }
    if (
        $componentExtension -ieq '.cls' -and
        $lines[$index] -match '^[ \t]*Attribute[ \t]+VB_(PredeclaredId|Exposed|Creatable|GlobalNameSpace)[ \t]*=[ \t]*True[ \t]*$'
    ) {
        throw (
            'Cette classe exige des attributs VBA speciaux que VBIDE ne peut pas ' +
            'recreer de facon sure. Creez-la dans le VBE natif.'
        )
    }
}
if ($moduleName -cnotmatch '^[A-Za-z_][A-Za-z0-9_]{0,30}$') {
    throw "Nom de composant VBA invalide : $moduleName"
}
if (-not [StringComparer]::OrdinalIgnoreCase.Equals($moduleName, $fileModuleName)) {
    throw (
        "Le nom Attribute VB_Name '$moduleName' ne correspond pas au fichier " +
        "'$componentFile'."
    )
}

if ($nameAttributeIndex -ge 0) {
    if ($nameAttributeIndex + 1 -ge $lines.Count) {
        throw 'Le fichier VBA ne contient aucun code apres ses attributs.'
    }
    $codeLines = @(
        $lines[($nameAttributeIndex + 1)..($lines.Count - 1)] |
            Where-Object {
                $_ -notmatch '^[ \t]*Attribute[ \t]+[A-Za-z_][A-Za-z0-9_.]*[ \t]*='
            }
    )
}
else {
    $headerLimit = [Math]::Min(10, $lines.Count)
    for ($index = 0; $index -lt $headerLimit; $index++) {
        if ($lines[$index] -match '^[ \t]*(VERSION|BEGIN)(?:[ \t]|$)') {
            throw 'Le preambule exporte VERSION/BEGIN exige un Attribute VB_Name valide.'
        }
    }
    $codeLines = $lines
}
$code = ($codeLines -join "`r`n").Trim()
if ([string]::IsNullOrWhiteSpace($code)) {
    throw 'Le code VBA est vide apres normalisation.'
}

$targetPath = [IO.Path]::ChangeExtension($sourcePath, '.xlsm')
$stagingPath = [IO.Path]::Combine(
    [IO.Path]::GetDirectoryName($sourcePath),
    (
        '.' +
        [IO.Path]::GetFileNameWithoutExtension($sourcePath) +
        '.excel-ai-vba-bootstrap-' +
        [Guid]::NewGuid().ToString('N') +
        '.xlsm'
    )
)
$expectedStagingPath = $stagingPath
Assert-LocalPathSyntax $targetPath | Out-Null
Assert-NoManagedBackupComponent $targetPath
Assert-NoReparsePointChain $targetPath | Out-Null
if (Test-Path -LiteralPath $targetPath) {
    throw "Le classeur cible existe deja : $targetPath. Utilisez directement ce fichier .xlsm."
}
Assert-LocalPathSyntax $stagingPath | Out-Null
Assert-NoManagedBackupComponent $stagingPath
Assert-NoReparsePointChain $stagingPath | Out-Null
if (Test-Path -LiteralPath $stagingPath) {
    throw "Le fichier temporaire existe deja : $stagingPath"
}
$sourceSha256Before = Get-Sha256 $sourcePath

$excel = $null
$workbooks = $null
$workbook = $null
$vbaProject = $null
$components = $null
$existingComponent = $null
$component = $null
$codeModule = $null
$stagingCreated = $false
$operationError = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.AutomationSecurity = 3
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.ScreenUpdating = $false
    $excel.Visible = $false

    $windowHandle = [IntPtr]([int64]$excel.Hwnd)
    [uint32]$ownedProcessId = 0
    [void][ExcelAiVbaStudio.NativeProcess]::GetWindowThreadProcessId(
        $windowHandle,
        [ref]$ownedProcessId
    )
    if ($ownedProcessId -eq 0) {
        throw 'Impossible de determiner le processus Excel controle.'
    }
    Write-Output "OWNED_EXCEL_PID|$ownedProcessId"

    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Open($sourcePath, 0, $true)
    $workbook.SaveAs($stagingPath, 52)
    $stagingCreated = $true

    try {
        $vbaProject = $workbook.VBProject
        $components = $vbaProject.VBComponents
    }
    catch {
        throw (
            "Excel refuse l'acces au projet VBA. Activez manuellement " +
            "'Acces approuve au modele d'objet du projet VBA', puis recommencez. " +
            "L'extension ne modifie jamais ce reglage."
        )
    }

    try {
        $existingComponent = $components.Item($moduleName)
    }
    catch {
        $existingComponent = $null
    }
    if ($null -ne $existingComponent) {
        throw "Un composant VBA nomme '$moduleName' existe deja."
    }

    $componentType = if ($componentExtension -ieq '.bas') { 1 } else { 2 }
    $component = $components.Add($componentType)
    $component.Name = $moduleName
    $codeModule = $component.CodeModule
    $codeModule.AddFromString($code)
    if (
        $component.Name -cne $moduleName -or
        [int]$codeModule.CountOfLines -le 0
    ) {
        throw "Excel n'a pas confirme le composant VBA ajoute."
    }
    $workbook.Save()
}
catch {
    $operationError = $_.Exception
}
finally {
    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
        }
        catch {
            if ($null -eq $operationError) {
                $operationError = $_.Exception
            }
        }
    }
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
            if ($null -eq $operationError) {
                $operationError = $_.Exception
            }
        }
    }

    Release-ComObject $codeModule
    Release-ComObject $component
    Release-ComObject $existingComponent
    Release-ComObject $components
    Release-ComObject $vbaProject
    Release-ComObject $workbook
    Release-ComObject $workbooks
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($null -ne $operationError) {
    if ($stagingCreated) {
        try {
            Remove-CreatedTarget $stagingPath $expectedStagingPath
        }
        catch {
            $operationError = New-Object System.Exception(
                "$($operationError.Message) Nettoyage cible incomplet : $($_.Exception.Message)"
            )
        }
    }
    Write-Error $operationError.Message
    exit 1
}

$verificationError = $null
try {
    Assert-NoReparsePointChain $stagingPath | Out-Null
    if (-not (Test-Path -LiteralPath $stagingPath -PathType Leaf)) {
        throw "Excel n'a pas cree le classeur XLSM attendu."
    }
    $archive = [IO.Compression.ZipFile]::OpenRead($stagingPath)
    try {
        $vbaEntry = $archive.GetEntry('xl/vbaProject.bin')
        if ($null -eq $vbaEntry -or $vbaEntry.Length -le 0) {
            throw 'Le classeur cree ne contient pas de projet VBA persistant.'
        }
    }
    finally {
        $archive.Dispose()
    }
    $sourceSha256After = Get-Sha256 $sourcePath
    if ($sourceSha256After -cne $sourceSha256Before) {
        throw 'Le classeur XLSX source a change pendant la conversion.'
    }
    Assert-NoReparsePointChain $targetPath | Out-Null
    if (Test-Path -LiteralPath $targetPath) {
        throw "Le classeur cible est apparu pendant la conversion : $targetPath"
    }
    [IO.File]::Move($stagingPath, $targetPath)
    $stagingCreated = $false
}
catch {
    $verificationError = $_.Exception
}

if ($null -ne $verificationError) {
    if ($stagingCreated) {
        try {
            Remove-CreatedTarget $stagingPath $expectedStagingPath
        }
        catch {
            $verificationError = New-Object System.Exception(
                "$($verificationError.Message) Nettoyage cible incomplet : $($_.Exception.Message)"
            )
        }
    }
    Write-Error $verificationError.Message
    exit 1
}

$workbookSha256 = Get-Sha256 $targetPath
[ordered]@{
    ok = $true
    targetWorkbookPath = [IO.Path]::GetFullPath($targetPath)
    sourceWorkbookPath = $sourcePath
    convertedToXlsm = $true
    changed = $true
    modifiedModules = @($moduleName)
    workbookSha256 = $workbookSha256
    macrosExecuted = $false
    accessVbomChanged = $false
} | ConvertTo-Json -Compress
