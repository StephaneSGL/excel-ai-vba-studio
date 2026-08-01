[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPathBase64
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:MaxOutputBytes = 262144
$script:MaxTrustedLocations = 64
$script:MaxWorkbookBytes = 536870912

function Decode-Base64Utf8 {
    param([string]$Value)

    try {
        return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Value))
    }
    catch {
        throw 'WorkbookPathBase64 is not valid UTF-8 base64.'
    }
}

function Resolve-SafeWorkbookPath {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value) -or -not [IO.Path]::IsPathRooted($Value)) {
        throw 'The workbook path must be an absolute local path.'
    }
    if (
        $Value.StartsWith('\\', [StringComparison]::Ordinal) -or
        $Value.StartsWith('//', [StringComparison]::Ordinal) -or
        $Value.StartsWith('\\?\', [StringComparison]::Ordinal) -or
        $Value.StartsWith('\\.\', [StringComparison]::Ordinal)
    ) {
        throw 'UNC, network, and device paths are not allowed.'
    }

    $fullPath = [IO.Path]::GetFullPath($Value)
    $root = [IO.Path]::GetPathRoot($fullPath)
    if ([string]::IsNullOrWhiteSpace($root)) {
        throw 'The workbook path has no local root.'
    }
    $drive = New-Object IO.DriveInfo($root)
    if (
        $drive.DriveType -eq [IO.DriveType]::Network -or
        $drive.DriveType -eq [IO.DriveType]::Unknown -or
        $drive.DriveType -eq [IO.DriveType]::NoRootDirectory
    ) {
        throw 'The workbook must be on a verified local drive.'
    }

    $current = $root
    $relative = $fullPath.Substring($root.Length)
    $parts = $relative.Split(
        [char[]]@([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar),
        [StringSplitOptions]::RemoveEmptyEntries
    )
    foreach ($part in $parts) {
        $current = [IO.Path]::Combine($current, $part)
        if (-not (Test-Path -LiteralPath $current)) {
            break
        }
        $item = Get-Item -LiteralPath $current -Force
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw 'Reparse points are not allowed in the workbook path.'
        }
    }
    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        throw 'The workbook does not exist or is not a file.'
    }
    $resolved = [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $fullPath).ProviderPath)
    $extension = [IO.Path]::GetExtension($resolved).ToLowerInvariant()
    if ($extension -notin @('.xlsx', '.xlsm', '.xlsb', '.xls', '.xltx', '.xltm', '.xlam', '.xlt', '.xla')) {
        throw 'The file extension is not a supported Excel workbook format.'
    }
    return $resolved
}

function Get-ZoneInformation {
    param([string]$Path)

    try {
        [void](Get-Item -LiteralPath $Path -Stream 'Zone.Identifier' -ErrorAction Stop)
    }
    catch [System.NotSupportedException] {
        return [ordered]@{ status = 'unsupported'; zoneId = $null }
    }
    catch [System.IO.FileNotFoundException] {
        return [ordered]@{ status = 'absent'; zoneId = $null }
    }
    catch {
        if (
            $_.FullyQualifiedErrorId -match 'AlternateDataStreamNotFound|PathNotFound|ItemNotFound' -or
            $_.Exception.Message -match '(?i)alternate data stream.*not found'
        ) {
            return [ordered]@{ status = 'absent'; zoneId = $null }
        }
        return [ordered]@{ status = 'unreadable'; zoneId = $null }
    }

    try {
        $lines = @(Get-Content -LiteralPath $Path -Stream 'Zone.Identifier' -ErrorAction Stop)
        foreach ($line in $lines) {
            if ($line -match '^ZoneId\s*=\s*([0-9]+)\s*$') {
                return [ordered]@{ status = 'read'; zoneId = [int]$matches[1] }
            }
        }
        return [ordered]@{ status = 'read'; zoneId = $null }
    }
    catch {
        return [ordered]@{ status = 'unreadable'; zoneId = $null }
    }
}

function Get-Sha256 {
	param([string]$Path)

    $stream = New-Object IO.FileStream(
        $Path,
        [IO.FileMode]::Open,
        [IO.FileAccess]::Read,
        [IO.FileShare]::ReadWrite
    )
	$sha = [Security.Cryptography.SHA256]::Create()
	try {
		$buffer = New-Object byte[] 81920
		$total = [int64]0
		while (($read = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
			$total += $read
			if ($total -gt $script:MaxWorkbookBytes) {
				throw 'The workbook exceeds the 512 MiB inspection limit.'
			}
			[void]$sha.TransformBlock($buffer, 0, $read, $buffer, 0)
		}
		[void]$sha.TransformFinalBlock($buffer, 0, 0)
		return ([BitConverter]::ToString($sha.Hash)).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
        $stream.Dispose()
    }
}

function Get-CompoundInventory {
    param([string]$Path)

    $entries = New-Object System.Collections.ArrayList
    $stream = New-Object IO.FileStream(
        $Path,
        [IO.FileMode]::Open,
        [IO.FileAccess]::Read,
        [IO.FileShare]::ReadWrite
    )
    $reader = New-Object IO.BinaryReader($stream, [Text.Encoding]::Unicode, $true)
    try {
        $header = $reader.ReadBytes(512)
        if ($header.Length -ne 512) {
            throw 'The compound file header is incomplete.'
        }
        $sectorShift = [BitConverter]::ToUInt16($header, 30)
        if ($sectorShift -ne 9 -and $sectorShift -ne 12) {
            throw 'The compound file sector size is invalid.'
        }
        $sectorSize = 1 -shl $sectorShift
        $firstDirectorySector = [BitConverter]::ToInt32($header, 48)
        $fatSectorCount = [BitConverter]::ToUInt32($header, 44)
        if ($fatSectorCount -gt 4096) {
            throw 'The compound file FAT exceeds the inspection limit.'
        }

        $difat = New-Object 'System.Collections.Generic.List[int]'
        for ($index = 0; $index -lt 109; $index++) {
            $sectorId = [BitConverter]::ToInt32($header, 76 + ($index * 4))
            if ($sectorId -ge 0) {
                $difat.Add($sectorId)
            }
        }
        $nextDifat = [BitConverter]::ToInt32($header, 68)
        $difatSectorCount = [BitConverter]::ToUInt32($header, 72)
        if ($difatSectorCount -gt 256) {
            throw 'The compound file DIFAT exceeds the inspection limit.'
        }
        for ($chainIndex = 0; $chainIndex -lt [int]$difatSectorCount; $chainIndex++) {
            if ($nextDifat -lt 0) { break }
            $offset = ([int64]$nextDifat + 1) * $sectorSize
            if ($offset -lt 0 -or $offset + $sectorSize -gt $stream.Length) {
                throw 'The compound file DIFAT points outside the file.'
            }
            $stream.Position = $offset
            $sector = $reader.ReadBytes($sectorSize)
            for ($entryIndex = 0; $entryIndex -lt (($sectorSize / 4) - 1); $entryIndex++) {
                $sectorId = [BitConverter]::ToInt32($sector, $entryIndex * 4)
                if ($sectorId -ge 0) { $difat.Add($sectorId) }
            }
            $nextDifat = [BitConverter]::ToInt32($sector, $sectorSize - 4)
        }
        if ($difat.Count -lt $fatSectorCount) {
            throw 'The compound file FAT is incomplete.'
        }

        $fat = New-Object 'System.Collections.Generic.List[int]'
        for ($fatIndex = 0; $fatIndex -lt [int]$fatSectorCount; $fatIndex++) {
            $fatSectorId = $difat[$fatIndex]
            $offset = ([int64]$fatSectorId + 1) * $sectorSize
            if ($fatSectorId -lt 0 -or $offset + $sectorSize -gt $stream.Length) {
                throw 'The compound file FAT points outside the file.'
            }
            $stream.Position = $offset
            $sector = $reader.ReadBytes($sectorSize)
            for ($entryIndex = 0; $entryIndex -lt ($sectorSize / 4); $entryIndex++) {
                $fat.Add([BitConverter]::ToInt32($sector, $entryIndex * 4))
            }
        }

        $seen = New-Object 'System.Collections.Generic.HashSet[int]'
        $directorySector = $firstDirectorySector
        while ($directorySector -ge 0) {
            if ($seen.Count -ge 4096 -or -not $seen.Add($directorySector)) {
                throw 'The compound file directory chain is cyclic or too large.'
            }
            if ($directorySector -ge $fat.Count) {
                throw 'The compound file directory chain is invalid.'
            }
            $offset = ([int64]$directorySector + 1) * $sectorSize
            if ($offset + $sectorSize -gt $stream.Length) {
                throw 'The compound file directory points outside the file.'
            }
            $stream.Position = $offset
            $sector = $reader.ReadBytes($sectorSize)
            for ($entryOffset = 0; $entryOffset -lt $sectorSize; $entryOffset += 128) {
                $nameBytes = [BitConverter]::ToUInt16($sector, $entryOffset + 64)
                if ($nameBytes -ge 2 -and $nameBytes -le 64 -and ($nameBytes % 2) -eq 0) {
                    $name = [Text.Encoding]::Unicode.GetString(
                        $sector,
                        $entryOffset,
                        $nameBytes - 2
                    )
                    if (-not [string]::IsNullOrWhiteSpace($name)) {
                        [void]$entries.Add([ordered]@{
                            name = $name
                            objectType = [int]$sector[$entryOffset + 66]
                        })
                    }
                }
            }
            $directorySector = $fat[$directorySector]
        }
    }
    finally {
        $reader.Dispose()
        $stream.Dispose()
    }

    $hasVbaStorage = $false
    $hasVbaProjectStream = $false
	$hasEncryptionInfo = $false
	$hasEncryptedPackage = $false
	$hasPackageSignature = $false
	$hasWorkbookStream = $false
	foreach ($entry in $entries) {
		if (($entry.name -ieq 'Workbook' -or $entry.name -ieq 'Book') -and $entry.objectType -eq 2) { $hasWorkbookStream = $true }
        if ($entry.name -ieq 'VBA' -and $entry.objectType -eq 1) { $hasVbaStorage = $true }
        if ($entry.name -ieq '_VBA_PROJECT' -and $entry.objectType -eq 2) { $hasVbaProjectStream = $true }
        if ($entry.name -ieq 'EncryptionInfo' -and $entry.objectType -eq 2) { $hasEncryptionInfo = $true }
        if ($entry.name -ieq 'EncryptedPackage' -and $entry.objectType -eq 2) { $hasEncryptedPackage = $true }
        if (
            ($entry.name -ieq '_signatures' -and $entry.objectType -eq 2) -or
            ($entry.name -ieq '_xmlsignatures' -and $entry.objectType -eq 1)
        ) {
            $hasPackageSignature = $true
        }
    }
    $hasVba = $hasVbaStorage -and $hasVbaProjectStream
    $officePackageEncrypted = $hasEncryptionInfo -and $hasEncryptedPackage
    $vbaSignatureStatus = if ($officePackageEncrypted -or $hasVba) { 'unknown' } else { 'absent' }
    $packageSignatureStatus = if ($hasPackageSignature) {
        'present'
    } elseif ($officePackageEncrypted) {
        'unknown'
    } else {
        'absent'
    }
    return [ordered]@{
        hasVbaProject = $hasVba
        hasVbaSignature = $false
        hasPackageSignature = $hasPackageSignature
        vbaSignatureStatus = $vbaSignatureStatus
        packageSignatureStatus = $packageSignatureStatus
		packageSignatureVerificationStatus = if ($packageSignatureStatus -eq 'absent') { 'notPresent' } else { 'unverifiable' }
		officePackageEncrypted = $officePackageEncrypted
		hasWorkbookPart = $hasWorkbookStream
        vbaProjectProtectionStatus = if ($hasVba) { 'unknown' } else { 'absent' }
        sensitivityLabelIds = @()
    }
}

function Read-BoundedZipEntryBytes {
    param(
        [IO.Compression.ZipArchiveEntry]$Entry,
        [int64]$MaximumBytes
    )

    if ($Entry.Length -gt $MaximumBytes) {
        throw 'The OPC XML part exceeds the inspection limit.'
    }
    $entryStream = $Entry.Open()
    $memory = New-Object IO.MemoryStream
    try {
        $buffer = New-Object byte[] 8192
        $total = 0
        while (($read = $entryStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $total += $read
            if ($total -gt $MaximumBytes) {
                throw 'The OPC XML part expands beyond the inspection limit.'
            }
            $memory.Write($buffer, 0, $read)
        }
        return $memory.ToArray()
    }
    finally {
        $memory.Dispose()
        $entryStream.Dispose()
    }
}

function Normalize-OpcPartName {
    param([AllowNull()][string]$Value)

    if (
        [string]::IsNullOrWhiteSpace($Value) -or
        $Value.Length -gt 2048 -or
        -not $Value.StartsWith('/', [StringComparison]::Ordinal) -or
        $Value.StartsWith('//', [StringComparison]::Ordinal) -or
        $Value.IndexOf('\') -ge 0 -or
        $Value.IndexOf('?') -ge 0 -or
        $Value.IndexOf('#') -ge 0 -or
        $Value.IndexOf([char]0) -ge 0
    ) {
        return $null
    }
    $segments = $Value.Substring(1).Split('/')
    if ($segments.Count -eq 0) { return $null }
    foreach ($segment in $segments) {
        if (
            [string]::IsNullOrWhiteSpace($segment) -or
            $segment -eq '.' -or
            $segment -eq '..' -or
            $segment -match '(?i)%2f|%5c'
        ) {
            return $null
        }
    }
    return '/' + ($segments -join '/')
}

function Resolve-OpcRelationshipTarget {
    param(
        [string]$SourcePartName,
        [AllowNull()][string]$Target
    )

    if (
        [string]::IsNullOrWhiteSpace($Target) -or
        $Target.Length -gt 2048 -or
        $Target.IndexOf('\') -ge 0 -or
        $Target.IndexOf([char]0) -ge 0
    ) {
        return $null
    }
    try {
        $basePath = if ($SourcePartName -eq '/') { '/' } else { $SourcePartName }
        $baseUri = New-Object Uri("http://opc.invalid$basePath", [UriKind]::Absolute)
        $resolvedUri = New-Object Uri($baseUri, $Target)
        if (
            $resolvedUri.Scheme -cne 'http' -or
            $resolvedUri.Host -cne 'opc.invalid' -or
            -not [string]::IsNullOrEmpty($resolvedUri.Query) -or
            -not [string]::IsNullOrEmpty($resolvedUri.Fragment)
        ) {
            return $null
        }
        return Normalize-OpcPartName $resolvedUri.AbsolutePath
    }
    catch {
        return $null
    }
}

function Get-OpcRelationshipPartName {
    param([string]$SourcePartName)

    if ($SourcePartName -eq '/') { return '/_rels/.rels' }
    $lastSlash = $SourcePartName.LastIndexOf('/')
    if ($lastSlash -lt 0 -or $lastSlash -ge $SourcePartName.Length - 1) { return $null }
    $directory = $SourcePartName.Substring(0, $lastSlash + 1)
    $fileName = $SourcePartName.Substring($lastSlash + 1)
    return "$directory`_rels/$fileName.rels"
}

function New-SafeXmlReaderSettings {
    $settings = New-Object Xml.XmlReaderSettings
    $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $settings.MaxCharactersInDocument = 1048576
    $settings.MaxCharactersFromEntities = 0
    $settings.IgnoreComments = $true
    return $settings
}

function Read-OpcContentTypes {
    param([AllowNull()][IO.Compression.ZipArchiveEntry]$Entry)

    $overrides = @{}
    $defaults = @{}
    if ($null -eq $Entry) {
        return [ordered]@{ valid = $false; overrides = $overrides; defaults = $defaults }
    }
    $valid = $true
    $rootSeen = $false
    try {
        $bytes = Read-BoundedZipEntryBytes $Entry 1048576
        $stream = New-Object IO.MemoryStream(, $bytes)
        $reader = [Xml.XmlReader]::Create($stream, (New-SafeXmlReaderSettings))
        try {
            while ($reader.Read()) {
                if ($reader.NodeType -ne [Xml.XmlNodeType]::Element) { continue }
                if ($reader.Depth -eq 0) {
                    if (
                        $reader.LocalName -cne 'Types' -or
                        $reader.NamespaceURI -cne 'http://schemas.openxmlformats.org/package/2006/content-types'
                    ) {
                        $valid = $false
                    }
                    $rootSeen = $true
                    continue
                }
                if ($reader.NamespaceURI -cne 'http://schemas.openxmlformats.org/package/2006/content-types') {
                    continue
                }
                if ($reader.LocalName -ceq 'Override') {
                    $partName = Normalize-OpcPartName $reader.GetAttribute('PartName')
                    $contentType = [string]$reader.GetAttribute('ContentType')
                    if (
                        $null -eq $partName -or
                        [string]::IsNullOrWhiteSpace($contentType) -or
                        $contentType.Length -gt 256 -or
                        $overrides.ContainsKey($partName)
                    ) {
                        $valid = $false
                    } else {
                        $overrides[$partName] = $contentType
                    }
                } elseif ($reader.LocalName -ceq 'Default') {
                    $extension = [string]$reader.GetAttribute('Extension')
                    $contentType = [string]$reader.GetAttribute('ContentType')
                    if (
                        [string]::IsNullOrWhiteSpace($extension) -or
                        $extension.Length -gt 64 -or
                        $extension.IndexOf('.') -ge 0 -or
                        [string]::IsNullOrWhiteSpace($contentType) -or
                        $contentType.Length -gt 256 -or
                        $defaults.ContainsKey($extension)
                    ) {
                        $valid = $false
                    } else {
                        $defaults[$extension.ToLowerInvariant()] = $contentType
                    }
                }
                if ($overrides.Count -gt 4096 -or $defaults.Count -gt 256) {
                    $valid = $false
                    break
                }
            }
        }
        finally {
            $reader.Dispose()
            $stream.Dispose()
        }
    }
    catch {
        $valid = $false
    }
    return [ordered]@{
        valid = $valid -and $rootSeen
        overrides = $overrides
        defaults = $defaults
    }
}

function Get-OpcContentType {
    param(
        [object]$ContentTypes,
        [string]$PartName
    )

    if ($ContentTypes.overrides.ContainsKey($PartName)) {
        return [string]$ContentTypes.overrides[$PartName]
    }
    $fileName = $PartName.Substring($PartName.LastIndexOf('/') + 1)
    $dot = $fileName.LastIndexOf('.')
    if ($dot -lt 0 -or $dot -eq $fileName.Length - 1) { return $null }
    $extension = $fileName.Substring($dot + 1).ToLowerInvariant()
    if ($ContentTypes.defaults.ContainsKey($extension)) {
        return [string]$ContentTypes.defaults[$extension]
    }
    return $null
}

function Read-OpcRelationships {
    param([AllowNull()][IO.Compression.ZipArchiveEntry]$Entry)

    $relationships = New-Object System.Collections.ArrayList
    if ($null -eq $Entry) {
        return [ordered]@{ valid = $true; relationships = @() }
    }
    $valid = $true
    $rootSeen = $false
    $ids = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    try {
        $bytes = Read-BoundedZipEntryBytes $Entry 1048576
        $stream = New-Object IO.MemoryStream(, $bytes)
        $reader = [Xml.XmlReader]::Create($stream, (New-SafeXmlReaderSettings))
        try {
            while ($reader.Read()) {
                if ($reader.NodeType -ne [Xml.XmlNodeType]::Element) { continue }
                if ($reader.Depth -eq 0) {
                    if (
                        $reader.LocalName -cne 'Relationships' -or
                        $reader.NamespaceURI -cne 'http://schemas.openxmlformats.org/package/2006/relationships'
                    ) {
                        $valid = $false
                    }
                    $rootSeen = $true
                    continue
                }
                if (
                    $reader.LocalName -cne 'Relationship' -or
                    $reader.NamespaceURI -cne 'http://schemas.openxmlformats.org/package/2006/relationships'
                ) {
                    continue
                }
                $id = [string]$reader.GetAttribute('Id')
                $type = [string]$reader.GetAttribute('Type')
                $target = [string]$reader.GetAttribute('Target')
                $targetMode = [string]$reader.GetAttribute('TargetMode')
                if (
                    [string]::IsNullOrWhiteSpace($id) -or $id.Length -gt 256 -or
                    [string]::IsNullOrWhiteSpace($type) -or $type.Length -gt 512 -or
                    [string]::IsNullOrWhiteSpace($target) -or $target.Length -gt 2048 -or
                    -not $ids.Add($id)
                ) {
                    $valid = $false
                    continue
                }
                [void]$relationships.Add([ordered]@{
                    id = $id
                    type = $type
                    target = $target
                    targetMode = $targetMode
                })
                if ($relationships.Count -gt 512) {
                    $valid = $false
                    break
                }
            }
        }
        finally {
            $reader.Dispose()
            $stream.Dispose()
        }
    }
    catch {
        $valid = $false
    }
    return [ordered]@{
        valid = $valid -and $rootSeen
        relationships = @($relationships)
    }
}

function Test-OpcSignatureXmlPart {
    param([IO.Compression.ZipArchiveEntry]$Entry)

    $rootSeen = $false
    $validRoot = $false
    try {
        $bytes = Read-BoundedZipEntryBytes $Entry 1048576
        $stream = New-Object IO.MemoryStream(, $bytes)
        $reader = [Xml.XmlReader]::Create($stream, (New-SafeXmlReaderSettings))
        try {
            while ($reader.Read()) {
                if ($reader.NodeType -eq [Xml.XmlNodeType]::Element -and $reader.Depth -eq 0) {
                    if ($rootSeen) { return $false }
                    $rootSeen = $true
                    $validRoot =
                        $reader.LocalName -ceq 'Signature' -and
                        $reader.NamespaceURI -ceq 'http://www.w3.org/2000/09/xmldsig#'
                }
            }
        }
        finally {
            $reader.Dispose()
            $stream.Dispose()
        }
    }
    catch {
        return $false
    }
    return $rootSeen -and $validRoot
}

function Get-OpcPackageSignatureInventory {
    param(
        [hashtable]$PartEntries,
        [System.Collections.Generic.HashSet[string]]$DuplicateParts,
        [bool]$ArchiveAmbiguous
    )

    $originRelationshipType = 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin'
    $signatureRelationshipType = 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature'
    $originContentType = 'application/vnd.openxmlformats-package.digital-signature-origin'
    $signatureContentType = 'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml'
    $ambiguous = $ArchiveAmbiguous -or $DuplicateParts.Count -gt 0

    $contentTypesEntry = if ($PartEntries.ContainsKey('/[Content_Types].xml')) {
        $PartEntries['/[Content_Types].xml']
    } else { $null }
    $contentTypes = Read-OpcContentTypes $contentTypesEntry
    if (-not $contentTypes.valid) { $ambiguous = $true }

    $signatureArtifacts = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($partName in $PartEntries.Keys) {
        $partContentType = Get-OpcContentType $contentTypes $partName
        if (
            $partContentType -ceq $originContentType -or
            $partContentType -ceq $signatureContentType -or
            $partName -match '(?i)(^|/)_xmlsignatures/(origin\.sigs|sig[^/]*\.xml)$'
        ) {
            [void]$signatureArtifacts.Add($partName)
        }
    }

    $rootRelationshipsEntry = if ($PartEntries.ContainsKey('/_rels/.rels')) {
        $PartEntries['/_rels/.rels']
    } else { $null }
    $rootRelationships = Read-OpcRelationships $rootRelationshipsEntry
    if (-not $rootRelationships.valid -and $null -ne $rootRelationshipsEntry) { $ambiguous = $true }
    $originRelationships = @(
        $rootRelationships.relationships | Where-Object { $_.type -ceq $originRelationshipType }
    )
    if ($originRelationships.Count -eq 0) {
        if ($signatureArtifacts.Count -gt 0 -or $ambiguous) {
            return [ordered]@{ status = 'unknown'; verificationStatus = 'unverifiable'; present = $false }
        }
        return [ordered]@{ status = 'absent'; verificationStatus = 'notPresent'; present = $false }
    }
    if ($originRelationships.Count -ne 1) { $ambiguous = $true }
    $originRelationship = $originRelationships[0]
    if (
        -not [string]::IsNullOrWhiteSpace($originRelationship.targetMode) -and
        $originRelationship.targetMode -ine 'Internal'
    ) { $ambiguous = $true }
    $originPartName = Resolve-OpcRelationshipTarget '/' $originRelationship.target
    if (
        $null -eq $originPartName -or
        -not $PartEntries.ContainsKey($originPartName) -or
        $DuplicateParts.Contains($originPartName) -or
        (Get-OpcContentType $contentTypes $originPartName) -cne $originContentType
    ) {
        $ambiguous = $true
    }
    if ($ambiguous) {
        return [ordered]@{ status = 'unknown'; verificationStatus = 'unverifiable'; present = $false }
    }

    $originRelationshipsPartName = Get-OpcRelationshipPartName $originPartName
    $originRelationshipsEntry = if (
        $null -ne $originRelationshipsPartName -and
        $PartEntries.ContainsKey($originRelationshipsPartName)
    ) { $PartEntries[$originRelationshipsPartName] } else { $null }
    $signatureRelationships = Read-OpcRelationships $originRelationshipsEntry
    if (-not $signatureRelationships.valid -and $null -ne $originRelationshipsEntry) { $ambiguous = $true }
    $signatureLinks = @(
        $signatureRelationships.relationships | Where-Object { $_.type -ceq $signatureRelationshipType }
    )
    if ($signatureLinks.Count -gt 128) { $ambiguous = $true }
    $referencedSignatureParts = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($signatureLink in $signatureLinks) {
        if (
            -not [string]::IsNullOrWhiteSpace($signatureLink.targetMode) -and
            $signatureLink.targetMode -ine 'Internal'
        ) {
            $ambiguous = $true
            continue
        }
        $signaturePartName = Resolve-OpcRelationshipTarget $originPartName $signatureLink.target
        if (
            $null -eq $signaturePartName -or
            -not $PartEntries.ContainsKey($signaturePartName) -or
            $DuplicateParts.Contains($signaturePartName) -or
            (Get-OpcContentType $contentTypes $signaturePartName) -cne $signatureContentType -or
            -not (Test-OpcSignatureXmlPart $PartEntries[$signaturePartName]) -or
            -not $referencedSignatureParts.Add($signaturePartName)
        ) {
            $ambiguous = $true
        }
    }

    foreach ($artifact in $signatureArtifacts) {
        if ($artifact -ieq $originPartName) { continue }
        if ((Get-OpcContentType $contentTypes $artifact) -ceq $signatureContentType) {
            if (-not $referencedSignatureParts.Contains($artifact)) { $ambiguous = $true }
        } elseif ($artifact -match '(?i)(^|/)_xmlsignatures/sig[^/]*\.xml$') {
            $ambiguous = $true
        }
    }
    if ($ambiguous) {
        return [ordered]@{ status = 'unknown'; verificationStatus = 'unverifiable'; present = $false }
    }
    if ($referencedSignatureParts.Count -eq 0) {
        return [ordered]@{ status = 'absent'; verificationStatus = 'notPresent'; present = $false }
    }
    return [ordered]@{ status = 'present'; verificationStatus = 'verified'; present = $true }
}

function Get-ZipInventory {
    param([string]$Path)

    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        if ($archive.Entries.Count -gt 20000) {
            throw 'The package contains too many parts to inspect safely.'
        }
        $hasVba = $false
        $hasVbaSignature = $false
		$hasContentTypes = $false
		$hasWorkbookXml = $false
		$hasWorkbookBinary = $false
        $archiveAmbiguous = $false
        $partEntries = @{}
        $duplicateParts = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        $customEntry = $null
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName.EndsWith('/', [StringComparison]::Ordinal)) { continue }
            if ($entry.FullName.IndexOf('\\') -ge 0 -or $entry.FullName.StartsWith('/', [StringComparison]::Ordinal)) {
                $archiveAmbiguous = $true
                continue
            }
            $partName = Normalize-OpcPartName ('/' + $entry.FullName)
            if ($null -eq $partName) {
                $archiveAmbiguous = $true
                continue
            }
            if ($partEntries.ContainsKey($partName)) {
                [void]$duplicateParts.Add($partName)
                $archiveAmbiguous = $true
                continue
            }
            $partEntries[$partName] = $entry
            $lowerPartName = $partName.ToLowerInvariant()
			if ($lowerPartName -eq '/[content_types].xml') { $hasContentTypes = $true }
			if ($lowerPartName -eq '/xl/workbook.xml') { $hasWorkbookXml = $true }
			if ($lowerPartName -eq '/xl/workbook.bin') { $hasWorkbookBinary = $true }
            if ($lowerPartName -eq '/xl/vbaproject.bin') { $hasVba = $true }
            if ($lowerPartName -match '^/xl/vbaprojectsignature(?:agile|v3)?\.bin$') {
                $hasVbaSignature = $true
            }
            if ($lowerPartName -eq '/docprops/custom.xml') { $customEntry = $entry }
        }

        $packageSignature = Get-OpcPackageSignatureInventory $partEntries $duplicateParts $archiveAmbiguous

        $labelIds = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        if ($null -ne $customEntry) {
            $xmlBytes = Read-BoundedZipEntryBytes $customEntry 1048576
            $xmlStream = New-Object IO.MemoryStream(, $xmlBytes)
            $xmlReader = [Xml.XmlReader]::Create($xmlStream, (New-SafeXmlReaderSettings))
            try {
                while ($xmlReader.Read()) {
                    if ($xmlReader.NodeType -eq [Xml.XmlNodeType]::Element -and $xmlReader.LocalName -eq 'property') {
                        $propertyName = $xmlReader.GetAttribute('name')
                        if ($propertyName -match '(?i)^MSIP_Label_([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})_Enabled(?:V2)?$') {
                            $labelId = $matches[1].ToLowerInvariant()
                            $propertyReader = $xmlReader.ReadSubtree()
                            $enabledValue = $null
                            try {
                                while ($propertyReader.Read()) {
                                    if (
                                        $propertyReader.NodeType -eq [Xml.XmlNodeType]::Text -or
                                        $propertyReader.NodeType -eq [Xml.XmlNodeType]::CDATA
                                    ) {
                                        $enabledValue = $propertyReader.Value.Trim()
                                    }
                                }
                            }
                            finally {
                                $propertyReader.Dispose()
                            }
                            if ($enabledValue -ieq 'true') {
                                [void]$labelIds.Add($labelId)
                            }
                        }
                    }
                }
            }
            finally {
                $xmlReader.Dispose()
                $xmlStream.Dispose()
            }
        }
        return [ordered]@{
            hasVbaProject = $hasVba
            hasVbaSignature = $hasVbaSignature
            hasPackageSignature = [bool]$packageSignature.present
            vbaSignatureStatus = if ($hasVbaSignature) { 'present' } else { 'absent' }
            packageSignatureStatus = [string]$packageSignature.status
            packageSignatureVerificationStatus = [string]$packageSignature.verificationStatus
			officePackageEncrypted = $false
			hasWorkbookPart = $hasContentTypes -and ($hasWorkbookXml -or $hasWorkbookBinary)
			hasWorkbookXml = $hasWorkbookXml
			hasWorkbookBinary = $hasWorkbookBinary
            vbaProjectProtectionStatus = if ($hasVba) { 'unknown' } else { 'absent' }
            sensitivityLabelIds = @($labelIds | Sort-Object)
        }
    }
    finally {
        $archive.Dispose()
    }
}

function Convert-RegistryValue {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [byte[]]) {
        $length = [Math]::Min($Value.Length, 256)
        $boundedBytes = New-Object byte[] $length
        if ($length -gt 0) { [Array]::Copy($Value, $boundedBytes, $length) }
        return [Convert]::ToBase64String($boundedBytes)
    }
    if ($Value -is [string[]]) {
        return (@($Value | Select-Object -First 16 | ForEach-Object {
            if ($_.Length -gt 256) { $_.Substring(0, 256) } else { $_ }
        }) -join [char]31)
    }
    if ($Value -is [string] -and $Value.Length -gt 2048) { return $Value.Substring(0, 2048) }
    return $Value
}

function Open-RegistryRoot {
    param(
        [string]$Hive,
        [Microsoft.Win32.RegistryView]$View
    )

    $baseHive = if ($Hive -eq 'HKLM') {
        [Microsoft.Win32.RegistryHive]::LocalMachine
    } else {
        [Microsoft.Win32.RegistryHive]::CurrentUser
    }
    return [Microsoft.Win32.RegistryKey]::OpenBaseKey($baseHive, $View)
}

function Add-RegistrySetting {
    param(
        [System.Collections.IList]$Target,
        [string]$Hive,
        [Microsoft.Win32.RegistryView]$View,
        [string]$KeyPath,
        [string]$Name,
        [string]$Id,
        [string]$Category,
        [string]$Source,
        [bool]$Managed
    )

    $root = Open-RegistryRoot $Hive $View
    $key = $null
    try {
        $key = $root.OpenSubKey($KeyPath, $false)
        if ($null -eq $key) { return }
        $sentinel = New-Object object
        $value = $key.GetValue(
            $Name,
            $sentinel,
            [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames
        )
        if ([object]::ReferenceEquals($value, $sentinel)) { return }
        $kind = $key.GetValueKind($Name).ToString()
        [void]$Target.Add([ordered]@{
            id = $Id
            category = $Category
            source = $Source
            managed = $Managed
            registryPath = "$Hive\$KeyPath"
            name = $Name
            value = Convert-RegistryValue $value
            valueKind = $kind
            registryView = if ($View -eq [Microsoft.Win32.RegistryView]::Registry64) { '64' } else { '32' }
        })
    }
    finally {
        if ($null -ne $key) { $key.Dispose() }
        $root.Dispose()
    }
}

function Test-RegistryKeyPresent {
    param(
        [string]$Hive,
        [Microsoft.Win32.RegistryView]$View,
        [string]$KeyPath
    )

    $root = Open-RegistryRoot $Hive $View
    $key = $null
    try {
        $key = $root.OpenSubKey($KeyPath, $false)
        return $null -ne $key
    }
    finally {
        if ($null -ne $key) { $key.Dispose() }
        $root.Dispose()
    }
}

function Add-TrustedLocations {
    param(
        [System.Collections.IList]$Target,
        [string]$Hive,
        [Microsoft.Win32.RegistryView]$View,
        [string]$KeyPath,
        [string]$Source,
        [bool]$Managed,
        [System.Collections.Generic.HashSet[string]]$Seen
    )

    if ($Target.Count -ge $script:MaxTrustedLocations) { return }
    $root = Open-RegistryRoot $Hive $View
    $locationsKey = $null
    try {
        $locationsKey = $root.OpenSubKey($KeyPath, $false)
        if ($null -eq $locationsKey) { return }
        for ($locationIndex = 0; $locationIndex -lt 64; $locationIndex++) {
            if ($Target.Count -ge $script:MaxTrustedLocations) { break }
            $locationName = "Location$locationIndex"
            $locationKey = $locationsKey.OpenSubKey($locationName, $false)
            try {
                if ($null -eq $locationKey) { continue }
                $sentinel = New-Object object
                $locationValue = $locationKey.GetValue(
                    'Path',
                    $sentinel,
                    [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames
                )
                if ([object]::ReferenceEquals($locationValue, $sentinel)) { continue }
                $locationPath = [string]$locationValue
                if ([string]::IsNullOrWhiteSpace($locationPath)) { continue }
                $allowSubfolders = $false
                $allowValue = $locationKey.GetValue('AllowSubfolders', $sentinel)
                if (-not [object]::ReferenceEquals($allowValue, $sentinel)) {
                    $allowSubfolders = [int64]$allowValue -ne 0
                }
                $registryPath = "$Hive\$KeyPath\$locationName"
                $deduplicationKey = "$Source|$registryPath|$locationPath|$allowSubfolders".ToLowerInvariant()
                if (-not $Seen.Add($deduplicationKey)) { continue }
                $item = [ordered]@{
                    source = $Source
                    managed = $Managed
                    registryPath = $registryPath
                    path = if ($locationPath.Length -gt 2048) { $locationPath.Substring(0, 2048) } else { $locationPath }
                    allowSubfolders = $allowSubfolders
                    registryView = if ($View -eq [Microsoft.Win32.RegistryView]::Registry64) { '64' } else { '32' }
                }
                $descriptionValue = $locationKey.GetValue('Description', $sentinel)
                if (-not [object]::ReferenceEquals($descriptionValue, $sentinel)) {
                    $description = [string]$descriptionValue
                    $item.description = if ($description.Length -gt 2048) { $description.Substring(0, 2048) } else { $description }
                }
                [void]$Target.Add($item)
            }
            finally {
                if ($null -ne $locationKey) { $locationKey.Dispose() }
            }
        }
    }
    finally {
        if ($null -ne $locationsKey) { $locationsKey.Dispose() }
        $root.Dispose()
    }
}

function Get-OfficeSecurity {
    $settings = New-Object System.Collections.ArrayList
    $trustedLocations = New-Object System.Collections.ArrayList
    $seenTrustedLocations = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $sources = @(
        @{ Hive = 'HKCU'; Prefix = 'Software\Policies\Microsoft\Cloud\Office\16.0'; Source = 'cloudPolicy'; Managed = $true },
        @{ Hive = 'HKLM'; Prefix = 'Software\Policies\Microsoft\Office\16.0'; Source = 'machinePolicy'; Managed = $true },
        @{ Hive = 'HKCU'; Prefix = 'Software\Policies\Microsoft\Office\16.0'; Source = 'userPolicy'; Managed = $true },
        @{ Hive = 'HKCU'; Prefix = 'Software\Microsoft\Office\16.0'; Source = 'userPreference'; Managed = $false },
        @{ Hive = 'HKLM'; Prefix = 'Software\Microsoft\Office\16.0'; Source = 'machinePreference'; Managed = $false }
    )
    $definitions = @(
        @{ Suffix = 'Excel\Security'; Name = 'VBAWarnings'; Id = 'vbaWarnings'; Category = 'vba' },
        @{ Suffix = 'Excel\Security'; Name = 'AccessVBOM'; Id = 'accessVbom'; Category = 'vba' },
        @{ Suffix = 'Excel\Security'; Name = 'BlockContentExecutionFromInternet'; Id = 'blockInternetMacros'; Category = 'vba' },
        @{ Suffix = 'Common\Security'; Name = 'DisableAllActiveX'; Id = 'disableAllActiveX'; Category = 'activeX' },
        @{ Suffix = 'Common\Security'; Name = 'UFIControls'; Id = 'ufiControls'; Category = 'activeX' },
        @{ Suffix = 'Excel\Security\ProtectedView'; Name = 'DisableInternetFilesInPV'; Id = 'disableInternetFilesInProtectedView'; Category = 'protectedView' },
        @{ Suffix = 'Excel\Security\ProtectedView'; Name = 'DisableUnsafeLocationsInPV'; Id = 'disableUnsafeLocationsInProtectedView'; Category = 'protectedView' },
        @{ Suffix = 'Excel\Security\ProtectedView'; Name = 'DisableAttachmentsInPV'; Id = 'disableAttachmentsInProtectedView'; Category = 'protectedView' },
        @{ Suffix = 'Excel\Security'; Name = 'DisableAllTrustedLocations'; Id = 'disableAllTrustedLocations'; Category = 'trustedLocations' },
        @{ Suffix = 'Excel\Security\Trusted Locations'; Name = 'AllowNetworkLocations'; Id = 'allowNetworkTrustedLocations'; Category = 'trustedLocations' },
        @{ Suffix = 'Common\Security\Trusted Locations'; Name = 'Allow User Locations'; Id = 'allowUserTrustedLocations'; Category = 'trustedLocations' }
    )
    $views = @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)
    foreach ($source in $sources) {
        foreach ($view in $views) {
            foreach ($definition in $definitions) {
                Add-RegistrySetting `
                    $settings `
                    $source.Hive `
                    $view `
                    "$($source.Prefix)\$($definition.Suffix)" `
                    $definition.Name `
                    $definition.Id `
                    $definition.Category `
                    $source.Source `
                    $source.Managed
            }
            Add-TrustedLocations `
                $trustedLocations `
                $source.Hive `
                $view `
                "$($source.Prefix)\Excel\Security\Trusted Locations" `
                $source.Source `
                $source.Managed `
                $seenTrustedLocations
            Add-TrustedLocations `
                $trustedLocations `
                $source.Hive `
                $view `
                "$($source.Prefix)\Common\Security\Trusted Locations" `
                $source.Source `
                $source.Managed `
                $seenTrustedLocations
        }
    }

    $cloudPolicyDetected = @(
        $settings | Where-Object { $_.source -eq 'cloudPolicy' }
    ).Count -gt 0 -or @(
        $trustedLocations | Where-Object { $_.source -eq 'cloudPolicy' }
    ).Count -gt 0
    $cloudPolicyServiceDetected = $false
    foreach ($view in $views) {
        if (Test-RegistryKeyPresent 'HKCU' $view 'Software\Microsoft\Office\16.0\Common\CloudPolicy') {
            $cloudPolicyServiceDetected = $true
        }
    }

    return [ordered]@{
        version = '16.0'
        settings = @($settings)
        trustedLocations = @($trustedLocations)
        cloudPolicyDetected = $cloudPolicyDetected
        cloudPolicyServiceDetected = $cloudPolicyServiceDetected
    }
}

function Get-WorkbookSecurity {
    param([string]$Path)

    $file = Get-Item -LiteralPath $Path -Force
    if ($file.Length -gt $script:MaxWorkbookBytes) {
        throw 'The workbook exceeds the 512 MiB inspection limit.'
    }
    $stream = New-Object IO.FileStream(
        $Path,
        [IO.FileMode]::Open,
        [IO.FileAccess]::Read,
        [IO.FileShare]::ReadWrite
    )
    try {
        $magic = New-Object byte[] 8
        $read = $stream.Read($magic, 0, $magic.Length)
    }
    finally {
        $stream.Dispose()
    }

    $zipMagic = $read -ge 4 -and $magic[0] -eq 0x50 -and $magic[1] -eq 0x4B -and (
        ($magic[2] -eq 0x03 -and $magic[3] -eq 0x04) -or
        ($magic[2] -eq 0x05 -and $magic[3] -eq 0x06) -or
        ($magic[2] -eq 0x07 -and $magic[3] -eq 0x08)
    )
    $compoundMagic = $read -ge 8 -and (
        [BitConverter]::ToString($magic, 0, 8) -eq 'D0-CF-11-E0-A1-B1-1A-E1'
    )
    if ($zipMagic) {
        $containerKind = 'zip'
        $inventory = Get-ZipInventory $Path
    }
    elseif ($compoundMagic) {
        $containerKind = 'compound'
        $inventory = Get-CompoundInventory $Path
    }
    else {
        $containerKind = 'other'
        $inventory = [ordered]@{
            hasVbaProject = $false
            hasVbaSignature = $false
			hasPackageSignature = $false
			vbaSignatureStatus = 'unknown'
			packageSignatureStatus = 'unknown'
			packageSignatureVerificationStatus = 'unverifiable'
			officePackageEncrypted = $false
			hasWorkbookPart = $false
			hasWorkbookXml = $false
			hasWorkbookBinary = $false
            vbaProjectProtectionStatus = 'absent'
            sensitivityLabelIds = @()
        }
    }

    $extension = $file.Extension.ToLowerInvariant()
    $zipExtensions = @('.xlsx', '.xlsm', '.xlsb', '.xltx', '.xltm', '.xlam')
    $compoundExtensions = @('.xls', '.xlt', '.xla')
	$zipWorkbookCompatible = $containerKind -eq 'zip' -and [bool]$inventory.hasWorkbookPart -and (
		($extension -eq '.xlsb' -and [bool]$inventory.hasWorkbookBinary) -or
		($extension -ne '.xlsb' -and [bool]$inventory.hasWorkbookXml)
	)
	$compatible =
		($zipWorkbookCompatible -and $extension -in $zipExtensions) -or
		($containerKind -eq 'compound' -and $extension -in $compoundExtensions -and [bool]$inventory.hasWorkbookPart) -or
        (
            $containerKind -eq 'compound' -and
            $extension -in $zipExtensions -and
            [bool]$inventory.officePackageEncrypted
        )
    if (-not $compatible -or $containerKind -eq 'other') {
        throw 'The workbook container is incompatible with its Excel extension.'
    }

    $zoneInformation = Get-ZoneInformation $Path

    return [ordered]@{
        path = $Path
        name = $file.Name
        extension = $extension
        sizeBytes = [int64]$file.Length
        readOnly = [bool]$file.IsReadOnly
        efsEncrypted = [bool](($file.Attributes -band [IO.FileAttributes]::Encrypted) -ne 0)
        zoneId = $zoneInformation.zoneId
        zoneStatus = $zoneInformation.status
        containerKind = $containerKind
        hasVbaProject = [bool]$inventory.hasVbaProject
        hasVbaSignature = [bool]$inventory.hasVbaSignature
        hasPackageSignature = [bool]$inventory.hasPackageSignature
        vbaSignatureStatus = [string]$inventory.vbaSignatureStatus
        packageSignatureStatus = [string]$inventory.packageSignatureStatus
        packageSignatureVerificationStatus = [string]$inventory.packageSignatureVerificationStatus
        officePackageEncrypted = [bool]$inventory.officePackageEncrypted
        vbaProjectProtectionStatus = [string]$inventory.vbaProjectProtectionStatus
        sensitivityLabelIds = @($inventory.sensitivityLabelIds)
    }
}

function Write-SingleJson {
    param(
        [object]$Value,
        [int]$ExitCode
    )

    $json = $Value | ConvertTo-Json -Depth 8 -Compress
    if ([Text.Encoding]::UTF8.GetByteCount($json) -gt $script:MaxOutputBytes) {
        $json = ([ordered]@{
            schemaVersion = 1
            error = [ordered]@{
                code = 'output_limit_exceeded'
                message = 'The security inspection result exceeded 256 KiB.'
            }
        } | ConvertTo-Json -Depth 4 -Compress)
        $ExitCode = 1
    }
    [Console]::Out.WriteLine($json)
    exit $ExitCode
}

try {
    $workbookPath = Resolve-SafeWorkbookPath (Decode-Base64Utf8 $WorkbookPathBase64)
    $workbookFile = Get-Item -LiteralPath $workbookPath -Force
    if ($workbookFile.Length -gt $script:MaxWorkbookBytes) {
        throw 'The workbook exceeds the 512 MiB inspection limit.'
    }
    $sha256Before = Get-Sha256 $workbookPath
    $workbookSecurity = Get-WorkbookSecurity $workbookPath
    $officeSecurity = Get-OfficeSecurity
    $sha256After = Get-Sha256 $workbookPath
    if ($sha256After -cne $sha256Before) {
        throw 'The workbook changed during security inspection.'
    }
    $workbookSecurity['sha256'] = $sha256After
    $result = [ordered]@{
        schemaVersion = 1
        inspectedAtUtc = [DateTime]::UtcNow.ToString('o')
        workbook = $workbookSecurity
        office = $officeSecurity
    }
    Write-SingleJson $result 0
}
catch {
    Write-SingleJson ([ordered]@{
        schemaVersion = 1
        error = [ordered]@{
            code = 'inspection_failed'
            message = $_.Exception.Message
        }
    }) 1
}
