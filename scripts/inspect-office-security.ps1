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
$script:MaxSensitivityLabels = 32
$script:MaxCompoundDirectoryEntries = 16384
$script:MaxCompoundHierarchyDepth = 128
$script:MaxCompoundPathChars = 2048

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
        $streamInfo = Get-Item -LiteralPath $Path -Stream 'Zone.Identifier' -ErrorAction Stop
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
        if ([int64]$streamInfo.Length -gt 65536) {
            return [ordered]@{ status = 'unreadable'; zoneId = $null }
        }
        $zoneText = Get-Content `
            -LiteralPath $Path `
            -Stream 'Zone.Identifier' `
            -Raw `
            -Encoding UTF8 `
            -ErrorAction Stop
        $currentSection = ''
        $zoneTransferSeen = $false
        $zoneIdSeen = $false
        $parsedZoneId = 0
        foreach ($line in @($zoneText -split '\r?\n')) {
            if ($line -match '^\s*\[([^\]]+)\]\s*$') {
                $currentSection = $matches[1].Trim()
                if ($currentSection -ieq 'ZoneTransfer') {
                    if ($zoneTransferSeen) {
                        return [ordered]@{ status = 'unreadable'; zoneId = $null }
                    }
                    $zoneTransferSeen = $true
                }
                continue
            }
            if (
                $currentSection -ieq 'ZoneTransfer' -and
                $line -match '^\s*ZoneId\s*=\s*([0-9]+)\s*$'
            ) {
                if (
                    $zoneIdSeen -or
                    -not [int]::TryParse($matches[1], [ref]$parsedZoneId) -or
                    $parsedZoneId -lt 0 -or
                    $parsedZoneId -gt 4
                ) {
                    return [ordered]@{ status = 'unreadable'; zoneId = $null }
                }
                $zoneIdSeen = $true
            }
        }
        if (-not $zoneTransferSeen -or -not $zoneIdSeen) {
            return [ordered]@{ status = 'unreadable'; zoneId = $null }
        }
        return [ordered]@{ status = 'read'; zoneId = $parsedZoneId }
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
        $directoryEntryIndex = 0
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
                if ($directoryEntryIndex -ge $script:MaxCompoundDirectoryEntries) {
                    throw 'The compound file directory contains too many entries.'
                }
                $nameBytes = [BitConverter]::ToUInt16($sector, $entryOffset + 64)
                if ($nameBytes -ge 2 -and $nameBytes -le 64 -and ($nameBytes % 2) -eq 0) {
                    $name = [Text.Encoding]::Unicode.GetString(
                        $sector,
                        $entryOffset,
                        $nameBytes - 2
                    )
                    if (-not [string]::IsNullOrWhiteSpace($name)) {
                        [void]$entries.Add([ordered]@{
                            id = $directoryEntryIndex
                            name = $name
                            objectType = [int]$sector[$entryOffset + 66]
                            leftSiblingId = [BitConverter]::ToInt32($sector, $entryOffset + 68)
                            rightSiblingId = [BitConverter]::ToInt32($sector, $entryOffset + 72)
                            childId = [BitConverter]::ToInt32($sector, $entryOffset + 76)
                        })
                    }
                }
                $directoryEntryIndex++
            }
            $directorySector = $fat[$directorySector]
        }
    }
    finally {
        $reader.Dispose()
        $stream.Dispose()
    }

    $entriesById = @{}
    foreach ($entry in $entries) { $entriesById[[int]$entry.id] = $entry }
    $pathsById = @{}
    $hierarchyValid = $true
    $rootEntry = @($entries | Where-Object { $_.objectType -eq 5 } | Select-Object -First 1)
    if ($rootEntry.Count -eq 1) {
        $visitedHierarchy = New-Object 'System.Collections.Generic.HashSet[int]'
        $storageQueue = New-Object 'System.Collections.Generic.Queue[object]'
        $storageQueue.Enqueue([ordered]@{ entry = $rootEntry[0]; path = ''; depth = 0 })
        while ($storageQueue.Count -gt 0) {
            $storageContext = $storageQueue.Dequeue()
            $childId = [int]$storageContext.entry.childId
            if ($childId -lt 0) { continue }
            $siblingStack = New-Object 'System.Collections.Generic.Stack[int]'
            $siblingStack.Push($childId)
            while ($siblingStack.Count -gt 0) {
                $entryId = $siblingStack.Pop()
                if (-not $entriesById.ContainsKey($entryId)) {
                    $hierarchyValid = $false
                    continue
                }
                if (-not $visitedHierarchy.Add($entryId)) {
                    $hierarchyValid = $false
                    continue
                }
                $childEntry = $entriesById[$entryId]
                $parentPath = [string]$storageContext.path
                $childName = [string]$childEntry.name
                $childDepth = [int]$storageContext.depth + 1
                $childPathLength = $parentPath.Length + 1 + $childName.Length
                if (
                    $childDepth -gt $script:MaxCompoundHierarchyDepth -or
                    $childPathLength -gt $script:MaxCompoundPathChars
                ) {
                    throw 'The compound file directory hierarchy exceeds the inspection limit.'
                }
                if ([int]$childEntry.leftSiblingId -ge 0) { $siblingStack.Push([int]$childEntry.leftSiblingId) }
                if ([int]$childEntry.rightSiblingId -ge 0) { $siblingStack.Push([int]$childEntry.rightSiblingId) }
                $entryPath = if ([string]::IsNullOrEmpty($parentPath)) {
                    "\$childName"
                } else {
                    "$parentPath\$childName"
                }
                $pathsById[$entryId] = $entryPath
                if ([int]$childEntry.objectType -eq 1) {
                    $storageQueue.Enqueue([ordered]@{
                        entry = $childEntry
                        path = $entryPath
                        depth = $childDepth
                    })
                }
            }
        }
    } else {
        $hierarchyValid = $false
    }

    $dataSpacesPath = "\$([char]6)DataSpaces"
    $dataSpaceMapPath = "$dataSpacesPath\DataSpaceMap"
    $transformInfoPath = "$dataSpacesPath\TransformInfo"
    $labelInfoPath = "$dataSpacesPath\TransformInfo\LabelInfo"
    $hasDataSpacesStorage = $false
    $hasDataSpaceMapStream = $false
    $hasTransformInfoStorage = $false
    $hasIrmLicenseStream = $false
    $hasExactLabelInfoStream = $false
    foreach ($entry in $entries) {
        $entryPath = if ($pathsById.ContainsKey([int]$entry.id)) { [string]$pathsById[[int]$entry.id] } else { '' }
        if ($entry.objectType -eq 1 -and $entryPath -ieq $dataSpacesPath) { $hasDataSpacesStorage = $true }
        if ($entry.objectType -eq 2 -and $entryPath -ieq $dataSpaceMapPath) { $hasDataSpaceMapStream = $true }
        if ($entry.objectType -eq 1 -and $entryPath -ieq $transformInfoPath) { $hasTransformInfoStorage = $true }
        if (
            $entry.objectType -eq 2 -and
            $entry.name -like 'EUL-*' -and
            $entryPath.StartsWith("$transformInfoPath\", [StringComparison]::OrdinalIgnoreCase)
        ) {
            $hasIrmLicenseStream = $true
        }
        if ($entry.objectType -eq 2 -and $entryPath -ieq $labelInfoPath) { $hasExactLabelInfoStream = $true }
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
    $irmProtected =
        $hierarchyValid -and
        $hasEncryptedPackage -and
        $hasDataSpacesStorage -and
        $hasDataSpaceMapStream -and
        $hasTransformInfoStorage -and
        $hasIrmLicenseStream
    $officePackageEncrypted = $hasEncryptedPackage -and ($hasEncryptionInfo -or $irmProtected)
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
		irmProtected = $irmProtected
		hasWorkbookPart = $hasWorkbookStream
        vbaProjectProtectionStatus = if ($hasVba) { 'unknown' } else { 'absent' }
        sensitivityLabelIds = @()
        sensitivityLabels = @()
        sensitivityMetadataStatus = 'unknown'
        sensitivityMetadataSource = if (-not $hierarchyValid) {
            'ambiguous'
        } elseif ($hasExactLabelInfoStream) {
            'labelInfoStream'
        } elseif ($officePackageEncrypted) {
            'encryptedContainer'
        } else {
            'unsupported'
        }
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
                        $rootSeen -or
                        $reader.LocalName -cne 'Types' -or
                        $reader.NamespaceURI -cne 'http://schemas.openxmlformats.org/package/2006/content-types'
                    ) {
                        $valid = $false
                    }
                    $rootSeen = $true
                    continue
                }
                if (
                    $reader.Depth -ne 1 -or
                    $reader.NamespaceURI -cne 'http://schemas.openxmlformats.org/package/2006/content-types' -or
                    $reader.LocalName -notin @('Override', 'Default')
                ) {
                    $valid = $false
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
                        $rootSeen -or
                        $reader.LocalName -cne 'Relationships' -or
                        $reader.NamespaceURI -cne 'http://schemas.openxmlformats.org/package/2006/relationships'
                    ) {
                        $valid = $false
                    }
                    $rootSeen = $true
                    continue
                }
                if (
                    $reader.Depth -ne 1 -or
                    $reader.LocalName -cne 'Relationship' -or
                    $reader.NamespaceURI -cne 'http://schemas.openxmlformats.org/package/2006/relationships'
                ) {
                    $valid = $false
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
    return [ordered]@{ status = 'present'; verificationStatus = 'structureVerified'; present = $true }
}

function Convert-OpcBoolean {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) { return $null }
    switch ($Value.Trim().ToLowerInvariant()) {
        'true' { return $true }
        '1' { return $true }
        'false' { return $false }
        '0' { return $false }
        default { return $null }
    }
}

function Convert-GuidText {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $parsed = [Guid]::Empty
    if (-not [Guid]::TryParse($Value, [ref]$parsed)) { return $null }
    return $parsed.ToString('D').ToLowerInvariant()
}

function Read-SensitivityLabelInfoPart {
    param([IO.Compression.ZipArchiveEntry]$Entry)

    $labels = New-Object System.Collections.ArrayList
    $coveredSiteIds = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $labelIds = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $valid = $true
    $rootSeen = $false
    $extensionsSeen = $false
    $labelElementCount = 0
    try {
        $bytes = Read-BoundedZipEntryBytes $Entry 1048576
        $stream = New-Object IO.MemoryStream(, $bytes)
        $reader = [Xml.XmlReader]::Create($stream, (New-SafeXmlReaderSettings))
        try {
            while ($reader.Read()) {
                if ($reader.NodeType -ne [Xml.XmlNodeType]::Element) { continue }
                if ($reader.Depth -eq 0) {
                    if (
                        $rootSeen -or
                        $reader.LocalName -cne 'labelList' -or
                        $reader.NamespaceURI -cne 'http://schemas.microsoft.com/office/2020/mipLabelMetadata'
                    ) {
                        $valid = $false
                    }
                    $rootSeen = $true
                    continue
                }
                if (
                    $reader.Depth -eq 1 -and
                    $reader.LocalName -ceq 'extLst' -and
                    $reader.NamespaceURI -ceq 'http://schemas.microsoft.com/office/2020/mipLabelMetadata'
                ) {
                    if ($extensionsSeen) { $valid = $false }
                    $extensionsSeen = $true
                    continue
                }
                if ($reader.Depth -gt 1 -and $extensionsSeen) { continue }
                if (
                    $reader.Depth -ne 1 -or
                    $reader.LocalName -cne 'label' -or
                    $reader.NamespaceURI -cne 'http://schemas.microsoft.com/office/2020/mipLabelMetadata' -or
                    $extensionsSeen
                ) {
                    $valid = $false
                    continue
                }
                $labelElementCount++
                if ($labelElementCount -gt $script:MaxSensitivityLabels) {
                    $valid = $false
                    continue
                }
                $id = Convert-GuidText $reader.GetAttribute('id')
                $siteId = Convert-GuidText $reader.GetAttribute('siteId')
                $enabled = Convert-OpcBoolean $reader.GetAttribute('enabled')
                $removed = Convert-OpcBoolean $reader.GetAttribute('removed')
                $methodAttribute = $reader.GetAttribute('method')
                $method = [string]$methodAttribute
                $contentBitsText = $reader.GetAttribute('contentBits')
                $contentBits = $null
                if ($null -ne $contentBitsText) {
                    $parsedBits = [uint32]0
                    if ([uint32]::TryParse($contentBitsText, [ref]$parsedBits)) {
                        $contentBits = [int64]$parsedBits
                    } else {
                        $valid = $false
                    }
                }
                if (
                    $null -eq $id -or
                    $null -eq $siteId -or
                    $null -eq $enabled -or
                    $null -eq $removed -or
                    $null -eq $methodAttribute -or
                    $method.Length -gt 128 -or
                    ([bool]$enabled -and -not [bool]$removed -and $method -cnotin @('Standard', 'Privileged')) -or
                    ([bool]$removed -and $method.Length -ne 0) -or
                    -not $labelIds.Add($id) -or
                    -not $coveredSiteIds.Add($siteId)
                ) {
                    $valid = $false
                    continue
                }
                if ([bool]$enabled -and -not [bool]$removed) {
                    [void]$labels.Add([ordered]@{
                        id = $id
                        enabled = $true
                        removed = $false
                        name = $null
                        method = $method
                        setDate = $null
                        contentBits = $contentBits
                        siteId = $siteId
                        source = 'labelInfoPart'
                        confidence = 'localDeclaration'
                    })
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
    if (-not $valid -or -not $rootSeen) {
        return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @(); coveredSiteIds = @() }
    }
    return [ordered]@{
        status = if ($labels.Count -gt 0) { 'present' } else { 'absent' }
        source = 'labelInfoPart'
        labels = @($labels)
        coveredSiteIds = @($coveredSiteIds)
    }
}

function Read-CustomSensitivityLabelsPart {
    param([IO.Compression.ZipArchiveEntry]$Entry)

    $metadata = @{}
    $propertyNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $propertyIds = New-Object 'System.Collections.Generic.HashSet[int]'
    $sensitivityId = $null
    $relevantSeen = $false
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
                        $rootSeen -or
                        $reader.LocalName -cne 'Properties' -or
                        $reader.NamespaceURI -cne 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
                    ) {
                        $valid = $false
                    }
                    $rootSeen = $true
                    continue
                }
                if (
                    $reader.Depth -ne 1 -or
                    $reader.LocalName -cne 'property' -or
                    $reader.NamespaceURI -cne 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
                ) {
                    continue
                }
                $propertyName = [string]$reader.GetAttribute('name')
                $isSensitivity = $propertyName -ieq 'Sensitivity'
                $mipPropertyMatch = [regex]::Match(
                    $propertyName,
                    '^MSIP_Label_([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})_(Enabled(?:V2)?|Name|Method|SetDate|ContentBits|SiteId|ActionId)$',
                    [Text.RegularExpressions.RegexOptions]::IgnoreCase
                )
                $isMipProperty = $mipPropertyMatch.Success
                $unknownEnabledVersionMatch = [regex]::Match(
                    $propertyName,
                    '^MSIP_Label_[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}_EnabledV([0-9]+)$',
                    [Text.RegularExpressions.RegexOptions]::IgnoreCase
                )
                $isUnknownEnabledVersion = $unknownEnabledVersionMatch.Success -and $unknownEnabledVersionMatch.Groups[1].Value -ne '2'
                if ($isUnknownEnabledVersion) {
                    $relevantSeen = $true
                    $valid = $false
                    continue
                }
                if (-not $isSensitivity -and -not $isMipProperty) { continue }
                $relevantSeen = $true
                $propertyId = 0
                if (
                    $reader.GetAttribute('fmtid') -ine '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}' -or
                    -not $propertyNames.Add($propertyName) -or
                    -not [int]::TryParse([string]$reader.GetAttribute('pid'), [ref]$propertyId) -or
                    $propertyId -lt 2 -or
                    -not $propertyIds.Add($propertyId)
                ) {
                    $valid = $false
                    continue
                }
                $propertyReader = $reader.ReadSubtree()
                $propertyValueBuilder = New-Object Text.StringBuilder
                $valueElementSeen = $false
                $valueTextSeen = $false
                try {
                    while ($propertyReader.Read()) {
                        if ($propertyReader.NodeType -eq [Xml.XmlNodeType]::Element -and $propertyReader.Depth -eq 1) {
                            if (
                                $valueElementSeen -or
                                $propertyReader.LocalName -cne 'lpwstr' -or
                                $propertyReader.NamespaceURI -cne 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
                            ) {
                                $valid = $false
                            }
                            $valueElementSeen = $true
                        }
                        if ($propertyReader.NodeType -eq [Xml.XmlNodeType]::Element -and $propertyReader.Depth -gt 1) {
                            $valid = $false
                        }
                        if (
                            $propertyReader.NodeType -eq [Xml.XmlNodeType]::Text -or
                            $propertyReader.NodeType -eq [Xml.XmlNodeType]::CDATA
                        ) {
                            if ($propertyReader.Depth -ne 2) { $valid = $false }
                            $valueTextSeen = $true
                            if ($propertyValueBuilder.Length + $propertyReader.Value.Length -gt 512) {
                                $valid = $false
                            } else {
                                [void]$propertyValueBuilder.Append($propertyReader.Value)
                            }
                        }
                    }
                }
                finally {
                    $propertyReader.Dispose()
                }
                $propertyValue = $propertyValueBuilder.ToString().Trim()
                if (-not $valueElementSeen -or -not $valueTextSeen -or $propertyValue.Length -gt 512) {
                    $valid = $false
                    continue
                }
                if ($isSensitivity) {
                    $sensitivityId = Convert-GuidText $propertyValue
                    if ($null -eq $sensitivityId) { $valid = $false }
                    continue
                }
                $labelId = $mipPropertyMatch.Groups[1].Value.ToLowerInvariant()
                $metadataName = $mipPropertyMatch.Groups[2].Value
                if (-not $metadata.ContainsKey($labelId)) {
                    $metadata[$labelId] = [ordered]@{
                        id = $labelId
                        enabledPresent = $false
                        enabled = $null
                        enabledV2Present = $false
                        enabledV2 = $null
                        name = $null
                        method = $null
                        setDate = $null
                        contentBits = $null
                        siteId = $null
                    }
                }
                $item = $metadata[$labelId]
                switch -Regex ($metadataName) {
                    '(?i)^Enabled$' {
                        $item.enabledPresent = $true
                        $item.enabled = Convert-OpcBoolean $propertyValue
                        if ($null -eq $item.enabled) { $valid = $false }
                    }
                    '(?i)^EnabledV2$' {
                        $item.enabledV2Present = $true
                        $item.enabledV2 = $propertyValue.ToLowerInvariant()
                        if ($item.enabledV2 -notin @('true', 'false', 'condition')) { $valid = $false }
                    }
                    '(?i)^Name$' { $item.name = $propertyValue }
                    '(?i)^Method$' { $item.method = $propertyValue }
                    '(?i)^SetDate$' { $item.setDate = $propertyValue }
                    '(?i)^ContentBits$' {
                        $parsedBits = [uint32]0
                        if ([uint32]::TryParse($propertyValue, [ref]$parsedBits)) {
                            $item.contentBits = [int64]$parsedBits
                        } else {
                            $valid = $false
                        }
                    }
                    '(?i)^SiteId$' {
                        $item.siteId = Convert-GuidText $propertyValue
                        if ($null -eq $item.siteId) { $valid = $false }
                    }
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
    if (-not $valid -or -not $rootSeen) {
        return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @(); coveredSiteIds = @() }
    }
    if (-not $relevantSeen) {
        return [ordered]@{ status = 'absent'; source = 'customProperties'; labels = @(); coveredSiteIds = @() }
    }
    if ($null -eq $sensitivityId -or -not $metadata.ContainsKey($sensitivityId)) {
        return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @(); coveredSiteIds = @() }
    }
    $activeLabels = New-Object System.Collections.ArrayList
    $activeSites = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $activeIds = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $metadata.Values) {
        $enabledState = if ([bool]$item.enabledV2Present) {
            [string]$item.enabledV2
        } elseif ([bool]$item.enabledPresent -and [bool]$item.enabled) {
            'true'
        } else {
            'false'
        }
        if ($enabledState -eq 'condition') {
            return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @(); coveredSiteIds = @() }
        }
        if ($enabledState -ne 'true') { continue }
        if (
            $null -eq $item.siteId -or
            -not $activeSites.Add([string]$item.siteId) -or
            -not $activeIds.Add([string]$item.id)
        ) {
            return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @(); coveredSiteIds = @() }
        }
        [void]$activeLabels.Add([ordered]@{
            id = [string]$item.id
            enabled = $true
            removed = $false
            name = if ([string]::IsNullOrWhiteSpace([string]$item.name)) { $null } else { [string]$item.name }
            method = if ([string]::IsNullOrWhiteSpace([string]$item.method)) { $null } else { [string]$item.method }
            setDate = if ([string]::IsNullOrWhiteSpace([string]$item.setDate)) { $null } else { [string]$item.setDate }
            contentBits = $item.contentBits
            siteId = [string]$item.siteId
            source = 'customProperties'
            confidence = 'localDeclaration'
        })
    }
    if ($activeLabels.Count -gt $script:MaxSensitivityLabels -or -not $activeIds.Contains($sensitivityId)) {
        return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @(); coveredSiteIds = @() }
    }
    return [ordered]@{
        status = if ($activeLabels.Count -gt 0) { 'present' } else { 'absent' }
        source = 'customProperties'
        labels = @($activeLabels)
        coveredSiteIds = @($activeSites)
    }
}

function Get-OpcSensitivityLabelInventory {
    param(
        [hashtable]$PartEntries,
        [System.Collections.Generic.HashSet[string]]$DuplicateParts,
        [bool]$ArchiveAmbiguous
    )

    $classificationRelationshipType = 'http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels'
    $customPropertiesRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
    $customPropertiesContentType = 'application/vnd.openxmlformats-officedocument.custom-properties+xml'
    $contentTypesEntry = if ($PartEntries.ContainsKey('/[Content_Types].xml')) { $PartEntries['/[Content_Types].xml'] } else { $null }
    $contentTypes = Read-OpcContentTypes $contentTypesEntry
    $rootRelationshipsEntry = if ($PartEntries.ContainsKey('/_rels/.rels')) { $PartEntries['/_rels/.rels'] } else { $null }
    $rootRelationships = Read-OpcRelationships $rootRelationshipsEntry
    if ($ArchiveAmbiguous -or $DuplicateParts.Count -gt 0 -or (-not $rootRelationships.valid -and $null -ne $rootRelationshipsEntry)) {
        return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @() }
    }

    $labelRelationships = @($rootRelationships.relationships | Where-Object { $_.type -ceq $classificationRelationshipType })
    $knownLabelArtifacts = @($PartEntries.Keys | Where-Object { $_ -match '(?i)^/docMetadata/LabelInfo\.xml$' })
    $customRelationships = @($rootRelationships.relationships | Where-Object { $_.type -ceq $customPropertiesRelationshipType })
    if ($customRelationships.Count -gt 1) {
        return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @() }
    }
    $customResult = [ordered]@{ status = 'absent'; source = 'none'; labels = @(); coveredSiteIds = @() }
    if ($customRelationships.Count -eq 1) {
        if (-not $contentTypes.valid) {
            return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @() }
        }
        $relationship = $customRelationships[0]
        $partName = Resolve-OpcRelationshipTarget '/' $relationship.target
        if (
            (-not [string]::IsNullOrWhiteSpace($relationship.targetMode) -and $relationship.targetMode -ine 'Internal') -or
            $null -eq $partName -or
            -not $PartEntries.ContainsKey($partName) -or
            $DuplicateParts.Contains($partName) -or
            (Get-OpcContentType $contentTypes $partName) -cne $customPropertiesContentType
        ) {
            return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @() }
        }
        $customResult = Read-CustomSensitivityLabelsPart $PartEntries[$partName]
    } elseif ($PartEntries.ContainsKey('/docProps/custom.xml')) {
        return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @() }
    }

    if ($labelRelationships.Count -eq 0) {
        if ($knownLabelArtifacts.Count -gt 0) {
            return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @() }
        }
        return $customResult
    }
    if ($labelRelationships.Count -ne 1 -or -not $contentTypes.valid) {
        return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @() }
    }
    $labelRelationship = $labelRelationships[0]
    $labelPartName = Resolve-OpcRelationshipTarget '/' $labelRelationship.target
    if (
        (-not [string]::IsNullOrWhiteSpace($labelRelationship.targetMode) -and $labelRelationship.targetMode -ine 'Internal') -or
        $null -eq $labelPartName -or
        -not $PartEntries.ContainsKey($labelPartName) -or
        $DuplicateParts.Contains($labelPartName) -or
        (Get-OpcContentType $contentTypes $labelPartName) -cne 'application/xml'
    ) {
        return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @() }
    }
    $labelInfoResult = Read-SensitivityLabelInfoPart $PartEntries[$labelPartName]
    if ($labelInfoResult.status -eq 'unknown' -or $customResult.status -eq 'unknown') {
        return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @() }
    }

    $coveredTenants = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($siteId in @($labelInfoResult.coveredSiteIds)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$siteId)) { [void]$coveredTenants.Add([string]$siteId) }
    }
    $combinedLabels = New-Object System.Collections.ArrayList
    foreach ($label in @($labelInfoResult.labels)) { [void]$combinedLabels.Add($label) }
    foreach ($legacyLabel in @($customResult.labels)) {
        $legacySiteId = [string]$legacyLabel.siteId
        if ([string]::IsNullOrWhiteSpace($legacySiteId)) {
            return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @() }
        }
        if (-not $coveredTenants.Contains($legacySiteId)) {
            [void]$combinedLabels.Add($legacyLabel)
        }
    }
    if ($combinedLabels.Count -gt $script:MaxSensitivityLabels) {
        return [ordered]@{ status = 'unknown'; source = 'ambiguous'; labels = @() }
    }
    $hasModernLabel = @($combinedLabels | Where-Object { $_.source -eq 'labelInfoPart' }).Count -gt 0
    $hasLegacyLabel = @($combinedLabels | Where-Object { $_.source -eq 'customProperties' }).Count -gt 0
    $source = if ($hasModernLabel -and $hasLegacyLabel) {
        'mixed'
    } elseif ($hasLegacyLabel) {
        'customProperties'
    } else {
        'labelInfoPart'
    }
    return [ordered]@{
        status = if ($combinedLabels.Count -gt 0) { 'present' } else { 'absent' }
        source = $source
        labels = @($combinedLabels)
    }
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
        }

        $packageSignature = Get-OpcPackageSignatureInventory $partEntries $duplicateParts $archiveAmbiguous
        $sensitivityInventory = Get-OpcSensitivityLabelInventory $partEntries $duplicateParts $archiveAmbiguous
        return [ordered]@{
            hasVbaProject = $hasVba
            hasVbaSignature = $hasVbaSignature
            hasPackageSignature = [bool]$packageSignature.present
            vbaSignatureStatus = if ($hasVbaSignature) { 'present' } else { 'absent' }
            packageSignatureStatus = [string]$packageSignature.status
            packageSignatureVerificationStatus = [string]$packageSignature.verificationStatus
			officePackageEncrypted = $false
			irmProtected = $false
			hasWorkbookPart = $hasContentTypes -and ($hasWorkbookXml -or $hasWorkbookBinary)
			hasWorkbookXml = $hasWorkbookXml
			hasWorkbookBinary = $hasWorkbookBinary
            vbaProjectProtectionStatus = if ($hasVba) { 'unknown' } else { 'absent' }
            sensitivityLabelIds = @($sensitivityInventory.labels | ForEach-Object { $_.id })
            sensitivityLabels = @($sensitivityInventory.labels)
            sensitivityMetadataStatus = [string]$sensitivityInventory.status
            sensitivityMetadataSource = [string]$sensitivityInventory.source
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
        [bool]$Managed,
        [bool]$SharedView,
        [ref]$Unreadable
    )

    $root = $null
    $key = $null
    try {
        $root = Open-RegistryRoot $Hive $View
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
            registryView = if ($SharedView) { $null } elseif ($View -eq [Microsoft.Win32.RegistryView]::Registry64) { '64' } else { '32' }
        })
    }
    catch {
        $Unreadable.Value = $true
    }
    finally {
        if ($null -ne $key) { $key.Dispose() }
        if ($null -ne $root) { $root.Dispose() }
    }
}

function Test-RegistryKeyPresent {
    param(
        [string]$Hive,
        [Microsoft.Win32.RegistryView]$View,
        [string]$KeyPath
    )

    $root = $null
    $key = $null
    try {
        $root = Open-RegistryRoot $Hive $View
        $key = $root.OpenSubKey($KeyPath, $false)
        if ($null -ne $key) { return 'detected' }
        return 'notDetected'
    }
    catch {
        return 'unreadable'
    }
    finally {
        if ($null -ne $key) { $key.Dispose() }
        if ($null -ne $root) { $root.Dispose() }
    }
}

function Get-MdmEnrollmentEvidence {
    param([Microsoft.Win32.RegistryView]$View)

    $root = $null
    $enrollmentsKey = $null
    $accountsKey = $null
    try {
        $root = Open-RegistryRoot 'HKLM' $View
        $enrollmentsKey = $root.OpenSubKey('Software\Microsoft\Enrollments', $false)
        $accountsKey = $root.OpenSubKey('Software\Microsoft\Provisioning\OMADM\Accounts', $false)
        if (
            ($null -ne $enrollmentsKey -and $enrollmentsKey.SubKeyCount -gt 64) -or
            ($null -ne $accountsKey -and $accountsKey.SubKeyCount -gt 64)
        ) {
            return [ordered]@{ status = 'unreadable'; provider = 'unknown' }
        }
        $enrollmentNames = @(
            if ($null -ne $enrollmentsKey) { $enrollmentsKey.GetSubKeyNames() }
        )
        $accountNames = @(
            if ($null -ne $accountsKey) { $accountsKey.GetSubKeyNames() }
        )
        if ($enrollmentNames.Count -gt 64) {
            return [ordered]@{ status = 'unreadable'; provider = 'unknown' }
        }
        if ($accountNames.Count -gt 64) {
            return [ordered]@{ status = 'unreadable'; provider = 'unknown' }
        }
        $enrollmentIds = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        $strongEnrollmentIds = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        $intuneCandidateIds = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($enrollmentName in $enrollmentNames) {
            $enrollmentId = Convert-GuidText $enrollmentName
            if ($null -eq $enrollmentId) { continue }
            [void]$enrollmentIds.Add($enrollmentId)
            $enrollmentKey = $null
            try {
                $enrollmentKey = $enrollmentsKey.OpenSubKey($enrollmentName, $false)
                if ($null -eq $enrollmentKey) { continue }
                $sentinel = New-Object object
                $discoveryUrl = $enrollmentKey.GetValue('DiscoveryServiceFullURL', $sentinel, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                $upn = $enrollmentKey.GetValue('UPN', $sentinel, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                $certificateThumbprint = $enrollmentKey.GetValue('DMPCertThumbPrint', $sentinel, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                if (
                    (-not [object]::ReferenceEquals($upn, $sentinel) -and -not [string]::IsNullOrWhiteSpace([string]$upn)) -or
                    (-not [object]::ReferenceEquals($certificateThumbprint, $sentinel) -and -not [string]::IsNullOrWhiteSpace([string]$certificateThumbprint))
                ) {
                    [void]$strongEnrollmentIds.Add($enrollmentId)
                }
                if (-not [object]::ReferenceEquals($discoveryUrl, $sentinel)) {
                    $uri = $null
                    if ([Uri]::TryCreate([string]$discoveryUrl, [UriKind]::Absolute, [ref]$uri)) {
                        $host = $uri.DnsSafeHost.ToLowerInvariant()
                        if ($host -eq 'manage.microsoft.com' -or $host.EndsWith('.manage.microsoft.com')) {
                            [void]$intuneCandidateIds.Add($enrollmentId)
                        }
                    }
                }
            }
            finally {
                if ($null -ne $enrollmentKey) { $enrollmentKey.Dispose() }
            }
        }
        $activeAccountIds = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($accountName in $accountNames) {
            $accountId = Convert-GuidText $accountName
            if ($null -eq $accountId) { continue }
            $accountKey = $null
            try {
                $accountKey = $accountsKey.OpenSubKey($accountName, $false)
                if ($null -eq $accountKey) { continue }
                $sentinel = New-Object object
                $serverId = $accountKey.GetValue('ServerId', $sentinel, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                $certificateReference = $accountKey.GetValue('SslClientCertReference', $sentinel, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                if (
                    (-not [object]::ReferenceEquals($serverId, $sentinel) -and -not [string]::IsNullOrWhiteSpace([string]$serverId)) -or
                    (-not [object]::ReferenceEquals($certificateReference, $sentinel) -and -not [string]::IsNullOrWhiteSpace([string]$certificateReference))
                ) {
                    [void]$activeAccountIds.Add($accountId)
                }
            }
            finally {
                if ($null -ne $accountKey) { $accountKey.Dispose() }
            }
        }
        $detectedIds = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($enrollmentId in $strongEnrollmentIds) { [void]$detectedIds.Add($enrollmentId) }
        foreach ($accountId in $activeAccountIds) {
            if ($enrollmentIds.Contains($accountId)) { [void]$detectedIds.Add($accountId) }
        }
        $microsoftIntune = $false
        foreach ($detectedId in $detectedIds) {
            if ($intuneCandidateIds.Contains($detectedId)) { $microsoftIntune = $true }
        }
        $detected = $detectedIds.Count -gt 0
        return [ordered]@{
            status = if ($detected) { 'detected' } else { 'notDetected' }
            provider = if ($microsoftIntune) { 'microsoftIntune' } elseif ($detected) { 'unknown' } else { 'none' }
        }
    }
    catch {
        return [ordered]@{ status = 'unreadable'; provider = 'unknown' }
    }
    finally {
        if ($null -ne $accountsKey) { $accountsKey.Dispose() }
        if ($null -ne $enrollmentsKey) { $enrollmentsKey.Dispose() }
        if ($null -ne $root) { $root.Dispose() }
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
        [bool]$SharedView,
        [System.Collections.Generic.HashSet[string]]$Seen,
        [ref]$Unreadable
    )

    if ($Target.Count -ge $script:MaxTrustedLocations) {
        $Unreadable.Value = $true
        return
    }
    $root = $null
    $locationsKey = $null
    try {
        $root = Open-RegistryRoot $Hive $View
        $locationsKey = $root.OpenSubKey($KeyPath, $false)
        if ($null -eq $locationsKey) { return }
        if ($locationsKey.SubKeyCount -gt 256) {
            $Unreadable.Value = $true
            return
        }
        $locationNames = @(
            $locationsKey.GetSubKeyNames() |
                Where-Object { $_ -match '^Location(?:0|[1-9][0-9]{0,8})$' } |
                Sort-Object { [int64]$_.Substring(8) }
        )
        if ($locationNames.Count -gt $script:MaxTrustedLocations) {
            $Unreadable.Value = $true
        }
        foreach ($locationName in $locationNames) {
            if ($Target.Count -ge $script:MaxTrustedLocations) {
                $Unreadable.Value = $true
                break
            }
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
                $rawLocationPath = [string]$locationValue
                if (
                    [string]::IsNullOrWhiteSpace($rawLocationPath) -or
                    $rawLocationPath.Length -gt 2048
                ) { continue }
                $locationPath = [Environment]::ExpandEnvironmentVariables($rawLocationPath)
                if (
                    [string]::IsNullOrWhiteSpace($locationPath) -or
                    $locationPath.Length -gt 2048 -or
                    $locationPath -match '%[^%]+%'
                ) {
                    $Unreadable.Value = $true
                    continue
                }
                $allowSubfolders = $false
                $allowValue = $locationKey.GetValue('AllowSubfolders', $sentinel)
                if (-not [object]::ReferenceEquals($allowValue, $sentinel)) {
                    $allowSubfolders = [int64]$allowValue -ne 0
                }
                $registryPath = "$Hive\$KeyPath\$locationName"
                $registryViewName = if ($SharedView) { $null } elseif ($View -eq [Microsoft.Win32.RegistryView]::Registry64) { '64' } else { '32' }
                $deduplicationKey = "$Source|$registryPath|$locationPath|$allowSubfolders|$registryViewName".ToLowerInvariant()
                if (-not $Seen.Add($deduplicationKey)) { continue }
                $item = [ordered]@{
                    source = $Source
                    managed = $Managed
                    registryPath = $registryPath
                    path = $locationPath
                    allowSubfolders = $allowSubfolders
                    registryView = $registryViewName
                }
                $descriptionValue = $locationKey.GetValue('Description', $sentinel)
                if (-not [object]::ReferenceEquals($descriptionValue, $sentinel)) {
                    $description = [string]$descriptionValue
                    $item.description = if ($description.Length -gt 2048) { $description.Substring(0, 2048) } else { $description }
                }
                [void]$Target.Add($item)
            }
            catch {
                $Unreadable.Value = $true
            }
            finally {
                if ($null -ne $locationKey) { $locationKey.Dispose() }
            }
        }
    }
    catch {
        $Unreadable.Value = $true
    }
    finally {
        if ($null -ne $locationsKey) { $locationsKey.Dispose() }
        if ($null -ne $root) { $root.Dispose() }
    }
}

function Add-DefaultTrustedLocations {
    param(
        [System.Collections.IList]$Target,
        [System.Collections.Generic.HashSet[string]]$Seen
    )

    $records = New-Object System.Collections.ArrayList
    $applicationData = [Environment]::GetFolderPath(
        [Environment+SpecialFolder]::ApplicationData
    )
    if (-not [string]::IsNullOrWhiteSpace($applicationData)) {
        [void]$records.Add(@{
            Path = [IO.Path]::Combine($applicationData, 'Microsoft', 'Templates')
            AllowSubfolders = $false
            Description = 'Modèles utilisateur Excel'
            RegistryView = $null
        })
        [void]$records.Add(@{
            Path = [IO.Path]::Combine($applicationData, 'Microsoft', 'Excel', 'XLSTART')
            AllowSubfolders = $false
            Description = 'Démarrage utilisateur Excel'
            RegistryView = $null
        })
    }
    $programRoots = @(
        @{
            Path = [Environment]::GetFolderPath(
                [Environment+SpecialFolder]::ProgramFiles
            )
            RegistryView = '64'
        },
        @{
            Path = [Environment]::GetFolderPath(
                [Environment+SpecialFolder]::ProgramFilesX86
            )
            RegistryView = '32'
        }
    )
    foreach ($programRoot in $programRoots) {
        if ([string]::IsNullOrWhiteSpace([string]$programRoot.Path)) { continue }
        foreach ($definition in @(
            @{ Relative = 'Microsoft Office\Root\Templates'; Description = "Modèles d’application Excel" },
            @{ Relative = 'Microsoft Office\Root\Office16\XLSTART'; Description = 'Démarrage Excel' },
            @{ Relative = 'Microsoft Office\Root\Office16\STARTUP'; Description = 'Démarrage Office' },
            @{ Relative = 'Microsoft Office\Root\Office16\Library'; Description = 'Compléments Excel' }
        )) {
            [void]$records.Add(@{
                Path = [IO.Path]::Combine(
                    [string]$programRoot.Path,
                    [string]$definition.Relative
                )
                AllowSubfolders = $true
                Description = [string]$definition.Description
                RegistryView = [string]$programRoot.RegistryView
            })
        }
    }

    foreach ($record in $records) {
        if ($Target.Count -ge $script:MaxTrustedLocations) { break }
        try {
            $locationPath = [IO.Path]::GetFullPath([string]$record.Path)
        }
        catch {
            continue
        }
        if (
            [string]::IsNullOrWhiteSpace($locationPath) -or
            $locationPath.Length -gt 2048
        ) { continue }
        $registryView = if ([string]::IsNullOrWhiteSpace([string]$record.RegistryView)) {
            $null
        } else {
            [string]$record.RegistryView
        }
        $deduplicationKey = "officeDefault|$locationPath|$($record.AllowSubfolders)|$registryView".ToLowerInvariant()
        if (-not $Seen.Add($deduplicationKey)) { continue }
        [void]$Target.Add([ordered]@{
            source = 'officeDefault'
            managed = $false
            registryPath = 'Excel default trusted location'
            path = $locationPath
            allowSubfolders = [bool]$record.AllowSubfolders
            description = [string]$record.Description
            registryView = $registryView
        })
    }
}

function Get-OfficeSecurity {
    $settings = New-Object System.Collections.ArrayList
    $unreadableSettings = New-Object System.Collections.ArrayList
    $clickToRunArchitectureSignals = New-Object System.Collections.ArrayList
    $outlookArchitectureSignals = New-Object System.Collections.ArrayList
    $trustedLocations = New-Object System.Collections.ArrayList
    $seenTrustedLocations = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $trustedLocationInspectionPartial = $false
    Add-DefaultTrustedLocations $trustedLocations $seenTrustedLocations
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
        @{ Suffix = 'Excel\Security'; Name = 'XL4MacroOff'; Id = 'xl4MacroOff'; Category = 'xlm' },
        @{ Suffix = 'Common\Security'; Name = 'DisableAllActiveX'; Id = 'disableAllActiveX'; Category = 'activeX' },
        @{ Suffix = 'Common\Security'; Name = 'UFIControls'; Id = 'ufiControls'; Category = 'activeX' },
        @{ Suffix = 'Excel\Security\ProtectedView'; Name = 'DisableInternetFilesInPV'; Id = 'disableInternetFilesInProtectedView'; Category = 'protectedView' },
        @{ Suffix = 'Excel\Security\ProtectedView'; Name = 'DisableUnsafeLocationsInPV'; Id = 'disableUnsafeLocationsInProtectedView'; Category = 'protectedView' },
        @{ Suffix = 'Excel\Security\ProtectedView'; Name = 'DisableAttachmentsInPV'; Id = 'disableAttachmentsInProtectedView'; Category = 'protectedView' },
        @{ Suffix = 'Excel\Security\Trusted Locations'; Name = 'AllLocationsDisabled'; Id = 'disableAllTrustedLocations'; Category = 'trustedLocations' },
        @{ Suffix = 'Excel\Security\Trusted Locations'; Name = 'AllowNetworkLocations'; Id = 'allowNetworkTrustedLocations'; Category = 'trustedLocations' },
        @{ Suffix = 'Common\Security\Trusted Locations'; Name = 'Allow User Locations'; Id = 'allowUserTrustedLocations'; Category = 'trustedLocations' }
    )
    $views = @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)
    $architectureRegistryUnreadable = $false
    foreach ($view in $views) {
        Add-RegistrySetting `
            $clickToRunArchitectureSignals `
            'HKLM' `
            $view `
            'Software\Microsoft\Office\ClickToRun\Configuration' `
            'Platform' `
            'officeArchitecture' `
            'installation' `
            'machinePreference' `
            $false `
            $false `
            ([ref]$architectureRegistryUnreadable)
        Add-RegistrySetting `
            $outlookArchitectureSignals `
            'HKLM' `
            $view `
            'Software\Microsoft\Office\16.0\Outlook' `
            'Bitness' `
            'officeArchitecture' `
            'installation' `
            'machinePreference' `
            $false `
            $false `
            ([ref]$architectureRegistryUnreadable)
    }
    $sourceUnreadable = @{}
    foreach ($source in $sources) {
        $readFailure = $false
        $sharedView = $source.Hive -eq 'HKCU'
        $sourceViews = if ($sharedView) { @($views[0]) } else { $views }
        foreach ($view in $sourceViews) {
            foreach ($definition in $definitions) {
                $settingUnreadable = $false
                Add-RegistrySetting `
                    $settings `
                    $source.Hive `
                    $view `
                    "$($source.Prefix)\$($definition.Suffix)" `
                    $definition.Name `
                    $definition.Id `
                    $definition.Category `
                    $source.Source `
                    $source.Managed `
                    $sharedView `
                    ([ref]$settingUnreadable)
                if ($settingUnreadable) {
                    $readFailure = $true
                    [void]$unreadableSettings.Add([ordered]@{
                        id = $definition.Id
                        source = $source.Source
                        registryView = if ($sharedView) { $null } elseif ($view -eq [Microsoft.Win32.RegistryView]::Registry64) { '64' } else { '32' }
                    })
                }
            }
            $locationReadFailure = $false
            Add-TrustedLocations `
                $trustedLocations `
                $source.Hive `
                $view `
                "$($source.Prefix)\Excel\Security\Trusted Locations" `
                $source.Source `
                $source.Managed `
                $sharedView `
                $seenTrustedLocations `
                ([ref]$locationReadFailure)
            Add-TrustedLocations `
                $trustedLocations `
                $source.Hive `
                $view `
                "$($source.Prefix)\Common\Security\Trusted Locations" `
                $source.Source `
                $source.Managed `
                $sharedView `
                $seenTrustedLocations `
                ([ref]$locationReadFailure)
            if ($locationReadFailure) {
                $readFailure = $true
                $trustedLocationInspectionPartial = $true
            }
        }
        $sourceUnreadable[$source.Source] = $readFailure
    }

    $cloudPolicyDetected = @(
        $settings | Where-Object { $_.source -eq 'cloudPolicy' }
    ).Count -gt 0 -or @(
        $trustedLocations | Where-Object { $_.source -eq 'cloudPolicy' }
    ).Count -gt 0
    $cloudPolicyServiceDetected = $false
    $cloudPolicyServiceUnreadable = $false
    $intuneManagementExtensionDetected = $false
    $intuneManagementExtensionUnreadable = $false
    $mdmEnrollmentArtifactsDetected = $false
    $mdmEnrollmentUnreadable = $false
    $mdmProvider = 'none'
    $groupPolicyHistoryDetected = $false
    $groupPolicyHistoryUnreadable = $false
    foreach ($view in $views) {
        if ($view -eq $views[0]) {
            $cloudStatus = Test-RegistryKeyPresent 'HKCU' $view 'Software\Microsoft\Office\16.0\Common\CloudPolicy'
            if ($cloudStatus -eq 'detected') { $cloudPolicyServiceDetected = $true }
            elseif ($cloudStatus -eq 'unreadable') { $cloudPolicyServiceUnreadable = $true }
        }
        foreach ($intunePath in @(
            'Software\Microsoft\IntuneManagementExtension',
            'System\CurrentControlSet\Services\IntuneManagementExtension'
        )) {
            $intuneStatus = Test-RegistryKeyPresent 'HKLM' $view $intunePath
            if ($intuneStatus -eq 'detected') { $intuneManagementExtensionDetected = $true }
            elseif ($intuneStatus -eq 'unreadable') { $intuneManagementExtensionUnreadable = $true }
        }
        $enrollmentEvidence = Get-MdmEnrollmentEvidence $view
        if ($enrollmentEvidence.status -eq 'detected') { $mdmEnrollmentArtifactsDetected = $true }
        elseif ($enrollmentEvidence.status -eq 'unreadable') { $mdmEnrollmentUnreadable = $true }
        if ($enrollmentEvidence.provider -eq 'microsoftIntune') { $mdmProvider = 'microsoftIntune' }
        elseif ($mdmProvider -eq 'none' -and $enrollmentEvidence.provider -eq 'unknown') { $mdmProvider = 'unknown' }
        foreach ($historySignal in @(
            @{ Hive = 'HKLM'; Path = 'Software\Microsoft\Windows\CurrentVersion\Group Policy\History' },
            @{ Hive = 'HKCU'; Path = 'Software\Microsoft\Windows\CurrentVersion\Group Policy\History' }
        )) {
            if ($historySignal.Hive -eq 'HKCU' -and $view -ne $views[0]) { continue }
            $historyStatus = Test-RegistryKeyPresent $historySignal.Hive $view $historySignal.Path
            if ($historyStatus -eq 'detected') { $groupPolicyHistoryDetected = $true }
            elseif ($historyStatus -eq 'unreadable') { $groupPolicyHistoryUnreadable = $true }
        }
    }
    $clickToRunArchitectures = @(
        $clickToRunArchitectureSignals |
            ForEach-Object {
                $value = ([string]$_.value).Trim().ToLowerInvariant()
                if ($value -in @('x64', '64', '64-bit', '64bit')) { 'x64' }
                elseif ($value -in @('x86', '32', '32-bit', '32bit')) { 'x86' }
            } |
            Select-Object -Unique
    )
    $outlookArchitectures = @(
        $outlookArchitectureSignals |
            ForEach-Object {
                $value = ([string]$_.value).Trim().ToLowerInvariant()
                if ($value -in @('x64', '64', '64-bit', '64bit')) { 'x64' }
                elseif ($value -in @('x86', '32', '32-bit', '32bit')) { 'x86' }
            } |
            Select-Object -Unique
    )
    $architecture = if ($clickToRunArchitectures.Count -eq 1) {
        [string]$clickToRunArchitectures[0]
    } elseif ($clickToRunArchitectures.Count -eq 0 -and $outlookArchitectures.Count -eq 1) {
        [string]$outlookArchitectures[0]
    } else {
        'unknown'
    }
    $cloudPolicyDetectionStatus = if ($cloudPolicyDetected) {
        'detected'
    } elseif ([bool]$sourceUnreadable.cloudPolicy) {
        'unknown'
    } else {
        'notDetected'
    }
    $cloudPolicyServiceStatus = if ($cloudPolicyServiceDetected) { 'detected' } elseif ($cloudPolicyServiceUnreadable) { 'unknown' } else { 'notDetected' }
    $intuneManagementExtensionStatus = if ($intuneManagementExtensionDetected) { 'detected' } elseif ($intuneManagementExtensionUnreadable) { 'unknown' } else { 'notDetected' }
    $mdmEnrollmentStatus = if ($mdmEnrollmentArtifactsDetected) { 'detected' } elseif ($mdmEnrollmentUnreadable) { 'unknown' } else { 'notDetected' }
    $groupPolicyHistoryStatus = if ($groupPolicyHistoryDetected) { 'detected' } elseif ($groupPolicyHistoryUnreadable) { 'unknown' } else { 'notDetected' }
    $windowsPolicyRegistryDetected = @(
        $settings | Where-Object { $_.source -in @('machinePolicy', 'userPolicy') }
    ).Count -gt 0 -or @(
        $trustedLocations | Where-Object { $_.source -in @('machinePolicy', 'userPolicy') }
    ).Count -gt 0
    $windowsPolicyRegistryStatus = if ($windowsPolicyRegistryDetected) {
        'detected'
    } elseif ([bool]$sourceUnreadable.machinePolicy -or [bool]$sourceUnreadable.userPolicy) {
        'unknown'
    } else {
        'notDetected'
    }

    return [ordered]@{
        version = '16.0'
        architecture = $architecture
        settings = @($settings)
        unreadableSettings = @($unreadableSettings)
        trustedLocations = @($trustedLocations)
        cloudPolicyDetected = $cloudPolicyDetected
        cloudPolicyServiceDetected = $cloudPolicyServiceDetected
        intuneManagementExtensionDetected = $intuneManagementExtensionDetected
        mdmEnrollmentArtifactsDetected = $mdmEnrollmentArtifactsDetected
        groupPolicyHistoryDetected = $groupPolicyHistoryDetected
        cloudPolicyDetectionStatus = $cloudPolicyDetectionStatus
        cloudPolicyServiceStatus = $cloudPolicyServiceStatus
        windowsPolicyRegistryStatus = $windowsPolicyRegistryStatus
        intuneManagementExtensionStatus = $intuneManagementExtensionStatus
        mdmEnrollmentStatus = $mdmEnrollmentStatus
        mdmProvider = $mdmProvider
        groupPolicyHistoryStatus = $groupPolicyHistoryStatus
        trustedLocationInspectionPartial = $trustedLocationInspectionPartial
        registryInspectionPartial = [bool](
            $architectureRegistryUnreadable -or
            $sourceUnreadable.Values -contains $true -or
            $cloudPolicyServiceUnreadable -or
            $intuneManagementExtensionUnreadable -or
            $mdmEnrollmentUnreadable -or
            $groupPolicyHistoryUnreadable
        )
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
			irmProtected = $false
			hasWorkbookPart = $false
			hasWorkbookXml = $false
			hasWorkbookBinary = $false
            vbaProjectProtectionStatus = 'absent'
            sensitivityLabelIds = @()
            sensitivityLabels = @()
            sensitivityMetadataStatus = 'unknown'
            sensitivityMetadataSource = 'unsupported'
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
        irmProtected = [bool]$inventory.irmProtected
        vbaProjectProtectionStatus = [string]$inventory.vbaProjectProtectionStatus
        sensitivityLabelIds = @($inventory.sensitivityLabelIds)
        sensitivityLabels = @($inventory.sensitivityLabels)
        sensitivityMetadataStatus = [string]$inventory.sensitivityMetadataStatus
        sensitivityMetadataSource = [string]$inventory.sensitivityMetadataSource
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

$inspectionLock = $null
try {
    $workbookPath = Resolve-SafeWorkbookPath (Decode-Base64Utf8 $WorkbookPathBase64)
    $workbookFile = Get-Item -LiteralPath $workbookPath -Force
    if ($workbookFile.Length -gt $script:MaxWorkbookBytes) {
        throw 'The workbook exceeds the 512 MiB inspection limit.'
    }
    $inspectionLock = New-Object IO.FileStream(
        $workbookPath,
        [IO.FileMode]::Open,
        [IO.FileAccess]::Read,
        [IO.FileShare]::Read
    )
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
    $inspectionLock.Dispose()
    $inspectionLock = $null
    Write-SingleJson $result 0
}
catch {
    if ($null -ne $inspectionLock) {
        $inspectionLock.Dispose()
        $inspectionLock = $null
    }
    Write-SingleJson ([ordered]@{
        schemaVersion = 1
        error = [ordered]@{
            code = 'inspection_failed'
            message = $_.Exception.Message
        }
    }) 1
}
