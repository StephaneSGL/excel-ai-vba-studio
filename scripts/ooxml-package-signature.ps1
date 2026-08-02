Set-StrictMode -Version 2.0

$script:OpcRelationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
$script:OpcContentTypesNamespace = 'http://schemas.openxmlformats.org/package/2006/content-types'
$script:OpcOriginRelationshipType = 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin'
$script:OpcSignatureRelationshipType = 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature'
$script:OpcOriginContentType = 'application/vnd.openxmlformats-package.digital-signature-origin'
$script:OpcSignatureContentType = 'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml'
$script:OpcXlmContentTypes = @(
    'application/vnd.ms-excel.macrosheet+xml',
    'application/vnd.ms-excel.intlmacrosheet+xml',
    'application/vnd.ms-excel.macrosheet',
    'application/vnd.ms-excel.intlmacrosheet'
)
$script:OpcXlmRelationshipTypes = @(
    'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet',
    'http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet'
)
$script:OpcMaxEntries = 20000
$script:OpcMaxMetadataBytes = 1MB
$script:OpcMaxRelationshipParts = 4096
$script:OpcMaxRelationshipBytes = 16MB

function Assert-OpcPercentEncoding {
    param([Parameter(Mandatory = $true)][string]$Value)

    for ($index = 0; $index -lt $Value.Length; $index++) {
        if ($Value[$index] -ne '%') { continue }
        if ($index + 2 -ge $Value.Length) {
            throw 'A package URI contains incomplete percent encoding.'
        }
        $pair = $Value.Substring($index + 1, 2)
        if ($pair -notmatch '^[0-9A-Fa-f]{2}$') {
            throw 'A package URI contains invalid percent encoding.'
        }
        $octet = [Convert]::ToInt32($pair, 16)
        if ($octet -eq 0 -or $octet -eq 0x2f -or $octet -eq 0x5c) {
            throw 'A package URI contains an encoded separator or NUL.'
        }
        $character = [char]$octet
        if ($character -match '^[A-Za-z0-9._~-]$') {
            throw 'A package URI percent-encodes an unreserved character.'
        }
        $index += 2
    }
}

function Assert-OpcPartName {
    param([Parameter(Mandatory = $true)][string]$PartName)

    if (
        -not $PartName.StartsWith('/', [StringComparison]::Ordinal) -or
        $PartName.Length -lt 2 -or
        $PartName.EndsWith('/', [StringComparison]::Ordinal) -or
        $PartName.Contains('\') -or
        $PartName.Contains([char]0) -or
        $PartName.Contains('?') -or
        $PartName.Contains('#') -or
        $PartName.Contains('//')
    ) {
        throw "Invalid OPC part name: $PartName"
    }
    Assert-OpcPercentEncoding $PartName
    foreach ($segment in $PartName.Substring(1).Split('/')) {
        if ($segment.Length -eq 0 -or $segment -eq '.' -or $segment -eq '..') {
            throw "Invalid OPC part name: $PartName"
        }
    }
    return $PartName
}

function Resolve-OpcRelationshipTarget {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePartName,
        [Parameter(Mandatory = $true)][string]$Target
    )

    if (
        $Target.Contains('\') -or
        $Target.Contains([char]0) -or
        $Target.Contains('?') -or
        $Target.Contains('#') -or
        $Target.StartsWith('//', [StringComparison]::Ordinal) -or
        $Target -match '^[A-Za-z][A-Za-z0-9+.-]*:'
    ) {
        throw "Ambiguous or external OPC relationship target: $Target"
    }
    Assert-OpcPercentEncoding $Target

    $baseDirectory = '/'
    if ($SourcePartName -ne '/') {
        $lastSlash = $SourcePartName.LastIndexOf('/')
        $baseDirectory = $SourcePartName.Substring(0, $lastSlash + 1)
    }
    $unresolved = if ($Target.StartsWith('/', [StringComparison]::Ordinal)) {
        $Target
    } else {
        $baseDirectory + $Target
    }
    $segments = New-Object 'Collections.Generic.List[string]'
    foreach ($segment in $unresolved.Split('/')) {
        if ($segment.Length -eq 0 -or $segment -eq '.') { continue }
        if ($segment -eq '..') {
            if ($segments.Count -eq 0) {
                throw 'OPC relationship target escapes the package root.'
            }
            $segments.RemoveAt($segments.Count - 1)
        } else {
            $segments.Add($segment)
        }
    }
    return Assert-OpcPartName ('/' + ($segments -join '/'))
}

function Get-OpcRelationshipsPartName {
    param([Parameter(Mandatory = $true)][string]$SourcePartName)

    [void](Assert-OpcPartName $SourcePartName)
    $lastSlash = $SourcePartName.LastIndexOf('/')
    $directory = $SourcePartName.Substring(0, $lastSlash + 1)
    $fileName = $SourcePartName.Substring($lastSlash + 1)
    return Assert-OpcPartName ($directory + '_rels/' + $fileName + '.rels')
}

function Read-OpcXmlEntry {
    param([Parameter(Mandatory = $true)][IO.Compression.ZipArchiveEntry]$Entry)

    if ($Entry.Length -gt $script:OpcMaxMetadataBytes) {
        throw "OPC metadata part exceeds the inspection limit: $($Entry.FullName)"
    }
    $stream = $null
    $reader = $null
    try {
        $stream = $Entry.Open()
        $settings = New-Object Xml.XmlReaderSettings
        $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver = $null
        $settings.MaxCharactersInDocument = $script:OpcMaxMetadataBytes
        $settings.IgnoreWhitespace = $true
        $reader = [Xml.XmlReader]::Create($stream, $settings)
        $document = New-Object Xml.XmlDocument
        $document.PreserveWhitespace = $false
        $document.XmlResolver = $null
        $document.Load($reader)
        return $document
    }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Get-OpcEntryMap {
    param([Parameter(Mandatory = $true)][IO.Compression.ZipArchive]$Archive)

    if ($Archive.Entries.Count -gt $script:OpcMaxEntries) {
        throw "Package contains more than $($script:OpcMaxEntries) ZIP entries."
    }
    $entries = New-Object 'Collections.Generic.Dictionary[string,IO.Compression.ZipArchiveEntry]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in $Archive.Entries) {
        $name = [string]$entry.FullName
        if ($name.Contains('\')) {
            throw "ZIP entry contains a backslash: $name"
        }
        if ($name.EndsWith('/', [StringComparison]::Ordinal)) {
            $directoryName = $name.Substring(0, $name.Length - 1)
            if ($directoryName.Length -eq 0) { throw 'Invalid ZIP directory entry.' }
            [void](Assert-OpcPartName ('/' + $directoryName))
            continue
        }
        if ($name.StartsWith('/', [StringComparison]::Ordinal)) {
            throw "ZIP entry starts with a slash: $name"
        }
        $partName = Assert-OpcPartName ('/' + $name)
        if ($entries.ContainsKey($partName)) {
            throw "Duplicate or case-colliding ZIP part: $partName"
        }
        $entries.Add($partName, $entry)
    }
    return $entries
}

function Get-OpcContentTypes {
    param(
        [Parameter(Mandatory = $true)][Collections.Generic.Dictionary[string,IO.Compression.ZipArchiveEntry]]$Entries
    )

    if (-not $Entries.ContainsKey('/[Content_Types].xml')) {
        throw '[Content_Types].xml is missing.'
    }
    $document = Read-OpcXmlEntry $Entries['/[Content_Types].xml']
    if (
        $null -eq $document.DocumentElement -or
        $document.DocumentElement.LocalName -cne 'Types' -or
        $document.DocumentElement.NamespaceURI -cne $script:OpcContentTypesNamespace
    ) {
        throw '[Content_Types].xml has an unexpected root element.'
    }
    $defaults = New-Object 'Collections.Generic.Dictionary[string,string]' ([StringComparer]::OrdinalIgnoreCase)
    $overrides = New-Object 'Collections.Generic.Dictionary[string,string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($node in $document.DocumentElement.ChildNodes) {
        if ($node.NodeType -ne [Xml.XmlNodeType]::Element) { continue }
        if ($node.NamespaceURI -cne $script:OpcContentTypesNamespace) {
            throw '[Content_Types].xml contains an unexpected namespace.'
        }
        if ($node.LocalName -ceq 'Default') {
            $extension = $node.GetAttribute('Extension')
            $contentType = $node.GetAttribute('ContentType')
            if (
                [string]::IsNullOrEmpty($extension) -or
                [string]::IsNullOrEmpty($contentType) -or
                $extension -match '[.\\/]'
            ) {
                throw 'Invalid Default content-type declaration.'
            }
            if ($defaults.ContainsKey($extension)) {
                throw 'Duplicate Default content-type declaration.'
            }
            $defaults.Add($extension, $contentType)
        }
        elseif ($node.LocalName -ceq 'Override') {
            $partName = $node.GetAttribute('PartName')
            $contentType = $node.GetAttribute('ContentType')
            if ([string]::IsNullOrEmpty($partName) -or [string]::IsNullOrEmpty($contentType)) {
                throw 'Invalid Override content-type declaration.'
            }
            $partName = Assert-OpcPartName $partName
            if ($overrides.ContainsKey($partName)) {
                throw 'Duplicate Override content-type declaration.'
            }
            $overrides.Add($partName, $contentType)
        }
        else {
            throw "Unexpected element in [Content_Types].xml: $($node.LocalName)"
        }
    }
    return [pscustomobject]@{ Defaults = $defaults; Overrides = $overrides }
}

function Get-OpcEffectiveContentType {
    param(
        [Parameter(Mandatory = $true)][object]$ContentTypes,
        [Parameter(Mandatory = $true)][string]$PartName
    )

    if ($ContentTypes.Overrides.ContainsKey($PartName)) {
        return $ContentTypes.Overrides[$PartName]
    }
    $fileName = $PartName.Substring($PartName.LastIndexOf('/') + 1)
    $dot = $fileName.LastIndexOf('.')
    if ($dot -lt 0) { return $null }
    $extension = $fileName.Substring($dot + 1)
    if ($ContentTypes.Defaults.ContainsKey($extension)) {
        return $ContentTypes.Defaults[$extension]
    }
    return $null
}

function Read-OpcRelationships {
    param([Parameter(Mandatory = $true)][IO.Compression.ZipArchiveEntry]$Entry)

    $document = Read-OpcXmlEntry $Entry
    if (
        $null -eq $document.DocumentElement -or
        $document.DocumentElement.LocalName -cne 'Relationships' -or
        $document.DocumentElement.NamespaceURI -cne $script:OpcRelationshipNamespace
    ) {
        throw "Relationship part has an unexpected root: $($Entry.FullName)"
    }
    $ids = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    $relationships = New-Object 'Collections.Generic.List[object]'
    foreach ($node in $document.DocumentElement.ChildNodes) {
        if ($node.NodeType -ne [Xml.XmlNodeType]::Element) { continue }
        if (
            $node.LocalName -cne 'Relationship' -or
            $node.NamespaceURI -cne $script:OpcRelationshipNamespace
        ) {
            throw "Unexpected element in relationship part: $($Entry.FullName)"
        }
        $id = $node.GetAttribute('Id')
        $type = $node.GetAttribute('Type')
        $target = $node.GetAttribute('Target')
        $targetMode = if ($node.HasAttribute('TargetMode')) { $node.GetAttribute('TargetMode') } else { $null }
        if (
            [string]::IsNullOrEmpty($id) -or
            [string]::IsNullOrEmpty($type) -or
            [string]::IsNullOrEmpty($target) -or
            -not $ids.Add($id)
        ) {
            throw "Invalid or duplicate relationship in $($Entry.FullName)"
        }
        $relationships.Add([pscustomobject]@{
            Id = $id
            Type = $type
            Target = $target
            TargetMode = $targetMode
        })
    }
    return $relationships.ToArray()
}

function Get-OpcPackageSignatureStatus {
    param([Parameter(Mandatory = $true)][IO.Compression.ZipArchive]$Archive)

    $entries = Get-OpcEntryMap $Archive
    $contentTypes = Get-OpcContentTypes $entries
    foreach ($overriddenPartName in $contentTypes.Overrides.Keys) {
        if (-not $entries.ContainsKey($overriddenPartName)) {
            throw "Content-type override targets a missing part: $overriddenPartName"
        }
    }
    $relationshipParts = @(
        $entries.Keys | Where-Object {
            $_ -ieq '/_rels/.rels' -or $_ -imatch '/_rels/[^/]+[.]rels$'
        }
    )
    if ($relationshipParts.Count -gt $script:OpcMaxRelationshipParts) {
        throw 'Package contains too many relationship parts.'
    }
    [long]$relationshipBytes = 0
    $relationshipsByPart = New-Object 'Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($partName in $relationshipParts) {
        $relationshipBytes += $entries[$partName].Length
        if ($relationshipBytes -gt $script:OpcMaxRelationshipBytes) {
            throw 'Relationship metadata exceeds the inspection limit.'
        }
        $relationshipsByPart.Add($partName, @(Read-OpcRelationships $entries[$partName]))
    }
    if (-not $relationshipsByPart.ContainsKey('/_rels/.rels')) {
        throw 'Root relationship part is missing.'
    }

    $originTypedParts = New-Object 'Collections.Generic.List[string]'
    $signatureTypedParts = New-Object 'Collections.Generic.List[string]'
    foreach ($partName in $entries.Keys) {
        $contentType = Get-OpcEffectiveContentType $contentTypes $partName
        if ($contentType -ceq $script:OpcOriginContentType) { $originTypedParts.Add($partName) }
        if ($contentType -ceq $script:OpcSignatureContentType) { $signatureTypedParts.Add($partName) }
    }

    $relevant = New-Object 'Collections.Generic.List[object]'
    foreach ($relationshipPart in $relationshipsByPart.Keys) {
        foreach ($relationship in @($relationshipsByPart[$relationshipPart])) {
            if (
                $relationship.Type -ceq $script:OpcOriginRelationshipType -or
                $relationship.Type -ceq $script:OpcSignatureRelationshipType
            ) {
                $relevant.Add([pscustomobject]@{
                    RelationshipPart = $relationshipPart
                    Relationship = $relationship
                })
            }
        }
    }
    $rootOrigins = @(
        @($relationshipsByPart['/_rels/.rels']) | Where-Object {
            $_.Type -ceq $script:OpcOriginRelationshipType
        }
    )
    if ($rootOrigins.Count -eq 0) {
        if ($originTypedParts.Count -ne 0 -or $signatureTypedParts.Count -ne 0 -or $relevant.Count -ne 0) {
            throw 'Orphaned digital-signature artifacts were found.'
        }
        return 'Absent'
    }
    if ($rootOrigins.Count -ne 1) {
        throw 'Package has multiple digital-signature origin relationships.'
    }
    $originRelationship = $rootOrigins[0]
    if (
        $null -ne $originRelationship.TargetMode -and
        $originRelationship.TargetMode.Length -gt 0 -and
        $originRelationship.TargetMode -cne 'Internal'
    ) {
        throw 'Digital-signature origin relationship is external.'
    }
    $originPartName = Resolve-OpcRelationshipTarget '/' $originRelationship.Target
    if (-not $entries.ContainsKey($originPartName)) {
        throw 'Digital-signature origin target is missing.'
    }
    if ((Get-OpcEffectiveContentType $contentTypes $originPartName) -cne $script:OpcOriginContentType) {
        throw 'Digital-signature origin content type is missing or inconsistent.'
    }
    if ($entries[$originPartName].Length -ne 0) {
        throw 'Digital-signature origin part must be empty.'
    }
    if ($originTypedParts.Count -ne 1 -or $originTypedParts[0] -ine $originPartName) {
        throw 'Orphaned or ambiguous digital-signature origin parts were found.'
    }
    foreach ($item in $relevant) {
        if (
            $item.Relationship.Type -ceq $script:OpcOriginRelationshipType -and
            $item.RelationshipPart -ine '/_rels/.rels'
        ) {
            throw 'Digital-signature origin relationship appears in an unexpected part.'
        }
    }
    $allSignatureRelationships = @(
        $relevant | Where-Object {
            $_.Relationship.Type -ceq $script:OpcSignatureRelationshipType
        }
    )
    $originRelationshipsPart = Get-OpcRelationshipsPartName $originPartName
    if (-not $relationshipsByPart.ContainsKey($originRelationshipsPart)) {
        if ($signatureTypedParts.Count -ne 0 -or $allSignatureRelationships.Count -ne 0) {
            throw 'Orphaned digital-signature artifacts were found.'
        }
        return 'Absent'
    }
    $signatureRelationships = @(
        @($relationshipsByPart[$originRelationshipsPart]) | Where-Object {
            $_.Type -ceq $script:OpcSignatureRelationshipType
        }
    )
    if ($signatureRelationships.Count -eq 0) {
        if ($signatureTypedParts.Count -ne 0 -or $allSignatureRelationships.Count -ne 0) {
            throw 'Orphaned digital-signature artifacts were found.'
        }
        return 'Absent'
    }
    $signatureTargets = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($relationship in $signatureRelationships) {
        if (
            $null -ne $relationship.TargetMode -and
            $relationship.TargetMode.Length -gt 0 -and
            $relationship.TargetMode -cne 'Internal'
        ) {
            throw 'Digital-signature relationship is external.'
        }
        $targetPartName = Resolve-OpcRelationshipTarget $originPartName $relationship.Target
        if (-not $entries.ContainsKey($targetPartName)) {
            throw 'Digital-signature target is missing.'
        }
        if ((Get-OpcEffectiveContentType $contentTypes $targetPartName) -cne $script:OpcSignatureContentType) {
            throw 'Digital-signature content type is missing or inconsistent.'
        }
        if (-not $signatureTargets.Add($targetPartName)) {
            throw 'Duplicate digital-signature targets were found.'
        }
    }
    if ($signatureTypedParts.Count -ne $signatureTargets.Count) {
        throw 'Orphaned or ambiguous digital-signature parts were found.'
    }
    foreach ($partName in $signatureTypedParts) {
        if (-not $signatureTargets.Contains($partName)) {
            throw 'Orphaned or ambiguous digital-signature parts were found.'
        }
    }
    foreach ($item in $relevant) {
        $expectedPart = if ($item.Relationship.Type -ceq $script:OpcOriginRelationshipType) {
            '/_rels/.rels'
        } else {
            $originRelationshipsPart
        }
        if ($item.RelationshipPart -ine $expectedPart) {
            throw 'Digital-signature relationship appears in an unexpected part.'
        }
    }
    return 'Present'
}

function Assert-OoxmlPackageUnsigned {
    param([Parameter(Mandatory = $true)][string]$Path)

    $archive = $null
    try {
        $archive = [IO.Compression.ZipFile]::OpenRead($Path)
        $status = Get-OpcPackageSignatureStatus $archive
    }
    catch {
        throw (
            'Package signature verification failed; write refused: ' +
            $_.Exception.Message
        )
    }
    finally {
        if ($null -ne $archive) { $archive.Dispose() }
    }
    if ($status -ceq 'Present') {
        throw 'Office package signature detected; write refused because modification would invalidate it.'
    }
    if ($status -cne 'Absent') {
        throw 'Package signature verification failed; write refused: unknown signature state.'
    }
}

function Get-OpcXlmMacroSheetStatus {
    param([Parameter(Mandatory = $true)][IO.Compression.ZipArchive]$Archive)

    $entries = Get-OpcEntryMap $Archive
    $contentTypes = Get-OpcContentTypes $entries
    foreach ($overriddenPartName in $contentTypes.Overrides.Keys) {
        if (-not $entries.ContainsKey($overriddenPartName)) {
            throw "Content-type override targets a missing part: $overriddenPartName"
        }
    }

    $found = $false
    foreach ($partName in $entries.Keys) {
        $contentType = Get-OpcEffectiveContentType $contentTypes $partName
        if (
            $partName -imatch '^/xl/macrosheets/' -or
            $script:OpcXlmContentTypes -ccontains $contentType
        ) {
            $found = $true
        }
    }

    $relationshipParts = @(
        $entries.Keys | Where-Object {
            $_ -ieq '/_rels/.rels' -or $_ -imatch '/_rels/[^/]+[.]rels$'
        }
    )
    if ($relationshipParts.Count -gt $script:OpcMaxRelationshipParts) {
        throw 'Package contains too many relationship parts.'
    }
    [long]$relationshipBytes = 0
    $hasRootRelationships = $false
    foreach ($partName in $relationshipParts) {
        $relationshipBytes += $entries[$partName].Length
        if ($relationshipBytes -gt $script:OpcMaxRelationshipBytes) {
            throw 'Relationship metadata exceeds the inspection limit.'
        }
        if ($partName -ieq '/_rels/.rels') { $hasRootRelationships = $true }
        foreach ($relationship in @(Read-OpcRelationships $entries[$partName])) {
            if ($script:OpcXlmRelationshipTypes -ccontains [string]$relationship.Type) {
                $found = $true
            }
        }
    }
    if (-not $hasRootRelationships) { throw 'Root relationship part is missing.' }
    if ($found) { return 'Present' }
    return 'Absent'
}

function Assert-OoxmlPackageHasNoXlmMacroSheets {
    param([Parameter(Mandatory = $true)][string]$Path)

    $archive = $null
    try {
        $archive = [IO.Compression.ZipFile]::OpenRead($Path)
        $status = Get-OpcXlmMacroSheetStatus $archive
    }
    catch {
        throw (
            'Excel 4.0 macro-sheet verification failed; automated open refused: ' +
            $_.Exception.Message
        )
    }
    finally {
        if ($null -ne $archive) { $archive.Dispose() }
    }
    if ($status -ceq 'Present') {
        throw 'Excel 4.0 macro sheet detected; automated open refused because AutomationSecurity does not disable XLM macros.'
    }
    if ($status -cne 'Absent') {
        throw 'Excel 4.0 macro-sheet verification failed; automated open refused: unknown XLM state.'
    }
}
