function Test-HuduPhotoSupportedFile {
    param([Parameter(Mandatory)][string]$Path)
    return ([IO.Path]::GetExtension($Path).ToLowerInvariant() -in @('.jpeg', '.jpg', '.png', '.gif', '.webp', '.heic'))
}

function Get-HuduPhotoCaptionIndex {
    param(
        [int]$CompanyId = 0,
        [Parameter(Mandatory)][hashtable]$PhotoCache
    )
    $cacheKey = if ($CompanyId -gt 0) { "Company:$CompanyId" } else { 'Global' }
    if (-not $PhotoCache.ContainsKey($cacheKey)) {
        $query = @{}
        if ($CompanyId -gt 0) { $query.CompanyId = $CompanyId }
        $photoResponse = Get-HuduPhotos @query -ErrorAction Stop
        $photos = Unwrap-HuduResultObject -InputObject $photoResponse -WrapperNames @('photos', 'photo')
        $index = @{}
        foreach ($photo in @($photos)) {
            $photo = Unwrap-HuduResultObject -InputObject $photo -WrapperNames @('photo')
            $photoCompanyId = Get-ObjectPropertyValue -InputObject $photo -Names @('company_id', 'companyId')
            $recordType = Get-ObjectPropertyValue -InputObject $photo -Names @('photoable_type', 'photoableType')
            $recordId = Get-ObjectPropertyValue -InputObject $photo -Names @('photoable_id', 'photoableId')
            # Index gallery photos only, excluding photos attached to Article/Asset/etc. records.
            $inScope = if ($CompanyId -gt 0) {
                $photoCompanyId -eq $CompanyId -and (
                    [string]::IsNullOrWhiteSpace($recordType) -or ($recordType -eq 'Company' -and $recordId -eq $CompanyId)
                )
            } else {
                -not $photoCompanyId -and [string]::IsNullOrWhiteSpace($recordType) -and -not $recordId
            }
            $caption = [string](Get-ObjectPropertyValue -InputObject $photo -Names @('caption'))
            if (-not $inScope -or [string]::IsNullOrWhiteSpace($caption)) { continue }
            if ($index.ContainsKey($caption)) { $index[$caption] += @($photo) } else { $index[$caption] = @($photo) }
        }
        # Commit only a completely retrieved index. Hashtable caption keys are case-insensitive.
        $PhotoCache[$cacheKey] = $index
    }
    return $PhotoCache[$cacheKey]
}

function Sync-HuduPhotoFromFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [string]$Caption,
        [int]$CompanyId = 0,
        [bool]$UpdateOnMatch = $true,
        [ValidateSet('date','filehash','none')][string]$UpdateStrategy = 'filehash',
        [datetime]$SourceMTimeUtc,
        [timespan]$Jitter = [timespan]::FromHours(12),
        [hashtable]$PhotoCache
    )

    $file = Get-Item -LiteralPath $FilePath -ErrorAction Stop
    if (-not (Test-HuduPhotoSupportedFile -Path $file.FullName)) {
        throw "Unsupported Hudu Photos format: $($file.Extension)"
    }
    foreach ($commandName in @('Get-HuduPhotos', 'New-HuduPhoto', 'Remove-HuduPhoto')) {
        if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) {
            throw "Photos migration requires $commandName. Update the HuduAPI module to the latest Hudu-Technologies-Inc fork."
        }
    }
    if ([string]::IsNullOrWhiteSpace($Caption)) { $Caption = $file.Name }
    if (-not $PSBoundParameters.ContainsKey('SourceMTimeUtc')) { $SourceMTimeUtc = $file.LastWriteTimeUtc }
    if ($null -eq $PhotoCache) { $PhotoCache = @{} }
    $photoIndex = Get-HuduPhotoCaptionIndex -CompanyId $CompanyId -PhotoCache $PhotoCache
    $matches = @(if ($photoIndex.ContainsKey($Caption)) { $photoIndex[$Caption] })
    if ($matches.Count -gt 1) { throw "Multiple photos match caption '$Caption' in the selected destination. Resolve duplicate captions before syncing." }
    $matchedPhoto = $matches | Select-Object -First 1
    $comparison = $null
    $action = 'Created'
    if ($matchedPhoto) {
        $action = 'Skipped'
        $shouldReplace = $false
        if ($UpdateOnMatch -and $UpdateStrategy -ne 'none') {
            $remoteDate = Get-HuduObjectDateUtc -InputObject $matchedPhoto
            $localNewer = -not $remoteDate -or $SourceMTimeUtc -gt $remoteDate.Add($Jitter)
            if ($UpdateStrategy -eq 'filehash') {
                $comparison = Compare-HuduRemoteFileHashWithFile -MediaKind Photo -RemoteId (Get-HuduObjectId -InputObject $matchedPhoto) -LocalFile $file.FullName
                $shouldReplace = -not $comparison.SameFile -and $localNewer
            } else {
                $shouldReplace = -not $remoteDate -or $SourceMTimeUtc -gt $remoteDate
            }
        }
        if (-not $shouldReplace) {
            return [pscustomobject]@{ Photo = $matchedPhoto; MatchedPhoto = $matchedPhoto; Action = $action; HashInfo = $comparison; Path = $file.FullName }
        }
        $action = 'Replaced'
    }

    $create = @{ Path = $file.FullName; Caption = $Caption; ErrorAction = 'Stop' }
    if ($CompanyId -gt 0) { $create.CompanyId = $CompanyId }
    if ($matchedPhoto) {
        $folderId = Get-ObjectPropertyValue -InputObject $matchedPhoto -Names @('folder_id', 'folderId')
        $pinned = Get-ObjectPropertyValue -InputObject $matchedPhoto -Names @('pinned')
        if ($folderId) { $create.FolderId = [int]$folderId }
        if ($null -ne $pinned) { $create.Pinned = [bool]$pinned }
    }
    # Photo metadata updates cannot replace bytes. Upload successfully before removing the old photo.
    try {
        $newPhoto = Unwrap-HuduResultObject -InputObject (New-HuduPhoto @create) -WrapperNames @('photo')
        $newId = Get-HuduObjectId -InputObject $newPhoto
        if (-not $newId) { throw "Hudu did not return a photo ID for '$Caption'; the existing photo has been retained." }
        if ($matchedPhoto) {
            $oldId = Get-HuduObjectId -InputObject $matchedPhoto
            if ($newId -eq $oldId) { throw "Hudu returned the existing photo ID while replacing '$Caption'; no photo was removed." }
            $removed = Remove-HuduPhoto -Id $oldId -Confirm:$false -ErrorAction Stop | Select-Object -Last 1
            if ($removed -ne $true) {
                throw "New photo $newId was created, but old photo $oldId could not be removed. Resolve the duplicate before retrying."
            }
        }
        $photoIndex[$Caption] = @($newPhoto)
    } catch {
        # A failed mutation can leave the server changed. Refresh this destination before the next file.
        $cacheKey = if ($CompanyId -gt 0) { "Company:$CompanyId" } else { 'Global' }
        $PhotoCache.Remove($cacheKey)
        throw
    }
    [pscustomobject]@{ Photo = $newPhoto; MatchedPhoto = $matchedPhoto; Action = $action; HashInfo = $comparison; Path = $file.FullName }
}

function Sync-HuduPdfPagePhotos {
    param(
        [Parameter(Mandatory)][string]$PdfPath,
        [Parameter(Mandatory)][object]$PdfData,
        [int]$CompanyId = 0,
        [bool]$UpdateOnMatch = $true,
        [ValidateSet('date','filehash','none')][string]$UpdateStrategy = 'filehash',
        [timespan]$Jitter = [timespan]::FromHours(12),
        [hashtable]$PhotoCache
    )
    if ($null -eq $PhotoCache) { $PhotoCache = @{} }
    $pdf = Get-Item -LiteralPath $PdfPath -ErrorAction Stop
    foreach ($image in @($PdfData.Images)) {
        Sync-HuduPhotoFromFile -FilePath $image -Caption "$($pdf.Name) - $([IO.Path]::GetFileName($image))" `
            -CompanyId $CompanyId -UpdateOnMatch $UpdateOnMatch -UpdateStrategy $UpdateStrategy `
            -SourceMTimeUtc $pdf.LastWriteTimeUtc -Jitter $Jitter -PhotoCache $PhotoCache
    }
}
