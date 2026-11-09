function Test-HuduPhotoSupportedFile {
    param([Parameter(Mandatory)][string]$Path)
    return ([IO.Path]::GetExtension($Path).ToLowerInvariant() -in @('.jpeg', '.jpg', '.png', '.gif', '.webp', '.heic'))
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
        [timespan]$Jitter = [timespan]::FromHours(12)
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
    $query = @{}
    if ($CompanyId -gt 0) { $query.CompanyId = $CompanyId }
    $photoResponse = Get-HuduPhotos @query -ErrorAction Stop
    $photos = Unwrap-HuduResultObject -InputObject $photoResponse -WrapperNames @('photos', 'photo')
    $matches = @($photos | ForEach-Object {
        Unwrap-HuduResultObject -InputObject $_ -WrapperNames @('photo')
    } | Where-Object {
        $photoCompanyId = Get-ObjectPropertyValue -InputObject $_ -Names @('company_id', 'companyId')
        $recordType = Get-ObjectPropertyValue -InputObject $_ -Names @('photoable_type', 'photoableType')
        $recordId = Get-ObjectPropertyValue -InputObject $_ -Names @('photoable_id', 'photoableId')
        # Match only gallery photos in this exact destination, never Article/Asset/etc. photos.
        $inScope = if ($CompanyId -gt 0) {
            $photoCompanyId -eq $CompanyId -and (
                [string]::IsNullOrWhiteSpace($recordType) -or ($recordType -eq 'Company' -and $recordId -eq $CompanyId)
            )
        } else {
            -not $photoCompanyId -and [string]::IsNullOrWhiteSpace($recordType) -and -not $recordId
        }
        $inScope -and (Get-ObjectPropertyValue -InputObject $_ -Names @('caption')) -ieq $Caption
    })
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
    [pscustomobject]@{ Photo = $newPhoto; MatchedPhoto = $matchedPhoto; Action = $action; HashInfo = $comparison; Path = $file.FullName }
}

function Sync-HuduPdfPagePhotos {
    param(
        [Parameter(Mandatory)][string]$PdfPath,
        [Parameter(Mandatory)][object]$PdfData,
        [int]$CompanyId = 0,
        [bool]$UpdateOnMatch = $true,
        [ValidateSet('date','filehash','none')][string]$UpdateStrategy = 'filehash',
        [timespan]$Jitter = [timespan]::FromHours(12)
    )
    $pdf = Get-Item -LiteralPath $PdfPath -ErrorAction Stop
    foreach ($image in @($PdfData.Images)) {
        Sync-HuduPhotoFromFile -FilePath $image -Caption "$($pdf.Name) - $([IO.Path]::GetFileName($image))" `
            -CompanyId $CompanyId -UpdateOnMatch $UpdateOnMatch -UpdateStrategy $UpdateStrategy `
            -SourceMTimeUtc $pdf.LastWriteTimeUtc -Jitter $Jitter
    }
}
