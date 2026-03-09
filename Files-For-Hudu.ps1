[CmdletBinding()]
param(
    # Root directory to scan for documents
    [Parameter(Mandatory = $true)]
    [string]$TargetDocumentDir,

    # Temporary working directory for conversions
    [Parameter(Mandatory = $false)]
    [string]$DocConversionTempDir,
    [Parameter(Mandatory = $false)]
    [string]$filter=$null,

    [Parameter(Mandatory = $false)]
    [bool]$updateFilesOnMatch=$true,

    [Parameter(Mandatory = $false)]
    [ValidateSet('filehash','date','skip','replace')]
    [string]$UpdateStrategy = 'filehash',

    # Destination strategy:
    # - VariousCompanies: prompt per-file
    # - SameCompany: choose once for all
    # - GlobalKB: no company attribution
    [Parameter(Mandatory = $false)]
    [ValidateSet('VariousCompanies','SameCompany','GlobalKB')]
    [string]$DestinationStrategy,

    # Source scan strategy:
    # - Recurse: search recursively
    # - TopLevel: only first-level
    [Parameter(Mandatory = $false)]
    [ValidateSet('Recurse','TopLevel')]
    [string]$SourceStrategy,

    [Parameter(Mandatory = $false)]
    [bool]$IncludeOriginals = $true,

    # Max number of items to process in one run (files + dirs, filtered)
    [Parameter(Mandatory = $false)]
    [int]$MaxItems = 500,

    # Max total bytes of all selected files combined
    [Parameter(Mandatory = $false)]
    [long]$MaxTotalBytes = 5GB,

    # Max recursion depth (only used if SourceStrategy = Recurse and PS supports -Depth)
    [Parameter(Mandatory = $false)]
    [int]$MaxDepth = 5,

    [Parameter(Mandatory = $false)]
    [bool]$PersistTempfiles = $false
)
    $VerbosePreference = 'SilentlyContinue'
    $WorkDir = $PSScriptRoot
    [long]$MaxItemBytes = 100MB
    $cacheValidityMinutes = 10

    # Load helper scripts
    foreach ($file in (Get-ChildItem -Path (Join-Path $WorkDir "helpers") -Filter "*.ps1" -File | Sort-Object Name)) {
        Write-Host "Importing helper: $($file.Name)" -ForegroundColor DarkBlue
        . $file.FullName
    }
    try {
        . .\files-config.ps1
    } catch {
        Write-Warning "Could not load files-config.ps1; proceeding with defaults and user prompts. Error: $($_.Exception.Message); Not to worry, using sane defaults."
        $EmbeddableImageExtensions = @(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".svg", ".apng", ".avif",".ico",".jfif",".pjpeg",".pjp")
        $DisallowedForConvert = [System.Collections.ArrayList]@(".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a",".dll", ".so", ".lib", ".bin", ".class", ".pyc", ".pyo", ".o", ".obj",".exe", ".msi", ".bat", ".cmd", ".sh", ".jar", ".app", ".apk", ".dmg", ".iso", ".img",".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".tgz", ".lz",".mp4", ".avi", ".mov", ".wmv", ".mkv", ".webm", ".flv",".psd", ".ai", ".eps", ".indd", ".sketch", ".fig", ".xd", ".blend", ".vsdx",".heic", ".eml", ".msg", ".esx", ".esxm")
        $SkipEntirely = [System.Collections.ArrayList]@(".tmp", ".log", ".ds_store", ".thumbs", ".lnk", ".ini", ".db", ".bak", ".old", ".partial", ".env", ".gitignore", ".gitattributes")
    }
    
    if (-not $HuduBaseUrl) {$HuduBaseUrl = Read-Host "Enter Hudu URL"}
    if (-not $HuduApiKey) {$HuduApiKey = Read-Host "Enter Hudu API Key"; clear-host;}
    # Pre-Flight checks and parameter fallbacks
    [version]$script:CurrentHuduVersion = [version]("$($(get-huduappinfo).version)")
    
    if ($script:CurrentHuduVersion -lt [version]("2.41.0") -and $UpdateStrategy -eq 'filehash') {
        Write-Warning "Your Hudu version ($script:CurrentHuduVersion) does not support filehash-based updates; falling back to date-based updates."
        $UpdateStrategy = 'date'
    }
    if (@('filehash','date','replace') -notcontains $UpdateStrategy) {$updateFilesOnMatch = $false; $UpdateStrategy = 'none'} else {$updateFilesOnMatch = $true}

    if ($DestinationStrategy -ne 'GlobalKB' -and ($null -eq $companiesLastIndexedDate -or $null -eq $availableCompanies -or (Get-Date) -gt $companiesLastIndexedDate.AddMinutes($cacheValidityMinutes))) {
        write-host "$(if ($null -eq $companiesLastIndexedDate) {'fetching + caching'} else {"refreshing company list from Hudu... company index is good for $($cacheValidityMinutes) minutes"})" -ForegroundColor DarkGray
        $companiesLastIndexedDate = Get-Date; $availableCompanies = Get-HuduCompanies;
    }
    if ($DestinationStrategy -ne 'GlobalKB' -and (-not $availableCompanies -or $availableCompanies.Count -lt 1)) {
        Write-Warning "No companies found in Hudu; defaulting to Global KB strategy."
        $DestinationStrategy = 'GlobalKB'; $availableCompanies = @();
    }
    if ($availableCompanies.count -eq 1 -and $DestinationStrategy -eq 'VariousCompanies') {
        Write-Host "Only one company found in Hudu; defaulting to SameCompany strategy with that company." -ForegroundColor Yellow
        $DestinationStrategy = 'SameCompany'; $sameCompanyTarget = $availableCompanies[0];
    }
    if (-not $TargetDocumentDir) {$TargetDocumentDir = Read-Host "Which directory contains documents"}
    if (([string]::IsNullOrWhiteSpace($TargetDocumentDir)) -or (not (Test-Path -LiteralPath $TargetDocumentDir))) {throw "Target document directory '$TargetDocumentDir' does not exist or is invalid."}
    if (-not $DocConversionTempDir) {$DocConversionTempDir = Join-Path -Path $WorkDir -ChildPath "Docs-Temp"}; Get-EnsuredPath -Path $DocConversionTempDir;
    if (-not $DestinationStrategy) {$DestinationStrategy = Select-ObjectFromList -Message "Will each file be for a unique company?" -Objects @("VariousCompanies","SameCompany","GlobalKB")}
    if (-not $SourceStrategy) {$SourceStrategy = $(Select-ObjectFromList -Message "Do you want to look for source documents in $TargetDocumentDir recursively?" -Objects @("Recurse","TopLevel"))}

    # check requested documents
    Write-Host "Discovering source documents with strategy $($SourceStrategy)..." -ForegroundColor Cyan
        if ($SourceStrategy -eq 'TopLevel') {
        $sourceObjects = Get-ChildItem -Path $TargetDocumentDir -Recurse:$false
    } else {
        try {
            $sourceObjects = Get-ChildItem -Path $TargetDocumentDir -Recurse -Depth $MaxDepth -ErrorAction Stop
        } catch {
            Write-Warning "Get-ChildItem -Depth is not supported in this PowerShell version; falling back to full recursion."
            $sourceObjects = Get-ChildItem -Path $TargetDocumentDir -Recurse -ErrorAction Stop
        }
    }
    # filter requested documents
    $SkipEntirely = $SkipEntirely ?? [System.Collections.ArrayList]@(".tmp", ".log", ".ds_store", ".thumbs", ".lnk", ".ini", ".db", ".bak", ".old", ".partial", ".env", ".gitignore", ".gitattributes")
    $sourceObjects = $sourceObjects | Where-Object {
        if ($SkipEntirely -contains $_.Extension.ToLower()) {return $false}
        return $true
    }

    $sourceObjects = $sourceObjects | Where-Object { -not $_.PSIsContainer -and $_.Length -lt $MaxItemBytes }

    if (-not [string]::IsNullOrEmpty($filter)) {
        Write-Host "Applying filter: $filter" -ForegroundColor DarkGray
        $sourceObjects = $sourceObjects | Where-Object { $_.Name -ilike "$filter" }
    }

    if (-not $sourceObjects -or $sourceObjects.count -lt 1 -or-not (Test-DocumentSetSafety -Items $sourceObjects -MaxItems $MaxItems -MaxTotalBytes $MaxTotalBytes -MaxItemBytes $MaxItemBytes)) {
        Write-Warning "Not enough viable source objects in your target directory after filtering; aborting."
        return
    }

    # initialize
    Get-PSVersionCompatible; Get-HuduModule; Set-HuduInstance -BaseUrl $HuduBaseUrl -ApiKey $HuduApiKey; Get-HuduVersionCompatible;
    $sofficePath = Get-LibreMSI -TmpFolder $DocConversionTempDir
    Write-Host "LibreOffice path: $sofficePath" -ForegroundColor DarkGray


    # region: destination company strategy
    $sameCompanyTarget = $null
    if ($DestinationStrategy -eq 'SameCompany') {
        $sameCompanyTarget = Select-ObjectFromList `
            -Objects $availableCompanies `
            -Message "Which company to attribute documents in $TargetDocumentDir to? Choose a company or select '0' for Global KB."

        if (-not $sameCompanyTarget) {
            Write-Host "No company selected; treating as Global KB." -ForegroundColor Yellow
            $DestinationStrategy = 'GlobalKB'
        }
    }
    $results = New-Object System.Collections.Generic.List[object]
    $script:DateCompareJitterHours = $script:DateCompareJitterHours ?? $([timespan]::FromHours(12))

    # region: main processing loop
    foreach ($sourceObject in $sourceObjects) {
        try {
            $articleFromResourceRequest = @{
                ResourceLocation = (Get-Item -LiteralPath $sourceObject.FullName)
                IncludeOriginals = ($IncludeOriginals ?? $true)
            }
            $alternativeTempPath = $(Resolve-Path ([IO.Path]::GetTempPath())).Path
            [IO.Directory]::CreateDirectory($alternativeTempPath) | Out-Null

            $articleFromResourceRequest.DocConversionTempDir = $DocConversionTempDir ?? $alternativeTempPath
            $articleFromResourceRequest.includeOriginals = $IncludeOriginals ?? $true
            if ($DisallowedForConvert) {$articleFromResourceRequest.DisallowedForConvert = $DisallowedForConvert}
            if ($EmbeddableImageExtensions){ $articleFromResourceRequest.EmbeddableImageExtensions = $EmbeddableImageExtensions }
            $articleFromResourceRequest.updateOnMatch = $updateFilesOnMatch
            $articleFromResourceRequest.UpdateStrategy = $UpdateStrategy

            switch ($DestinationStrategy) {
                'VariousCompanies' {
                    $target = Select-ObjectFromList -Objects (Get-HuduCompanies) -allownull $true -Message "Which company to attribute `"$($articleFromResourceRequest.ResourceLocation)`" to? (0Cancel for Global KB)"
                    if ($target -and $target.name) {$articleFromResourceRequest.companyName = $target.name} else {write-host "No company selected; treating as Global KB for this file." -ForegroundColor Yellow}
                }

                'SameCompany' {
                    if ($sameCompanyTarget -and $sameCompanyTarget.name) {$articleFromResourceRequest.companyName = $sameCompanyTarget.name} else {write-host "Single company target does not seem valid; treating as Global KB for this file." -ForegroundColor Yellow}
                }
            }
            write-host "article processing parameters:`n$($($articleFromResourceRequest | format-list | Out-String))" -ForegroundColor DarkGray
            $result = New-HuduArticleFromLocalResource @articleFromResourceRequest
            $result.GetEnumerator()
            $results.Add($result)

            Write-Host "Created article from $($sourceObject.FullName)" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to create article from $($sourceObject.FullName): $($_.Exception.Message)"
            $results.Add([pscustomobject]@{
                Path   = $sourceObject.FullName
                Error  = $_.Exception.Message
                Status = 'Failed'
            })
        }
    }

    Write-Host "Completed processing $($results.Count) items. Results will be written to $resultsFile" -ForegroundColor Cyan
    if ($true -eq $PersistTempfiles) {
        Write-Host "Temporary files have been preserved at $DocConversionTempDir" -ForegroundColor Yellow
    } else {
        Write-Host "Cleaning up temporary files at $DocConversionTempDir" -ForegroundColor DarkGray
        Remove-Item -LiteralPath $DocConversionTempDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    $VerbosePreference = 'SilentlyContinue'
    return $results