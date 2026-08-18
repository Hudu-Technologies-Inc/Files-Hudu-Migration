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
    [ValidateSet('date','filehash','none')]
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

    # Include directories as "resources" to convert
    [Parameter(Mandatory = $false)]
    [switch]$IncludeDirectories,

    [Parameter(Mandatory = $false)]
    [bool]$IncludeOriginals = $true,

    [Parameter(Mandatory = $false)]
    [bool]$PlainTextPdfConversion = $true,

    [Parameter(Mandatory = $false)]
    [AllowEmptyString()]
    [ValidateSet('','plaintext','html','rasterized')]
    [string]$PdfMigrationStrategy = '',

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
    [bool]$PersistTempfiles = $false,
    [Parameter(Mandatory = $false)]
    [string]$HuduBaseUrl,

    [Parameter(Mandatory = $false)]
    [securestring]$HuduApiKeySecure,

    [Parameter(Mandatory = $false)]
    [string]$HuduApiKeyPath,

    [Parameter(Mandatory = $false)]
    [string]$HuduApiKey,

    [Parameter(Mandatory = $false)]
    [string]$SameCompanyName,

    [Parameter(Mandatory = $false)]
    [string[]]$ConvertExtensions = @(),

    [Parameter(Mandatory = $false)]
    [string[]]$UploadAsArticleExtensions = @()
)
    $WorkDir = $PSScriptRoot
    $VerbosePreference = 'SilentlyContinue'
    $requestedConvertExtensions = @($ConvertExtensions)
    $requestedUploadAsArticleExtensions = @($UploadAsArticleExtensions)
    $ConvertExtensions = $null
    $UploadAsArticleExtensions = $null
    $effectivePdfMigrationStrategy = if ([string]::IsNullOrWhiteSpace($PdfMigrationStrategy)) {
        if ($PlainTextPdfConversion) { 'plaintext' } else { 'html' }
    } else {
        $PdfMigrationStrategy
    }

    $defaultEmbeddableImageExtensions = @(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".svg", ".apng", ".avif",".ico",".jfif",".pjpeg",".pjp")
    $defaultDisallowedForConvert = [System.Collections.ArrayList]@(".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a",".dll", ".so", ".lib", ".bin", ".class", ".pyc", ".pyo", ".o", ".obj",".exe", ".msi", ".bat", ".cmd", ".sh", ".jar", ".app", ".apk", ".dmg", ".iso", ".img",".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".tgz", ".lz",".mp4", ".avi", ".mov", ".wmv", ".mkv", ".webm", ".flv",".psd", ".ai", ".eps", ".indd", ".sketch", ".fig", ".xd", ".blend", ".vsdx",".heic", ".eml", ".msg", ".esx", ".esxm")
    $defaultSkipEntirely = [System.Collections.ArrayList]@(".tmp", ".log", ".ds_store", ".thumbs", ".lnk", ".ini", ".db", ".bak", ".old", ".partial", ".env", ".gitignore", ".gitattributes")

    # Load config before helper scripts because helpers/init.ps1 normalizes these values at import time.
    try {
        . (Join-Path $WorkDir "files-config.ps1")
    } catch {
        Write-Warning "Could not load files-config.ps1; proceeding with defaults and user prompts. Error: $($_.Exception.Message); Not to worry, using sane defaults."
    }

    if (-not (Get-Variable -Name EmbeddableImageExtensions -Scope Local -ErrorAction SilentlyContinue) -or $null -eq $EmbeddableImageExtensions) {
        $EmbeddableImageExtensions = $defaultEmbeddableImageExtensions
    }
    if (-not (Get-Variable -Name DisallowedForConvert -Scope Local -ErrorAction SilentlyContinue) -or $null -eq $DisallowedForConvert) {
        $DisallowedForConvert = $defaultDisallowedForConvert
    }
    if (-not (Get-Variable -Name SkipEntirely -Scope Local -ErrorAction SilentlyContinue) -or $null -eq $SkipEntirely) {
        $SkipEntirely = $defaultSkipEntirely
    }
    if (-not (Get-Variable -Name ConvertExtensions -Scope Local -ErrorAction SilentlyContinue) -or $null -eq $ConvertExtensions) {
        $ConvertExtensions = @()
    }
    if (-not (Get-Variable -Name UploadAsArticleExtensions -Scope Local -ErrorAction SilentlyContinue) -or $null -eq $UploadAsArticleExtensions) {
        $UploadAsArticleExtensions = @()
    }

    $configConvertExtensions = @($ConvertExtensions)
    $configUploadAsArticleExtensions = @($UploadAsArticleExtensions)

    # Load helper scripts
    foreach ($file in (Get-ChildItem -Path (Join-Path $WorkDir "helpers") -Filter "*.ps1" -File | Sort-Object Name)) {
        Write-Host "Importing helper: $($file.Name)" -ForegroundColor DarkBlue
        . $file.FullName
    }

    # Ensure or prompt for params and directories
    $null = Get-EnsuredPath -Path $DocConversionTempDir
    if (-not $TargetDocumentDir) {$TargetDocumentDir = Read-Host "Which directory contains documents"}
    if (-not (Test-Path -LiteralPath $TargetDocumentDir)) {throw "Target document directory '$TargetDocumentDir' does not exist."}
    if (-not $DocConversionTempDir) {$DocConversionTempDir = Join-Path -Path $WorkDir -ChildPath "Docs-Temp"}
    if (-not $HuduBaseUrl) {$HuduBaseUrl = Read-Host "Enter Hudu URL"}
    if (-not $HuduApiKey -and $HuduApiKeyPath) {$HuduApiKey = Read-HuduApiKeyFromTransientFile -Path $HuduApiKeyPath}
    if (-not $HuduApiKey -and $HuduApiKeySecure) {$HuduApiKey = Convert-HuduSecureStringToPlainText -SecureString $HuduApiKeySecure}
    if (-not $HuduApiKey) {$HuduApiKeySecure = Read-Host -Prompt "Enter Hudu API Key" -AsSecureString; $HuduApiKey = Convert-HuduSecureStringToPlainText -SecureString $HuduApiKeySecure; clear-host;}
    if (-not $DestinationStrategy) {$DestinationStrategy = Select-ObjectFromList -Message "Will each file be for a unique company?" -Objects @("VariousCompanies","SameCompany","GlobalKB")}
    if (-not $SourceStrategy) {$SourceStrategy = $(if ($IncludeDirectories.IsPresent) {'TopLevel'} else {Select-ObjectFromList -Message "Do you want to look for source documents in $TargetDocumentDir recursively?" -Objects @("Recurse","TopLevel")})}
    [long]$MaxItemBytes = 100MB

    # check requested documents
    Write-Host "Discovering source documents with strategy $($SourceStrategy)..." -ForegroundColor Cyan    
    if ($IncludeDirectories.IsPresent -and $SourceStrategy -ne 'TopLevel') {
        $SourceStrategy = 'TopLevel'; Write-Host "Directories will be included, however recursion is limited to TopLevel when including directories." -ForegroundColor Yellow;
    }    
    
    if ($SourceStrategy -eq 'TopLevel') {
        $sourceObjects = @(Get-ChildItem -Path $TargetDocumentDir -Recurse:$false)
    } else {
        try {
            $sourceObjects = @(Get-ChildItem -Path $TargetDocumentDir -Recurse -Depth $MaxDepth -ErrorAction Stop)
        } catch {
            Write-Warning "Get-ChildItem -Depth is not supported in this PowerShell version; falling back to full recursion."
            $sourceObjects = @(Get-ChildItem -Path $TargetDocumentDir -Recurse -ErrorAction Stop)
        }
    }
    # filter requested documents
    $SkipEntirely = $SkipEntirely ?? [System.Collections.ArrayList]@(".tmp", ".log", ".ds_store", ".thumbs", ".lnk", ".ini", ".db", ".bak", ".old", ".partial", ".env", ".gitignore", ".gitattributes")
    $sourceObjects = @($sourceObjects | Where-Object {
        if ($SkipEntirely -contains $_.Extension.ToLower()) {return $false}
        return $true
    })

    if ($IncludeDirectories.IsPresent) {
        $sourceObjects = @($sourceObjects |
            Where-Object { $_.PSIsContainer -or (-not $_.PSIsContainer -and $_.Length -lt $MaxItemBytes) })
    } else {
        $sourceObjects = @($sourceObjects |
            Where-Object { -not $_.PSIsContainer -and $_.Length -lt $MaxItemBytes })
    }
    if (-not [string]::IsNullOrEmpty($filter)) {
        Write-Host "Applying filter: $filter" -ForegroundColor DarkGray
        $sourceObjects = @($sourceObjects | Where-Object { $_.Name -ilike "$filter" })
    }

    if (-not $sourceObjects -or $sourceObjects.Count -lt 1 -or-not (Test-DocumentSetSafety -Items $sourceObjects -MaxItems $MaxItems -MaxTotalBytes $MaxTotalBytes -MaxItemBytes $MaxItemBytes)) {
        Write-Warning "Not enough viable source objects in your target directory after filtering; aborting."
        return
    }    

    # initialize
    Get-PSVersionCompatible
    $currentVersionResult = Set-HuduModuleInitialized -HuduBaseURL $HuduBaseUrl -HuduAPIKey $HuduApiKey
    $currentVersionResult = $($currentVersionResult ?? $([version]((get-huduappinfo).version))); $MinAllowedVersion = ([version]"2.45.0"); $DisallowedVersions = @([version]("2.37.0")); if ($currentVersionResult -lt $MinAllowedVersion){Write-Host "Sorry, your Hudu version $currentVersionResult is not supported. You'll need to upgrade to $($MinAllowedVersion) in order to continue."; exit 1;}; if ($DisallowedVersions -contains [version]($currentVersionResult)) {write-host "disallowed version $($currentVersionResult); Please upgrade or downgrade if possible first." -ForegroundColor Red; exit 1;};

    [version]$script:CurrentHuduVersion = [version]("$($currentVersionResult | Select-Object -Last 1)")
    $sofficePath = Get-LibreMSI -TmpFolder $DocConversionTempDir
    try {$migrationRecord = Set-MigrationRecord} catch {}
    Write-Host "LibreOffice path: $sofficePath" -ForegroundColor DarkGray


    # region: destination company strategy
    $sameCompanyTarget = $null
    if ($DestinationStrategy -eq 'SameCompany' -and -not [string]::IsNullOrWhiteSpace($SameCompanyName)) {
        $sameCompanyTarget = ChoseBest-ByName -Name $SameCompanyName -choices @(Get-HuduCompanies -name $SameCompanyName)
        if (-not $sameCompanyTarget) {
            Write-Warning "Could not match SameCompanyName '$SameCompanyName' to a Hudu company; choose from the list."
        }
        $companies = @($sameCompanyTarget)
    } else {
        $companies = Get-HuduCompanies
    }

    $results = New-Object System.Collections.Generic.List[object]
    if (-not (Get-Variable -Name DateCompareJitterHours -Scope Script -ErrorAction SilentlyContinue) -or $null -eq $script:DateCompareJitterHours) {
        $script:DateCompareJitterHours = [timespan]::FromHours(12)
    }

    # region: main processing loop
    foreach ($sourceObject in $sourceObjects) {
        try {
            $articleFromResourceRequest = @{
                ResourceLocation = (Get-Item -LiteralPath $sourceObject.FullName)
                IncludeOriginals = ($IncludeOriginals ?? $true)
                PlainTextPdfConversion = ($PlainTextPdfConversion ?? $true)
                PdfMigrationStrategy = $effectivePdfMigrationStrategy
            }
            $alternativeTempPath = $(Resolve-Path ([IO.Path]::GetTempPath())).Path
            [IO.Directory]::CreateDirectory($alternativeTempPath) | Out-Null

            $articleFromResourceRequest.DocConversionTempDir = $DocConversionTempDir ?? $alternativeTempPath
            $articleFromResourceRequest.includeOriginals = $IncludeOriginals ?? $true
            if ($DisallowedForConvert) {$articleFromResourceRequest.DisallowedForConvert = $DisallowedForConvert}
            if ($EmbeddableImageExtensions){ $articleFromResourceRequest.EmbeddableImageExtensions = $EmbeddableImageExtensions }
            $sourceExtension = if ($sourceObject.PSIsContainer) { "" } else { [IO.Path]::GetExtension($sourceObject.Name).ToLowerInvariant() }
            $resourcePreferences = Resolve-ResourceConversionPreferences `
                -Extension $sourceExtension `
                -BaseDisallowedForConvert $DisallowedForConvert `
                -BaseEmbeddableImageExtensions $EmbeddableImageExtensions
            $articleFromResourceRequest.DisallowedForConvert = $resourcePreferences.DisallowedForConvert
            $articleFromResourceRequest.EmbeddableImageExtensions = $resourcePreferences.EmbeddableImageExtensions
            if ($true -eq $updateFilesOnMatch) {
                $articleFromResourceRequest.updateOnMatch = $true
                $articleFromResourceRequest.UpdateStrategy = $UpdateStrategy
            } else {
                $articleFromResourceRequest.UpdateStrategy = 'none'
                $articleFromResourceRequest.updateOnMatch = $false
            }

            switch ($DestinationStrategy) {
                'VariousCompanies' {
                    $target = Select-ObjectFromList `
                        -Objects (Get-HuduCompanies) -allownull $true `
                        -Message "Which company to attribute `"$($articleFromResourceRequest.ResourceLocation)`" to? (Cancel for Global KB)"

                    $targetName = Get-HuduObjectName -InputObject (Unwrap-HuduResultObject -InputObject $target -WrapperNames @('company'))
                    if ($target -and $targetName) {
                        $articleFromResourceRequest.companyName = $targetName
                    }
                }

                'SameCompany' {
                    $sameCompanyTargetName = Get-HuduObjectName -InputObject $sameCompanyTarget
                    if ($sameCompanyTarget -and $sameCompanyTargetName) {
                        $articleFromResourceRequest.companyName = $sameCompanyTargetName
                    }
                }

                'GlobalKB' {
                    # No companyName => global KB in your New-HuduArticleFromLocalResource logic
                }
            }
            # $VerbosePreference = 'Continue'
            $newArticleCommand = Get-Command -Name New-HuduArticleFromLocalResource -ErrorAction Stop
            foreach ($paramName in @($articleFromResourceRequest.Keys)) {
                if (-not $newArticleCommand.Parameters.ContainsKey($paramName)) {
                    $articleFromResourceRequest.Remove($paramName)
                }
            }
            write-host "article processing parameters:`n$($($articleFromResourceRequest | format-list | Out-String))" -ForegroundColor DarkGray
            $result = New-HuduArticleFromLocalResource @articleFromResourceRequest
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

    Write-Host "Completed processing $($results.Count) items." -ForegroundColor Cyan
    if ($true -eq $PersistTempfiles) {
        Write-Host "Temporary files have been preserved at $DocConversionTempDir" -ForegroundColor Yellow
    } else {
        Write-Host "Cleaning up temporary files at $DocConversionTempDir" -ForegroundColor DarkGray
        Remove-Item -LiteralPath $DocConversionTempDir -Recurse -Force -ErrorAction SilentlyContinue
    }

