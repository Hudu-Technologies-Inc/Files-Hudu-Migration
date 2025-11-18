[CmdletBinding()]
param(
    # Root directory to scan for documents
    [Parameter(Mandatory = $false)]
    [string]$TargetDocumentDir,

    # Temporary working directory for conversions
    [Parameter(Mandatory = $false)]
    [string]$DocConversionTempDir,

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

    # Whether to also upload original docs (you already had this; still exposed for future logic)
    [Parameter(Mandatory = $false)]
    [bool]$IncludeOriginals = $true,

    # ---- GUARDRAILS / SAFETY LIMITS ----

    # Max number of items to process in one run (files + dirs, filtered)
    [Parameter(Mandatory = $false)]
    [int]$MaxItems = 500,

    # Max total bytes of all selected files combined
    [Parameter(Mandatory = $false)]
    [long]$MaxTotalBytes = 2GB,

    # Max size of any single item
    [Parameter(Mandatory = $false)]
    [long]$MaxItemBytes = 100MB,

    # Max recursion depth (only used if SourceStrategy = Recurse and PS supports -Depth)
    [Parameter(Mandatory = $false)]
    [int]$MaxDepth = 5
)
    $WorkDir = $PSScriptRoot
    # Load helper scripts
    foreach ($file in (Get-ChildItem -Path (Join-Path $WorkDir "helpers") -Filter "*.ps1" -File | Sort-Object Name)) {
        Write-Host "Importing helper: $($file.Name)" -ForegroundColor DarkBlue
        . $file.FullName
    }
    . .\files-config.ps1

    # Ensure or prompt for params and directories
    Get-EnsuredPath -Path $DocConversionTempDir
    if (-not $TargetDocumentDir) {$TargetDocumentDir = Read-Host "Which directory contains documents"}
    if (-not (Test-Path -LiteralPath $TargetDocumentDir)) {throw "Target document directory '$TargetDocumentDir' does not exist."}
    if (-not $DocConversionTempDir) {$DocConversionTempDir = Join-Path -Path $WorkDir -ChildPath "Docs-Temp"}
    if (-not $HuduBaseUrl) {$HuduBaseUrl = Read-Host "Enter Hudu URL"}
    if (-not $HuduApiKey) {$HuduApiKey = Read-Host "Enter Hudu API Key"; clear-host;}
    if (-not $DestinationStrategy) {$DestinationStrategy = Select-ObjectFromList -Message "Will each file be for a unique company?" -Objects @("VariousCompanies","SameCompany","GlobalKB")}
    if (-not $SourceStrategy) {$SourceStrategy = Select-ObjectFromList -Message "Do you want to look for source documents in $TargetDocumentDir recursively?" -Objects @("Recurse","TopLevel")}
    
    # check requested documents
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
    if ($IncludeDirectories.IsPresent) {
        $sourceObjects = $sourceObjects |
            Where-Object { $_.PSIsContainer -or (-not $_.PSIsContainer -and $_.Length -lt $MaxItemBytes) }
    } else {
        $sourceObjects = $sourceObjects |
            Where-Object { -not $_.PSIsContainer -and $_.Length -lt $MaxItemBytes }
    }
    if (-not (Test-DocumentSetSafety -Items $sourceObjects -MaxItems $MaxItems -MaxTotalBytes $MaxTotalBytes -MaxItemBytes $MaxItemBytes)) {
        return
    }    

    # initialize
    Get-PSVersionCompatible; Get-HuduModule; Set-HuduInstance -BaseUrl $HuduBaseUrl -ApiKey $HuduApiKey; Get-HuduVersionCompatible;
    $sofficePath = Get-LibreMSI -TmpFolder $DocConversionTempDir
    Write-Host "LibreOffice path: $sofficePath" -ForegroundColor DarkGray
    Write-Host "Discovering source documents..." -ForegroundColor Cyan

    # region: destination company strategy
    $sameCompanyTarget = $null
    if ($DestinationStrategy -eq 'SameCompany') {
        $sameCompanyTarget = Select-ObjectFromList `
            -Objects (Get-HuduCompanies) `
            -Message "Which company to attribute documents in $TargetDocumentDir to? Choose a company or cancel for Global KB."

        if (-not $sameCompanyTarget) {
            Write-Host "No company selected; treating as Global KB." -ForegroundColor Yellow
            $DestinationStrategy = 'GlobalKB'
        }
    }

    $results = New-Object System.Collections.Generic.List[object]

    # endregion destination company strategy

    # region: main processing loop
    foreach ($sourceObject in $sourceObjects) {
        try {
            $articleFromResourceRequest = @{
                ResourceLocation = (Get-Item -LiteralPath $sourceObject.FullName)
            }
            if ($EmbeddableImageExtensions){ $articleFromResourceRequest.EmbeddableImageExtensions = $EmbeddableImageExtensions }

            switch ($DestinationStrategy) {
                'VariousCompanies' {
                    $target = Select-ObjectFromList `
                        -Objects (Get-HuduCompanies) -allownull $true `
                        -Message "Which company to attribute `"$($articleFromResourceRequest.ResourceLocation)`" to? (Cancel for Global KB)"

                    if ($target -and $target.name) {
                        $articleFromResourceRequest.companyName = $target.name
                    }
                }

                'SameCompany' {
                    if ($sameCompanyTarget -and $sameCompanyTarget.name) {
                        $articleFromResourceRequest.companyName = $sameCompanyTarget.name
                    }
                }

                'GlobalKB' {
                    # No companyName => global KB in your New-HuduArticleFromLocalResource logic
                }
            }

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
    # endregion main processing loop

    Write-Host "Completed processing $($results.Count) items." -ForegroundColor Cyan
    # You can also output $results for pipeline consumption
    return $results