[CmdletBinding()]
param(
    # Root directory to scan for documents
    [Parameter(Mandatory = $false)]
    [string]$TargetDocumentDir,

    # Temporary working directory for conversions
    [Parameter(Mandatory = $false)]
    [string]$DocConversionTempDir,

    # Hudu connection info
    [Parameter(Mandatory = $false)]
    [string]$HuduBaseUrl,

    [Parameter(Mandatory = $false)]
    [string]$HuduApiKey,

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

begin {
    # region: helper – safety check for document sets
    function Test-DocumentSetSafety {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [System.IO.FileSystemInfo[]]$Items,

            [int]$MaxItems,
            [long]$MaxTotalBytes,
            [long]$MaxItemBytes
        )

        if (-not $Items -or $Items.Count -eq 0) {
            Write-Warning "No source items found after filtering."
            return $false
        }

        $files = $Items | Where-Object { -not $_.PSIsContainer }

        $count       = $Items.Count
        $fileCount   = $files.Count
        $totalBytes  = ($files | Measure-Object Length -Sum).Sum
        $largestItem = ($files  | Measure-Object Length -Maximum).Maximum

        $tooMany      = $count      -gt $MaxItems
        $tooLargeTotal= $totalBytes -gt $MaxTotalBytes
        $tooLargeItem = $largestItem -gt $MaxItemBytes

        Write-Host "Selected items: $count (files: $fileCount)" -ForegroundColor Cyan
        Write-Host ("Total size   : {0:N0} bytes" -f $totalBytes) -ForegroundColor Cyan
        Write-Host ("Largest item : {0:N0} bytes" -f $largestItem) -ForegroundColor Cyan

        if (-not ($tooMany -or $tooLargeTotal -or $tooLargeItem)) {
            return $true
        }

        Write-Warning "One or more safety limits were exceeded:"
        if ($tooMany) {
            Write-Warning " - Item count $count exceeds MaxItems $MaxItems"
        }
        if ($tooLargeTotal) {
            Write-Warning (" - Total size {0:N0} exceeds MaxTotalBytes {1:N0}" -f $totalBytes, $MaxTotalBytes)
        }
        if ($tooLargeItem) {
            Write-Warning (" - Largest item {0:N0} exceeds MaxItemBytes {1:N0}" -f $largestItem, $MaxItemBytes)
        }


        $answer = Read-Host "Type 'YES' to proceed anyway (anything else will abort)"
        if ($answer -eq 'YES') {
            Write-Warning "Proceeding despite safety warnings."
            return $true
        } else {
            Write-Host "Aborting per user choice." -ForegroundColor Yellow
            return $false
        }
    }
    # endregion helper
}

process {
    Clear-Host

    # region: basic setup
    $WorkDir = $PSScriptRoot
    if (-not $WorkDir) {
        # Fallback if run in ISE/interactive without $PSScriptRoot
        $WorkDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
    }

    # Ask for missing params interactively (keeps your existing UX)
    if (-not $TargetDocumentDir) {
        $TargetDocumentDir = Read-Host "Which directory contains documents"
    }
    if (-not (Test-Path -LiteralPath $TargetDocumentDir)) {
        throw "Target document directory '$TargetDocumentDir' does not exist."
    }

    if (-not $DocConversionTempDir) {
        $DocConversionTempDir = Join-Path -Path $WorkDir -ChildPath "Docs-Temp"
    }

    if (-not $HuduBaseUrl) {
        $HuduBaseUrl = Read-Host "Enter Hudu URL"
    }
    if (-not $HuduApiKey) {
        $HuduApiKey = Read-Host "Enter Hudu API Key"
    }

    if (-not $DestinationStrategy) {
        $DestinationStrategy = Select-ObjectFromList `
            -Message "Will each file be for a unique company?" `
            -Objects @("VariousCompanies","SameCompany","GlobalKB")
    }

    if (-not $SourceStrategy) {
        $SourceStrategy = Select-ObjectFromList `
            -Message "Do you want to look for source documents in $TargetDocumentDir recursively?" `
            -Objects @("Recurse","TopLevel")
    }

    # Load helper scripts
    foreach ($file in (Get-ChildItem -Path (Join-Path $WorkDir "helpers") -Filter "*.ps1" -File | Sort-Object Name)) {
        Write-Host "Importing helper: $($file.Name)" -ForegroundColor DarkBlue
        . $file.FullName
    }
    . .\files-config.ps1

    # Ensure temp dir
    Get-EnsuredPath -Path $DocConversionTempDir

    # Hudu + Libre setup
    Get-PSVersionCompatible
    Get-HuduModule
    Set-HuduInstance -BaseUrl $HuduBaseUrl -ApiKey $HuduApiKey
    Get-HuduVersionCompatible

    $sofficePath = Get-LibreMSI -TmpFolder $DocConversionTempDir
    Write-Host "LibreOffice path: $sofficePath" -ForegroundColor DarkGray
    Write-Host "Environment ready to use." -ForegroundColor Green

    # endregion basic setup

    # region: build source object list (with safety)
    Write-Host "Discovering source documents..." -ForegroundColor Cyan

    if ($SourceStrategy -eq 'TopLevel') {
        # depth 0 = just the folder itself (in PS 7+). We want just immediate children.
        # Use -Depth 0 + -Recurse:$false as a cross-version-ish compromise:
        $sourceObjects = Get-ChildItem -Path $TargetDocumentDir -Recurse:$false
    }
    else {
        # Recurse with optional MaxDepth (only works natively on PS 7+)
        try {
            $sourceObjects = Get-ChildItem -Path $TargetDocumentDir -Recurse -Depth $MaxDepth -ErrorAction Stop
        } catch {
            # Fallback for PS versions without -Depth
            Write-Warning "Get-ChildItem -Depth is not supported in this PowerShell version; falling back to full recursion."
            $sourceObjects = Get-ChildItem -Path $TargetDocumentDir -Recurse -ErrorAction Stop
        }
    }

    # Filter based on IncludeDirectories + size
    if ($IncludeDirectories.IsPresent) {
        $sourceObjects = $sourceObjects |
            Where-Object { $_.PSIsContainer -or (-not $_.PSIsContainer -and $_.Length -lt $MaxItemBytes) }
    } else {
        $sourceObjects = $sourceObjects |
            Where-Object { -not $_.PSIsContainer -and $_.Length -lt $MaxItemBytes }
    }

    # Apply guardrails (count/size checks)
    if (-not (Test-DocumentSetSafety -Items $sourceObjects -MaxItems $MaxItems -MaxTotalBytes $MaxTotalBytes -MaxItemBytes $MaxItemBytes)) {
        return
    }

    # endregion build source list

    # region: destination company strategy
    $sameCompanyTarget = $null
    if ($DestinationStrategy -eq 'SameCompany') {
        $sameCompanyTarget = Select-ObjectFromList `
            -Objects (Get-HuduCompanies) `
            -Message "Which company to attribute documents in $TargetDocumentDir to? Choose a company or cancel for Global KB."

        # Interpret a $null selection as GlobalKB
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

            switch ($DestinationStrategy) {
                'VariousCompanies' {
                    $target = Select-ObjectFromList `
                        -Objects (Get-HuduCompanies) `
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
    $results
}
