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

function Test-ShouldUpdateUpload {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][bool]$UpdateOnMatch,
    [Parameter(Mandatory)][ValidateSet('date','filehash','none')][string]$Strategy,

    [Parameter(Mandatory)][datetime]$SourceMTimeUtc,
    [string]$SourceSha256,

    # destination (may be $null if no upload yet)
    [object]$DestUpload
  )

  if (-not $UpdateOnMatch) { return $false }
  if ($Strategy -eq 'none') { return $false }
  if ($null -eq $DestUpload) { return $true } # nothing exists yet => upload

  # normalize dest updated time to UTC
  $destUpdatedUtc = $null
  if ($DestUpload.PSObject.Properties.Name -contains 'created_at' -and $DestUpload.updated_at) {
    try { $destUpdatedUtc = ([datetime]$DestUpload.updated_at).ToUniversalTime() } catch {}
  }

  switch ($Strategy) {
    'date' {
      if ($null -eq $destUpdatedUtc) { return $true }           # can’t compare => choose update
      return ($SourceMTimeUtc -gt $destUpdatedUtc)
    }

    'filehash' {
      if ([string]::IsNullOrWhiteSpace($SourceSha256)) { return $true } # can’t compare => update

      $destHash = $null
      foreach ($p in @('sha256','checksum','hash')) {
        if ($DestUpload.PSObject.Properties.Name -contains $p -and $DestUpload.$p) { $destHash = $DestUpload.$p; break }
      }

      # fallback if no hash (as in folder / dir upload strategy) is to compare by date if available, otherwise update
      if ([string]::IsNullOrWhiteSpace($destHash)) {
        if ($null -ne $destUpdatedUtc) { return ($SourceMTimeUtc -gt $destUpdatedUtc) }
        return $true
      }

      return ($SourceSha256.ToUpperInvariant() -ne $destHash.ToUpperInvariant())
    }
  }
}

function Test-ShouldUpdateUpload {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][bool]$UpdateOnMatch,
    [Parameter(Mandatory)][ValidateSet('date','filehash','none')][string]$Strategy,
    # local
    [Parameter(Mandatory)][datetime]$SourceMTimeUtc,
    [string]$SourceSha256,
    [object]$DestUpload
  )

  if (-not $UpdateOnMatch) { return $false }
  if ($Strategy -eq 'none') { return $false }
  if ($null -eq $DestUpload) { return $true }

  # normalize dest updated time to UTC
  $destUpdatedUtc = $null
  if ($DestUpload.PSObject.Properties.Name -contains 'updated_at' -and $DestUpload.updated_at) {
    try { $destUpdatedUtc = ([datetime]$DestUpload.updated_at).ToUniversalTime() } catch {}
  }

  switch ($Strategy) {
    'date' {
      if ($null -eq $destUpdatedUtc) { return $true }           # can’t compare => choose update
      return ($SourceMTimeUtc -gt $destUpdatedUtc)
    }

    'filehash' {
      if ([string]::IsNullOrWhiteSpace($SourceSha256)) { return $true } # can’t compare => update

      # If your Hudu upload object includes a hash/checksum field, use it here.
      $destHash = $null
      foreach ($p in @('sha256','checksum','hash')) {
        if ($DestUpload.PSObject.Properties.Name -contains $p -and $DestUpload.$p) { $destHash = $DestUpload.$p; break }
      }

      # fall back to date if no hash is available (folder) 
      if ([string]::IsNullOrWhiteSpace($destHash)) {
        if ($null -ne $destUpdatedUtc) { return ($SourceMTimeUtc -gt $destUpdatedUtc) }
        return $true
      }

      return ($SourceSha256.ToUpperInvariant() -ne $destHash.ToUpperInvariant())
    }
  }
}
function New-HuduArticleFromLocalResource {
  param (
    [string]$resourceLocation,
    [string]$companyName=$null,
    [array]$companyDocs=$null,
    [bool]$updateOnMatch=$true,
    [Parameter(Mandatory)][ValidateSet('date','filehash','none')][string]$UpdateStrategy,
    [bool]$includeOriginals=$true,
    [Parameter(Mandatory)][string]$DocConversionTempDir,
    [array]$EmbeddableImageExtensions=@(".jpg", ".jpeg",".png",".gif",".bmp",".webp",".svg",".apng",".avif",".ico",".jfif",".pjpeg",".pjp"),
    [System.Collections.ArrayList]$DisallowedForConvert=[System.Collections.ArrayList]@(".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a",".dll", ".so", ".lib", ".bin", ".class", ".pyc", ".pyo", ".o", ".obj",".exe", ".msi", ".bat", ".cmd", ".sh", ".jar", ".app", ".apk", ".dmg", ".iso", ".img",".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".tgz", ".lz",".mp4", ".avi", ".mov", ".wmv", ".mkv", ".webm", ".flv",".psd", ".ai", ".eps", ".indd", ".sketch", ".fig", ".xd", ".blend", ".vsdx",".ds_store", ".thumbs", ".lnk", ".heic", ".eml", ".msg", ".esx", ".esxm")
  )
    Get-EnsuredPath -Path $DocConversionTempDir
    [version]$script:CurrentHuduVersion = $script:CurrentHuduVersion ?? $([version]("$($(get-huduappinfo).version)"))
    $script:DateCompareJitterHours = $script:DateCompareJitterHours ?? $([timespan]::FromHours(12))

    $companyDocs = $null; $MatchedDocs = $null;
    $results = [pscustomobject]@{
        RequestParams = @{DisallowedForConvert=$DisallowedForConvert; EmbeddableImageExtensions = $EmbeddableImageExtensions; includeOriginals=$includeOriginals; updateOnMatch=$updateOnMatch; companyName=$companyName; UpdateStrategy = $UpdateStrategy;}
        Company=$null; Result=$null; Action=$null; Error=$null; Global=$null; IsPDF = $null; IsImage = $null; Results = $null; FileHash = $null; AllowedToConvertFile = $null; OriginalName = $null; ShouldConvert = $null; MatchedDoc = $null; IsGlobalKB = $null; ArticleResult = $null; Strategy = $null; SourceLastModified = $null; IsDirectory=$null; Images = @(); OriginalEXT = $null; loggedMessages = @(); OutputDir = $null; HTMLPath = $null; isScript =$null; 
        attachmentStatus = "No attachment info yet.";
        NewDoc = $null; OriginalDoc = $null; Upload = $null; CalculateEmbedHashes = ([bool]($script:CurrentHuduVersion -ge [version]("2.39.0")))
    }

    if (([string]::IsNullOrWhiteSpace($resourceLocation)) -or -not $(test-path $resourceLocation)){
        $results.Error= "resource location $resourceLocation does not appear to be a valid path"; Write-Warning $results.Error; 
        return $results
    }
    if (-not ([string]::IsNullOrEmpty($companyName))){
      $results.Company = $(ChoseBest-ByName -Name $companyName -choices $(get-huducompanies)) ?? $null
    }
    $results.IsGlobalKB = [bool]$($null -eq $results.Company)
    write-host "$(if ($results.IsGlobalKB) {'Global KB'} else {"Company '$($results.Company.name)' KB"}) will be target for this article" -ForegroundColor Green

    $companyDocs = $companyDocs ?? $(if ($true -eq $results.IsGlobalKB) {Get-HuduArticles} else {Get-HuduArticles -companyId $results.Company.id})
    $results.OriginalDoc = Get-Item -LiteralPath $resourceLocation
    $results.originalExt  = [IO.Path]::GetExtension($results.OriginalDoc.Name).ToLowerInvariant()
    $results.originalName = [IO.Path]::GetFileNameWithoutExtension($results.OriginalDoc.Name)    
    $results.SourceLastModified = $results.OriginalDoc.LastWriteTimeUtc; write-host "source document $($results.originalName) last modified (UTC): $($results.SourceLastModified)";
    # determine if we're looking at a file or directory and set strategy
    if ($results.OriginalDoc.PSIsContainer) {
        $results.isDirectory = $true
        $results.Strategy = "user-supplied path appears to be a directory. proccing it as a resource itself (gallery of photos, index of files)"; Write-host $results.Strategy -ForegroundColor Green
        try {$results.NewDoc = $(
            if ($null -ne $results.Company) {
                Set-HuduArticleFromResourceFolder -resourcesFolder $results.OriginalDoc -companyName $results.Company.name
                } else {
                Set-HuduArticleFromResourceFolder -resourcesFolder $results.OriginalDoc
            })} catch {
                $results.Error="Error creating article from resource folder $_"; return $results 
            }
    } else {$results.isDirectory = $false}

      $results.Strategy = "user-supplied path appears to be a file. determining strategy for single-file"; Write-host $results.Strategy -ForegroundColor Green
      $results.AllowedToConvertFile = $DisallowedForConvert -contains $results.originalExt
      $results.isPdf        = ($results.originalExt -eq '.pdf')
      $results.isImage      = ($results.originalExt -in $EmbeddableImageExtensions)
      $results.isScript     = ($results.originalExt -in @(".sh", ".expect", ".ps1", ".bat", ".cmd", ".py", ".js", ".vbs", ".wsf", ".psm1", ".psd1"))
        $results.FileHash = "$(Get-FileHash -LiteralPath $results.OriginalDoc.FullName -Algorithm SHA256).Hash"
        if ($results.originalName -and $companyDocs?.Count) {
        $MatchedDocs = $companyDocs | Where-Object {
            (Test-Equiv -A $_.name -B $results.originalName) -or
            (Compare-StringsIgnoring -A $_.name -B $results.originalName)
        }

            if ($MatchedDocs?.Count) {
            $results.MatchedDoc = $MatchedDocs | Select-Object -First 1; $results.MatchedDoc = $results.MatchedDoc.article ?? $results.MatchedDoc;

            if (-not $updateOnMatch) {
                $results.Action = "SkippedMatch(updateOnMatch=false)"; Write-Host $results.Action -ForegroundColor Yellow
                return $results
            } else {
                if ($results.UpdateStrategy -eq 'date') {
                    $shouldUpdate = Test-ShouldUpdateUpload -UpdateOnMatch $updateOnMatch -Strategy $results.UpdateStrategy -SourceMTimeUtc $results.SourceLastModified -DestUpload $results.MatchedDoc.attachments[0]
                    $results.Action = if ($shouldUpdate) { "Matched existing article '$($results.MatchedDoc.name)' but source is newer; proceeding with update." } else { "Matched existing article '$($results.MatchedDoc.name)' and source is not newer; skipping update." }
                    Write-Host $results.Action -ForegroundColor Yellow
                    if (-not $shouldUpdate) { return $results }
                } elseif ($results.UpdateStrategy -eq 'filehash') {
                    $shouldUpdate = Test-ShouldUpdateUpload -UpdateOnMatch $updateOnMatch -Strategy $results.UpdateStrategy -SourceMTimeUtc $results.SourceLastModified -SourceSha256 $results.FileHash .Hash -DestUpload $results.MatchedDoc.attachments[0]
                    $results.Action = if ($shouldUpdate) { "Matched existing article '$($results.MatchedDoc.name)' but file hash differs; proceeding with update." } else { "Matched existing article '$($results.MatchedDoc.name)' and file hash matches; skipping update." }
                    Write-Host $results.Action -ForegroundColor Yellow
                    if (-not $shouldUpdate) { return $results }
                } else {
                    # strategy 'none' should have been handled by the earlier check, but just in case:
                    $results.Action = "Matched existing article '$($results.MatchedDoc.name)' but UpdateStrategy is 'none'; skipping update."
                    Write-Host $results.Action -ForegroundColor Yellow
                    return $results
                }
            }
        }
    }
      
    try {

    if ($true -eq $results.isScript) {
        $safeName = ($results.originalName -replace '[^\w\.-]', '_')
        $results.HtmlPath = [IO.Path]::Combine($DocConversionTempDir,"$safeName-$(Get-Date -Format 'yyyyMMddHHmmss').html")
        $html = Get-HTMLTemplatedScriptContent -FilePath $results.OriginalDoc.FullName -Heading $results.originalName -OutputPath $results.HtmlPath
        write-host "HTML from script generated at $($results.HtmlPath) with contents $($html | Out-String)" -ForegroundColor Green
        $results.NewDoc = Set-HuduArticleFromHtml -ImagesArray @() -CompanyName $(if ($results.IsGlobalKB) { '' } else { $results.Company.name }) -Title $results.originalName -HtmlContents $html -CalculateHashes $results.CalculateEmbedHashes
    } elseif ($true -eq $results.isImage) {
        $results.Strategy = "Processing as single-informatic image, to be embedded in Article"; Write-Host $results.Strategy -ForegroundColor Green
        $results.NewDoc = $(Set-HuduArticleFromHtml -ImagesArray @($results.OriginalDoc.FullName) -Title $results.originalName -CompanyName $(if ($results.IsGlobalKB) { '' } else { $company.name }) -HtmlContents "<img src='$($results.OriginalDoc.Name)' alt='$results.originalName' />")
      }  elseif ($true -eq $results.isPdf) {
        $results.Strategy = "Processing as singular PDF to convert and attach as Article."; Write-Host $results.Strategy -ForegroundColor Green
    # conversion process - pdf [convert to html and attach graphics]
        $results.NewDoc = Set-HuduArticleFromPDF -PdfPath $results.OriginalDoc.FullName -CompanyName $(if ($true -eq $results.IsGlobalKB) {''} else {$CompanyName}) -Title $results.originalName -includeOriginal $includeOriginals -CalculateHashes $results.CalculateEmbedHashes
        $results.NewDoc = $results.NewDoc.HuduArticle;
      } elseif ($true -eq $results.AllowedToConvertFile) {
    # conversion process - non-pdf [but convertable]
        $results.Strategy = "Processing as singular file to convert to and attach as Article."; Write-Host $results.Strategy -ForegroundColor Green
            $results.outputDir = Join-Path $DocConversionTempDir ([guid]::NewGuid().ToString())
            $null = New-Item -ItemType Directory -Path $results.outputDir -Force

            $localIn = Join-Path $results.outputDir $doc.Name
            Copy-Item -LiteralPath $results.OriginalDoc.FullName -Destination $localIn -Force

            $results.htmlpath = Convert-WithLibreOffice -InputFile $localIn -OutputDir $results.outputDir -SofficePath $sofficePath
            if ([string]::IsNullOrWhiteSpace($results.htmlpath) -or -not (Test-Path -LiteralPath $results.htmlpath)) {
                $results.htmlpath = get-childitem -Path $results.outputDir -Filter "*.xhtml" -File | Select-Object -First 1
                $results.htmlpath = $results.htmlpath ?? $(get-childitem -Path $results.outputDir -Filter "*.html" -File | Select-Object -First 1)
            }
            if ([string]::IsNullOrWhiteSpace($results.htmlpath) -or -not (Test-Path -LiteralPath $results.htmlpath)) {
                $results.Error = "Conversion to HTML failed for $($results.OriginalDoc.FullName); no HTML output found.";
                return $results
            }
            $results.Images = Get-ChildItem -LiteralPath $results.outputDir -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '^\.(png|jpg|jpeg|gif|bmp|tif|tiff)$' } | Select-Object -ExpandProperty FullName
            $results.LoggedMessages += "$($results.Images.count) images extracted during conversion."
            $results.NewDoc = Set-HuduArticleFromHtml -ImagesArray ($results.Images ?? @()) -CompanyName $(if ($true -eq $results.IsGlobalKB) {''} else {$CompanyName}) -Title $results.originalName -HtmlContents (Get-Content -Encoding utf8 -Raw $results.htmlpath)  -CalculateHashes $results.CalculateEmbedHashes
            $results.NewDoc = $results.NewDoc.HuduArticle
    # standalone article-as-attachment process [not pdf or convertable]
      } else {
        $results.Strategy = "Processing as Attachment to Reference Article, as file cannot be converted and $(if ($null -ne $results.MatchedDoc){"Article with id $($results.MatchedDoc.id) will be updated"} else {"a new article will be created"})."; Write-Host $results.Strategy -ForegroundColor Green

        $results.NewDoc = $results.MatchedDoc ?? 
            $(if ($true -eq $results.IsGlobalKB) {
                New-HuduArticle -name $results.originalName -content "Attaching Upload"
            } else {
                New-HuduArticle -name $results.originalName -companyId $matchedCompany.id -content "Attaching Upload"
            })
        $results.NewDoc = $results.NewDoc.article ?? $results.NewDoc
        $results.Upload = New-HuduUpload -Uploadable_Id $results.NewDoc.id -Uploadable_Type 'Article' -FilePath $results.OriginalDoc.FullName; $results.Upload = $results.Upload.upload ?? $results.Upload;
        $results.NewDoc = if ($true -eq $results.IsGlobalKB) {
            Set-HuduArticle -id $results.NewDoc.id -content "<h2>$($results.OriginalDoc.Name)</h2><br><a href='$($results.Upload.url)'>See Attached Document, $($results.OriginalDoc.Name)</a> $(Get-MetadataArticleBlock -filePath $results.OriginalDoc.FullName)"
        } else {
            Set-HuduArticle -id $results.NewDoc.id -companyId $matchedCompany.id -content "<a href='$($results.Upload.url)'>See Attached Document, $($results.OriginalDoc.Name)</a>"
        }
        return $results
    }
    $results.ArticleResult = $results.NewDoc
    $results.NewDoc = $results.NewDoc.HuduArticle ?? $results.NewDoc.article ?? $results.NewDoc            

    if ($null -eq $results.NewDoc -or -not $results.NewDoc.id) {
        $results.Error = "New Document object $($results.NewDoc | Out-String) unexpectedly came back empty"
        Write-Error $results.Error
        return $results
    }

    if ($true -eq $includeOriginals -or $true -eq $results.isScript) {
        $uploadHashResult = $null; $localNewer = $false;
        $existingupload = get-huduuploads | where-object {$_.uploadable_id -eq $results.NewDoc.id -and $_.uploadable_type -eq 'Article' -and $_.name -ieq $results.OriginalDoc.Name} | select-object -first 1; $existingupload = $existingupload.upload ?? $existingupload;
        if ($existingupload){
            write-host "An existing upload (attachment) was found." -ForegroundColor DarkGray
            if ($script:CurrentHuduVersion -lt [version]("2.39.0")){
                $results.attachmentStatus =  "Existing attachment upload found for article, but current Hudu version $script:CurrentHuduVersion does not support hash comparison. Using existing attachment/upload as-is. Update to hudu version 2.39.0 or newer to enable hash comparison."; write-host $results.attachmentStatus -ForegroundColor Yellow;
            } else {
                $uploadHashResult = Compare-UploadHashWithFile -uploadId $existingupload.id -FilePath $results.OriginalDoc.FullName
                $uploadUniversalTime = (([datetime]$existingupload.created_date).add($script:DateCompareJitterHours)).ToUniversalTime()
                $localNewer = $results.SourceLastModified -gt $uploadUniversalTime
                if (((-not $uploadHashResult.SameFile) -and $localNewer) -or ([string]::IsNullOrWhiteSpace($uploadHashResult.localPath))) {
                    $results.attachmentStatus = "Existing attachment upload is older $($uploadUniversalTime) and has different hash. Deleting existing upload to replace with new version."; write-host $results.attachmentStatus -ForegroundColor Yellow;
                    Remove-HuduUpload -id $existingupload.id -confirm:$false
                } else {
                    $results.attachmentStatus = "Existing attachment upload appears newest. No need to replace."; write-host $results   .attachmentStatus -ForegroundColor Green;
                    $results.Upload = $existingupload
                    return $results
                }
            }
        } else {$results.attachmentStatus = "No existing upload found. Proceeding to upload new file."; write-host $results.attachmentStatus -ForegroundColor Green;}
        $results.Upload = $existingupload ?? $(New-HuduUpload -Uploadable_Id $results.NewDoc.id -Uploadable_Type 'Article' -FilePath $results.OriginalDoc.FullName)
        $results.Upload = $results.Upload.upload ?? $results.Upload
    }

        return $results
    } catch {
        $results.Error =  "Article from Resource Error-- $_. $($_.Exception.Message) $($_.ScriptStackTrace)"; Write-Error $results.Error
        return $results
    }
}
    
function Convert-WithLibreOffice {
    param (
        [string]$inputFile,
        [string]$outputDir,
        [string]$sofficePath
    )

    try {
        $extension = [System.IO.Path]::GetExtension($inputFile).ToLowerInvariant()
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputFile)

        switch ($extension.ToLowerInvariant()) {
            # Word processors
            ".doc"      { $intermediateExt = "odt" }
            ".docx"     { $intermediateExt = "odt" }
            ".docm"     { $intermediateExt = "odt" }
            ".rtf"      { $intermediateExt = "odt" }
            ".txt"      { $intermediateExt = "odt" }
            ".md"       { $intermediateExt = "odt" }
            ".wpd"      { $intermediateExt = "odt" }

            # Spreadsheets
            ".xls"      { $intermediateExt = "ods" }
            ".xlsx"     { $intermediateExt = "ods" }
            ".csv"      { $intermediateExt = "ods" }

            # Presentations
            ".ppt"      { $intermediateExt = "odp" }
            ".pptx"     { $intermediateExt = "odp" }
            ".pptm"     { $intermediateExt = "odp" }

            # Already OpenDocument
            ".odt"      { $intermediateExt = $null }
            ".ods"      { $intermediateExt = $null }
            ".odp"      { $intermediateExt = $null }

            default { $intermediateExt = $null }
        }
        if ($intermediateExt) {
            $intermediatePath = Join-Path $outputDir "$baseName.$intermediateExt"
            write-host "Step 1: Converting to .$intermediateExt..." 

            Start-Process -FilePath "$sofficePath" `
                -ArgumentList "--headless", "--convert-to", $intermediateExt, "--outdir", "`"$outputDir`"", "`"$inputFile`"" `
                -Wait -NoNewWindow

            if (-not (Test-Path $intermediatePath)) {
                throw "$intermediateExt conversion failed for $inputFile"
            }
        } else {
            # No conversion needed
            $intermediatePath = $inputFile
        }

        write-host "Step $(if ($intermediateExt) {'2'} else {'1'}): Converting .$intermediateExt to XHTML..."

        Start-Process -FilePath "$sofficePath" `
            -ArgumentList "--headless", "--convert-to", "xhtml", "--outdir", "`"$outputDir`"", "`"$intermediatePath`"" `
            -Wait -NoNewWindow

        $htmlPath = Join-Path $outputDir "$baseName.xhtml"

        if (-not (Test-Path $htmlPath)) {
            throw "XHTML conversion failed for $intermediatePath"
        }

        return $htmlPath
    }
    catch {
       write-host $_
        return $null
    }
}

function Get-EmbeddedFilesFromHtml {
    param (
        [string]$htmlPath,
        [int32]$resolution=5
    )

    if (-not (Test-Path $htmlPath)) {
        Write-Warning "HTML file not found: $htmlPath"
        return @{}
    }

    $htmlContent = Get-Content $htmlPath -Raw
    $baseDir = Split-Path -Path $htmlPath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($htmlPath)
    $trimmedBaseName = if ($baseName.Length -gt $resolution) {
        $baseName.Substring(0, $baseName.Length - $resolution).ToLower()
    } else {
        $baseName.ToLower()
    }
    $results = @{
        ExternalFiles        = @()
        Base64Images         = @()
        Base64ImagesWritten  = @()
        UpdatedHTMLContent   = $null
    }

    $guid = [guid]::NewGuid().ToString()
    $uuidSuffix = ($guid -split '-')[0]

    $counter = 0
    $htmlContent = [regex]::Replace($htmlContent, '(?i)<img([^>]+?)src\s*=\s*["'']data:image/(?<type>[a-z]+);base64,(?<b64data>[^"'']+)["'']', {
        param($match)

        $type = $match.Groups["type"].Value
        $b64  = $match.Groups["b64data"].Value

        $ext = switch ($type) {
            'png'  { 'png' }
            'jpeg' { 'jpg' }
            'jpg'  { 'jpg' }
            'gif'  { 'gif' }
            'svg'  { 'svg' }
            'bmp'  { 'bmp' }
            default { 'bin' }
        }

        $counter++
        $filename = "${baseName}_embedded_${uuidSuffix}_$counter.$ext"
        $filepath = Join-Path $baseDir $filename

        try {
            [IO.File]::WriteAllBytes($filepath, [Convert]::FromBase64String($b64))
            $results.ExternalFiles += $filepath
            $results.Base64Images  += "data:image/$type;base64,..."
            $results.Base64ImagesWritten += $filepath

            return "<img$($match.Groups[1].Value)src='$filename'"
        } catch {
            Write-Warning "Failed to decode embedded image: $($_.Exception.Message)"
            return "<img$($match.Groups[1].Value)src='$filename'"
        }
    })
    $skipExts = @(
        ".doc", ".docx", ".docm", ".rtf", ".txt", ".md", ".wpd",
        ".xls", ".xlsx", ".csv", ".ppt", ".pptx", ".pptm",
        ".odt", ".ods", ".odp", ".xhtml", ".xml", ".html", ".json", ".htm"
    )

    $allFiles = Get-ChildItem -Path $baseDir -File
    foreach ($file in $allFiles) {
        $fullFilePath = [IO.Path]::GetFullPath($file.FullName).ToLowerInvariant()
        $htmlPathNormalized = [IO.Path]::GetFullPath($htmlPath).ToLowerInvariant()

        if ($fullFilePath -eq $htmlPathNormalized) {
            continue
        }

        if ($file.Extension.ToLowerInvariant() -in $skipExts) {
            continue
        }

        $otherBaseName = $file.BaseName.ToLower()
        if ($otherBaseName.StartsWith($trimmedBaseName)) {
            $results.ExternalFiles += "$fullFilePath"
        }
    }
        
        
    $results.UpdatedHTMLContent = $htmlContent
    return $results
}

function Convert-PdfXmlToHtml {
    param (
        [Parameter(Mandatory)][string]$XmlPath,
        [string]$OutputHtmlPath = "$XmlPath.html"
    )

    if (-not (Test-Path $XmlPath)) {
        throw "Input XML not found: $XmlPath"
    }

    [xml]$doc = Get-Content $XmlPath
    $html = @()
    $html += '<!DOCTYPE html>'
    $html += '<html><head><meta charset="UTF-8">'
    $html += '<style>body{font-family:sans-serif;font-size:12pt;line-height:1.4}</style></head><body>'

    foreach ($page in $doc.pdf2xml.page) {
        $html += "<div class='page' style='margin-bottom:2em'>"
        foreach ($text in $page.text) {
            $content = ($text.'#text' -replace '\s+', ' ').Trim()
            if ($content) {
                $html += "<p>$content</p>"
            }
        }
        $html += "</div>"
    }

    $html += '</body></html>'
    Set-Content -Path $OutputHtmlPath -Value ($html -join "`n") -Encoding UTF8
    Set-PrintAndLog -message  "Generated slim HTML: $OutputHtmlPath" -Color Green
}
function Convert-PdfToHtml {
    param (
        [string]$inputPath,
        [string]$outputDir = (Split-Path $inputPath),
        [string]$pdftohtmlPath = "C:\tools\poppler\bin\pdftohtml.exe",
        [bool]$includeHiddenText = $true,
        [bool]$complexLayoutMode = $true
    )

    $filename = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
    $outputHtml = Join-Path $outputDir "$filename.html"

    $popplerArgs = @()

    # Preserve layout with less nesting
    if ($complexLayoutMode) {
        $popplerArgs += "-c"            # complex layout mode
    }

    # Enable image extraction
    $popplerArgs += "-p"                # extract images
    $popplerArgs += "-zoom 1.0"         # avoid automatic zoom bloat

    # Output options
    $popplerArgs += "-noframes"        # single HTML file instead of one per page
    $popplerArgs += "-nomerge"         # don't merge text blocks (more control)
    $popplerArgs += "-enc UTF-8"       # UTF-8 encoding
    $popplerArgs += "-nodrm"           # ignore any DRM restrictions

    if ($includeHiddenText) {
        $popplerArgs += "-hidden"
    }

    # Wrap file paths
    $popplerArgs += "`"$inputPath`""
    $popplerArgs += "`"$outputHtml`""

    Start-Process -FilePath $pdftohtmlPath `
        -ArgumentList $popplerArgs -Wait -NoNewWindow

    return (Test-Path $outputHtml) ? $outputHtml : $null
}


function Save-Base64ToFile {
    param (
        [Parameter(Mandatory)]
        [string]$Base64String,

        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    # Remove data URI prefix if present (e.g., "data:image/png;base64,...")
    if ($Base64String -match '^data:.*?;base64,') {
        $Base64String = $Base64String -replace '^data:.*?;base64,', ''
    }

    $bytes = [System.Convert]::FromBase64String($Base64String)
    [System.IO.File]::WriteAllBytes($OutputPath, $bytes)

    Set-PrintAndLog -message  "Saved Base64 content to: $OutputPath" -Color Cyan
}


function Get-FileMagicBytes {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [int]$Count = 16
    )

    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $buffer = New-Object byte[] $Count
        $fs.Read($buffer, 0, $Count) | Out-Null
        return $buffer
    }
    finally {
        $fs.Dispose()
    }
}
function Test-IsPdf {
    param($Bytes)

    # %PDF-
    return ($Bytes[0] -eq 0x25 -and
            $Bytes[1] -eq 0x50 -and
            $Bytes[2] -eq 0x44 -and
            $Bytes[3] -eq 0x46 -and
            $Bytes[4] -eq 0x2D)
}
function Test-IsDocx {
    param([string]$Path, $Bytes)

    # ZIP header
    if (-not ($Bytes[0] -eq 0x50 -and $Bytes[1] -eq 0x4B)) {
        return $false
    }

    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        $found = $zip.Entries | Where-Object { $_.FullName -ieq 'word/document.xml' }
        return [bool]$found
    }
    catch {
        return $false
    }
    finally {
        if ($zip) { $zip.Dispose() }
    }
}
function Test-IsPlainText {
    param([string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)

    # Reject if NULL bytes found
    if ($bytes -contains 0) { return $false }

    try {
        [System.Text.Encoding]::UTF8.GetString($bytes) | Out-Null
        return $true
    }
    catch {
        return $false
    }
}
function Get-FileType {
    param([string]$Path)

    $magic = Get-FileMagicBytes $Path

    if (Test-IsPdf $magic) {
        return 'PDF'
    }

    if (Test-IsDocx $Path $magic) {
        return 'DOCX'
    }

    if (Test-IsPlainText $Path) {
        return 'PlainText'
    }

    return 'UnknownBinary'
}
