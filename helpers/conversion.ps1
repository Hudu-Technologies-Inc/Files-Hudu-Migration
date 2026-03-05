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

function New-HuduArticleFromLocalResource {
  param (
    [string]$resourceLocation,
    [string]$companyName=$null,
    [array]$companyDocs=$null,
    [bool]$updateOnMatch=$true,
    [bool]$includeOriginals=$true,
    [array]$EmbeddableImageExtensions=@(".jpg", ".jpeg",".png",".gif",".bmp",".webp",".svg",".apng",".avif",".ico",".jfif",".pjpeg",".pjp"),
    [System.Collections.ArrayList]$DisallowedForConvert=[System.Collections.ArrayList]@(".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a",".dll", ".so", ".lib", ".bin", ".class", ".pyc", ".pyo", ".o", ".obj",".exe", ".msi", ".bat", ".cmd", ".sh", ".jar", ".app", ".apk", ".dmg", ".iso", ".img",".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".tgz", ".lz",".mp4", ".avi", ".mov", ".wmv", ".mkv", ".webm", ".flv",".psd", ".ai", ".eps", ".indd", ".sketch", ".fig", ".xd", ".blend", ".vsdx",".ds_store", ".thumbs", ".lnk", ".heic", ".eml", ".msg", ".esx", ".esxm")
  )
  $companyDocs = $null; $MatchedDocs = $null;
  $results = @{
    RequestParams = @{DisallowedForConvert=$DisallowedForConvert; EmbeddableImageExtensions = $EmbeddableImageExtensions; includeOriginals=$includeOriginals; updateOnMatch=$updateOnMatch; companyName=$companyName;}
    Company=$null; Result=$null; Action=$null; Error=$null; Global=$null; IsPDF = $null; IsImage = $null; Results = $null; AllowedToConvertFile = $null; OriginalName = $null; ShouldConvert = $null;
    IsGlobalKB = $null; ArticleResult = $null; Strategy = $null; SourceLastModified = $null; IsDirectory=$null; Images = @(); OriginalEXT = $null; loggedMessages = @(); OutputDir = $null; HTMLPath = $null;
    NewDoc = $null; OriginalDoc = $null; Upload = $null;
    }

    if (([string]::IsNullOrWhiteSpace($resourceLocation)) -or -not $(test-path $resourceLocation)){
        $results.Error= "resource location $resourceLocation does not appear to be a valid path"; Write-Warning $results.Error; 
        return $results
    }

    write-host "user-supplied path seems to exist" -ForegroundColor Green
    if ([string]::IsNullOrEmpty($companyName)){
    } else {
      $results.Company = $(ChoseBest-ByName -Name $companyName -choices $(get-huducompanies)) ?? $null
    }
    $results.IsGlobalKB = [bool]$($null -eq $results.Company)
    write-host "$(if ($results.IsGlobalKB) {'Global KB'} else {"Company '$($results.Company.name)' KB"}) will be target for this article" -ForegroundColor Green


    $companyDocs = $companyDocs ?? $(if ($true -eq $results.IsGlobalKB) {Get-HuduArticles} else {Get-HuduArticles -companyId $results.Company.id})
    $results.OriginalDoc = Get-Item -LiteralPath $resourceLocation
    $results.SourceLastModified = $results.OriginalDoc.LastWriteTimeUtc; write-host "source document $($results.originalName) last modified (UTC): $($results.SourceLastModified)";

    # determine if we're looking at a file or directory and set strategy
    if ($results.OriginalDoc.PSIsContainer) {
        $results.isDirectory = $true
        $results.Strategy = "user-supplied path appears to be a file. determining strategy for single-file"; Write-host $results.Strategy -ForegroundColor Green
        try {$results.NewDoc = $(
            if ($null -ne $results.Company) {
                Set-HuduArticleFromResourceFolder -resourcesFolder $results.OriginalDoc -companyName $results.Company.name
                } else {
                Set-HuduArticleFromResourceFolder -resourcesFolder $results.OriginalDoc
            })} catch {
                $results.Error="Error creating article from resource folder $_"; return $results 
            }
    } else {$results.isDirectory = $false}

    try {
      $results.Strategy = "user-supplied path appears to be a file. determining strategy for single-file"; Write-host $results.Strategy -ForegroundColor Green
      $results.originalExt  = [IO.Path]::GetExtension($results.OriginalDoc.Name).ToLowerInvariant()
      $results.originalName = [IO.Path]::GetFileNameWithoutExtension($results.OriginalDoc.Name)
      $results.AllowedToConvertFile = $DisallowedForConvert -contains $results.originalExt
      $results.isPdf        = ($results.originalExt -eq '.pdf')
      $results.isImage      = ($results.originalExt -in $EmbeddableImageExtensions)
      $results.Shouldconvert = -not $results.AllowedToConvertFile

      
      if (-not ([string]::IsNullOrEmpty($results.originalName)) -and $companyDocs -and $companyDocs.count -gt 0){
        $MatchedDocs = $companyDocs | Where-Object {(Test-Equiv -A $_.name -B $results.originalName) -or $(Compare-StringsIgnoring -A $_.name -B $results.originalName)}
        if ($MatchedDocs -or $MatchedDocs.count -gt 0){
          if ($false -eq $updateOnMatch){
              $result.Action = "Skipped on basis of $($results.originalName) matched existing documents: $($MatchedDocs.name -join ', ') and updateonmatch set to false.";
              return $result;
          } else {
              $MatchedDocs = $($MatchedDocs | Select-Object -First 1) ?? $MatchedDocs
              $result.Action = "Article $($MatchedDocs.Name) matched and set to be SKIPPED from $($results.OriginalDoc.FullName)"; 
              return $result
          }
        }
      }

      if ($true -eq $results.isImage) {
        $results.Strategy = "Processing as single-informatic image, to be embedded in Article"; Write-Host $results.Strategy -ForegroundColor Green
        $results.NewDoc = $(Set-HuduArticleFromHtml -ImagesArray @($results.OriginalDoc.FullName) -Title $results.originalName -CompanyName $(if ($results.IsGlobalKB) { '' } else { $company.name }) -HtmlContents "<img src='$($results.OriginalDoc.Name)' alt='$results.originalName' />")
      }  elseif ($true -eq $results.isPdf) {
        $results.Strategy = "Processing as singular PDF to convert and attach as Article."; Write-Host $results.Strategy -ForegroundColor Green
    # conversion process - pdf [convert to html and attach graphics]
        $results.NewDoc = Set-HuduArticleFromPDF -PdfPath $results.OriginalDoc.FullName -CompanyName $(if ($true -eq $results.IsGlobalKB) {''} else {$CompanyName}) -Title $results.originalName; $results.NewDoc = $results.NewDoc.HuduArticle;
      } elseif ($true -eq $results.Shouldconvert) {
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
            $results.NewDoc = Set-HuduArticleFromHtml -ImagesArray ($results.Images ?? @()) -CompanyName $(if ($true -eq $results.IsGlobalKB) {''} else {$CompanyName}) `
                                            -Title $results.originalName -HtmlContents (Get-Content -Encoding utf8 -Raw $results.htmlpath)
            $results.NewDoc = $results.NewDoc.HuduArticle
    # standalone article-as-attachment process [not pdf or convertable]
      } else {
        $results.Strategy = "Processing as Attachment to Reference Article, as file cannot be converted."; Write-Host $results.Strategy -ForegroundColor Green

        $results.NewDoc = $MatchedDocs ?? 
            $(if ($true -eq $results.IsGlobalKB) {
                New-HuduArticle -name $results.originalName -content "Attaching Upload"
            } else {
                New-HuduArticle -name $results.originalName -companyId $matchedCompany.id -content "Attaching Upload"
            })
        $results.NewDoc = $results.NewDoc.article ?? $results.NewDoc
        $results.Upload = New-HuduUpload -Uploadable_Id $results.NewDoc.id -Uploadable_Type 'Article' -FilePath $results.OriginalDoc.FullName; $results.Upload = $results.Upload.upload ?? $results.Upload;
        $results.NewDoc = if ($true -eq $results.IsGlobalKB) {
            Set-HuduArticle -id $results.NewDoc.id -content "<a href='$($results.Upload.url)'>See Attached Document, $($results.OriginalDoc.Name)</a>"
        } else {
            Set-HuduArticle -id $results.NewDoc.id -companyId $matchedCompany.id -content "<a href='$($results.Upload.url)'>See Attached Document, $($results.OriginalDoc.Name)</a>"
        }
        return $results
    }

      if ($null -eq $results.NewDoc -or -not $results.NewDoc.id) {
          $result.Error = 'Exception'; error="New Document object $results.NewDoc unexpectedly came back empty"; Write-Error $result.Error;
          return $result
      }
      $results.From=$results.OriginalDoc.FullName;

      if ($includeOriginals) {
        $results.Upload = New-HuduUpload -Uploadable_Id $results.NewDoc.id -Uploadable_Type 'Article' -FilePath $results.OriginalDoc.FullName
        $results.Upload = $results.Upload.upload ?? $results.Upload
      }

      

    } catch {
      $results.Error =  "Error processing '$($results.OriginalDoc.FullName)': $($($_ | convertto-json -depth 99).ToString()))"; Write-Error $results.Error;
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
