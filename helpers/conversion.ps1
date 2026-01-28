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
    $company = $null
    if ([string]::IsNullOrEmpty($companyName)){
      write-host "No company specified, assuming global kb."
    } else {
      $company = $(ChoseBest-ByName -Name $companyName -choices $(get-huducompanies)) ?? $null
    }

    $IsGlobalKB = [bool]$($null -eq $company)
    $companyDocs = $companyDocs ?? $(if ($true -eq $IsGlobalKB) {Get-HuduArticles} else {Get-HuduArticles -companyId $company.id})
    if (-not $(test-path $resourceLocation)){
      Write-Error "resource location $resourceLocation does not appear to be a valid path"
        return @{ company=$CompanyName; from=$doc.FullName; to="resource path specified $resourceLocation is not valid"; }
    } else {write-host "user-supplied path seems to exist"}
    $doc = Get-Item -LiteralPath $resourceLocation

    if ($doc.PSIsContainer) {
      Write-Host "user-supplied resource appears to be a directory. processing as such"
      try {
        $resourceFolderResult = if ($company) {
          Set-HuduArticleFromResourceFolder -resourcesFolder $doc -companyName $company.name
        } else {
          Set-HuduArticleFromResourceFolder -resourcesFolder $doc
        }
        return @{ company=$company?.name; from=$doc.FullName; to=$resourceFolderResult }
      } catch {
        return @{ company=$company?.name; from=$doc.FullName; to="Error creating article from resource folder"; Exception=$_ }
      }
    }

    write-host "user-supplied path appears to be a file. determining strategy for single-file"
    try {
      $originalExt  = [IO.Path]::GetExtension($doc.Name).ToLowerInvariant()
      $originalName = [IO.Path]::GetFileNameWithoutExtension($doc.Name)
      $isDisallowed = $DisallowedForConvert -contains $originalExt
      $isPdf        = ($originalExt -eq '.pdf')
      $isImage      = $originalExt -in $EmbeddableImageExtensions


      $shouldConvert = -not $isDisallowed
      Write-Host "• $($doc.Name) — ext=$originalExt; disallowed=$isDisallowed; pdf=$isPdf; convert=$shouldConvert"
      $MatchedDocs = $null
      if (-not ([string]::IsNullOrEmpty($originalName)) -and $companyDocs -and $companyDocs.count -gt 0){
        $MatchedDocs = $companyDocs | Where-Object {
          (Test-Equiv -A $_.name -B $originalName) -or $(Compare-StringsIgnoring -A $_.name -B $originalName)
        }
        if ($MatchedDocs -or $MatchedDocs.count -gt 0){
          if ($false -eq $updateOnMatch){
              $skipReason = "Skipped on basis of $originalName matched existing documents: $($MatchedDocs.name -join ', ')"
              return @{ company=$CompanyName; from=$doc.FullName; to='Skipped'; Explain=$skipReason; Global=$IsGlobalKB; }
              continue
          } else {
              $MatchedDocs = $($MatchedDocs | Select-Object -First 1) ?? $MatchedDocs
              Write-Host "Article $($MatchedDocs.Name) matched and set to be SKIPPED from $($doc.FullName)"
              continue
          }
        }
      } else {$MatchedDocs=$null}

      $newDoc = $null
      if ($isImage) {
        Write-Host "Processing as single-informatic image"
        return $(Set-HuduArticleFromHtml `
                    -ImagesArray @($doc.FullName) `
                    -Title $originalName `
                    -CompanyName $(if ($IsGlobalKB) { '' } else { $company.name }) `
                    -HtmlContents "<img src='$($doc.Name)' alt='$originalName' />")
      }  elseif ($isPdf) {
        Write-Host "Processing as singular pdf"        
    # conversion process - pdf [convert to html and attach graphics]
        $newDoc = Set-HuduArticleFromPDF -PdfPath $doc.FullName -CompanyName $(if ($true -eq $IsGlobalKB) {''} else {$CompanyName}) -Title $originalName
        Write-Host "Hudu response:" ($newDoc | ConvertTo-Json -Depth 5)
        $newDoc = $newDoc.HuduArticle
      } elseif ($true -eq $shouldConvert) {
    # conversion process - non-pdf [but convertable]
            $outputDir = Join-Path $DocConversionTempDir ([guid]::NewGuid().ToString())
            $null = New-Item -ItemType Directory -Path $outputDir -Force

            $localIn = Join-Path $outputDir $doc.Name
            Copy-Item -LiteralPath $doc.FullName -Destination $localIn -Force

            $htmlPath = Convert-WithLibreOffice -InputFile $localIn `
                                                -OutputDir $outputDir `
                                                -SofficePath $sofficePath
            if ([string]::IsNullOrWhiteSpace($htmlPath) -or -not (Test-Path -LiteralPath $htmlPath)) {
                $htmlPath = get-childitem -Path $outputDir -Filter "*.xhtml" -File | Select-Object -First 1
                $htmlPath = $htmlPath ?? $(get-childitem -Path $outputDir -Filter "*.html" -File | Select-Object -First 1)
            }
            if ([string]::IsNullOrWhiteSpace($htmlPath) -or -not (Test-Path -LiteralPath $htmlPath)) {
                throw "Conversion to HTML failed for $($doc.FullName); no HTML output found."
            }
            $images = Get-ChildItem -LiteralPath $outputDir -File -Recurse -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -match '^\.(png|jpg|jpeg|gif|bmp|tif|tiff)$' } |
                Select-Object -ExpandProperty FullName
            write-host "$($images.count) images extracted during conversion."
            $newDoc = Set-HuduArticleFromHtml -ImagesArray ($images ?? @()) `
                                            -CompanyName $(if ($true -eq $IsGlobalKB) {''} else {$CompanyName}) `
                                            -Title $originalName `
                                            -HtmlContents (Get-Content -Encoding utf8 -Raw $htmlPath)
            $newDoc = $newDoc.HuduArticle
    # standalone article-as-attachment process [not pdf or convertable]
      } else {
        $newDoc = $MatchedDocs ?? 
            $(if ($true -eq $IsGlobalKB) {
                New-HuduArticle -name $originalName -content "Attaching Upload"
            } else {
                New-HuduArticle -name $originalName -companyId $matchedCompany.id -content "Attaching Upload"
            })
        $newdoc = $newdoc.article ?? $newdoc
        $upload = New-HuduUpload -Uploadable_Id $newdoc.id -Uploadable_Type 'Article' -FilePath $doc.FullName; $upload = $upload.upload ?? $upload;
        $newDoc = if ($true -eq $IsGlobalKB) {
            Set-HuduArticle -id $newDoc.id -content "<a href='$($upload.url)'>See Attached Document, $($DOC.Name)</a>"
        } else {
            Set-HuduArticle -id $newDoc.id -companyId $matchedCompany.id -content "<a href='$($upload.url)'>See Attached Document, $($DOC.Name)</a>"
        }
        return @{ company=$CompanyName; from=$doc.FullName; to='Company Upload'; result=($upload.upload ?? $upload) }
      }

      if ($null -eq $newDoc -or -not $newDoc.id) {
          $completed += @{ company=$CompanyName; from=$doc.FullName; to='Exception'; error="New Document object $newdoc unexpectedly came back empty" }
          return @{ company=$CompanyName; from=$doc.FullName; to='Error'; error="article was not created or updated" }
      }

      if ($includeOriginals) {
        $upload = New-HuduUpload -Uploadable_Id $newDoc.id -Uploadable_Type 'Article' -FilePath $doc.FullName
        $upload = $upload.upload ?? $upload
      }

      return @{ company=$CompanyName; from=$doc.FullName; to='Article'; result=$newDoc; Action=$(if ($null -ne $MatchedDocs){"Updated"} else {"Created"}); Global=$IsGlobalKB; }

    } catch {
      Write-Warning "Error processing '$($doc.FullName)': $($_.Exception.Message)"
      return @{ company=$CompanyName; from=$doc.FullName; to='Error'; error=$_.Exception.Message }
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
