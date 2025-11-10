# upload original documents alongside converted counterparts
$WorkDir = $PSScriptRoot
$includeOriginals=$includeOriginals ?? $true
$TargetDocumentDir = $TargetDocumentDir ?? $(read-host "which directory contains documents")
$EmbeddableImageExtensions = 
$DocConversionTempDir = $DocConversionTempDir ?? "c:\conversion-tempdir"
$hudubaseurl = $hudubaseurl ?? $(read-host "Enter Hudu URL")
$huduapikey = $huduapikey ?? $(read-host "Enter Hudu API Key")
$globalkbfoldername = $globalkbfoldername ?? "$(get-random -Minimum 111111 -Maximum 999999)" # default to random value if not specified
clear-host
## completed documents array for logging
##
foreach ($file in $(Get-ChildItem -Path ".\helpers" -Filter "*.ps1" -File | Sort-Object Name)) {
    Write-Host "Importing: $($file.Name)" -ForegroundColor DarkBlue
    . $file.FullName
}
Get-EnsuredPath -path $DocConversionTempDir
Get-PSVersionCompatible; Get-HuduModule; Set-HuduInstance -baseurl $hudubaseurl -apikey $huduapikey; Get-HuduVersionCompatible;
$sofficePath=$(Get-LibreMSI -tmpfolder $tmpfolder)
try {Stop-LibreOffice} catch {}


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
    } else {
      write-host "user-supplied path appears to be a file. determining strategy for single-file"
      $doc = $(Get-Item $resourceLocation)
    }

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
              Write-Host "Article $($MatchedDocs.Name) matched and set to be updated from $($doc.FullName)"
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
            $images = Get-ChildItem -LiteralPath $outputDir -File -Recurse -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -match '^\.(png|jpg|jpeg|gif|bmp|tif|tiff)$' } |
                Select-Object -ExpandProperty FullName

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
