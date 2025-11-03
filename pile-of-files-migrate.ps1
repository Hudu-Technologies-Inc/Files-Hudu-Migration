# upload original documents alongside converted counterparts
$includeOriginals=$includeOriginals ?? $true
$StartDir = $StartDir ?? $(read-host "which directory contains documents")
$EmbeddableImageExtensions = $EmbeddableImageExtensions ?? @(".jpg", ".jpeg",".png",".gif",".bmp",".webp",".svg",".apng",".avif",".ico",".jfif",".pjpeg",".pjp")
$DisallowedForConvert = $DisallowedForConvert ?? [System.Collections.ArrayList]@(".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a",".dll", ".so", ".lib", ".bin", ".class", ".pyc", ".pyo", ".o", ".obj",".exe", ".msi", ".bat", ".cmd", ".sh", ".jar", ".app", ".apk", ".dmg", ".iso", ".img",".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".tgz", ".lz",".mp4", ".avi", ".mov", ".wmv", ".mkv", ".webm", ".flv",".psd", ".ai", ".eps", ".indd", ".sketch", ".fig", ".xd", ".blend", ".vsdx",".ds_store", ".thumbs", ".lnk", ".heic", ".eml", ".msg", ".esx", ".esxm")
$tempDir = $tempdir ?? "c:\conversion-tempdir"
$hudubaseurl = $hudubaseurl ?? $(read-host "Enter Hudu URL")
$huduapikey = $huduapikey ?? $(read-host "Enter Hudu API Key")
clear-host
## completed documents array for logging
$completed = @()
##
$WorkDir = $PSScriptRoot
foreach ($file in $(Get-ChildItem -Path ".\helpers" -Filter "*.ps1" -File | Sort-Object Name)) {
    Write-Host "Importing: $($file.Name)" -ForegroundColor DarkBlue
    . $file.FullName
}
Get-EnsuredPath -path $tempdir
Get-PSVersionCompatible; Get-HuduModule; Set-HuduInstance -baseurl $hudubaseurl -apikey $huduapikey; Get-HuduVersionCompatible;

$companyFolders = Get-ChildItem -Path $StartDir -Depth 1 | where-object {$_.PSIsContainer -eq $true}
$sofficePath=$(Get-LibreMSI -tmpfolder $tmpfolder)
try {Stop-LibreOffice} catch {}
$huduCompanies = Get-HuduCompanies
foreach ($companyFolder in $companyFolders) {

  $CompanyName = $companyFolder.Name.TrimEnd('\'); $CompanyName = "$("$($companyName)" -replace "__",'')".Trim();

  # find or create company
  $matchedCompany =
      $huduCompanies | Where-Object { $_.name -eq $CompanyName } | Select-Object -First 1
  if (-not $matchedCompany) {
    $matchedCompany = $huduCompanies | Where-Object {
      (Test-Equiv -A $_.name -B $CompanyName) -or (Test-Equiv -A $_.nickname -B $CompanyName) -or `
          $(Compare-StringsIgnoring -A $_.name -B $CompanyName) -or `
          $($(Get-Similarity -A $(Normalize-Text $_.name) -B $(Normalize-Text $companyName) -gt 0.975))
    } | Select-Object -First 1
  }
  if (-not $matchedCompany) {
    Write-Host "No match for '$CompanyName'. Creating…"
    $created = New-HuduCompany -Name $CompanyName
    $matchedCompany = $created.company ?? $created
    # refresh cache
    $huduCompanies = Get-HuduCompanies
  }
  Write-Host "Starting '$CompanyName' (KB folder: $($companyFolder?.FullName)) — Hudu company id: $($matchedCompany.id)"
  $companyDocs = Get-HuduArticles -companyId $matchedCompany.id

  foreach ($doc in (Get-ChildItem -LiteralPath $companyFolder.FullName -File -depth 1)) {
    # determine upload strategy from filename and whether article exists or not
    try {
      $originalExt  = [IO.Path]::GetExtension($doc.Name).ToLowerInvariant()
      $originalName = [IO.Path]::GetFileNameWithoutExtension($doc.Name)
      $isDisallowed = $DisallowedForConvert -contains $originalExt
      $isPdf        = ($originalExt -eq '.pdf')
      $shouldConvert = -not $isDisallowed
      Write-Host "• $($doc.Name) — ext=$originalExt; disallowed=$isDisallowed; pdf=$isPdf; convert=$shouldConvert"
      $MatchedDocs = $null
      $MatchedDocs = $companyDocs | Where-Object {
        (Test-Equiv -A $_.name -B $originalName) -or $(Compare-StringsIgnoring -A $_.name -B $originalName)
      }
      if ($MatchedDocs -or $MatchedDocs.count -gt 0){
        $skipReason = "Skipped on basis of $originalName matched existing documents: $($MatchedDocs.name -join ', ')"
        $completed += @{ company=$CompanyName; from=$doc.FullName; to='Skipped'; Explain=$skipReason }
        continue
      }

      $newDoc = $null

      if ($isPdf) {
    # conversion process - pdf
        $newDoc = Set-HuduArticleFromPDF -PdfPath $doc.FullName -CompanyName $CompanyName -Title $originalName
        Write-Host "Hudu response:" ($newDoc | ConvertTo-Json -Depth 5)
        $newDoc = $newDoc.HuduArticle
      } elseif ($true -eq $shouldConvert) {
    # conversion process - non-pdf
            $outputDir = Join-Path $tempDir ([guid]::NewGuid().ToString())
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
                                            -CompanyName $CompanyName `
                                            -Title $originalName `
                                            -HtmlContents (Get-Content -Encoding utf8 -Raw $htmlPath) `
                                            -CreateCompanyIfMissing
        $newDoc = $newDoc.HuduArticle
    # standalone article-as-attachment process
      } else {
        $newDoc = New-HuduArticle -name $originalName -companyId $matchedCompany.id -content "Attaching Upload"; $newdoc = $newdoc.article ?? $newdoc;
        $upload = New-HuduUpload -Uploadable_Id $newdoc.id -Uploadable_Type 'Article' -FilePath $doc.FullName; $upload = $upload.upload ?? $upload;
        $newDoc = Set-HuduArticle -id $newDoc.id -companyId $matchedCompany.id -content "<a href='$($upload.url)'>See Attached Document, $($DOC.Name)</a>"
        $completed += @{ company=$CompanyName; from=$doc.FullName; to='Company Upload'; result=($upload.upload ?? $upload) }
        continue
      }

      if ($null -eq $newDoc -or -not $newDoc.id) {
          $completed += @{ company=$CompanyName; from=$doc.FullName; to='Exception'; error="New Document object $newdoc unexpectedly came back empty" }
        throw "Failed to create article for '$($doc.Name)'."
      }

    # attach original document to article if user configured to do so
      if ($includeOriginals) {
        $upload = New-HuduUpload -Uploadable_Id $newDoc.id -Uploadable_Type 'Article' -FilePath $doc.FullName
        $upload = $upload.upload ?? $upload
      }

      $completed += @{ company=$CompanyName; from=$doc.FullName; to='Article'; result=$newDoc }

    } catch {
      Write-Warning "Error processing '$($doc.FullName)': $($_.Exception.Message)"
      $completed += @{ company=$CompanyName; from=$doc.FullName; to='Error'; error=$_.Exception.Message }
      continue
    }
  }
}
$logFile="$($workdir)\pile-of-files-$(Get-Date -Format 'yyyy-MM-dd_hh-mmtt').json"
Write-Host "Completed Upload/Sync for $($completed.count) Articles; Writing detailed results to logfile: $logFile"
$completed | convertto-json -depth 99 | out-file $logFile