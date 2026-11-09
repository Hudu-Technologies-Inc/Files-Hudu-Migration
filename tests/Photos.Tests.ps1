BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    # Load definitions without running the helper's interactive examples or module installer.
    foreach ($helper in @('articlesanywhere.ps1', 'init.ps1', 'photos.ps1')) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot "helpers/$helper"), [ref]$null, [ref]$null)
        foreach ($statement in $ast.EndBlock.Statements) {
            if ($statement -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
                . ([scriptblock]::Create($statement.Extent.Text))
            }
        }
    }
    # Use the bundled module's actual public signatures for mocks.
    $zip = [IO.Compression.ZipFile]::OpenRead((Join-Path $repoRoot 'HAPI.zip'))
    try {
        foreach ($name in @('Get-HuduPhotos', 'New-HuduPhoto', 'Remove-HuduPhoto', 'Get-HuduCompanies', 'Get-HuduArticles', 'Get-HuduUploads', 'Get-HuduBaseURL')) {
            $entry = $zip.Entries | Where-Object FullName -Like "*/Public/$name.ps1" | Select-Object -First 1
            $reader = [IO.StreamReader]::new($entry.Open())
            try { . ([scriptblock]::Create($reader.ReadToEnd())) } finally { $reader.Dispose() }
        }
    } finally { $zip.Dispose() }
    function New-TestPhoto {
        param($Id = 11, $CompanyId = 7, $Caption = 'picture.png', $RecordType = 'Company', $RecordId = 7)
        [pscustomobject]@{ id = $Id; company_id = $CompanyId; caption = $Caption; photoable_type = $RecordType; photoable_id = $RecordId; updated_at = '2020-01-01T00:00:00Z'; folder_id = 3; pinned = $true }
    }
}

Describe 'Photos destinations and download hash updates' {
    BeforeEach {
        $script:CurrentHuduVersion = [version]'2.44.0'
        $script:DateCompareJitterHours = [timespan]::FromHours(12)
        $script:localFile = Join-Path $TestDrive 'picture.png'
        [IO.File]::WriteAllBytes($script:localFile, [byte[]](1, 2, 3))
        (Get-Item $script:localFile).LastWriteTimeUtc = [datetime]'2025-01-01T00:00:00Z'
        $script:remotePhotos = @()
        $script:remoteBytes = [byte[]](1, 2, 3)
        $script:downloadMissing = $false
        $script:uploadFails = $false
        $script:uploadEmpty = $false
        $script:deleteFails = $false
        $script:events = [Collections.Generic.List[string]]::new()
        Mock Get-HuduPhotos {
            if (-not $Download) { return [pscustomobject]@{ photos = $script:remotePhotos } }
            if ($script:downloadMissing) { return [pscustomobject]@{ id = $Id } }
            $path = Join-Path $OutDir 'remote.png'
            [IO.File]::WriteAllBytes($path, $script:remoteBytes)
            return [pscustomobject]@{ photo = [pscustomobject]@{ id = $Id; localPath = $path } }
        }
        Mock New-HuduPhoto {
            $script:events.Add('create')
            if ($script:uploadFails) { throw 'Upload failed' }
            if ($script:uploadEmpty) { return $null }
            [pscustomobject]@{ photo = [pscustomobject]@{ id = 22; caption = $Caption; company_id = $CompanyId } }
        }
        Mock Remove-HuduPhoto { $script:events.Add('delete'); return (-not $script:deleteFails) }
        Mock Get-HuduCompanies { [pscustomobject]@{ id = 7; name = 'Example' } }
        Mock Get-HuduArticles { @() }
        Mock Get-HuduUploads { @() }
        Mock Set-HuduArticleFromHtml { [pscustomobject]@{ HuduArticle = [pscustomobject]@{ id = 80 }; HuduImages = $ImagesArray } }
    }

    It 'creates a company photo using the filename caption' {
        $result = Sync-HuduPhotoFromFile -FilePath $script:localFile -CompanyId 7
        $result.Action | Should -Be Created
        $result.Photo.id | Should -Be 22
        Should -Invoke New-HuduPhoto -Times 1 -Exactly -ParameterFilter { $CompanyId -eq 7 -and $Caption -eq 'picture.png' }
        Should -Invoke Remove-HuduPhoto -Times 0
    }

    It 'leaves CompanyId unbound when creating a global photo' {
        Sync-HuduPhotoFromFile -FilePath $script:localFile | Out-Null
        Should -Invoke New-HuduPhoto -Times 1 -ParameterFilter { -not $PesterBoundParameters.ContainsKey('CompanyId') }
    }

    It 'ignores photos belonging to other companies or records' {
        $script:remotePhotos = @((New-TestPhoto -CompanyId 8 -RecordId 8), (New-TestPhoto -RecordType Article -RecordId 99))
        (Sync-HuduPhotoFromFile -FilePath $script:localFile -CompanyId 7).Action | Should -Be Created
        Should -Invoke Remove-HuduPhoto -Times 0
    }

    It 'does not match company or record photos in the global gallery' {
        $script:remotePhotos = @((New-TestPhoto), (New-TestPhoto -CompanyId $null -RecordType Article -RecordId 99))
        (Sync-HuduPhotoFromFile -FilePath $script:localFile).Action | Should -Be Created
    }

    It 'downloads identical photos and skips replacement' {
        $script:remotePhotos = @(New-TestPhoto)
        $result = Sync-HuduPhotoFromFile -FilePath $script:localFile -CompanyId 7
        $result.Action | Should -Be Skipped
        $result.HashInfo.SameFile | Should -BeTrue
        Should -Invoke Get-HuduPhotos -Times 1 -Exactly -ParameterFilter { $Download -and $Id -eq 11 }
        Should -Invoke New-HuduPhoto -Times 0
        Should -Invoke Remove-HuduPhoto -Times 0
    }

    It 'reuses matching global photos without confusing company galleries' {
        $script:remotePhotos = @(New-TestPhoto -CompanyId $null -RecordType $null -RecordId $null)
        (Sync-HuduPhotoFromFile -FilePath $script:localFile).Action | Should -Be Skipped
        Should -Invoke New-HuduPhoto -Times 0
    }

    It 'supports literal filenames containing wildcard characters during hash comparison' {
        $path = Join-Path $TestDrive 'picture[1].png'
        [IO.File]::WriteAllBytes($path, [byte[]](1, 2, 3))
        $script:remotePhotos = @(New-TestPhoto -Caption 'picture[1].png')
        (Sync-HuduPhotoFromFile -FilePath $path -CompanyId 7).Action | Should -Be Skipped
        Should -Invoke New-HuduPhoto -Times 0
    }

    It 'uploads changed newer bytes before deleting and preserves folder and pinned metadata' {
        $script:remotePhotos = @(New-TestPhoto)
        $script:remoteBytes = [byte[]](9, 8, 7)
        $result = Sync-HuduPhotoFromFile -FilePath $script:localFile -CompanyId 7
        $result.Action | Should -Be Replaced
        ($script:events -join ',') | Should -Be 'create,delete'
        Should -Invoke New-HuduPhoto -Times 1 -ParameterFilter { $FolderId -eq 3 -and $Pinned -eq $true }
        Should -Invoke Remove-HuduPhoto -Times 1 -ParameterFilter { $Id -eq 11 -and $Confirm -eq $false }
    }

    It 'retains changed remote bytes when the local source is older or within the tolerance' -TestCases @(
        @{ Date = '2019-01-01T00:00:00Z' }, @{ Date = '2020-01-01T06:00:00Z' }
    ) {
        param($Date)
        $script:remotePhotos = @(New-TestPhoto)
        $script:remoteBytes = [byte[]](9)
        (Sync-HuduPhotoFromFile -FilePath $script:localFile -CompanyId 7 -SourceMTimeUtc ([datetime]$Date)).Action | Should -Be Skipped
        Should -Invoke New-HuduPhoto -Times 0
    }

    It 'uses timestamps without a download for date strategy' {
        $script:remotePhotos = @(New-TestPhoto)
        (Sync-HuduPhotoFromFile -FilePath $script:localFile -CompanyId 7 -UpdateStrategy date).Action | Should -Be Replaced
        Should -Invoke Get-HuduPhotos -Times 0 -ParameterFilter { $Download }
    }

    It 'skips older sources under date strategy' {
        $script:remotePhotos = @(New-TestPhoto)
        (Sync-HuduPhotoFromFile -FilePath $script:localFile -CompanyId 7 -UpdateStrategy date -SourceMTimeUtc ([datetime]'2019-01-01')).Action | Should -Be Skipped
        Should -Invoke New-HuduPhoto -Times 0
    }

    It 'retains matched photos when updates are disabled' -TestCases @(
        @{ Update = $false; Strategy = 'filehash' }, @{ Update = $true; Strategy = 'none' }
    ) {
        param($Update, $Strategy)
        $script:remotePhotos = @(New-TestPhoto)
        (Sync-HuduPhotoFromFile -FilePath $script:localFile -CompanyId 7 -UpdateOnMatch $Update -UpdateStrategy $Strategy).Action | Should -Be Skipped
        Should -Invoke Get-HuduPhotos -Times 0 -ParameterFilter { $Download }
        Should -Invoke New-HuduPhoto -Times 0
    }

    It 'creates missing photos even when update strategy is none' {
        (Sync-HuduPhotoFromFile -FilePath $script:localFile -UpdateStrategy none).Action | Should -Be Created
    }

    It 'preserves the existing photo when the download is missing' {
        $script:remotePhotos = @(New-TestPhoto)
        $script:downloadMissing = $true
        { Sync-HuduPhotoFromFile -FilePath $script:localFile -CompanyId 7 } | Should -Throw '*Could not download photo*'
        Should -Invoke New-HuduPhoto -Times 0
        Should -Invoke Remove-HuduPhoto -Times 0
    }

    It 'preserves the existing photo when upload fails or returns no ID' -TestCases @(
        @{ Fails = $true; Empty = $false }, @{ Fails = $false; Empty = $true }
    ) {
        param($Fails, $Empty)
        $script:remotePhotos = @(New-TestPhoto)
        $script:remoteBytes = [byte[]](9)
        $script:uploadFails = $Fails
        $script:uploadEmpty = $Empty
        { Sync-HuduPhotoFromFile -FilePath $script:localFile -CompanyId 7 } | Should -Throw
        Should -Invoke Remove-HuduPhoto -Times 0
    }

    It 'reports failed deletion with both IDs after creating the replacement' {
        $script:remotePhotos = @(New-TestPhoto)
        $script:remoteBytes = [byte[]](9)
        $script:deleteFails = $true
        { Sync-HuduPhotoFromFile -FilePath $script:localFile -CompanyId 7 } | Should -Throw '*New photo 22*old photo 11*'
        ($script:events -join ',') | Should -Be 'create,delete'
    }

    It 'rejects ambiguous matching captions' {
        $script:remotePhotos = @((New-TestPhoto), (New-TestPhoto -Id 12))
        { Sync-HuduPhotoFromFile -FilePath $script:localFile -CompanyId 7 } | Should -Throw '*Multiple photos*'
        Should -Invoke New-HuduPhoto -Times 0
    }

    It 'routes HEIC standalone images to Photos without querying articles' {
        $heic = Join-Path $TestDrive 'camera.heic'
        [IO.File]::WriteAllBytes($heic, [byte[]](1))
        $result = New-HuduArticleFromLocalResource -ResourceLocation $heic -CompanyName Example -StandaloneImagesAsPhotos $true -DocConversionTempDir $TestDrive
        $result.PhotoResult.Action | Should -Be Created
        Should -Invoke Get-HuduArticles -Times 0
        Should -Invoke Set-HuduArticleFromHtml -Times 0
    }

    It 'retains article handling for unsupported photo formats' {
        $svg = Join-Path $TestDrive 'diagram.svg'
        Set-Content -LiteralPath $svg -Value '<svg />'
        $result = New-HuduArticleFromLocalResource -ResourceLocation $svg -StandaloneImagesAsPhotos $true -IncludeOriginals $false -DocConversionTempDir $TestDrive
        $result.Result.id | Should -Be 80
        Should -Invoke New-HuduPhoto -Times 0
        Should -Invoke Set-HuduArticleFromHtml -Times 1
    }

    It 'keeps image article handling by default' {
        $result = New-HuduArticleFromLocalResource -ResourceLocation $script:localFile -IncludeOriginals $false -DocConversionTempDir $TestDrive
        $result.Result.id | Should -Be 80
        Should -Invoke New-HuduPhoto -Times 0
    }

    Context 'Rasterized PDF page photos' {
        BeforeEach {
            $script:pdfPath = Join-Path $TestDrive 'manual.pdf'
            Set-Content -LiteralPath $script:pdfPath -Value 'PDF fixture'
            (Get-Item $script:pdfPath).LastWriteTimeUtc = [datetime]'2019-01-01T00:00:00Z'
            $script:pagePath = Join-Path $TestDrive 'page-01.png'
            [IO.File]::WriteAllBytes($script:pagePath, [byte[]](1, 2, 3))
            $script:pdfData = [pscustomobject]@{ Html = '<img src="page-01.png" />'; Images = @($script:pagePath) }
            Mock Get-HTMLFromPDFRasterized { $script:pdfData }
            Mock Get-HuduBaseURL { 'https://example.test' }
        }

        It 'adds page photos even when the matched PDF article is skipped' {
            Mock Get-HuduArticles { [pscustomobject]@{ id = 80; name = 'manual'; company_id = $null } }
            $result = New-HuduArticleFromLocalResource -ResourceLocation $script:pdfPath -IncludePdfPagePhotos $true -PdfMigrationStrategy rasterized -UpdateOnMatch $false -DocConversionTempDir $TestDrive
            $result.PdfPagePhotos.Count | Should -Be 1
            $result.Action | Should -Be 'SkippedMatch(updateOnMatch=false)'
            Should -Invoke New-HuduPhoto -Times 1 -ParameterFilter { $Caption -eq 'manual.pdf - page-01.png' }
        }

        It 'compares using PDF timestamps rather than newly rendered page timestamps' {
            $script:remotePhotos = @(New-TestPhoto -Caption 'manual.pdf - page-01.png')
            $script:remoteBytes = [byte[]](9)
            $result = Sync-HuduPdfPagePhotos -PdfPath $script:pdfPath -PdfData $script:pdfData -CompanyId 7
            $result.Action | Should -Be Skipped
            Should -Invoke New-HuduPhoto -Times 0
        }

        It 'renders once and retains the PDF article with page photos enabled' {
            $result = New-HuduArticleFromLocalResource -ResourceLocation $script:pdfPath -IncludePdfPagePhotos $true -PdfMigrationStrategy rasterized -IncludeOriginals $false -DisallowedForConvert @() -DocConversionTempDir $TestDrive
            $result.PdfPagePhotos.Count | Should -Be 1
            $result.Result.id | Should -Be 80
            Should -Invoke Get-HTMLFromPDFRasterized -Times 1 -Exactly -ParameterFilter { $OutputDir.StartsWith($TestDrive) }
            Should -Invoke Set-HuduArticleFromHtml -Times 1 -ParameterFilter { $ImagesArray.Count -eq 1 }
        }

        It 'does not publish page photos when disabled' {
            Set-HuduArticleFromPDF -PdfPath $script:pdfPath -PdfMigrationStrategy rasterized -IncludeOriginal $false | Out-Null
            Should -Invoke New-HuduPhoto -Times 0
            Should -Invoke Set-HuduArticleFromHtml -Times 1
        }

        It 'supports direct PDF helper calls with company page photos' {
            $result = Set-HuduArticleFromPDF -PdfPath $script:pdfPath -CompanyName Example -IncludePdfPagePhotos $true -PdfMigrationStrategy rasterized -IncludeOriginal $false
            $result.PdfPagePhotos.Count | Should -Be 1
            Should -Invoke New-HuduPhoto -Times 1 -ParameterFilter { $CompanyId -eq 7 -and $Caption -eq 'manual.pdf - page-01.png' }
        }
    }
}

Describe 'CLI Photos option precedence and forwarding' {
    BeforeEach {
        $script:driverRoot = Join-Path $TestDrive 'driver'
        $helpers = Join-Path $script:driverRoot 'helpers'
        $script:inputRoot = Join-Path $TestDrive 'input'
        New-Item -ItemType Directory -Path $helpers, $script:inputRoot -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $repoRoot 'Files-For-Hudu.ps1') -Destination $script:driverRoot
        Set-Content -LiteralPath (Join-Path $script:inputRoot 'picture.png') -Value 'fixture'
        # Replace external initialization/conversion with a capture, and exercise the actual CLI driver.
        Set-Content -LiteralPath (Join-Path $helpers 'capture.ps1') -Value @'
function Get-EnsuredPath { param($Path) return $Path }
function Test-DocumentSetSafety { param($Items, $MaxItems, $MaxTotalBytes, $MaxItemBytes) return $true }
function Get-PSVersionCompatible {}
function Set-HuduModuleInitialized { param($HuduBaseURL, $HuduAPIKey) return [version]'2.44.0' }
function Get-LibreMSI { param($TmpFolder) return 'fake-libreoffice' }
function Set-MigrationRecord {}
function Get-HuduCompanies { @() }
function Resolve-ResourceConversionPreferences {
    param($Extension, $BaseDisallowedForConvert, $BaseEmbeddableImageExtensions)
    @{ DisallowedForConvert = $BaseDisallowedForConvert; EmbeddableImageExtensions = $BaseEmbeddableImageExtensions }
}
function New-HuduArticleFromLocalResource {
    param($ResourceLocation, $IncludeOriginals, $PlainTextPdfConversion, $PdfMigrationStrategy,
        $StandaloneImagesAsPhotos, $IncludePdfPagePhotos, $DocConversionTempDir,
        $DisallowedForConvert, $EmbeddableImageExtensions, $UpdateOnMatch, $UpdateStrategy)
    $PSBoundParameters | Export-Clixml -LiteralPath (Join-Path $WorkDir 'captured.xml')
    [pscustomobject]@{ Status = 'Captured' }
}
'@
        $script:driverArgs = @{
            TargetDocumentDir = $script:inputRoot
            DocConversionTempDir = (Join-Path $TestDrive 'conversion')
            HuduBaseUrl = 'https://example.test'
            HuduApiKey = 'fake-test-key'
            DestinationStrategy = 'GlobalKB'
            SourceStrategy = 'TopLevel'
            PersistTempfiles = $true
        }
    }

    It 'defaults both photo options to false' {
        Set-Content -LiteralPath (Join-Path $script:driverRoot 'files-config.ps1') -Value '# no overrides'
        & (Join-Path $script:driverRoot 'Files-For-Hudu.ps1') @script:driverArgs 6>$null
        $captured = Import-Clixml -LiteralPath (Join-Path $script:driverRoot 'captured.xml')
        $captured.StandaloneImagesAsPhotos | Should -BeFalse
        $captured.IncludePdfPagePhotos | Should -BeFalse
    }

    It 'uses configured photo defaults when parameters are omitted' {
        Set-Content -LiteralPath (Join-Path $script:driverRoot 'files-config.ps1') -Value '$StandaloneImagesAsPhotos = $true; $IncludePdfPagePhotos = $true'
        & (Join-Path $script:driverRoot 'Files-For-Hudu.ps1') @script:driverArgs 6>$null
        $captured = Import-Clixml -LiteralPath (Join-Path $script:driverRoot 'captured.xml')
        $captured.StandaloneImagesAsPhotos | Should -BeTrue
        $captured.IncludePdfPagePhotos | Should -BeTrue
    }

    It 'lets explicit false and true override config independently and forwards disabled updates' {
        Set-Content -LiteralPath (Join-Path $script:driverRoot 'files-config.ps1') -Value '$StandaloneImagesAsPhotos = $true; $IncludePdfPagePhotos = $false'
        & (Join-Path $script:driverRoot 'Files-For-Hudu.ps1') @script:driverArgs -StandaloneImagesAsPhotos $false -IncludePdfPagePhotos $true -UpdateFilesOnMatch $false 6>$null
        $captured = Import-Clixml -LiteralPath (Join-Path $script:driverRoot 'captured.xml')
        $captured.StandaloneImagesAsPhotos | Should -BeFalse
        $captured.IncludePdfPagePhotos | Should -BeTrue
        $captured.UpdateStrategy | Should -Be none
        $captured.UpdateOnMatch | Should -BeFalse
    }
}
