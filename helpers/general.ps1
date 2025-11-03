
function Normalize-Text {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $s = $s.Trim().ToLowerInvariant()
    $s = [regex]::Replace($s, '[\s_-]+', ' ')  # "primary_email" -> "primary email"
    # strip diacritics (prénom -> prenom)
    $formD = $s.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $formD.ToCharArray()){
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne
            [System.Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
    }
    ($sb.ToString()).Normalize([System.Text.NormalizationForm]::FormC)
}
function Compare-StringsIgnoring {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$A,
        [Parameter(Mandatory)] [string]$B,
        $ignore = @(
                '\bthe\b',
                '\borg\b',
                '\binc\b',
                '\bpc\b',
                '\band\b',
                '\bltd\b',
                '[\.,/&]'
            ))
    function _Normalize($s) {

        if (-not $s) { return '' }
        $t = Normalize-Text $s
        $t = $t -replace '\p{P}+', ''

        foreach ($pattern in $ignore) {
            $t = $t -replace $pattern, ''
        }
        $t = ($t -replace '\s+', ' ').Trim()
        return $t
    }

    $normA = _Normalize $A
    $normB = _Normalize $B

    return ($normA -eq $normB)
}
function Get-Similarity {
    param([string]$A, [string]$B)

    $a = [string](Normalize-Text $A)
    $b = [string](Normalize-Text $B)
    if ([string]::IsNullOrEmpty($a) -and [string]::IsNullOrEmpty($b)) { return 1.0 }
    if ([string]::IsNullOrEmpty($a) -or  [string]::IsNullOrEmpty($b))  { return 0.0 }

    $n = [int]$a.Length
    $m = [int]$b.Length
    if ($n -eq 0) { return [double]($m -eq 0) }
    if ($m -eq 0) { return 0.0 }

    $d = New-Object 'int[,]' ($n+1), ($m+1)
    for ($i = 0; $i -le $n; $i++) { $d[$i,0] = $i }
    for ($j = 0; $j -le $m; $j++) { $d[0,$j] = $j }

    for ($i = 1; $i -le $n; $i++) {
        $im1 = ([int]$i) - 1
        $ai  = $a[$im1]
        for ($j = 1; $j -le $m; $j++) {
            $jm1 = ([int]$j) - 1
            $cost = if ($ai -eq $b[$jm1]) { 0 } else { 1 }

            $del = [int]$d[$i,  $j]   + 1
            $ins = [int]$d[$i,  $jm1] + 1
            $sub = [int]$d[$im1,$jm1] + $cost

            $d[$i,$j] = [Math]::Min($del, [Math]::Min($ins, $sub))
        }
    }

    $dist   = [double]$d[$n,$m]
    $maxLen = [double][Math]::Max($n,$m)
    return 1.0 - ($dist / $maxLen)
}
function Export-DocPropertyJson {
    param (
        [Parameter(Mandatory)][PSCustomObject]$Doc,
        [Parameter(Mandatory)][string]$Property,
        [int]$Depth = 45
    )

    if (-not ($Doc.PSObject.Properties.Name -contains $Property)) {
        throw "Property '$Property' does not exist on the provided document object."
    }

    $value = $Doc.$Property

    $dir  = [System.IO.Path]::GetDirectoryName($Doc.LocalPath)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Doc.LocalPath)
    $outPath = [System.IO.Path]::Combine($dir, "$base-$($Property.ToLower()).json")

    $value | ConvertTo-Json -Depth $Depth | Out-File -FilePath $outPath -Encoding UTF8

    return $outPath
}
function Get-EnsuredPath {
    param([string]$path)
    $outpath = if (-not $path -or [string]::IsNullOrWhiteSpace($path)) { $(join-path $(Resolve-Path .).path "debug") } else {$path}
    if (-not (Test-Path $outpath)) {
        Get-ChildItem -Path "$outpath" -File -Recurse -Force | Remove-Item -Force
        New-Item -ItemType Directory -Path $outpath -Force -ErrorAction Stop | Out-Null
        write-host "path is now present: $outpath"
    } else {write-host "path is present: $outpath"}
    return $outpath
}

function Write-ErrorObjectsToFile {
    param (
        [Parameter(Mandatory)]
        [object]$ErrorObject,

        [Parameter()]
        [string]$Name = "unnamed",

        [Parameter()]
        [ValidateSet("Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White")]
        [string]$Color
    )

    $stringOutput = try {
        $ErrorObject | Format-List -Force | Out-String
    } catch {
        "Failed to stringify object: $_"
    }

    $propertyDump = try {
        $props = $ErrorObject | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
        $lines = foreach ($p in $props) {
            try {
                "$p = $($ErrorObject.$p)"
            } catch {
                "$p = <unreadable>"
            }
        }
        $lines -join "`n"
    } catch {
        "Failed to enumerate properties: $_"
    }

    $logContent = @"
==== OBJECT STRING ====
$stringOutput

==== PROPERTY DUMP ====
$propertyDump
"@

    if ($ErroredItemsFolder -and (Test-Path $ErroredItemsFolder)) {
        $SafeName = ($Name -replace '[\\/:*?"<>|]', '_') -replace '\s+', ''
        if ($SafeName.Length -gt 60) {
            $SafeName = $SafeName.Substring(0, 60)
        }
        $filename = "${SafeName}_error_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        $fullPath = Join-Path $ErroredItemsFolder $filename
        Set-Content -Path $fullPath -Value $logContent -Encoding UTF8
        if ($Color) {
            Write-Host "Error written to $fullPath" -ForegroundColor $Color
        } else {
            Write-Host "Error written to $fullPath"
        }
    }

    if ($Color) {
        Write-Host "$logContent" -ForegroundColor $Color
    } else {
        Write-Host "$logContent"
    }
}


function Save-HtmlSnapshot {
    param (
        [Parameter(Mandatory)][string]$PageId,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Content,
        [Parameter(Mandatory)][string]$Suffix,
        [Parameter(Mandatory)][string]$OutDir
    )

    $safeTitle = ($Title -replace '[^\w\d\-]', '_') -replace '_+', '_'
    $filename = "${PageId}_${safeTitle}_${Suffix}.html"
    $path = Join-Path -Path $OutDir -ChildPath $filename

    try {
        $Content | Out-File -FilePath $path -Encoding UTF8
        Write-Host "Saved HTML snapshot: $path"
    } catch {
        Write-ErrorObjectsToFile -Name "$($_.safeTitle ?? "unnamed")" -ErrorObject @{
            Error       = $_
            PageId      = $PageId 
            Content     = $Content
            Message     ="Error Saving HTML Snapshot"
            OutDir      = $OutDir
        }
    }
}
function Get-PercentDone {
    param (
        [int]$Current,
        [int]$Total
    )
    if ($Total -eq 0) {
        return 100}
    $percentDone = ($Current / $Total) * 100
    if ($percentDone -gt 100){
        return 100
    }
    $rounded = [Math]::Round($percentDone, 2)
    return $rounded
}   
function Set-PrintAndLog {
    param (
        [string]$message,
        [Parameter()]
        [Alias("ForegroundColor")]
        [ValidateSet("Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White")]
        [string]$Color
    )
    $logline = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $message"
    if ($Color) {
        Write-Host $logline -ForegroundColor $Color
    } else {
        Write-Host $logline
    }
    Add-Content -Path $LogFile -Value $logline
}
function Select-ObjectFromList($objects,$message,$allowNull = $false) {
    $validated=$false
    while ($validated -eq $false){
        if ($allowNull -eq $true) {
            Write-Host "0: None/Custom"
        }
        for ($i = 0; $i -lt $objects.Count; $i++) {
            $object = $objects[$i]
            if ($null -ne $object.OptionMessage) {
                Write-Host "$($i+1): $($object.OptionMessage)"
            } elseif ($null -ne $object.name) {
                Write-Host "$($i+1): $($object.name)"
            } else {
                Write-Host "$($i+1): $($object)"
            }
        }
        $choice = Read-Host $message
        if ($null -eq $choice -or $choice -lt 0 -or $choice -gt $objects.Count +1) {
            Set-PrintAndLog -message "Invalid selection. Please enter a number from above"
        }
        if ($choice -eq 0 -and $true -eq $allowNull) {
            return $null
        }
        if ($null -ne $objects[$choice - 1]){
            return $objects[$choice - 1]
        }
    }
}
function Get-YesNoResponse($message) {
    do {
        $response = Read-Host "$message (y/n)"
        $response = if($null -ne $response) {$response.ToLower()} else {""}
        if ($response -eq 'y' -or $response -eq 'yes') {
            return $true
        } elseif ($response -eq 'n' -or $response -eq 'no') {
            return $false
        } else {
            Set-PrintAndLog -message "Invalid input. Please enter 'y' for Yes or 'n' for No."
        }
    }
    while ($true)
}

function Get-ArticlePreviewBlock {
    param (
        [string]$Title,
        [string]$docId,
        [string]$Content,
        [int]$MaxLength = 200
    )
    $descriptor = "ID: $docId, titled $Title"
    $snippet = if ($Content.Length -gt $MaxLength) {
        $Content.Substring(0, $MaxLength) + "..."
    } else {
        $Content
    }

@"
Mapping Sharepoint Page $descriptor ---
Title: $Title
Snippet: $snippet
"@
}


function Get-SafeFilename {
    param([string]$Name,
        [int]$MaxLength=25
    )

    # If there's a '?', take only the part before it
    $BaseName = $Name -split '\?' | Select-Object -First 1

    # Extract extension (including the dot), if present
    $Extension = [System.IO.Path]::GetExtension($BaseName)
    $NameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($BaseName)

    # Sanitize name and extension
    $SafeName = $NameWithoutExt -replace '[\\\/:*?"<>|]', '_'
    $SafeExt = $Extension -replace '[\\\/:*?"<>|]', '_'

    # Truncate base name to 25 chars
    if ($SafeName.Length -gt $MaxLength) {
        $SafeName = $SafeName.Substring(0, $MaxLength)
    }

    return "$SafeName$SafeExt"
}
function New-HuduStubArticle {
    param (
        [string]$Title,
        [string]$Content,
        [nullable[int]]$CompanyId,
        [nullable[int]]$FolderId
    )

    $params = @{
        Name    = $Title
        Content = $Content
    }

    if ($CompanyId -ne $null -and $CompanyId -ne -1) {
        $params.CompanyId = $CompanyId
    }

    if ($FolderId -ne $null -and $FolderId -ne 0) {
        $params.FolderId = $FolderId
    }

    return (New-HuduArticle @params).article
}

function Get-SafeTitle {
    param ([string]$Name)

    if (-not $Name) {
        return "untitled"
    }
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    $decoded = [uri]::UnescapeDataString($baseName)
    $safe = $decoded -replace '[\\/:*?"<>|]', ' '
    $safe = ($safe -replace '\s{2,}', ' ').Trim()
    return $safe
}

function Test-Equiv {
    param([string]$A, [string]$B)
    $a = Normalize-Text $A; $b = Normalize-Text $B
    if (-not $a -or -not $b) { return $false }
    if ($a -eq $b) { return $true }
    $reA = "(^| )$([regex]::Escape($a))( |$)"
    $reB = "(^| )$([regex]::Escape($b))( |$)"
    if ($b -match $reA -or $a -match $reB) { return $true } 
    if ($a.Replace(' ', '') -eq $b.Replace(' ', '')) { return $true }
    return $false
}
function Get-Similarity {
    param([string]$A, [string]$B)

    $a = [string](Normalize-Text $A)
    $b = [string](Normalize-Text $B)
    if ([string]::IsNullOrEmpty($a) -and [string]::IsNullOrEmpty($b)) { return 1.0 }
    if ([string]::IsNullOrEmpty($a) -or  [string]::IsNullOrEmpty($b))  { return 0.0 }

    $n = [int]$a.Length
    $m = [int]$b.Length
    if ($n -eq 0) { return [double]($m -eq 0) }
    if ($m -eq 0) { return 0.0 }

    $d = New-Object 'int[,]' ($n+1), ($m+1)
    for ($i = 0; $i -le $n; $i++) { $d[$i,0] = $i }
    for ($j = 0; $j -le $m; $j++) { $d[0,$j] = $j }

    for ($i = 1; $i -le $n; $i++) {
        $im1 = ([int]$i) - 1
        $ai  = $a[$im1]
        for ($j = 1; $j -le $m; $j++) {
            $jm1 = ([int]$j) - 1
            $cost = if ($ai -eq $b[$jm1]) { 0 } else { 1 }

            $del = [int]$d[$i,  $j]   + 1
            $ins = [int]$d[$i,  $jm1] + 1
            $sub = [int]$d[$im1,$jm1] + $cost

            $d[$i,$j] = [Math]::Min($del, [Math]::Min($ins, $sub))
        }
    }

    $dist   = [double]$d[$n,$m]
    $maxLen = [double][Math]::Max($n,$m)
    return 1.0 - ($dist / $maxLen)
}