# upload original documents alongside converted counterparts
$WorkDir = $PSScriptRoot
$includeOriginals=$includeOriginals ?? $true
$TargetDocumentDir = $TargetDocumentDir ?? $(read-host "which directory contains documents")

$DocConversionTempDir = $DocConversionTempDir ?? "c:\conversion-tempdir"
$hudubaseurl = $hudubaseurl ?? $(read-host "Enter Hudu URL")
$huduapikey = $huduapikey ?? $(read-host "Enter Hudu API Key")
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
# try {Stop-LibreOffice} catch {}



$destinationStrategy = $(select-objectfromlist -message "will each file be for a unique company?" -objects @("various companies","same company","global KB"))
$sourceStrategy = $(select-objectfromlist -message "do you want to look for source documents in $targetdocumentdir recursively?" -objects @("search recursively","only first-level"))

$sourceObjects = $(if ($sourceStrategy -eq "only first-level"){$(get-childitem -path $TargetDocumentDir -depth 0)} else {$(get-childitem -path $TargetDocumentDir -recurse)})

$sourceItemsCanBeFolders = Select-ObjectFromList -Message "Include directories?" -Objects @(
    "directories included",
    "only process files"
)

if ($sourceItemsCanBeFolders -eq "only process files") {
    $sourceObjects = $sourceObjects | Where-Object { -not $_.PSIsContainer }
} 


if ('same company' -eq $destinationStrategy){
  $sameCompanyTarget = select-objectfromlist -objects $(get-huducompanies) "Which company to attribute documents in $TargetDocumentDir to? Enter 0 for Global-KB."
} else {
  $sameCompanyTarget = $null
}
$results = @()


foreach ($sourceObject in $sourceObjects){
  $articleFromResourceRequest = @{
    ResourceLocation = $(Get-Item -LiteralPath $sourceObject)
  }
  if ($destinationStrategy -eq "various companies") {
    $target = $(select-objectfromlist -objects $(get-huducompanies) "Which company to attribute $($articleFromResourceRequest.ResourceLocation) to? Enter 0 for Global-KB.")
    if ($target && $target.name){
      $articleFromResourceRequest.companyName = $target.name
    }
  } elseif ('same company' -eq $destinationStrategy){
    if ($null -ne $sameCompanyTarget -and $sameCompanyTarget.name){
      $articleFromResourceRequest.companyName = $sameCompanyTarget.name
    }
  }
  $result = New-HuduArticleFromLocalResource @articleFromResourceRequest
  $results+=$result

}
