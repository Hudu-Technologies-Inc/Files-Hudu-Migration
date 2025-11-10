$UseAZVault = $false                                             # use AZ Keyvault if running noninteractively
$AzVault_HuduSecretName = "HuduAPIKeySecretName"                 # Name of your secret in AZure Keystore for your Hudu API key
$AzVault_Name           = "MyVaultName"                          # Name of your Azure Keyvault
$HuduBaseUrl = "https://my-huduinstance.huducloud.com"
$TargetDocumentDir = "c:\yourdocsfolder"                         # folder containing documents for Hudu KB articles
$DocConversionTempDir = "c:\temporary"                           # temp folder for file conversion
$includeOriginals= $true                                         # upload original documents alongside converted counterparts
$updateOnMatch = $false                                          # if article with same name exists, skip (false) or update (true)
$globalkbfoldername = ""                                         # name of subdirectory in $TargetDocumentDir that contains global kb articles 


# image formats that can be uploaded and referenced in Hudu KB articles
$EmbeddableImageExtensions = @(
    ".jpg", ".jpeg",  # JPEG
    ".png",           # Portable Network Graphics
    ".gif",           # GIF (including animated)
    ".bmp",           # Bitmap (support varies by browser)
    ".webp",          # WebP (modern, compressed)
    ".svg",           # Scalable Vector Graphics
    ".apng",          # Animated PNG (limited support)
    ".avif",          # AV1 Image File Format (modern)
    ".ico",           # Icon files (used in favicons)
    ".jfif",          # JPEG File Interchange Format
    ".pjpeg",         # Progressive JPEG
    ".pjp"            # Alternative JPEG extension
)

# formats that we don't convert and instead are uploaded as standalone documents
$DisallowedForConvert = [System.Collections.ArrayList]@(
    ".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a",
    ".dll", ".so", ".lib", ".bin", ".class", ".pyc", ".pyo", ".o", ".obj",
    ".exe", ".msi", ".bat", ".cmd", ".sh", ".jar", ".app", ".apk", ".dmg", ".iso", ".img",
    ".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".tgz", ".lz",
    ".mp4", ".avi", ".mov", ".wmv", ".mkv", ".webm", ".flv",
    ".psd", ".ai", ".eps", ".indd", ".sketch", ".fig", ".xd", ".blend", ".vsdx",
    ".ds_store", ".thumbs", ".lnk", ".heic", ".eml", ".msg", ".esx", ".esxm"
)

if ($true -eq $UseAZVault){
    $huduapikey = $(read-host "Enter Hudu API Key")
    clear-host 
} else {
    if ($true -eq $UseAZVault) {
    foreach ($module in @('Az.KeyVault')) {if (Get-Module -ListAvailable -Name $module) { Write-Host "Importing module, $module..."; Import-Module $module } else {Write-Host "Installing and importing module $module..."; Install-Module $module -Force -AllowClobber; Import-Module $module }}
    if (-not (Get-AzContext)) { Connect-AzAccount };
        $HuduAPIKey = "$(Get-AzKeyVaultSecret -VaultName "$AzVault_Name" -Name "$AzVault_HuduSecretName" -AsPlainText)"
    }
}


. .\pile-of-files-migrate.ps1