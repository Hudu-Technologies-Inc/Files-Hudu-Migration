# upload original documents alongside converted counterparts
$includeOriginals= $true
# if article with same name exists, skip (false) or update (true)
$updateOnMatch = $false

# folder containing documents for Hudu KB articles
$StartDir = "c:\yourdocsfolder"
# temp folder for file conversion
$tempDir = "c:\temporary"

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