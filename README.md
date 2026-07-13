
# Files → Hudu Articles Sync / Migration

This project provides a unified, highly-extensible workflow for generating Hudu articles from many different types of source material (directories, files, PDFs, Office docs, HTML pages, authenticated web content, etc.).

It uses the **community-supported [Articles-From-Anything](./Client-Libraries/Articles-From-Anything/README.md) methods**, but this wrapper adds additional structure, pre-processing, extract/convert logic, sane defaults, and guardrails.

*It supports remote filestores as well, including SharePoint, OneDrive, and any resource accessible in Windows Explorer through a mapped drive.*

---

## Requirements
| Component | Required Version |
|----------|------------------|
| PowerShell | 7.5.1+ |
| Hudu | 2.40.2+ |
| Windows | Any |
| Hudu API Key | — |
| LibreOffice | Latest MSI* |
| Documents | ≤100 MB each |

#### If Libreoffice is not installed on your device yet, an installer will be launched for you.
---

## Summary
This project serves as a **safe foundation** for:

- Large-scale document migration
- Directory ingestion
- File-to-article conversion
- Syncing Document Folders, Network Shares, Mounted Cloud Storages

There is also now a GUI which allows for easy, one-off file syncs, conversions, article updates.

---

# GUI Section

To get the latest GUI, simply download the exe from releases section. The exe itself includes a self-extracting folder with everything we need, so that is completely sufficient.

As you can see, the frontend is just a friendly and simple way of doing the same article syncs/conversions/creates as the CLI.

<img width="1245" height="960" alt="image" src="https://github.com/user-attachments/assets/345af08e-bf9a-4ff5-9b23-3ace88b6a938" />

There are friendly defaults included here, which can *optionally* be saved to a local settings file. It will not save your API key to this settings file and does not directly pass in your API key for security reasons. This means that after you close your GUI session, you'll have to re-enter your API key. You do not need to save settings in order to run files migration, but it does save time if you plan on undergoing the same operation more than once.

It is generally reccomended to use recurse strategy so that any files you point at can be effectively read-in and tracked, but this is not necessary.

like the CLI invocation style, temporary files are cleaned up on exit and all output is logged to file.
It's quite simple. Fill in the form. When you have entered your API key, the invocation parameters that will be used are shown in preview (and can be refreshed)
old or new log files can be accressed next to the launch button

<img width="1199" height="862" alt="image" src="https://github.com/user-attachments/assets/09d5e803-b757-4c93-bb96-83a2bd7b2241" />

# CLI Section

To get the latest CLI, clone or download a zipfile of this repo.

## Script Parameters

| Parameter | Description |
|----------|-------------|
| **TargetDocumentDir** *(required)* | File or directory containing the articles to process. |
| **DocConversionTempDir** | Temporary directory for PDF/HTML/LibreOffice conversions. Defaults to `Docs-Temp` under the project directory. |
| **filter** | Case-insensitive file or directory filter. Supports wildcards (e.g., `*.pdf`, `keep*`). |
| **updateFilesOnMatch** | When `true`, matched existing articles can be updated according to `UpdateStrategy`. When `false`, matched articles are skipped. Default: **true**. |
| **UpdateStrategy** | Match-update comparison mode: `filehash`, `date`, or `none`. Default: **filehash**. If `updateFilesOnMatch` is `false`, this is forced to `none` for processing. |
| **DestinationStrategy** | Determines where articles are added: `GlobalKB`, `SameCompany`, or `VariousCompanies`. Optional; prompts if omitted. |
| **SourceStrategy** | Controls discovery depth: `TopLevel` only scans the first level; `Recurse` scans subdirectories up to `MaxDepth`. Optional; prompts if omitted unless `IncludeDirectories` is used. |
| **IncludeDirectories** | Treat top-level directories as resources that become directory-listing articles. When enabled, discovery is forced to `TopLevel`. |
| **IncludeOriginals** | Include original documents in the article along with converted versions. Default: **true**. |
| **PdfMigrationStrategy** | PDF conversion mode: `plaintext`, `html`, or `rasterized`. Defaults to `plaintext`. |
| **MaxItems** | Maximum number of files/directories allowed in a batch. Default: **500**. |
| **MaxTotalBytes** | Maximum allowed total size of incoming documents. Default: **5 GB**. |
| **MaxDepth** | Maximum recursion depth when using `Recurse`. Default: **5** levels. |
| **PersistTempfiles** | Keep conversion temp files after the run instead of deleting `DocConversionTempDir`. Default: **false**. |
| **HuduBaseUrl** | Hudu base URL. If omitted, the script prompts. Useful for frontend or unattended invocation. |
| **HuduAPIKeySecure** | SecureString containing Hudu API key
| **SameCompanyName** | Company name to use when `DestinationStrategy` is `SameCompany`. If omitted or not matched, no company is assigned and the article path behaves as Global KB. |
| **ConvertExtensions** | Optional explicit conversion allow-list, e.g. `@(".docx",".xlsx",".csv")`. If non-empty, matching extensions are removed from the effective deny-list before processing and files outside this list are uploaded as attachment-only articles. If empty, falls back to `files-config.ps1`; if config is also empty, `DisallowedForConvert` controls conversion eligibility. |
| **UploadAsArticleExtensions** | Optional explicit force-upload list, e.g. `@(".xlsx",".csv")`. Matching non-image files are added to the effective deny-list before processing and uploaded as attachment-only articles instead of converted. If empty, falls back to `files-config.ps1`; if config is also empty, no additional extensions are forced to upload-only. |

### Parameter and Config Precedence

For `ConvertExtensions` and `UploadAsArticleExtensions`, non-empty script parameters are explicit instructions and win over `files-config.ps1`. If the parameter is empty or omitted, the matching value from `files-config.ps1` is used. If both are empty, the script uses its normal defaults.

Before each call to `New-HuduArticleFromLocalResource`, the main script resolves these preferences into effective `DisallowedForConvert` and `EmbeddableImageExtensions` values for that file. Explicit `UploadAsArticleExtensions` has highest priority. Explicit `ConvertExtensions` is next and can allow an extension even when config-level upload or deny preferences would otherwise block it. Config-level `UploadAsArticleExtensions`, config-level `ConvertExtensions`, and then `DisallowedForConvert` apply after that.

File handling order is: `SkipEntirely` removes files from discovery first; size, recursion, and `filter` rules select the batch; then per-file strategy is chosen. Images listed in `EmbeddableImageExtensions` are embedded in articles. PDFs use the PDF conversion path. Scripts are rendered as code-style articles. Other files are converted unless blocked by `DisallowedForConvert`, excluded by a non-empty `ConvertExtensions` allow-list, or included in `UploadAsArticleExtensions`.

Extension lists are normalized before use: values are lowercased, leading dots are added when missing, blank values are ignored, and duplicates are removed.


---

> **Permissions Notice**
>
> Some scripts may require elevated permissions. If you encounter access-related errors, consider launching PowerShell (`pwsh`) with **Run as Administrator**.
>
> Please note that administrative privileges do not override Windows Rights Management or similarly enforced file protection mechanisms.

---

## Examples:

#### Recursing Through a Path of various document types (Bulk File → Article)

```powershell
. .\Files-For-Hudu.ps1 -TargetDocumentDir C:\Users\Administrator\Downloads\ -SourceStrategy Recurse
```

#### Specifying only Docx files in a SharePoint or OneDrive mount
```powershell
. .\Files-For-Hudu.ps1 -TargetDocumentDir X:\Billing\ -SourceStrategy Recurse -Filter "*.docx"
```

#### Adding Articles for Critical PNG images on a Digital Camera
```powershell
. .\Files-For-Hudu.ps1 -TargetDocumentDir N:\DCIM -SourceStrategy Recurse -Filter "*.png"
```

---

## Document Conversion and Resulting Types of Articles

It all comes down to specified resource... We'll iterate through the specified resource if it's a directory and handle every file or directory in the best way possible. 

### PDF Format

If you wanted to upload all eligible .pdf documents in c:\path to a single company, for example, you might do something like this, below.
```
 . .\Files-For-Hudu.ps1 -TargetDocumentDir C:\Path\ -SourceStrategy Recurse -Filter "*.pdf"
```

By default PDFs use `plaintext`, which runs Poppler `pdftotext` and creates a searchable text article with the original PDF attached. You can also set `-PdfMigrationStrategy html` to use `pdftohtml` and extract HTML/images, or `-PdfMigrationStrategy rasterized` to render every page as a PNG image with `pdftoppm` for maximum visual fidelity.

<img width="1224" height="448" alt="image" src="https://github.com/user-attachments/assets/e1298f57-0e3e-4903-8425-ccbcafd109a7" />

### Typical Office Formats

similar to pdf documents, common (and even uncommon) document formats are converted to html and have any embedded images extracted during processing
A docx file will look similar to the PDF example, above, but may have slightly more basic formatting for paragraphs and sections

<img width="1301" height="496" alt="image" src="https://github.com/user-attachments/assets/edbf8e08-02d1-4fd7-8912-2d3fc3ea6dcc" />
Excel and CSV files are converted into an HTML table


### Plaintext files and Scripts

It depends on the content of plaintext files, but most often their contents are read as a 'codeblock', with the original file as an attachment 

<img width="1299" height="890" alt="image" src="https://github.com/user-attachments/assets/5b63485e-a7e3-4a0f-8c81-b7bf62131d72" />


##### If directories-as-articles is enabled and your directory contains an .html file-
it will process the article as the html article and any images therein as images to attache to article (and replace links/src for)

---

## Idempotence, Updates & Storage Considerations

Articles are created and updated idempotently. This is facilitated with both attachment/embed file-hasing and content-hashing with SHA256 algorithm. Embedded images within converted documents are reused when possible. If an attached/embed file is updated locally and remote file has a different file hash, the remote file will be updated if the local file's date is newer.

---

## Personalizing File Type Preferences

To configure any files that you wish to skip conversion for, upload as standalone documents, or image types to ignore, simply add/remove them to/from the respective list in your `files-config.ps1` file in this project directory.

`EmbeddableImageExtensions` - these are image files that can be loaded into Hudu inside articles. Chances are, you won't need to/want to fuss with these. You aren't likely to encounter some of the more exotic image formats here, but they do work, should you encounter them directly in a folder or subsequent to a file conversion.

`DisallowedForConvert` These are common non-document extensions/formats that we will not try to convert to a READABLE Hudu Article. However, if we encounter these while working in a directory listing or a per-file basis, they can be uploaded to Hudu as an article attachment.

If there is a specific format that you don't like to convert, like xlsx or xlsm, for example, you can add it to this array.

`SkipEntirely` is an array of extensions that we simply want to try to avoid touching. These may be partially downloaded files, sensitive files, or files that we simply don't need or want in Hudu. There are some sane defaults in-place if you aren't sure.

`ConvertExtensions` is optional. When populated, it becomes a conversion allow-list: only matching extensions, plus embeddable images, are converted or embedded. Everything else is uploaded as an attachment-only article. Leave it empty to use `DisallowedForConvert` as the broader conversion gate.

`UploadAsArticleExtensions` is optional. When populated, matching non-image extensions are forced to upload-only article behavior even if LibreOffice could convert them.

`PdfMigrationStrategy` is optional (default `plaintext`). Use `plaintext` for Poppler `pdftotext`, `html` for Poppler `pdftohtml`, or `rasterized` for Poppler `pdftoppm` page images. `PlainTextPdfConversion` remains supported for older callers: `$true` maps to `plaintext`, and `$false` maps to `html` when `PdfMigrationStrategy` is omitted.

## Community & Socials

[![Hudu Community](https://img.shields.io/badge/Community-Forum-blue?logo=discourse)](https://community.hudu.com/)
[![Reddit](https://img.shields.io/badge/Reddit-r%2Fhudu-FF4500?logo=reddit)](https://www.reddit.com/r/hudu)
[![YouTube](https://img.shields.io/badge/YouTube-Hudu-red?logo=youtube)](https://www.youtube.com/@hudu1715)
[![X (Twitter)](https://img.shields.io/badge/X-@HuduHQ-black?logo=x)](https://x.com/HuduHQ)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Hudu_Technologies-0A66C2?logo=linkedin)](https://www.linkedin.com/company/hudu-technologies/)
[![Facebook](https://img.shields.io/badge/Facebook-HuduHQ-1877F2?logo=facebook)](https://www.facebook.com/HuduHQ/)
[![Instagram](https://img.shields.io/badge/Instagram-@huduhq-E4405F?logo=instagram)](https://www.instagram.com/huduhq/)
[![Feature Requests](https://img.shields.io/badge/Feedback-Feature_Requests-brightgreen?logo=github)](https://hudu.canny.io/)


