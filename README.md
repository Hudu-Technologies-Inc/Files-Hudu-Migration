
# Files → Hudu Articles Sync / Migration

This project provides a unified, highly-extensible workflow for generating Hudu articles from many different types of source material (directories, files, PDFs, Office docs, HTML pages, authenticated web content, etc.).

It uses the **community-supported [Articles-From-Anything](./Client-Libraries/Articles-From-Anything/README.md) methods**, but this wrapper adds additional structure, pre-processing, extract/convert logic, sane defaults, and guardrails.

*It supports remote filestores as well, including SharePoint, OneDrive, and any resource accessible in Windows Explorer through a mapped drive.*

---

## Requirements
| Component | Required Version |
|----------|------------------|
| PowerShell | 7.5.1+ |
| Hudu | 2.39.6+ |
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

---

## Script Parameters

| Parameter | Description |
|----------|-------------|
| **TargetDocumentDir** *(required)* | Directory containing the articles to process. |
| **DocConversionTempDir** | Temporary directory for PDF/HTML/LibreOffice conversions. |
| **filter** | Case-insensitive file or directory filter. Supports wildcards (e.g., `*.pdf`, `keep*`). |
| **DestinationStrategy** | Determines how articles are added to Hudu: `GlobalKb`, `SameCompany`, or `VariousCompanies`. Optional; prompts if omitted. |
| **SourceStrategy** | Controls recursion: use `Recurse` to search subdirectories; omit to stay at a single level (will prompt if missing). |
| **IncludeDirectories** | Whether to treat directories as a resource. Recommended when *not* using recursive source strategy. |
| **IncludeOriginals** | Include original documents in the article along with converted versions. Default: **true**. |
| **MaxItems** | Maximum number of files/directories allowed in a batch. Default: **500**. |
| **MaxTotalBytes** | Maximum allowed total size of incoming documents. Default: **5 GB**. |
| **MaxDepth** | Maximum recursion depth when using `Recurse`. Default: **5** levels. |

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

#### Every folder has various documents for a given company. Upload documents and create single 'directory listing' of all files as an article
```powershell
. .\Files-For-Hudu.ps1 -TargetDocumentDir Z:\Companies\ -SourceStrategy TopLevel -IncludeDirectories -DestinationStrategy 'VariousCompanies'
```

#### Adding Articles for Critical PNG images on a Digital Camera
```powershell
. .\Files-For-Hudu.ps1 -TargetDocumentDir N:\DCIM -SourceStrategy Recurse -Filter "*.png"
```

---

## Document Conversion and Resulting Types of Articles

It all comes down to specified resource... We'll iterate through the specified resource if it's a directory and handle every file or directory in the best way possible. 

### Directory Listing article
Requirement - *the specified resource is a directory containing at least one file under 100mb*

if directory is provided as a resource, any images in directory are uploaded and placed into a 'gallery' section, each linked to the upload object in Hudu.

Any non-image items within directory resource are uploaded (if under 100mb) and also linked in a section, below
<img width="617" height="971" alt="image" src="https://github.com/user-attachments/assets/38335c2f-fc6d-4f41-b8b3-84ce5f6d938b" />

If chosen directory does not include images, you just get an article with a link to attached files
<img width="921" height="244" alt="image" src="https://github.com/user-attachments/assets/2c86a1f6-b6ef-4a9a-8f37-5e6b428f1b5d" />

To enable Directory Listing option, you can include the `-IncludeDirectories` switch in your command. Be wary of using this in conjunction with -Recurse param, as you might get more than you bargained for. Best to use these in a mutually-exclusive manner.

### PDF Format

If you wanted to upload all eligible .pdf documents in c:\path to a single company, for example, you might do something like this, below.
```
 . .\Files-For-Hudu.ps1 -TargetDocumentDir C:\Path\ -SourceStrategy Recurse -Filter "*.pdf"
```

Doing so will use pdftohtml to extract html/images from every found pdf, create a native article from html extracted from pdf, attach and relink images in pdf, then upload original pdf document as an attachment to the article. The extracted HTML is almost indistinguishable from the source pdf and can be edited, searched for, etc.

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
Articles are created or updated idempotently. Embedded images within converted documents are reused when possible.

### ⚠️ Directory Listings Are *Not* Fully Idempotent
If you sync a directory multiple times:

- All new files will be re-uploaded
- Old attachments will *not* be automatically removed
- Storage will grow unnecessarily

### Why Not Automatically Remove Old Files?
Because distinguishing “obsolete” vs. “intentionally retained” attachments requires:

- File hash comparison
- Full download of server content
- Authentication cookie / session-based download logic

This raises complexity and security implications—so it is not enabled at present

---

## Personalizing File Type Preferences

To configure any files that you wish to skip conversion for, upload as standalone documents, or image types to ignore, simply add/remove them to/from the respective list in your `files-config.ps1` file in this project directory.

`EmbeddableImageExtensions` - these are image files that can be loaded into Hudu inside articles. Chances are, you won't need to/want to fuss with these. You aren't likely to encounter some of the more exotic image formats here, but they do work, should you encounter them directly in a folder or subsequent to a file conversion.

`DisallowedForConvert` These are common non-document extensions/formats that we will not try to convert to a READABLE Hudu Article. However, if we encounter these while working in a directory listing or a per-file basis, they can be uploaded to Hudu as an article attachment.

If there is a specific format that you don't like to convert, like xlsx or xlsm, for example, you can add it to this array.

`SkipEntirely` is an array of extensions that we simply want to try to avoid touching. These may be partially downloaded files, sensitive files, or files that we simply don't need or want in Hudu. There are some sane defaults in-place if you aren't sure.

## Community & Socials

[![Hudu Community](https://img.shields.io/badge/Community-Forum-blue?logo=discourse)](https://community.hudu.com/)
[![Reddit](https://img.shields.io/badge/Reddit-r%2Fhudu-FF4500?logo=reddit)](https://www.reddit.com/r/hudu)
[![YouTube](https://img.shields.io/badge/YouTube-Hudu-red?logo=youtube)](https://www.youtube.com/@hudu1715)
[![X (Twitter)](https://img.shields.io/badge/X-@HuduHQ-black?logo=x)](https://x.com/HuduHQ)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Hudu_Technologies-0A66C2?logo=linkedin)](https://www.linkedin.com/company/hudu-technologies/)
[![Facebook](https://img.shields.io/badge/Facebook-HuduHQ-1877F2?logo=facebook)](https://www.facebook.com/HuduHQ/)
[![Instagram](https://img.shields.io/badge/Instagram-@huduhq-E4405F?logo=instagram)](https://www.instagram.com/huduhq/)
[![Feature Requests](https://img.shields.io/badge/Feedback-Feature_Requests-brightgreen?logo=github)](https://hudu.canny.io/)


