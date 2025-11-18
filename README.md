
# Pile of Files Sync / Migration

This project provides a unified, highly-extensible workflow for generating Hudu articles from many different types of source material (directories, files, PDFs, Office docs, HTML pages, authenticated web content, etc.).

It uses the **community-supported “articles anywhere” methods**, but this wrapper adds additional structure, pre-processing, extract/convert logic, and guardrails.

*You can even use it on remote filestores, sharepoint/onedrive mounts, just about anything you can access in Windows Explorer.*

---

## Configuring File Types

To configure any files that you wish to skip conversion for, upload as standalone documents, or image types to ignore, simply add/remove them to/from the respective list-

`EmbeddableImageExtensions` - these are image files that can be loaded into Hudu inside articles. Chances are, you won't need to/want to fuss with these. You aren't likely to encounter some of the more exotic image formats here, but they do work, should you encounter them directly in a folder or subsequent to a file conversion.

`DisallowedForConvert` These are common non-document extensions/formats that we will not try to convert to a READABLE Hudu Article. However, if we encounter these while working in a directory listing or a per-file basis, they can be uploaded to Hudu as an article attachment.

If there is a specific format that you don't like to convert, like xlsx or xlsm, for example, you can add it to this array.

`SkipEntirely` is an array of extensions that we simply want to try to avoid touching. These may be partially downloaded files, sensitive files, or files that we simply don't need or want in Hudu.

---

## Script Params / Options

`TargetDocumentDir` - This is the only required parameter - the directory where your desired articles are located

`DocConversionTempDir` - Temporary Directory for File Conversion (Pdf2Html and LibreOffice)

`filter` - Case-Insensitive File (or directory) naming filter [can use wildcards]. For example, to target just PDF files, you might specify `-filter "*.pdf"` or to specify directories starting with 'keep', you might specify `-filter "keep*"`

`DestinationStrategy` - This is for how you want to add articles to Hudu - If you want to upload all docs in Global/Central KB, you can specify `GlobalKb`. To add articles under a single company, specify `SameCompany`. Otherwise, you can specify `VariousCompanies`. This param is optional, so you will be prompted for destination info if not provided.

`SourceStrategy` - To allow for recursing into sub-directories when searching for resources, you can specify this with `-SourceStrategy Recurse`, otherwise to stay at a single level in your `TargetDocumentDir`, you can specify `-SourceStrategy Recurse` or omit this param and you will be prompted.

`IncludeDirectories` - Whether or not to treat directories as a resource. Reccomended to be used WITHOUT a recursive source strategy. See below per-document examples for what this looks like.

`IncludeOriginals` - Include original documents attached to articles, alongside converted counterparts? Default is true.

`MaxItems` - Max allowed Files/Directories to handle at once. Default is 500. If your `TargetDocumentDir`, `SourceStrategy`, and/or `filter` results in more than this number of files, you will be prompted to continue or exit and refine your command.

`MaxTotalBytes` - Max allowed total filesize for pre-converted documents/attachments. Default is 5gb

`MaxDepth` - Max recursion depth (if using 'Recurse' for your `SourceStrategy`)- default is 5 levels of recursion

---

## Supported Resource Types

### 1. Directory Listing → Article
If the provided resource is a **directory** containing at least one file under 100 MB:

- Images are uploaded and displayed in a *gallery* section
- Files are uploaded as attachments
- Links are auto-generated to each item

---

### 2. Recursing Through a Path (Bulk File → Article)

```powershell
. .\pile-of-files-migrate.ps1 -TargetDocumentDir C:\Users\Administrator\Downloads\ -SourceStrategy Recurse
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
 . .\pile-of-files-migrate.ps1 -TargetDocumentDir C:\Path\ -SourceStrategy Recurse -Filter "*.pdf"
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

## Requirements
| Component | Required Version |
|----------|------------------|
| PowerShell | 7.5.1+ |
| Hudu | 2.39.4+ |
| Windows | Any |
| Hudu API Key | — |
| LibreOffice | Latest MSI |
| Documents | ≤100 MB each |

---

## Summary
This project serves as a **safe foundation** for:

- Large-scale document migration
- Directory ingestion
- File-to-article conversion
- HTML normalization & relinking

Guardrails prevent:

- Recursion explosions
- Upload storms
- Storage blowouts
- Redundant multi-sync loops
