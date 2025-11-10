# Pile of Files Sync/Migration

*for now, this is just an internal basis /starting point for migrating authenticated webpages, directory listings, or any document into Hudu as an article*

It's too much power to be weilded by users without more gaurdrails, however.

**Sync or Create Articles from Files to Hudu**

###### *PDFs, Word Documents, Zipfiles, come one, come all*

---
It all comes down to specified resource... 

##### Directory Listing article
If your specified resource is a directory containing at least one file under 100mb

if directory is provided as a resource, any images in directory are uploaded and placed into a 'gallery' section, each linked to the full upload object.
Any non-image items within directory resource are uploaded (if under 100mb) and also linked in a section, below
<img width="617" height="971" alt="image" src="https://github.com/user-attachments/assets/38335c2f-fc6d-4f41-b8b3-84ce5f6d938b" />

If chosen directory does not include images, you just get an article with a link to attached files
<img width="921" height="244" alt="image" src="https://github.com/user-attachments/assets/2c86a1f6-b6ef-4a9a-8f37-5e6b428f1b5d" />

##### Recursing through a given path for files by filter
you can recurse a base directory with the intention of uploading each matched file (or directory) to a given company, to a new company each time, or to global kb
If you wanted to upload all eligible .pdf documents in c:\path to a single company, for example, you might
```
foreach ($f in $(get-childitem c:\path -filter "*.pdf" -File -Recurse | where-object { $_.Length -lt 100MB )){
  New-HuduArticleFromLocalResource -ResourceLocation $f.fullname -companyName "gregs supply"
}
```
Doing so will use pdftohtml to extract html/images from every found pdf, create a native article from html extracted from pdf, attach and relink images in pdf, then upload original pdf document as an attachment to the article. The extracted HTML is almost indistinguishable from the source pdf and can be edited, searched for, etc
<img width="1224" height="448" alt="image" src="https://github.com/user-attachments/assets/e1298f57-0e3e-4903-8425-ccbcafd109a7" />

##### Typical Office Formats

similar to pdf documents, common (and even uncommon) document formats are converted to html and have any embedded images extracted during processing
A docx file will look similar to the PDF example, above, but may have slightly more basic formatting for paragraphs and sections

<img width="1301" height="496" alt="image" src="https://github.com/user-attachments/assets/edbf8e08-02d1-4fd7-8912-2d3fc3ea6dcc" />

##### If directories-as-articles is enabled and your directory contains an .html file-
it will process the article as the html article and any images therein as images to attache to article (and replace links/src for)


--- 

### Requirements

- Powershell 7.5.1 or newer
- Hudu 2.39.4 or newer
- Windows machine for script execution
- Hudu API Key
- LibreOffice [for non-pdf documents conversion]
- Documents to upload to Hudu

---

### Getting Started - Target Folder/Files

If you have a target directory in mind, make sure that it is consistent and ready for migration.
During startup, you'll be able to choose

$destinationStrategy - will all articles be matched to a single company? will each article need to be matched to various companies? or will all articles be designated as global kb?

$sourceStrategy - search for article candidates recursively or at limited depth? 

$sourceItemsCanBeFolders - can a folder be considered an article? directory listing or resource folder?


### Getting Started - Environment File

Make a copy of Environment.example.ps1
```
copy-item environment.example.ps1  my-environment.ps1
```
Then edit the configuration / setup items

```
notepad.exe .\my-environment.ps1
```

Once you have your settings placed as desired, you can execute your environment file to kick things off [preferably via dot-sourcing]

```
. .\my-environment.ps1
```

