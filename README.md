# Pile of Files Sync/Migration

*for now, this is just an internal basis /starting point for migrating authenticated webpages, directory listings, or any document into Hudu as an article*
The 'articles anywhere' methods that this uses, however, are community-published and readily available.

It's too much power to be weilded by users without more gaurdrails, however. So, for now, it can just be a standardized basis for creating articles from a myriad of different source materials.

###### A Word On Idemnopotence

Articles are created/updated wth idemnopotence and any embedded images are reused. HOWEVER, directory listings will update the article of same-name with ALL NEW FILES. this can result in a lot of storage space being wasted. Better to download the resource, upload new resource if file hash has changed, and remvoe older resource. However, downloading the file in question does require session cookie / web authentication, so this is not built in (and part of the reason it maybe shouldnt be available to users). Otherwise, if they sync their C:\ drive 10x recursively, they might have 10 versions of every 1000 files that match their criteria

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


##### Plaintext files and Scripts

It depends on the content of plaintext files, but most often their contents are read as a 'codeblock', with the original file as an attachment 

<img width="1299" height="890" alt="image" src="https://github.com/user-attachments/assets/5b63485e-a7e3-4a0f-8c81-b7bf62131d72" />


##### If directories-as-articles is enabled and your directory contains an .html file-
it will process the article as the html article and any images therein as images to attache to article (and replace links/src for)

##### Web URI / Link

you can also capture webpages and set those as articles, even authenticated webpages.
<img width="691" height="542" alt="image" src="https://github.com/user-attachments/assets/fd553a85-c1dd-4629-8943-8fe960f900bb" />

Your mileage may vary without some pre-processing of html results, but essentially it just downloads the html, any image resources that can be downloaded, and treats it as a 'directory listing containing html file'

<img width="901" height="704" alt="image" src="https://github.com/user-attachments/assets/40333c00-e1a4-4726-bffc-47a51625c072" />



--- 

### Requirements

- Powershell 7.5.1 or newer
- Hudu 2.39.4 or newer
- Windows machine for script execution
- Hudu API Key
- LibreOffice [for non-pdf documents conversion]
- Documents to upload to Hudu

---


