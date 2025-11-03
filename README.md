# Pile of Files Sync/Migration

**Sync or Create Articles from Files to Hudu**

###### *PDFs, Word Documents, Zipfiles, come one, come all*

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

You'll want to prepare your target directory and files like so:

```
TargetDocumentDir
    company-name-1
        annual-report.xlsx
        note-from-diana.docx
        deprovisioning-guide.pdf
        user-provisioning-wizard.ps1
    company-name-2
        security-information.pdf
        agent-install.exe
        FLIR-cameras.txt
        buildings-info.docx
```
Each folder in `$TargetDocumentDir` corresponds to a company and each company folder is only searched one level deep (no nested folders for now), so just be mindful of your directory tree

If you set a `$globalkbfoldername`, documents within this folder will be placed in Global/Central KB


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

