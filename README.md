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

