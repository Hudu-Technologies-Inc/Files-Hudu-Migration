
# Pile of Files Sync / Migration
*A standardized internal foundation for converting files, directories, and authenticated web resources into Hudu Articles.*

This project provides a unified, highly-extensible workflow for generating Hudu articles from many different types of source material (directories, files, PDFs, Office docs, HTML pages, authenticated web content, etc.).

It uses the **community-supported “articles anywhere” methods**, but this wrapper adds additional structure, pre-processing, extract/convert logic, and guardrails.

---

## ⚠️ Important Note on Power & Guardrails
This tool is extremely powerful:

- It can recursively ingest large directory trees
- Convert files in bulkx
- Capture authenticated web pages
- Create hundreds or thousands of Hudu articles automatically
- Attach extracted images, upload originals, and expand storage usage quickly

Because of this, *default guardrails are in place* (file size caps, recursion limits, max item counts, safety prompts).

This project is **not** intended for direct end-user use without additional UI, workflow limits, or policy logic.
It serves as a *foundation* for safe, controlled migration and ingestion processes.

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

This raises complexity and security implications—so it is not enabled by default.

---

## Supported Resource Types

### 1. Directory Listing → Article
If the provided resource is a **directory** containing at least one file under 100 MB:

- Images are uploaded and displayed in a *gallery* section
- Files are uploaded as attachments
- Links are auto-generated to each item

---

### 2. Recursing Through a Path (Bulk File → Article)


---

### 3. Office Documents (docx, pptx, xlsx)
Non-PDF documents go through LibreOffice conversion.

---

### 4. Plaintext Files & Scripts
Rendered as:

- Syntax-highlighted code
- Original file uploaded as attachment

---

### 5. Directories Containing HTML Files
If a directory contains an `.html` file, that becomes the article body.

---

### 6. Web URI / Link → Article
HTML is downloaded, images fetched if possible, and converted as an HTML-based article.

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
