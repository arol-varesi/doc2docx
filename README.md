# DOC → DOCX Converter (PowerShell + Microsoft Word)


PowerShell script for **massive and automatic conversion** of Microsoft Word files
from the legacy **`.doc` (Word 97–2003)** format to the modern **`.docx`** format.

The conversion is performed using **Microsoft Word COM Automation**, ensuring
maximum fidelity for layouts, tables, images, and OLE content.

---

## Requirements

- **Windows** operating system
- **Microsoft Word installed** (any modern version)
- PowerShell 5.x or higher
- Read/write permissions on the folders involved

> ⚠️ The script **does not work** on machines without Microsoft Word.
> In server environments without Office, LibreOffice headless is recommended.

---

## Key Features

- Automatic conversion of **all `.doc` files**
- **Recursive** support for subfolders
- Exclusion of Word temporary files (`~$*`)
- Options to:
  - overwrite the original structure
  - or replicate it to a separate destination folder
- Silent execution (Word not visible, no popups)

---

## Usage

### Syntax

```powershell
.\Doc2Docx.ps1 -SourceFolder <source_folder> [-DestFolder <destination_folder>]
```

### Parameters

| Name           | Required | Description                        |
| -------------- | -------- | ---------------------------------- |
| `SourceFolder` | YES      | Folder containing the `.doc` files |
| `DestFolder`   | NO       | Output folder for `.docx` files    |

If `DestFolder` **is not specified**, the `.docx` files will be created
in the **same folder** as the original `.doc` files.

---

### Examples

#### Convert keeping the same folder

```powershell
.\Doc2Docx.ps1 -SourceFolder "C:\temp\olddocs"
```

#### Convert replicating the folder structure

```powershell
.\Doc2Docx.ps1 -SourceFolder "C:\temp\olddocs" -DestFolder "C:\temp\newdocx"
```

---

## How It Works Internally

1. The script recursively scans the source folder
2. For each `.doc` file:
    - opens the document in **read-only mode**
    - uses `SaveAs2` with format `wdFormatXMLDocument (16)`
    - saves the file as `.docx`

3. Each document is closed properly
4. The Word COM instance is released at the end of processing

`SaveAs2` is used instead of `SaveAs` to avoid COM binding issues
with PowerShell (e.g., `PSObject` errors).

---

## Reliability and Stability

- Word runs in invisible mode
- Alerts and popups disabled
- Explicit release of COM objects (`ReleaseComObject`)
- Forced garbage collection at script end
- Prevention of "Word.exe zombie" processes in memory

---

## Important Notes

- `.docx` files **do not automatically overwrite** existing ones
  (can be easily added if needed).
- VBA macros present in `.doc` files **are not migrated**
- Files with special protections may not be converted

---

## License

This project is distributed under the MIT license. See the [LICENSE](LICENSE) file for details.
