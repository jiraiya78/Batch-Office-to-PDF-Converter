# Office to PDF Batch Converter

**Batch convert Microsoft Word and Excel documents to PDF with flexible folder structure options and a modern, easy-to-use interface.**  
Created by **Mohd Azri Afifi**, 2025

![App Screenshot](images/Screenshot%202025-07-04%20192636.png)

---

## Features

- **Batch Conversion:** Converts all `.docx`, `.doc`, `.xlsx`, and `.xls` files in a selected folder (and its subfolders) to PDF.
- **Flexible Output:**  
  - Save PDFs in the same folder as the source files.
  - Or, save all PDFs in a dedicated `PDF` folder at the root—optionally preserving the original folder structure.
- **Overwrite Warnings:** Clearly warns if a PDF will overwrite an existing file.
- **Folder Awareness:** Status messages and result lists display subfolder information.
- **Progress Tracking:** Real-time logs, progress bar, and elapsed time.
- **Drag & Drop:** Easily drag a folder from Explorer into the app.
- **Cancel Support:** Stop conversion mid-process, safely closing all Office processes.
- **Modern, Resizable UI:** Clean appearance with minimum size set for usability.

---

## Installation & Requirements

- **Platform:** Windows 10/11  
- **Requires:**  
  - [Microsoft Word](https://www.microsoft.com/en-us/microsoft-365/word) and [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel) (for conversion)
  - No Python required for `.exe` release (see below)

---

## Usage

### 1. Download

- [Download the latest release](https://github.com/yourusername/yourrepo/releases/latest/download/office2pdf.exe)  
  *(Replace with your actual repo link)*

### 2. Run

- Double-click `office2pdf.exe`.

### 3. Convert

1. **Select a folder:** Use the Browse button or drag&drop a folder into the top path field.
2. **Choose PDF location:**  
   - “Same folder as source file” — PDFs appear next to originals.
   - “PDF folder” — All PDFs go into a new `PDF` folder. You can keep the subfolder structure or flatten all PDFs into the root `PDF` folder.
3. **Start:** Click **Convert**.  
   - A progress bar and status log will track conversion.
   - Overwrite warnings are shown in orange.
   - Stop anytime with the **Stop** button.
4. **Results:**  
   - When finished, see a summary (and `ConversionLog.txt` in your folder).

---

## Screenshots

*(Add screenshots here if available, e.g. UI overview, progress, and result logs)*

---

## Credits

- **Author:** Mohd Azri Afifi
- **Email:** [azri.kun@gmail.com](mailto:azri.kun@gmail.com)
- **Year:** 2025
- **Version:** v1.2

---

## License

All rights reserved.  
For questions or licensing, contact [azri.kun@gmail.com](mailto:azri.kun@gmail.com).

---

## Troubleshooting

- **Missing `tkinterdnd2` error?**
  - If building your own executable, ensure you add `--hidden-import=tkinterdnd2` to your PyInstaller command.
- **Office COM errors?**
  - Make sure Microsoft Word and Excel are installed and not already running with open dialogs.
- **Other issues?**
  - Please open an issue on GitHub or email the author.
