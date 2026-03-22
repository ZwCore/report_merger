# Report Merger (Word Document Merging Tool)

This is a lightweight web application built with Flask, specifically designed to automatically merge multiple Word documents (.docx) into a master template. By recognizing specific placeholders in the template, the application can seamlessly assemble scattered report content while preserving the original formatting and hierarchical structure.

## ✨ Key Features

- **Template-driven placeholder replacement**：Use {{filename}} (e.g., {{Test Report}}) as placeholders in the main document. The application will automatically locate and insert the corresponding file (e.g., Test Report.docx).
- **Perfect format preservation (w:altChunk)**：Leveraging OpenXML’s altChunk feature to directly embed document streams, ensuring that tables, images, and complex layouts in sub-documents are preserved 100% in their original format.
- **Intelligent heading level inference**：When inserting sub-reports into the template, the tool scans nearby headings in the template and automatically generates corresponding sub-headings, ensuring a well-structured and consistent document outline.
- **Friendly missing file alerts**：If a placeholder (sub-report) referenced in the template is not uploaded, the application will mark the position with a red “To Be Added” label for easy manual follow-up.
- **One-click web interaction**：The front end provides an intuitive interface that supports uploading the template and all dependent sub-reports in batch. With one click, the merged final report is generated and downloaded automatically.
- **Standalone packaging support**：Compatible with packaging tools like PyInstaller. The application can be bundled into a single .exe file (e.g., ReportAuto.exe), allowing it to run on any Windows machine without requiring a Python environment. Double-click to launch and automatically open in a browser.

## 🛠️ Tech Stack

- **Backend**：Python, Flask, python-docx (for document parsing and manipulation), OpenXML
- **Frontend**：HTML/CSS/JavaScript (located in the templates and static directories)

## 🚀 How It Works？

1. **Prepare the template**：Create a master Word document and insert placeholders like {{SubReportName}} on separate lines where sub-reports should be included.
2. **Prepare sub-documents**：Ensure you have corresponding Word files (e.g., SubReportName.docx).
3. **Upload and merge**：Upload the template and all sub-documents via the web interface, then click merge.
4. **Underlying processing logic**：
   - Parse the template and extract all {{xyz}} placeholders.
   - Read the binary data of the corresponding sub-files.
   - Create OpenXML Parts and establish relationships within the template.
   - Insert w:altChunk tags to achieve native-level document merging.
   - Finally, clean up all placeholders automatically.

## ⚙️ Local Run Guide

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Run the application:：
   ```bash
   python app.py
   ```
3. The application will automatically start a local server (default at http://127.0.0.1:5000) and open the default system browser after 1.5 seconds.

## 📦 Packaging as a Standalone Executable

The project already includes the necessary configuration. You can package it into a standalone executable using the following command:
```bash
pyinstaller ReportAuto.spec
```
The generated executable file can be found in the dist directory.
