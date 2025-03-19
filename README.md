# Rename PDF Invoices by "Bill to" Name

This repository contains a PowerShell script that automatically renames PDF invoice files based on the content found under the "Bill to" section within each PDF.

The script is especially helpful for batch-renaming large numbers of invoice PDFs quickly and consistently.

## How It Works

- Extracts text from each PDF using the `pdftotext` utility from Xpdf tools.
- Uses regex to reliably find the "Bill to" section and captures the subsequent line as the customer name.
- Renames the PDF file to a clean, safe filename formatted as `[Customer Name] 2025 calendar ad invoice.pdf`.

## Prerequisites

You must install the Xpdf command-line tools (which include `pdftotext.exe`).

### Installing Xpdf CLI Tools

1. **Download Xpdf**:
   - Visit [Xpdf Download Page](https://www.xpdfreader.com/download.html).
   - Under the "Xpdf command line tools" section, download the ZIP file (`xpdf-tools-win-*.zip`).

2. **Extract Files**:
   - Extract the downloaded ZIP file to a convenient location (e.g., `C:\xpdf-tools`).

3. **Add Xpdf to Windows PATH**:
   - Press `Win + R`, type `SystemPropertiesAdvanced`, and press Enter.
   - Click on "Environment Variables".
   - Under "System Variables", select `Path`, then click "Edit".
   - Click "New" and add the path to your extracted Xpdf directory (e.g., `C:\xpdf-tools\bin64`).
   - Click OK on each dialog to apply the changes.

4. **Verify Installation**:
   - Open a new PowerShell or CMD prompt and type:
     ```shell
     pdftotext.exe -v
     ```
   - If installed correctly, this command should display the version information.

## Using the Script

### Run the Script

1. Place the script file (`RenameInvoices.ps1`) in the directory containing your PDF files.
2. Open PowerShell in that directory.
3. Execute the script:
   ```powershell
   .\RenameInvoices.ps1
   ```

### Script Behavior

- The script scans all PDFs in the current directory.
- It outputs verbose messages during file renaming.
- Provides warnings for:
  - Files where "Bill to" information is missing.
  - Potential filename conflicts.

## Customization

- Modify the `$newFileName` format if you prefer a different naming convention.

## Contributing

Feel free to open issues or submit pull requests to improve this script further.

## License

This project is open-source and free to use. See the repository license for more details.
