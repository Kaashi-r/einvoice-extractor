# eInvoice Extractor

🇮🇳

## Overview

In India, GST department has made e-Invoicing mandatory for most of it's registered businesses. As of now, for whatever reason, you can only download generated einvoices in bulk as a single PDF file within an archive.

`eInvoice Extractor` is a Python script designed to handle and process e-invoice archives. It performs the following tasks:

1. **Folder Selection**: Allows users to select a folder containing e-invoice archives.
2. **Folder Management**: Deletes and recreates target folders to ensure a clean workspace.
3. **File Extraction**: Extracts PDF files from ZIP archives and renames them based on the archive name.
4. **PDF Processing**: Processes extracted PDFs to split them based on invoice details and renames the split files appropriately.

## Features

- Selects a folder via a graphical interface.
- Handles ZIP file extraction and PDF processing.
- Renames and organizes PDF files based on extracted data.
- Provides detailed error handling for file operations.

## Requirements

- Python 3.6 or higher
- `PyPDF2` library for PDF processing
- `tkinter` for the graphical folder selection
- `zipfile` for handling ZIP archives

To install the required Python libraries, use the following command:

    pip install PyPDF2

## Usage

1. **Run the Script**: Execute the script to start the process. You will be prompted to select a folder containing e-invoice archives.

2. **Processing**:

   - The script will reset the `einvoices` and `output` folders by deleting their contents and recreating them.
   - It will extract PDF files from ZIP archives found in the selected folder.
   - The extracted PDFs will be processed to split them based on invoice details.
   - The split files will be renamed and saved in the `output` folder.

3. **Check Results**: The processed and renamed files will be available in the `output` folder.

## Error Handling

The script includes error handling for common issues such as:

- Invalid ZIP files
- Missing or permission-denied files
- Issues during PDF extraction and processing

## Code Structure

- **`select_folder()`**: Opens a dialog to select the folder containing e-invoice archives.
- **`reset_folder(folder: Path)`**: Deletes and recreates the specified folder.
- **`extract_files()`**: Extracts PDF files from ZIP archives and renames them.
- **`extract_reg(text, pattern)`**: Extracts specific substrings using regular expressions.
- **`convert_date_format(date_str)`**: Converts date formats for file naming.
- **`escape_filename(filename)`**: Sanitizes filenames by replacing invalid characters.
- **`split_invoices()`**: Processes PDFs to split and rename based on invoice details.

## Contributing

If you'd like to contribute to this project, please fork the repository and create a pull request with your changes. Make sure to follow the coding style and include relevant tests.

## License

This project is licensed under the MPL 2.0 License. See the [LICENSE](LICENSE) file for details.

## Contact

For any questions or support, please contact [Anuresh](mailto:kombanmoonga@gmail.com).

## Repository

You can find the repository at: [https://github.com/KombanMoonga/einvoice-extractor.git](https://github.com/KombanMoonga/einvoice-extractor.git)
