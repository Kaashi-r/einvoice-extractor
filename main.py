"""
This script performs the following tasks:
1. Selects a folder containing e-invoice archives.
2. Resets the target folders by deleting their contents and recreating them.
3. Extracts PDF files from ZIP archives in the selected folder.
4. Processes the extracted PDFs to split them based on invoice details.
5. Renames and organizes the split PDF files.
"""

from datetime import datetime
import os
import re
import shutil
from PyPDF2 import PdfReader, PdfWriter
import zipfile
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

# Constants
einv_det: str = "e-Invoice Details"
doc_no_ptn: str = r"Document No. :\s*([\w\s\W]+?)\s*IGST"
doc_dt_ptn: str = r"Document Date : (\d{2}-\d{2}-\d{4})"
ack_dt_ptn: str = r"Ack Date : (\d{2}-\d{2}-\d{4})"
einv_name = "einv1"
op_folder = "output"
ar_folder = "archives"
ei_folder = "einvoices"
cwd_folder = Path.cwd()

# Define paths using Path objects
ei_path = cwd_folder.joinpath(ei_folder)
op_path = cwd_folder.joinpath(op_folder)
ar_path = cwd_folder.joinpath(ar_folder)


def select_folder():
    """
    Opens a dialog to select a folder containing e-invoice archives.
    Returns:
        Path: The path to the selected folder.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    folder_path = filedialog.askdirectory(
        title="Select the folder containing e-invoice Archives"
    )
    return Path(folder_path)


def reset_folder(folder: Path):
    """
    Deletes a folder and its contents if it exists, then recreates it.
    Args:
        folder (Path): The path to the folder to reset.
    """
    if folder.exists():
        try:
            shutil.rmtree(folder)  # Delete folder and contents
            print(f"Deleted folder: {folder}")
        except OSError as e:
            print(f"Error deleting folder: {e}")
    try:
        folder.mkdir(parents=True, exist_ok=True)  # Recreate folder
        print(f"Created folder: {folder}")
    except OSError as e:
        print(f"Error creating folder: {e}")


def extract_files():
    """
    Extracts PDF files from ZIP archives in the specified archive folder.
    Renames the extracted files based on the ZIP file's name.
    """
    for zip_file in ar_path.glob("*.zip"):
        try:
            with zipfile.ZipFile(zip_file, "r") as archive:
                try:
                    archive.extract(f"{einv_name}.pdf", path=ei_path)
                    extracted_file = ei_path.joinpath(f"{einv_name}.pdf")
                    new_file_name = f"{zip_file.stem}.pdf"
                    extracted_file.rename(ei_path.joinpath(new_file_name))
                except FileNotFoundError:
                    print(f"{einv_name}.pdf not found in {zip_file}.")
                except PermissionError:
                    print(
                        f"Permission denied while extracting or renaming file from {zip_file}."
                    )
                except Exception as e:
                    print(f"Error during ZIP file extraction: {e}")
        except zipfile.BadZipFile:
            print(f"File {zip_file} is not a valid ZIP file.")
        except Exception as e:
            print(f"Error opening ZIP file {zip_file}: {e}")


def extract_reg(text, pattern):
    """
    Extracts a substring from text that matches a given regular expression pattern.
    Args:
        text (str): The text to search in.
        pattern (str): The regular expression pattern.
    Returns:
        str or None: The matched substring, or None if no match is found.
    """
    match = re.search(pattern, text)
    return match.group(1) if match else None


def convert_date_format(date_str):
    """
    Converts a date string from 'dd-mm-yyyy' to 'yyyymmdd'.
    Args:
        date_str (str): The date string in 'dd-mm-yyyy' format.
    Returns:
        str: The date string in 'yyyymmdd' format.
    """
    day, month, year = date_str.split("-")
    return year + month + day


def escape_filename(filename):
    """
    Replaces invalid characters in filenames with a hyphen.
    Args:
        filename (str): The filename to sanitize.
    Returns:
        str: The sanitized filename.
    """
    invalid_chars = r'[<>:"/\\|?*]'
    return re.sub(invalid_chars, "-", filename)


def split_invoices():
    """
    Processes PDF files in the e-invoices folder to split them based on invoice details.
    Renames and organizes the split PDF files.
    """
    for file in ei_path.glob("*.pdf"):
        try:
            pdf = PdfReader(file)
        except FileNotFoundError:
            print(f"File {file} not found.")
            continue
        except PermissionError:
            print(f"Permission denied while opening file {file}.")
            continue
        except Exception as e:
            print(f"Error opening PDF file {file}: {e}")
            continue

        output = PdfWriter()
        new_file_name = file.stem
        doc_no = "einvoice"

        for i, page in enumerate(pdf.pages):
            try:
                txt = page.extract_text()
            except Exception as e:
                print(f"Error extracting text from page {i} of {file}: {e}")
                continue

            if einv_det in txt:
                ack_date = extract_reg(txt, ack_dt_ptn)
                doc_no = escape_filename(extract_reg(txt, doc_no_ptn))
                if i == 0:
                    new_file_name = convert_date_format(ack_date)
                else:
                    if output.pages:
                        try:
                            with open(op_path.joinpath(doc_no + ".pdf"), "wb") as p:
                                output.write(p)
                        except IOError as e:
                            print(f"Error writing file {doc_no}.pdf: {e}")
                            continue
                        output = PdfWriter()
                output.add_page(page)
            else:
                output.add_page(page)

        if output.pages:
            try:
                with open(op_path.joinpath(doc_no + ".pdf"), "wb") as p:
                    output.write(p)
            except IOError as e:
                print(f"Error writing final file {doc_no}.pdf: {e}")

        try:
            file.rename(ei_path.joinpath(new_file_name + ".pdf"))
        except FileNotFoundError:
            print(f"File {file} not found for renaming.")
        except PermissionError:
            print(f"Permission denied while renaming file {file}.")
        except Exception as e:
            print(f"Error renaming file {file}: {e}")


if __name__ == "__main__":
    """
    Main execution block of the script.
    """
    ar_path = select_folder()
    if ar_path:
        ar_path = Path(ar_path)  # Ensure the selected path is a Path object
        print(f"Selected folder: {ar_path}")
        reset_folder(ei_path)
        reset_folder(op_path)
        extract_files()
        split_invoices()
        print("Done")
    else:
        print("No folder selected")
