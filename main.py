import os
import shutil
import fitz  # PyMuPDF
from datetime import datetime



"""
This script automates the process of renaming PDF files based on specific text content extracted from them, focusing on a use case involving payment date extraction from payslip documents. The script operates within its current directory and follows these key steps:

1. Back Up All Files:
   Before any processing, the script creates backups of all files in the current directory to prevent data loss. The backups are stored in a subdirectory named 'backup'. This ensures that original files remain unaltered and safe throughout the script's execution.

2. Convert PDFs to Text:
   Each PDF file in the directory is scanned, and its textual content is extracted. This conversion process is facilitated by the PyMuPDF library (also known as fitz). The extracted text for each PDF is saved to a new text file in a subdirectory named 'output_text'. This step is crucial for analyzing the content of PDF files without altering the originals.

3. Extract Payment Dates and Rename:
   The script reads the extracted text files and searches for payment dates located on specific lines (by default, lines 29 and 35, based on the provided use case). These dates are expected to be in the format 'DD-MMM-YYYY'. Upon successful extraction of a valid payment date, the script then renames the original PDF file to reflect this date, formatted as 'YYYY-MMM-DD'. This renaming scheme facilitates organized storage and easy retrieval of documents based on chronological order.

Precautions and Assumptions:
- The script assumes that the payment date appears in a consistent position within the text extracted from the PDFs (lines 29 and 35 in the provided use case).
- It is designed to run in the same directory as the PDF files to be processed. Users must ensure the script is placed correctly before execution.
- Error handling includes skipping over directories during the backup process and handling files that do not conform to the expected format or structure, ensuring the script runs smoothly without interruption.

This script is particularly useful for individuals or organizations looking to automate the organization of payslip documents or similar PDF files that contain date-based information in a consistent format and location within the document.
"""




# Define the root directory where the script and PDFs are located
root_dir = os.path.dirname(os.path.realpath(__file__))
backup_dir = os.path.join(root_dir, 'backup')
output_text_dir = os.path.join(root_dir, 'output_text')

# Make sure backup and output_text directories exist
os.makedirs(backup_dir, exist_ok=True)
os.makedirs(output_text_dir, exist_ok=True)

def create_backup():
    # Backup all files in the directory before processing
    for filename in os.listdir(root_dir):
        if filename in ['backup', 'output_text']:  # Skip the backup and output_text directories
            continue
        file_path = os.path.join(root_dir, filename)
        if os.path.isfile(file_path):  # Check if it is a file
            backup_path = os.path.join(backup_dir, filename)
            shutil.copy2(file_path, backup_path)
        else:
            print(f"Skipping directory: {filename}")


def convert_pdfs_to_text():
    # Convert all PDF files in the directory to text files in output_text folder
    for filename in os.listdir(root_dir):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(root_dir, filename)
            with fitz.open(pdf_path) as doc:
                text = ''
                for page in doc:
                    text += page.get_text()
                text_file_name = filename + '.txt'
                text_file_path = os.path.join(output_text_dir, text_file_name)
                with open(text_file_path, 'w') as text_file:
                    text_file.write(text)


from datetime import datetime
import os


# Assuming other necessary imports and function definitions are already in place

def extract_date_and_prepare_renaming():
    renaming_info = []

    for text_filename in os.listdir(output_text_dir):
        if text_filename.endswith('.txt'):
            with open(os.path.join(output_text_dir, text_filename), 'r') as file:
                lines = file.readlines()
                period_line = lines[1].strip()
                start_date_str, end_date_str = period_line.replace('For the Period: ', '').split(' thru ')

                # Sortable date formatting
                start_date = datetime.strptime(start_date_str, "%m/%d/%y")
                end_date = datetime.strptime(end_date_str, "%m/%d/%y")

                original_pdf_filename = text_filename[:-4]
                original_pdf_path = os.path.join(root_dir, original_pdf_filename)

                new_pdf_base_name = f"{start_date.strftime('%Y-%m-%d')}to{end_date.strftime('%Y-%m-%d')}"
                new_pdf_name = new_pdf_base_name + ".pdf"

                renaming_info.append((start_date, original_pdf_path, new_pdf_name))

    return renaming_info


def rename_files_chronologically():
    renaming_info = extract_date_and_prepare_renaming()
    renaming_info_sorted = sorted(renaming_info, key=lambda x: x[0])

    for _, original_pdf_path, new_pdf_name in renaming_info_sorted:
        new_pdf_path = os.path.join(root_dir, new_pdf_name)
        counter = 1

        # Extract base name for use in duplicate handling
        new_pdf_base_name = new_pdf_name.rsplit('.', 1)[0]

        while os.path.exists(new_pdf_path):
            # Increment the name with a counter if duplicates exist
            new_pdf_name = f"{new_pdf_base_name}_{counter}.pdf"
            new_pdf_path = os.path.join(root_dir, new_pdf_name)
            counter += 1

        os.rename(original_pdf_path, new_pdf_path)
        print(f"Renamed '{original_pdf_path}' to '{new_pdf_name}'")


# Run the whole process
create_backup()
convert_pdfs_to_text()
rename_files_chronologically()
