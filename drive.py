import os
import re
import pandas as pd
import gdown
import fitz  # PyMuPDF

# Function to extract data from resume text
def extract_resume_data(resume_text):
    patterns = {
        'Name': r'\b(?:Name|BISWOJIT BISWAL)\b[:\s]*([A-Za-z\s]+)',
        'Phone': r'(\b\d{10}\b)',
        'Email': r'\b[\w\.-]+@[\w\.-]+\.\w{2,4}\b',
        'LinkedIn': r'(https:\/\/www\.linkedin\.com\/in\/[\w-]+)',
        'Location': r'Location[:\s]*([\w\s,]+)',
        'Summary': r'SUMMARY(.+?)(?=\bEXPERIENCE\b)',
        'Experience': r'EXPERIENCE(.+?)(?=\bEDUCATION\b)',
        'Education': r'EDUCATION(.+?)(?=\bSKILLS\b)',
        'Skills': r'SKILLS(.+?)(?=\bPROJECTS\b)',
        'Projects': r'PROJECTS(.+?)(?=\bACHIEVEMENTS\b)',
        'Achievements': r'ACHIEVEMENTS(.+?)(?=\bwww\.enhancv\.com\b)'
    }

    extracted_data = {}
    for field, pattern in patterns.items():
        match = re.search(pattern, resume_text, re.DOTALL | re.IGNORECASE)
        if match:
            extracted_data[field] = match.group(1).strip()
        else:
            extracted_data[field] = None
            print(f"Could not extract {field}")

    return extracted_data

# Function to read PDF file and extract text
def read_pdf(file):
    pdf_reader = fitz.open(stream=file.read(), filetype="pdf")
    text = ""
    for page_num in range(len(pdf_reader)):
        page = pdf_reader.load_page(page_num)
        text += page.get_text()
    return text

# Function to download files from Google Drive
def download_files_from_drive(drive_links, download_folder):
    if not os.path.exists(download_folder):
        os.makedirs(download_folder)
    for link in drive_links:
        file_id = link.split('/')[-2]
        gdown.download(f"https://drive.google.com/uc?export=download&id={file_id}", os.path.join(download_folder, file_id), quiet=False)

# Function to process a folder of resumes
def process_resumes_folder(folder_path, excel_file_path, sheet_name):
    # Load the existing Excel sheet
    sheet_data = pd.read_excel(excel_file_path, sheet_name=sheet_name)

    # Get a list of already processed emails
    processed_emails = sheet_data['Email'].dropna().unique().tolist()

    # Process each file in the folder
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        
        # Skip non-PDF and non-TXT files
        if not (filename.endswith('.pdf') or filename.endswith('.txt')):
            continue

        try:
            # Read the resume text based on file type
            if filename.endswith('.pdf'):
                with open(file_path, 'rb') as file:
                    resume_text = read_pdf(file)
            else:
                with open(file_path, 'r', encoding='utf-8') as file:
                    resume_text = file.read()

            # Extract data from the resume
            extracted_data = extract_resume_data(resume_text)

            # Check if the resume has already been processed
            if extracted_data['Email'] in processed_emails:
                print(f"Resume {filename} already processed. Skipping.")
                continue

            # Prepare the data to be appended
            new_row = pd.DataFrame([extracted_data])

            # Append the new data to the sheet
            sheet_data = pd.concat([sheet_data, new_row], ignore_index=True)

        except Exception as e:
            print(f"An error occurred while processing {filename}: {e}")

    # Save the updated Excel file (overwrite the existing one)
    sheet_data.to_excel(excel_file_path, sheet_name=sheet_name, index=False)

    print("All resumes have been processed and the Excel sheet has been updated.")

# Example usage
def main():
    drive_links = [
        # Add your Google Drive links here
        'https://drive.google.com/file/d/1XXXXXXXXXXXX/view?usp=sharing',
        'https://drive.google.com/file/d/2XXXXXXXXXXXX/view?usp=sharing',
        # Add more links as needed
    ]
    download_folder = 'downloaded_resumes'
    excel_file_path = 'Bulk Upload Sheet-3.xlsx'
    sheet_name = 'Sheet1'

    # Download files from Google Drive
    download_files_from_drive(drive_links, download_folder)

    # Process downloaded resumes
    process_resumes_folder(download_folder, excel_file_path, sheet_name)

if __name__ == "__main__":
    main()
