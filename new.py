import sys
import os
import re
import pandas as pd
import fitz  # PyMuPDF
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QVBoxLayout, QWidget, QDesktopWidget
)
from PyQt5.QtGui import QIcon, QFont
import spacy

class ResumeApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set the window title and initial size
        self.setWindowTitle('Resume Data to Excel')
        self.resize(400, 300)  # Set initial size of the window
        self.center()  # Center the window on the screen

        # Create layout and add buttons and label
        layout = QVBoxLayout()
        layout.setSpacing(10)

        self.upload_button = QPushButton('Upload Resume')
        self.upload_button.setIcon(QIcon('icons/upload.png'))  # Add appropriate icon path
        self.upload_button.clicked.connect(self.upload_resume)
        layout.addWidget(self.upload_button)

        self.upload_folder_button = QPushButton('Upload Resume Folder')
        self.upload_folder_button.setIcon(QIcon('icons/folder.png'))  # Add appropriate icon path
        self.upload_folder_button.clicked.connect(self.upload_resume_folder)
        layout.addWidget(self.upload_folder_button)

        self.update_button = QPushButton('Update Sheet')
        self.update_button.setIcon(QIcon('icons/update.png'))  # Add appropriate icon path
        self.update_button.clicked.connect(self.update_sheet)
        layout.addWidget(self.update_button)

        self.status_label = QLabel('')
        self.status_label.setFont(QFont('Arial', 10))
        layout.addWidget(self.status_label)

        # Set layout to the central widget
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Initialize the list to store resume files
        self.resume_files = []

        # Load NLP model
        self.nlp = spacy.load('en_core_web_sm')

    def center(self):
        """
        Center the window on the screen.
        """
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def upload_resume(self):
        """
        Open a file dialog to select resume files (PDF or TXT).
        """
        options = QFileDialog.Options()
        files, _ = QFileDialog.getOpenFileNames(self, "Select Resume Files", "", "PDF Files (*.pdf);;Text Files (*.txt)", options=options)
        if files:
            self.resume_files.extend(files)
            self.status_label.setText(f"Selected {len(files)} files")

    def upload_resume_folder(self):
        """
        Open a file dialog to select a folder containing resume files.
        """
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            for filename in os.listdir(folder):
                if filename.endswith('.pdf') or filename.endswith('.txt'):
                    self.resume_files.append(os.path.join(folder, filename))
            self.status_label.setText(f"Selected {len(self.resume_files)} files from folder")

    def update_sheet(self):
        """
        Process the selected resumes and update the Excel sheet.
        """
        if not self.resume_files:
            self.status_label.setText("No resumes uploaded.")
            return

        excel_file_path = 'Bulk Upload Sheet-3.xlsx'
        sheet_name = 'Sheet1'

        try:
            sheet_data = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        except FileNotFoundError:
            sheet_data = pd.DataFrame()

        # Get a list of already processed emails
        processed_emails = sheet_data['Email'].dropna().unique().tolist()

        for file_path in self.resume_files:
            try:
                if file_path.endswith('.pdf'):
                    with open(file_path, 'rb') as file:
                        resume_text = read_pdf(file)
                else:
                    with open(file_path, 'r', encoding='utf-8') as file:
                        resume_text = file.read()

                # Extract data from the resume
                extracted_data = self.extract_resume_data(resume_text)

                # Check if the resume has already been processed
                if extracted_data['Email'] in processed_emails:
                    print(f"Resume {os.path.basename(file_path)} already processed. Skipping.")
                    continue

                # Append the new data to the sheet
                new_row = pd.DataFrame([extracted_data])
                sheet_data = pd.concat([sheet_data, new_row], ignore_index=True)
            except Exception as e:
                print(f"An error occurred while processing {os.path.basename(file_path)}: {e}")

        # Save the updated Excel file
        sheet_data.to_excel(excel_file_path, sheet_name=sheet_name, index=False)
        self.status_label.setText("All resumes have been processed and the Excel sheet has been updated.")
        self.resume_files = []

    def extract_resume_data(self, resume_text):
        """
        Extract data from resume text using NLP and regex patterns.
        """
        doc = self.nlp(resume_text)

        extracted_data = {
            'Name': self.extract_name(doc),
            'Number': self.extract_phone_number(resume_text),
            'Email': self.extract_email(resume_text),
            'DOB': self.extract_dob(resume_text),
            'Gender': self.extract_gender(resume_text),
            'Pincode': self.extract_pincode(resume_text),
            'Address': self.extract_address(resume_text),
            'Qualification': self.extract_qualification(resume_text),
            'Specialization': self.extract_specialization(resume_text),
            'Experience': self.extract_experience(resume_text),
            'Sectors': self.extract_sectors(resume_text),
            'Skills': self.extract_skills(resume_text),
            'Mark': self.extract_mark(resume_text),
            'College': self.extract_college(resume_text),
            'Year Gap': self.extract_year_gap(resume_text),
            'Passing Year': self.extract_passing_year(resume_text),
            'Preferred Location': self.extract_preferred_location(resume_text)
        }

        return extracted_data

    def extract_name(self, doc):
        for ent in doc.ents:
            if ent.label_ == 'PERSON':
                return ent.text
        return None

    def extract_phone_number(self, text):
        match = re.search(r'\b\d{10}\b', text)
        return match.group(0) if match else None

    def extract_email(self, text):
        match = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', text)
        return match.group(0) if match else None

    def extract_dob(self, text):
        match = re.search(r'\b\d{2}-\d{2}-\d{4}\b', text)
        return match.group(0) if match else None

    def extract_gender(self, text):
        match = re.search(r'\b(Male|Female|Other)\b', text, re.IGNORECASE)
        return match.group(0) if match else None

    def extract_pincode(self, text):
        match = re.search(r'\b\d{6}\b', text)
        return match.group(0) if match else None

    def extract_address(self, text):
        # Simplified pattern for address extraction
        match = re.search(r'Address[:\s]*(.+)', text)
        return match.group(1).strip() if match else None

    def extract_qualification(self, text):
        match = re.search(r'Qualification[:\s]*(.+)', text)
        return match.group(1).strip() if match else None

    def extract_specialization(self, text):
        match = re.search(r'Specialization[:\s]*(.+)', text)
        return match.group(1).strip() if match else None

    def extract_experience(self, text):
        match = re.search(r'Experience[:\s]*(.+)', text)
        return match.group(1).strip() if match else None

    def extract_sectors(self, text):
        match = re.search(r'Sectors[:\s]*(.+)', text)
        return match.group(1).strip() if match else None

    def extract_skills(self, text):
        match = re.search(r'Skills[:\s]*(.+)', text)
        return match.group(1).strip() if match else None

    def extract_mark(self, text):
        match = re.search(r'Mark[:\s]*(.+)', text)
        return match.group(1).strip() if match else None

    def extract_college(self, text):
        match = re.search(r'College[:\s]*(.+)', text)
        return match.group(1).strip() if match else None

    def extract_year_gap(self, text):
        match = re.search(r'Year Gap[:\s]*(.+)', text)
        return match.group(1).strip() if match else None

    def extract_passing_year(self, text):
        match = re.search(r'Passing Year[:\s]*(\d{4})', text)
        return match.group(1).strip() if match else None

    def extract_preferred_location(self, text):
        match = re.search(r'Preferred Location[:\s]*(.+)', text)
        return match.group(1).strip() if match else None

def read_pdf(file):
    """
    Extract text from a PDF file.
    """
    pdf_document = fitz.open(stream=file.read(), filetype='pdf')
    text = ""
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text += page.get_text()
    return text

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ResumeApp()
    ex.show()
    sys.exit(app.exec_())
