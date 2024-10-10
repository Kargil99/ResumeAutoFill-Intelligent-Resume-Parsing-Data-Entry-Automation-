import sys
import os
import re
import pandas as pd
import fitz  # PyMuPDF
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QVBoxLayout, QWidget, QDesktopWidget
from io import StringIO

class ResumeApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set the window title and initial size
        self.setWindowTitle('Resume Data to Excel')
        self.resize(800, 600)  # Set initial size of the window
        self.center()  # Center the window on the screen

        # Create layout and add buttons and label
        layout = QVBoxLayout()

        self.upload_button = QPushButton('Upload Resume', self)
        self.upload_button.clicked.connect(self.upload_resume)
        layout.addWidget(self.upload_button)

        self.upload_folder_button = QPushButton('Upload Resume Folder', self)
        self.upload_folder_button.clicked.connect(self.upload_resume_folder)
        layout.addWidget(self.upload_folder_button)

        self.update_button = QPushButton('Update Sheet', self)
        self.update_button.clicked.connect(self.update_sheet)
        layout.addWidget(self.update_button)

        self.status_label = QLabel('', self)
        layout.addWidget(self.status_label)

        # Set layout to the central widget
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Initialize the list to store resume files
        self.resume_files = []

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
                extracted_data = extract_resume_data(resume_text)

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

def extract_resume_data(resume_text):
    """
    Extract data from resume text using regex patterns.
    """
    patterns = {
        'Name': r'Name[:\s]*([A-Za-z\s]+)',
        'Number': r'Number[:\s]*([\d\s]+)',
        'Email': r'Email[:\s]*([\w\.-]+@[\w\.-]+)',
        'DOB': r'DOB[:\s]*(\d{2}-\d{2}-\d{2})',
        'Gender': r'Gender[:\s]*(Male|Female|Other)',
        'Pincode': r'Pincode[:\s]*(\d{6})',
        'Address': r'Address[:\s]*(.+)',
        'Qualification': r'Qualification[:\s]*(.+)',
        'Specialization': r'Specialization[:\s]*(.+)',
        'Experience': r'Experience[:\s]*(.+)',
        'Sectors': r'Sectors[:\s]*(.+)',
        'Skills': r'Skills[:\s]*(.+)',
        'Mark': r'Mark[:\s]*(.+)',
        'College': r'College[:\s]*(.+)',
        'YearGap': r'YearGap[:\s]*(\d+)',
        'PassingYear': r'PassingYear[:\s]*(\d{4})',
        'Preferred Location': r'Preferred Location[:\s]*(.+)'
    }

    extracted_data = {}
    for field, pattern in patterns.items():
        match = re.search(pattern, resume_text, re.IGNORECASE)
        if match and match.group(1):
            extracted_data[field] = match.group(1).strip()
        else:
            extracted_data[field] = None
            print(f"Could not extract {field}")

    return extracted_data

def read_pdf(file):
    """
    Read a PDF file and extract its text content.
    """
    pdf_reader = fitz.open(stream=file.read(), filetype="pdf")
    text = ""
    for page_num in range(len(pdf_reader)):
        page = pdf_reader.load_page(page_num)
        text += page.get_text()
    return text

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ResumeApp()
    ex.show()
    sys.exit(app.exec_())
