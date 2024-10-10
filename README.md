# ResumeAutoFill: Intelligent Resume Parsing & Data Entry Automation

## Project Overview
**ResumeAutoFill** is a Python-based automation model that simplifies the process of extracting key information from resumes and populating a spreadsheet with accurate details such as name, email address, phone number, LinkedIn profile, and more. This project is designed to reduce manual data entry tasks for recruiters, saving time and ensuring data consistency.

## Key Features
- **Automatic Resume Parsing**: Scans and extracts key details like name, email, phone number, LinkedIn URL, and more from uploaded resumes.
- **Spreadsheet Automation**: Automatically populates a spreadsheet with the extracted information in pre-defined columns.
- **Error Handling**: Detects and handles incomplete or malformed resumes, ensuring smooth data processing.
- **Data Accuracy**: Uses advanced regex and parsing techniques to ensure data is correctly extracted and placed in the appropriate spreadsheet fields.

## How It Works
1. Upload a resume file (in formats like PDF, DOCX, etc.).
2. The model parses the document and extracts relevant information using Python libraries like `PyPDF2`, `docx`, and `regex`.
3. The extracted data is automatically filled into a pre-structured spreadsheet with exact column mappings (e.g., Name, Email, LinkedIn, etc.).
4. The completed spreadsheet can be downloaded or used for further recruitment processes.

## Tech Stack
- **Language**: Python
- **Libraries**: PyPDF2, docx, regex, pandas, openpyxl
- **Output Format**: Excel (XLSX) / CSV

## Project Structure
- `resume_parser.py`: Core script for parsing resumes and extracting information.
- `spreadsheet_fill.py`: Script that fills the extracted data into the spreadsheet.
- `sample_resume.pdf`: Example resume used for testing the model.
- `output_spreadsheet.xlsx`: Sample output showing extracted data.

## Limitations
- The model is designed for standard resumes. Unconventional formats may require additional tweaking.
- Data extraction depends on the resume's structure and format.

## Future Improvements
- Adding support for more file formats.
- Improving the accuracy of data extraction for non-standard resumes.
- Integrating with an applicant tracking system (ATS) for seamless data transfer.

## Usage
1. Clone this repository.
2. Run the `resume_parser.py` script and upload the resume file.
3. The parsed data will automatically fill into the `output_spreadsheet.xlsx`.

## License
This project is open for personal and educational use. Commercial use or modifications for proprietary tools are subject to approval.

# ResumeAutoFill-Intelligent-Resume-Parsing-Data-Entry-Automation-
