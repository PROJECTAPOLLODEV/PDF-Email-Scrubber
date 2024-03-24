import re
import fitz  # PyMuPDF
from openpyxl import Workbook

def extract_emails_from_pdf(pdf_path):
    emails = []

    # Open the PDF file
    pdf_document = fitz.open(pdf_path)

    for page_number in range(pdf_document.page_count):
        page = pdf_document[page_number]

        # Extract text from the page
        text = page.get_text()

        # Use regex to find email addresses in the text
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        page_emails = re.findall(email_pattern, text)

        # Add the found emails to the list
        emails.extend(page_emails)

    # Close the PDF document
    pdf_document.close()

    return emails

def export_emails_to_excel(emails, excel_path):
    # Create a new Excel workbook and get the active sheet
    workbook = Workbook()
    sheet = workbook.active

    # Add headers to the Excel sheet
    sheet.append(['Email'])

    # Add emails to the sheet
    for email in emails:
        sheet.append([email])

    # Save the Excel workbook
    workbook.save(excel_path)

if __name__ == "__main__":
    # Define the paths for the PDF to scrub and the Excel file to export
    pdf_path_to_scrub = "C:/Users/yourpdfdochere.pdf"
    excel_export_path = "C:/Users/output 1337.xlsx"

    # Extract emails from the PDF
    scrubbed_emails = extract_emails_from_pdf(pdf_path_to_scrub)

    # Export emails to Excel
    export_emails_to_excel(scrubbed_emails, excel_export_path)

    print(f"Scrubbed emails exported to {excel_export_path}")
