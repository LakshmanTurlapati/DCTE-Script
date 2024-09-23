import sys
import json
import re
import argparse
import os
import tempfile
import shutil
from datetime import datetime
from tabulate import tabulate  # For displaying structured tables

# File type specific imports
from pdfminer.high_level import extract_text as extract_text_from_pdf
from pytesseract import image_to_string
from PIL import Image
import pandas as pd
from bs4 import BeautifulSoup

# For GUI file selection
import tkinter as tk
from tkinter import filedialog

# For parsing .eml files
from email import policy
from email.parser import BytesParser

# For OCR on PDFs
from pdf2image import convert_from_path


# Function to clean extracted text
def clean_extracted_text(text):
    text = re.sub(r'\s+', ' ', text)  # Replace multiple spaces and newlines with a single space
    text = re.sub(r'<.*?>', '', text)  # Remove HTML tags
    return text.strip()


# Function to extract text from images
def extract_text_from_image(image_path):
    try:
        text = image_to_string(Image.open(image_path))
        return clean_extracted_text(text)
    except Exception as e:
        print(f"Error extracting text from image {image_path}: {e}")
        return ""


# Function to extract text from Excel files
def extract_text_from_excel(excel_path):
    try:
        df = pd.read_excel(excel_path, engine='openpyxl')
        text = df.to_string()
        return clean_extracted_text(text)
    except Exception as e:
        print(f"Error extracting text from Excel {excel_path}: {e}")
        return ""


# Function to extract text from CSV files
def extract_text_from_csv(csv_path):
    try:
        df = pd.read_csv(csv_path)
        text = df.to_string()
        return clean_extracted_text(text)
    except Exception as e:
        print(f"Error extracting text from CSV {csv_path}: {e}")
        return ""


# Function to extract text from PDFs
def extract_text_from_pdf_file(pdf_path):
    try:
        text = extract_text_from_pdf(pdf_path)
        if not text.strip():
            pages = convert_from_path(pdf_path)
            text = ''.join(image_to_string(page) for page in pages)
        return clean_extracted_text(text)
    except Exception as e:
        print(f"Error extracting text from PDF {pdf_path}: {e}")
        return ""


# Function to extract text from EML files
def extract_text_from_eml(eml_path):
    try:
        with open(eml_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)
        subject = msg['subject'] if msg['subject'] else ''
        from_email = msg['from'] if msg['from'] else ''
        body = ''
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type in ['text/plain', 'text/html']:
                    body += part.get_payload(decode=True).decode(errors='ignore')
        else:
            body = msg.get_payload(decode=True).decode(errors='ignore')
        return subject, from_email, clean_extracted_text(body)
    except Exception as e:
        print(f"Error extracting text from EML {eml_path}: {e}")
        return '', '', ''


# Function to extract text from any supported file type
def extract_text_from_file(file_path):
    extension = os.path.splitext(file_path)[1].lower()
    if extension == '.pdf':
        return extract_text_from_pdf_file(file_path)
    elif extension in ['.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif']:
        return extract_text_from_image(file_path)
    elif extension in ['.xlsx', '.xls']:
        return extract_text_from_excel(file_path)
    elif extension == '.csv':
        return extract_text_from_csv(file_path)
    elif extension == '.txt':
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                text = f.read()
                return clean_extracted_text(text)
        except Exception as e:
            print(f"Error reading text file {file_path}: {e}")
            return ""
    else:
        print(f"Unsupported file type: {extension}")
        return ""


# Function to identify document type based on keywords
def identify_document_type(subject, from_email, body):
    document_type = {
        'IsRFQ': False,
        'IsPO': False,
        'IsDelivery': False
    }

    rfq_keywords = r'\brfq\b|\bquote\b|\bmrr\b|\bquotation\b'
    po_keywords = r'\bpurchase order\b|\bpo\b|\bspo\b'
    delivery_keywords = r'\bdelivery\b|\btracking\b|\bshipment\b|\bshipping\b'

    if re.search(rfq_keywords, subject, re.IGNORECASE) or re.search(rfq_keywords, body, re.IGNORECASE):
        document_type['IsRFQ'] = True
    elif re.search(po_keywords, subject, re.IGNORECASE) or re.search(po_keywords, body, re.IGNORECASE):
        document_type['IsPO'] = True
    elif re.search(delivery_keywords, subject, re.IGNORECASE) or re.search(delivery_keywords, body, re.IGNORECASE):
        document_type['IsDelivery'] = True

    return document_type


# Function to extract and classify content for RFQ
def extract_rfq_content(body):
    content = {
        'RFQ Number': re.search(r'RFQ\s*(?:No\.?|Number)?\s*[:\-]?\s*(\w+)', body, re.IGNORECASE),
        'Agency': re.search(r'Agency\s*[:\-]?\s*(.+?)(?=\s*(?:Request|Attachments|Requirements|$))', body,
                            re.IGNORECASE),
        'Email': re.search(r'Email\s*[:\-]?\s*(\S+@\S+)', body, re.IGNORECASE),
        'Release Date': re.search(r'Release Date\s*[:\-]?\s*(\d{4}-\d{2}-\d{2})', body),
        'Due By': re.search(r'Due By\s*[:\-]?\s*(\d{4}-\d{2}-\d{2})', body),
        'Line Items': re.search(r'Line Items\s*[:\-]?\s*(\d+)', body),
        'Quotes': re.search(r'Quotes\s*[:\-]?\s*(\d+)', body)
    }
    for key in content:
        if isinstance(content[key], re.Match):
            content[key] = content[key].group(1).strip()
        else:
            content[key] = 'N/A'
    return content


# Function to extract and classify content for PO
def extract_po_content(body):
    content = {
        'SPO Number': re.search(r'SPO\s*(?:No\.?|Number)?\s*[:\-]?\s*(\w+)', body, re.IGNORECASE),
        'Related RFQ Number': re.search(r'Related RFQ No\.?\s*[:\-]?\s*(\w+)', body, re.IGNORECASE),
        'Agency': re.search(r'Agency\s*[:\-]?\s*(.+?)(?=\s*(?:Request|Attachments|Requirements|$))', body,
                            re.IGNORECASE),
        'Contracting Officer': re.search(r'Contracting Officer\s*[:\-]?\s*(.+)', body, re.IGNORECASE),
        'Email': re.search(r'Email\s*[:\-]?\s*(\S+@\S+)', body, re.IGNORECASE),
        'Awarded Date': re.search(r'Awarded Date\s*[:\-]?\s*(\d{4}-\d{2}-\d{2})', body),
        'Line Items': re.search(r'Line Items\s*[:\-]?\s*(\d+)', body),
        'Attachments': re.search(r'Attachments\s*[:\-]?\s*(\d+)', body),
    }
    for key in content:
        if isinstance(content[key], re.Match):
            content[key] = content[key].group(1).strip()
        else:
            content[key] = 'N/A'
    return content


# Function to present data concisely
def present_data(content, classification_results):
    print("\nClassification Results:")
    for key, value in classification_results.items():
        if value:
            print(f"  - {key}: {value}")

    print("\nExtracted Information:")
    for key, value in content.items():
        print(f"  - {key}: {value}")


# Main function to process files
def process_file(file_path, subject='', from_email=''):
    print(f"\nProcessing file: {file_path}")

    # Extract text from file
    extension = os.path.splitext(file_path)[1].lower()
    if extension == '.eml':
        subject, from_email, body = extract_text_from_eml(file_path)
    else:
        body = extract_text_from_file(file_path)

    if not body:
        print("No text extracted from the file.")
        return

    # Identify document type
    classification_results = identify_document_type(subject, from_email, body)

    # Extract content based on document type
    if classification_results['IsRFQ']:
        content = extract_rfq_content(body)
    elif classification_results['IsPO']:
        content = extract_po_content(body)
    else:
        print("No specific document type identified.")
        return

    # Present data in a structured table format
    present_data(content, classification_results)


def main():
    # Try to get the file paths from command-line arguments
    parser = argparse.ArgumentParser(description='Process files for text extraction and classification.')
    parser.add_argument('--files', nargs='*', help='Paths to the files to process')
    parser.add_argument('--subject', default='', help='Email subject (optional)')
    parser.add_argument('--from_email', default='', help='Sender email address (optional)')

    args = parser.parse_args()

    file_paths = args.files
    subject = args.subject
    from_email = args.from_email

    # If file paths are not provided, open a file dialog for multiple files
    if not file_paths:
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        file_paths = filedialog.askopenfilenames(title='Select files to process')
        if not file_paths:
            print("No files selected. Exiting.")
            return

    # Process each file
    for file_path in file_paths:
        process_file(file_path, subject, from_email)


if __name__ == '__main__':
    main()