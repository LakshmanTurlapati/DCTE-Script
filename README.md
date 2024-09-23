### README: Document Classification and Text Extraction Script

#### Overview
This script is designed to process and classify business documents such as RFQs (Request for Quotation) and POs (Purchase Orders) from various file formats, including `.eml`, `.pdf`, `.xlsx`, `.csv`, and image files. It extracts relevant information based on predefined patterns and presents it in a concise, structured format. Additionally, the script processes and extracts text from attachments found within `.eml` files.

#### Prerequisites
Before running the script, ensure that you have the following installed:

1. **Python 3.6 or higher**
2. **Required Python packages**:
   - pdfminer.six
   - pytesseract
   - Pillow (PIL)
   - pandas
   - bs4 (BeautifulSoup)
   - tabulate
   - tkinter (for file dialog)
   - openpyxl (for Excel files)

You can install these packages using the following command:
```bash
pip install pdfminer.six pytesseract Pillow pandas bs4 tabulate tkinter openpyxl
```

> **Note:** For extracting text from images, you need to have `Tesseract-OCR` installed. You can download it from [here](https://github.com/tesseract-ocr/tesseract) and ensure it's accessible in your system's PATH.

#### Script Explanation
1. **Text Extraction Functions**:
   - Extracts text from different file types (`.pdf`, `.xlsx`, `.csv`, `.eml`, `.txt`, images).
   - Cleans and processes the extracted text to remove unnecessary artifacts and formats it for easier parsing.

2. **Classification**:
   - Identifies the document type (RFQ, PO, or Delivery) based on keywords found in the subject or body of the document.

3. **Content Extraction**:
   - Extracts specific fields like RFQ Number, Agency, Email, and others based on predefined patterns for each document type.
   - The extracted information is stored in a structured format (dictionary).

4. **Attachment Handling**:
   - For `.eml` files, the script also extracts and processes any attachments, extracting text and presenting relevant content.

5. **Output Presentation**:
   - Displays the classification results and extracted information in a structured and easy-to-read format.

#### How to Run the Script
1. **Command Line Execution**:
   You can run the script from the command line by specifying the files you want to process:
   ```bash
   python text_extraction.py --files /path/to/your/file1.eml /path/to/your/file2.pdf
   ```

   You can also specify additional optional parameters:
   - `--subject`: Specify an email subject if not present in the file.
   - `--from_email`: Specify a sender's email address if not present in the file.

2. **Graphical File Selection**:
   If no file paths are provided via command line arguments, a file selection dialog will open, allowing you to select multiple files to process.

#### Expected Output
- **Classification Results**:
  Indicates whether the document is classified as an RFQ, PO, or Delivery, based on the detected keywords.

- **Extracted Information**:
  Displays extracted fields such as RFQ Number, Agency, Email, Release Date, Due By, Line Items, and Quotes. If any of these fields are not found, they will be marked as `N/A`.

- **Attachment Processing**:
  If the document is an `.eml` file with attachments, the script will process each attachment and attempt to extract relevant text. The first 500 characters of the extracted text will be displayed.

#### Example Output
```
Processing file: /path/to/your/file.eml

Classification Results:
  - IsRFQ: True

Extracted Information:
  - RFQ Number: 305431
  - Agency: Department of Defense
  - Email: john.doe@defense.gov
  - Release Date: 2024-04-22
  - Due By: 2024-05-01
  - Line Items: 10
  - Quotes: 5

Processing attachment: /path/to/attachment.pdf
Extracted text from attachment:
"This request for quotation (RFQ) is to support the Defense Information System Agency (DISA)..."
```

#### Additional Notes
- Ensure that the extracted patterns (`re.search(...)`) in the script match the expected format in your documents. You may need to adjust these patterns based on your document's structure.
- Fine-tuning is still in progress and will be improved over time for better accuracy. For more specific use cases, you may need to modify the regex patterns or add additional content extraction logic to meet your requirements.
