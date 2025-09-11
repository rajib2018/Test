import streamlit as st
import pytesseract
from PIL import Image
import pandas as pd
import os
import re
import spacy
import uuid # To generate unique filenames
import io # To handle in-memory files

# --- Function Definitions (from previous steps) ---

# Load a pre-trained English language model for spaCy
@st.cache_resource # Cache the model loading for efficiency
def load_spacy_model():
    try:
        nlp = spacy.load("en_core_web_sm")
    except OSError:
        st.warning("Downloading en_core_web_sm model...")
        from spacy.cli import download
        download("en_core_web_sm")
        nlp = spacy.load("en_core_web_sm")
    return nlp

nlp = load_spacy_model()

def ocr_image_to_text(image_file):
    """
    Performs OCR on an image file-like object and extracts text content.

    Args:
        image_file: A file-like object of the image or scanned document.

    Returns:
        str: The extracted text content, or None if an error occurs.
    """
    try:
        img = Image.open(image_file)
        text = pytesseract.image_to_string(img)
        return text
    except pytesseract.TesseractNotFoundError:
        st.error("Error: Tesseract is not installed or not in your PATH.")
        st.error("Please install Tesseract OCR engine.")
        return None
    except Exception as e:
        st.error(f"An error occurred during OCR: {e}")
        return None

def extract_invoice_data(text):
    """
    Extracts data fields from invoice text using spaCy and rule-based methods.
    """
    doc = nlp(text)
    extracted_data = {
        "Invoice Number": None,
        "Invoice Date": None,
        "Vendor Name": None,
        "Vendor Address": None,
        "Customer Name": None,
        "Customer Address": None,
        "Subtotal Amount": None,
        "Tax Amount": None,
        "Total Amount": None,
        "Currency": None,
        "Line Items": []
    }

    # Use spaCy for basic entity recognition (Names, Dates, Money)
    # Refined: Use entities as hints, but rely more on robust regex for specific fields
    dates = [ent.text for ent in doc.ents if ent.label_ == "DATE"]
    if dates:
        # Simple heuristic: try to find a date near "invoice date" or "date"
        date_match = re.search(r'(invoice date|date)[:\s]*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}|\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4})', text, re.IGNORECASE)
        if date_match:
            extracted_data["Invoice Date"] = date_match.group(2).strip()
        elif dates:
             # Fallback to first identified date if keyword match fails
             extracted_data["Invoice Date"] = dates[0]


    # Rule-based extraction using regular expressions and keywords
    # Invoice Number - More flexible regex
    invoice_number_match = re.search(r'(invoice number|invoice no|invoice #|inv #|bill no)[:\s]*([A-Z0-9\-\/]+)', text, re.IGNORECASE)
    if invoice_number_match:
        extracted_data["Invoice Number"] = invoice_number_match.group(2).strip()

    # Currency Symbol - Look for common symbols globally
    currency_match = re.search(r'([\$\£\€])', text)
    if currency_match:
        extracted_data["Currency"] = currency_match.group(1)

    # Amounts - More robust regex to capture currency symbol and various number formats
    amount_pattern = r'([\$\£\€]?\s*\d{1,3}(?:,\d{3})*(?:\.\d{2})?)' # Matches $1,234.56, €10.00, 50.00 etc.

    # Subtotal Amount
    subtotal_match = re.search(r'(subtotal)[:\s]*' + amount_pattern, text, re.IGNORECASE)
    if subtotal_match:
        extracted_data["Subtotal Amount"] = subtotal_match.group(1).strip()
        if extracted_data["Currency"] is None and re.search(r'[\$\£\€]', subtotal_match.group(1)):
             extracted_data["Currency"] = re.search(r'[\$\£\€]', subtotal_match.group(1)).group(0)


    # Tax Amount
    tax_match = re.search(r'(tax|vat)[:\s]*' + amount_pattern, text, re.IGNORECASE)
    if tax_match:
        extracted_data["Tax Amount"] = tax_match.group(1).strip()
        if extracted_data["Currency"] is None and re.search(r'[\$\£\€]', tax_match.group(1)):
             extracted_data["Currency"] = re.search(r'[\$\£\€]', tax_match.group(1)).group(0)


    # Total Amount
    # Prioritize keywords like "total", "grand total", "amount due"
    total_match = re.search(r'(total|grand total|amount due)[:\s]*' + amount_pattern, text, re.IGNORECASE)
    if total_match:
        extracted_data["Total Amount"] = total_match.group(1).strip()
        if extracted_data["Currency"] is None and re.search(r'[\$\£\€]', total_match.group(1)):
             extracted_data["Currency"] = re.search(r'[\$\£\€]', total_match.group(1)).group(0)
    elif extracted_data["Subtotal Amount"] and extracted_data["Tax Amount"]:
        # Simple fallback: if total not found but subtotal and tax are, sum them (requires cleaning amounts first)
        # This is a basic heuristic and may not be accurate if text includes total before tax/subtotal
        pass # More complex logic needed here to parse and sum amounts


    # Vendor and Customer Names/Addresses - More challenging without layout info.
    # Basic heuristic: Look for names/orgs near keywords like "Vendor", "Customer", "Bill To", "Ship To"
    vendor_match = re.search(r'(vendor|supplier|from)[:\s]*(.+)', text, re.IGNORECASE)
    if vendor_match and extracted_data["Vendor Name"] is None:
        # Try to capture the first line after the keyword as the name
        extracted_data["Vendor Name"] = vendor_match.group(2).split('\n')[0].strip()

    customer_match = re.search(r'(customer|client|bill to)[:\s]*(.+)', text, re.IGNORECASE)
    if customer_match and extracted_data["Customer Name"] is None:
         # Try to capture the first line after the keyword as the name
         extracted_data["Customer Name"] = customer_match.group(2).split('\n')[0].strip()


    # Line Items - Improved regex based on observed patterns.
    # This regex assumes columns like Description, Quantity, Unit Price, Line Total exist.
    # It's still a simplification and may fail on complex layouts.
    # It tries to be more flexible with spacing and currency symbols.
    line_item_pattern = re.compile(
        r'(.+?)\s+' # Description (non-greedy) followed by spaces
        r'(\d+)\s+' # Quantity (one or more digits) followed by spaces
        + amount_pattern + r'\s+' # Unit Price (using amount pattern) followed by spaces
        + amount_pattern, # Line Item Total (using amount pattern)
        re.IGNORECASE | re.DOTALL # Ignore case, allow . to match newline
    )

    # Look for a potential "line item" section (basic heuristic)
    lines = text.split('\n')
    line_item_section = False
    section_text = ""
    for i, line in enumerate(lines):
        if re.search(r'(description|item|qty|quantity|unit price|price per unit|amount|total)', line, re.IGNORECASE):
            line_item_section = True
            # Start searching for line items from this point
            section_text = "\n".join(lines[i:])
            break

    if line_item_section:
        for match in line_item_pattern.finditer(section_text):
            extracted_data["Line Items"].append({
                "Item Description": match.group(1).strip(),
                "Quantity": int(match.group(2).strip()),
                "Unit Price": match.group(3).strip(),
                "Line Item Total": match.group(4).strip()
            })


    return extracted_data


def extract_contract_data(text):
    """
    Refined function to extract data fields from contract text using spaCy and rule-based methods.
    """
    doc = nlp(text)
    extracted_data = {
        "Contract Title": None,
        "Party Names": [],
        "Effective Date": None,
        "Expiration Date": None,
        "Governing Law": None,
        "Key Clauses": {
            "Payment Terms": None,
            "Termination Clause": None,
            "Confidentiality Clause": None,
            # Add more common clauses here if needed
            "Intellectual Property Clause": None,
            "Limitation of Liability Clause": None,
            "Dispute Resolution Clause": None
        }
    }

    # Contract Title - Look for common contract titles near the beginning of the document
    title_match = re.search(r'^(.*?)(agreement|contract|service level agreement|partnership contract|consulting contract)', text, re.IGNORECASE | re.DOTALL)
    if title_match:
         # Capture the line containing the keyword as the title
         lines_before_keyword = title_match.group(1).split('\n')
         if lines_before_keyword:
             extracted_data["Contract Title"] = (lines_before_keyword[-1] + title_match.group(2)).strip()
         else:
             extracted_data["Contract Title"] = title_match.group(2).strip()


    # Party Names - Use spaCy entities near keywords like "between", "parties", "and"
    party_keywords = re.search(r'(between|parties|and)[:\s]*(.+)', text, re.IGNORECASE | re.DOTALL)
    if party_keywords:
        # Process the text following the party keyword
        party_text = party_keywords.group(2).split('\n')[0:5] # Look at the next few lines
        party_doc = nlp("\n".join(party_text))
        for ent in party_doc.ents:
            if (ent.label_ == "ORG" or ent.label_ == "PERSON") and ent.text.strip() not in extracted_data["Party Names"]:
                 extracted_data["Party Names"].append(ent.text.strip())

    # Dates - More robust date extraction and association with keywords
    date_pattern_flexible = r'(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}|\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2}(?:st|nd|rd|th)?(?:,|\s+)\s+\d{4})'

    # Effective Date
    effective_date_match = re.search(r'(effective date|date of effect)[:\s]*' + date_pattern_flexible, text, re.IGNORECASE)
    if effective_date_match:
        extracted_data["Effective Date"] = effective_date_match.group(1).strip()
    elif [ent.text for ent in doc.ents if ent.label_ == "DATE"]:
         # Fallback to first date entity if keyword not found
         extracted_data["Effective Date"] = [ent.text for ent in doc.ents if ent.label_ == "DATE"][0]


    # Expiration Date
    expiration_date_match = re.search(r'(expiration date|expires on|term ends)[:\s]*' + date_pattern_flexible, text, re.IGNORECASE)
    if expiration_date_match:
        extracted_data["Expiration Date"] = expiration_date_match.group(1).strip()


    # Governing Law - Look for keywords and capture the relevant sentence/phrase
    governing_law_match = re.search(r'(governing law|jurisdiction|this agreement shall be governed by)[:\s]*(.+?[\.\n])', text, re.IGNORECASE | re.DOTALL)
    if governing_law_match:
        extracted_data["Governing Law"] = (governing_law_match.group(1) + governing_law_match.group(2)).strip()


    # Key Clauses - More flexible keyword search and extraction of a surrounding text block
    keywords = {
        "Payment Terms": ["payment terms", "billing", "fees", "compensation"],
        "Termination Clause": ["termination", "expiration", "term and termination", "terminate"],
        "Confidentiality Clause": ["confidentiality", "non-disclosure", "confidential information"],
        "Intellectual Property Clause": ["intellectual property", "ip rights"],
        "Limitation of Liability Clause": ["limitation of liability", "liability"],
        "Dispute Resolution Clause": ["dispute resolution", "arbitration", "governing law"] # Governing law often near dispute resolution
    }

    for clause, terms in keywords.items():
        for term in terms:
            # Search for the term and extract a larger section of text around it
            # This regex attempts to capture a paragraph or multiple lines around the keyword
            clause_match = re.search(r'(.{0,300}' + re.escape(term) + r'.{0,500})', text, re.IGNORECASE | re.DOTALL)
            if clause_match:
                extracted_data["Key Clauses"][clause] = clause_match.group(1).strip()
                break # Stop searching for this clause once a match is found


    return extracted_data


def structure_invoice_data_for_excel(extracted_invoice_data):
    """
    Structures extracted invoice data into a list of dictionaries for Excel export.
    """
    structured_data = []
    # Extract common invoice details
    invoice_number = extracted_invoice_data.get("Invoice Number")
    invoice_date = extracted_invoice_data.get("Invoice Date")
    vendor_name = extracted_invoice_data.get("Vendor Name")
    vendor_address = extracted_invoice_data.get("Vendor Address")
    customer_name = extracted_invoice_data.get("Customer Name")
    customer_address = extracted_invoice_data.get("Customer Address")
    subtotal_amount = extracted_invoice_data.get("Subtotal Amount")
    tax_amount = extracted_invoice_data.get("Tax Amount")
    total_amount = extracted_invoice_data.get("Total Amount")
    currency = extracted_invoice_data.get("Currency")

    line_items = extracted_invoice_data.get("Line Items", [])

    if not line_items:
        # If no line items are found, create one row with just the main details
        structured_data.append({
            "Invoice Number": invoice_number,
            "Invoice Date": invoice_date,
            "Vendor Name": vendor_name,
            "Vendor Address": vendor_address,
            "Customer Name": customer_name,
            "Customer Address": customer_address,
            "Subtotal Amount": subtotal_amount,
            "Tax Amount": tax_amount,
            "Total Amount": total_amount,
            "Currency": currency,
            "Line Item Description": None,
            "Quantity": None,
            "Unit Price": None,
            "Line Item Total": None
        })
    else:
        # Create a row for each line item, repeating common details
        for item in line_items:
            structured_data.append({
                "Invoice Number": invoice_number,
                "Invoice Date": invoice_date,
                "Vendor Name": vendor_name,
                "Vendor Address": vendor_address,
                "Customer Name": customer_name,
                "Customer Address": customer_address,
                "Subtotal Amount": subtotal_amount,
                "Tax Amount": tax_amount,
                "Total Amount": total_amount,
                "Currency": currency,
                "Line Item Description": item.get("Item Description"),
                "Quantity": item.get("Quantity"),
                "Unit Price": item.get("Unit Price"),
                "Line Item Total": item.get("Line Item Total")
            })

    return structured_data

def structure_contract_data_for_excel(extracted_contract_data):
    """
    Structures extracted contract data into a list containing a single dictionary
    for Excel export.
    """
    # Contract data typically fits into a single row
    structured_data = [{
        "Contract Title": extracted_contract_data.get("Contract Title"),
        "Party Names": ", ".join(extracted_contract_data.get("Party Names", [])), # Join party names into a string
        "Effective Date": extracted_contract_data.get("Effective Date"),
        "Expiration Date": extracted_contract_data.get("Expiration Date"),
        "Governing Law": extracted_contract_data.get("Governing Law"),
        "Payment Terms Clause": extracted_contract_data.get("Key Clauses", {}).get("Payment Terms"),
        "Termination Clause": extracted_contract_data.get("Key Clauses", {}).get("Termination Clause"),
        "Confidentiality Clause": extracted_contract_data.get("Key Clauses", {}).get("Confidentiality Clause"),
        "Intellectual Property Clause": extracted_contract_data.get("Key Clauses", {}).get("Intellectual Property Clause"),
        "Limitation of Liability Clause": extracted_contract_data.get("Key Clauses", {}).get("Limitation of Liability Clause"),
        "Dispute Resolution Clause": extracted_contract_data.get("Key Clauses", {}).get("Dispute Resolution Clause")
        # Add other contract fields as needed
    }]

    return structured_data

def generate_excel_file(invoice_data_list, contract_data_list):
    """
    Generates an Excel file from structured invoice and contract data in memory.

    Args:
        invoice_data_list (list): A list of dictionaries for structured invoice data.
        contract_data_list (list): A list containing a single dictionary for structured contract data.

    Returns:
        bytes: The content of the Excel file as bytes, or None if an error occurs.
    """
    try:
        # Create pandas DataFrames
        invoice_df = pd.DataFrame(invoice_data_list)
        contract_df = pd.DataFrame(contract_data_list)

        # Use BytesIO to save the Excel file in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if not invoice_df.empty:
                invoice_df.to_excel(writer, sheet_name='Invoices', index=False)
            if not contract_df.empty:
                contract_df.to_excel(writer, sheet_name='Contracts', index=False)

        output.seek(0) # Go to the beginning of the stream
        return output.getvalue() # Return the content as bytes

    except Exception as e:
        st.error(f"An error occurred while generating the Excel file: {e}")
        return None

# --- Streamlit App ---

st.title("Intelligent Document Processor")
st.write("Upload an invoice or contract image/PDF to extract data into an Excel file.")

uploaded_file = st.file_uploader("Choose a document (Image or PDF)", type=["png", "jpg", "jpeg", "pdf"])

if uploaded_file is not None:
    file_details = {"FileName": uploaded_file.name, "FileType": uploaded_file.type}
    st.write(file_details)

    # To read file as string:
    # stringio = io.StringIO(uploaded_file.getvalue().decode("utf-8"))
    # st.write(stringio.read())

    # Check file type and process
    if file_details["FileType"] in ["image/png", "image/jpeg"]:
        st.info("Processing as an image (Invoice assumed for now)...")
        # Pass the file-like object directly to the OCR function
        extracted_text = ocr_image_to_text(uploaded_file)

        if extracted_text:
            st.subheader("Extracted Text:")
            st.text(extracted_text)

            st.info("Extracting Invoice Data...")
            extracted_data = extract_invoice_data(extracted_text)
            structured_invoice_data = structure_invoice_data_for_excel(extracted_data)
            structured_contract_data = [] # No contract data from image (for this simplified example)

            st.subheader("Extracted and Structured Data (Invoice):")
            if structured_invoice_data:
                 st.dataframe(pd.DataFrame(structured_invoice_data))
            else:
                 st.write("No structured invoice data extracted.")


            excel_bytes = generate_excel_file(structured_invoice_data, structured_contract_data)

            if excel_bytes:
                st.download_button(
                    label="Download Extracted Data (Excel)",
                    data=excel_bytes,
                    file_name=f"extracted_invoice_{uuid.uuid4().hex}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    elif file_details["FileType"] == "application/pdf":
        st.info("Processing as a PDF (Contract assumed for now)...")
        # For PDF, you would typically need a library to extract text from PDF
        # or use a cloud OCR/Document AI service that handles PDFs.
        # As a placeholder, we will simulate text extraction from PDF.

        # --- Placeholder for PDF Text Extraction ---
        st.warning("PDF text extraction is not fully implemented in this example.")
        st.warning("Using dummy text for PDF processing.")

        # Simulate text extraction for PDF
        # In a real application, use a library like pdfminer.six or a cloud service
        dummy_pdf_text = """
        CONFIDENTIAL AGREEMENT

        Effective Date: January 1, 2024

        Parties:
        Acme Corporation ("Acme")
        Beta Solutions ("Beta")

        This agreement outlines the terms of collaboration...

        Payment Terms: Acme will pay Beta within 30 days of service completion.

        Term and Termination: This agreement is for a period of one year and can be terminated with 60 days notice.

        Confidentiality: Both parties agree to keep all information confidential.

        Governing Law: This Agreement shall be governed by the laws of the State of California.
        """
        extracted_text = dummy_pdf_text # Use dummy text

        # --- End of Placeholder ---


        if extracted_text:
            st.subheader("Simulated Extracted Text from PDF:")
            st.text(extracted_text)

            st.info("Extracting Contract Data...")
            extracted_data = extract_contract_data(extracted_text)
            structured_contract_data = structure_contract_data_for_excel(extracted_data)
            structured_invoice_data = [] # No invoice data from contract (for this simplified example)

            st.subheader("Extracted and Structured Data (Contract):")
            if structured_contract_data:
                 st.dataframe(pd.DataFrame(structured_contract_data))
            else:
                 st.write("No structured contract data extracted.")


            excel_bytes = generate_excel_file(structured_invoice_data, structured_contract_data)

            if excel_bytes:
                st.download_button(
                    label="Download Extracted Data (Excel)",
                    data=excel_bytes,
                    file_name=f"extracted_contract_{uuid.uuid4().hex}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


    else:
        st.error("Unsupported file type. Please upload an image (PNG, JPG) or PDF.")

st.markdown("---")
st.write("Note: This is a basic implementation. For production use, consider:")
st.write("- More robust OCR for various layouts and document types.")
st.write("- Advanced NLP for more accurate and comprehensive data extraction.")
st.write("- Handling multi-page documents.")
st.write("- Error handling for poor quality scans.")
st.write("- More sophisticated UI/UX.")
