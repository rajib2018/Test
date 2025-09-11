import streamlit as st
import pytesseract
from PIL import Image
import pandas as pd
import os
import re
import uuid
import io

# --- Function Definitions ---

def ocr_image_to_text(image_file):
    """
    Performs OCR on an image file-like object and extracts text content.
    Requires Tesseract to be installed and in PATH.

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

def extract_invoice_data_rule_based(text):
    """
    Extracts data fields from invoice text using rule-based methods (regex).
    This is a simplified, non-LLM approach.
    """
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

    # Rule-based extraction using regular expressions and keywords
    # Invoice Number
    invoice_number_match = re.search(r'(invoice number|invoice no|invoice #|inv #|bill no)[:\s]*([A-Z0-9\-\/]+)', text, re.IGNORECASE)
    if invoice_number_match:
        extracted_data["Invoice Number"] = invoice_number_match.group(2).strip()

    # Invoice Date - Flexible date formats
    date_match = re.search(r'(invoice date|date)[:\s]*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}|\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2}(?:st|nd|rd|th)?(?:,|\s+)\s+\d{4})', text, re.IGNORECASE)
    if date_match:
        extracted_data["Invoice Date"] = date_match.group(2).strip()

    # Currency Symbol
    currency_match = re.search(r'([\$\£\€])', text)
    if currency_match:
        extracted_data["Currency"] = currency_match.group(1)

    # Amounts - Capture currency symbol and various number formats
    amount_pattern = r'([\$\£\€]?\s*\d{1,3}(?:,\d{3})*(?:\.\d{2})?)'

    # Subtotal Amount
    subtotal_match = re.search(r'(subtotal)[:\s]*' + amount_pattern, text, re.IGNORECASE)
    if subtotal_match:
        extracted_data["Subtotal Amount"] = subtotal_match.group(1).strip()

    # Tax Amount
    tax_match = re.search(r'(tax|vat)[:\s]*' + amount_pattern, text, re.IGNORECASE)
    if tax_match:
        extracted_data["Tax Amount"] = tax_match.group(1).strip()

    # Total Amount - Prioritize keywords
    total_match = re.search(r'(total|grand total|amount due)[:\s]*' + amount_pattern, text, re.IGNORECASE)
    if total_match:
        extracted_data["Total Amount"] = total_match.group(1).strip()


    # Vendor and Customer Names (simplified rule-based)
    vendor_match = re.search(r'(vendor|supplier|from)[:\s]*(.+)', text, re.IGNORECASE)
    if vendor_match:
        extracted_data["Vendor Name"] = vendor_match.group(2).split('\n')[0].strip()

    customer_match = re.search(r'(customer|client|bill to)[:\s]*(.+)', text, re.IGNORECASE)
    if customer_match:
         extracted_data["Customer Name"] = customer_match.group(2).split('\n')[0].strip()


    # Line Items - Basic regex for structured lines
    line_item_pattern = re.compile(
        r'(.+?)\s+' # Description
        r'(\d+)\s+' # Quantity
        + amount_pattern + r'\s+' # Unit Price
        + amount_pattern, # Line Item Total
        re.IGNORECASE | re.DOTALL
    )

    # Look for a potential line item section
    lines = text.split('\n')
    line_item_section = False
    section_text = ""
    for i, line in enumerate(lines):
        if re.search(r'(description|item|qty|quantity|unit price|price per unit|amount|total)', line, re.IGNORECASE):
            line_item_section = True
            section_text = "\n".join(lines[i:])
            break

    if line_item_section:
        for match in line_item_pattern.finditer(section_text):
            try:
                extracted_data["Line Items"].append({
                    "Item Description": match.group(1).strip(),
                    "Quantity": int(match.group(2).strip()),
                    "Unit Price": match.group(3).strip(),
                    "Line Item Total": match.group(4).strip()
                })
            except ValueError:
                # Handle cases where quantity is not a valid integer
                continue # Skip this line item if quantity is invalid


    return extracted_data

def extract_contract_data_rule_based(text):
    """
    Extracts data fields from contract text using rule-based methods (regex).
    This is a simplified, non-LLM approach.
    """
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
        }
    }

    # Contract Title
    title_match = re.search(r'^(.*?)(agreement|contract|service level agreement|partnership contract|consulting contract)', text, re.IGNORECASE | re.DOTALL)
    if title_match:
         lines_before_keyword = title_match.group(1).split('\n')
         if lines_before_keyword:
             extracted_data["Contract Title"] = (lines_before_keyword[-1] + title_match.group(2)).strip()
         else:
             extracted_data["Contract Title"] = title_match.group(2).strip()


    # Party Names (Simplified rule-based - look for names near keywords)
    party_keywords = re.search(r'(between|parties|and)[:\s]*(.+)', text, re.IGNORECASE | re.DOTALL)
    if party_keywords:
        potential_parties_text = party_keywords.group(2).split('\n')[0:5]
        # Basic heuristic: look for capitalized words or phrases as potential names
        for line in potential_parties_text:
            potential_names = re.findall(r'[A-Z][a-zA-Z]*(?:\s+[A-Z][a-zA-Z]*)*', line)
            for name in potential_names:
                 if len(name.split()) > 1 and name.strip() not in extracted_data["Party Names"]: # Filter short single words
                     extracted_data["Party Names"].append(name.strip())


    # Dates
    date_pattern_flexible = r'(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}|\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2}(?:st|nd|rd|th)?(?:,|\s+)\s+\d{4})'

    # Effective Date
    effective_date_match = re.search(r'(effective date|date of effect)[:\s]*' + date_pattern_flexible, text, re.IGNORECASE)
    if effective_date_match:
        extracted_data["Effective Date"] = effective_date_match.group(1).strip()

    # Expiration Date
    expiration_date_match = re.search(r'(expiration date|expires on|term ends)[:\s]*' + date_pattern_flexible, text, re.IGNORECASE)
    if expiration_date_match:
        extracted_data["Expiration Date"] = expiration_date_match.group(1).strip()


    # Governing Law
    governing_law_match = re.search(r'(governing law|jurisdiction|this agreement shall be governed by)[:\s]*(.+?[\.\n])', text, re.IGNORECASE | re.DOTALL)
    if governing_law_match:
        extracted_data["Governing Law"] = (governing_law_match.group(1) + governing_law_match.group(2)).strip()


    # Key Clauses (Simplified rule-based - look for keywords and extract surrounding text)
    keywords = {
        "Payment Terms": ["payment terms", "billing", "fees", "compensation"],
        "Termination Clause": ["termination", "expiration", "term and termination", "terminate"],
        "Confidentiality Clause": ["confidentiality", "non-disclosure", "confidential information"],
        # Add more common clauses here if needed with their keywords
    }

    for clause, terms in keywords.items():
        for term in terms:
            # Search for the term and extract a section of text around it
            clause_match = re.search(r'(.{0,200}' + re.escape(term) + r'.{0,200})', text, re.IGNORECASE | re.DOTALL)
            if clause_match:
                extracted_data["Key Clauses"][clause] = clause_match.group(1).strip()
                break # Stop searching for this clause once a match is found


    return extracted_contract_data_rule_based


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

st.title("Intelligent Document Processor (Rule-Based)")
st.write("Upload an invoice or contract image/PDF to extract data into an Excel file.")
st.write("Note: This version uses rule-based extraction, not LLMs.")

uploaded_file = st.file_uploader("Choose a document (Image or PDF)", type=["png", "jpg", "jpeg", "pdf"])

if uploaded_file is not None:
    file_details = {"FileName": uploaded_file.name, "FileType": uploaded_file.type}
    st.write(file_details)

    # To read file as string:
    # stringio = io.StringIO(uploaded_file.getvalue().decode("utf-8"))
    # st.write(stringio.read())

    extracted_text = None
    extracted_data = None
    structured_invoice_data = []
    structured_contract_data = []

    # Check file type and process
    if file_details["FileType"] in ["image/png", "image/jpeg"]:
        st.info("Processing as an image (assuming Invoice or Contract)...")
        extracted_text = ocr_image_to_text(uploaded_file)

    elif file_details["FileType"] == "application/pdf":
        st.info("Processing as a PDF (assuming Invoice or Contract)...")
        # --- Placeholder for PDF Text Extraction (Rule-Based) ---
        st.warning("PDF text extraction requires additional libraries (like pdfminer.six).")
        st.warning("Using dummy text for PDF processing.")

        # Simulate text extraction for PDF (replace with actual PDF text extraction)
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

    else:
        st.error("Unsupported file type. Please upload an image (PNG, JPG) or PDF.")


    if extracted_text:
        st.subheader("Extracted Text:")
        st.text(extracted_text)

        # Basic check to assume document type for rule-based extraction
        # In a real app, you might need more sophisticated classification
        document_type = None
        if re.search(r'invoice', extracted_text, re.IGNORECASE):
            document_type = "invoice"
            st.info("Document identified as Invoice.")
            extracted_data = extract_invoice_data_rule_based(extracted_text)
            structured_invoice_data = structure_invoice_data_for_excel(extracted_data)

        elif re.search(r'agreement|contract', extracted_text, re.IGNORECASE):
             document_type = "contract"
             st.info("Document identified as Contract.")
             # Note: The function call below was incorrect in the previous response. Corrected.
             extracted_data = extract_contract_data_rule_based(extracted_text)
             structured_contract_data = structure_contract_data_for_excel(extracted_data)


        else:
            st.warning("Could not confidently determine document type (Invoice or Contract).")
            st.info("Attempting extraction for both types...")
            # Attempt both extractions if type is ambiguous
            extracted_invoice_data = extract_invoice_data_rule_based(extracted_text)
            structured_invoice_data = structure_invoice_data_for_excel(extracted_invoice_data)

            extracted_contract_data = extract_contract_data_rule_based(extracted_text)
            structured_contract_data = structure_contract_data_for_excel(extracted_contract_data)
            extracted_data = {"Invoice Data": extracted_invoice_data, "Contract Data": extracted_contract_data}


        st.subheader("Extracted and Structured Data:")
        if structured_invoice_data:
             st.write("Invoice Data:")
             st.dataframe(pd.DataFrame(structured_invoice_data))
        if structured_contract_data:
             st.write("Contract Data:")
             st.dataframe(pd.DataFrame(structured_contract_data))

        # Generate Excel file if any data was extracted and structured
        if structured_invoice_data or structured_contract_data:
             excel_bytes = generate_excel_file(structured_invoice_data, structured_contract_data)

             if excel_bytes:
                 st.download_button(
                     label="Download Extracted Data (Excel)",
                     data=excel_bytes,
                     file_name=f"extracted_data_{uuid.uuid4().hex}.xlsx",
                     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                 )
        else:
            st.warning("No relevant data extracted for either Invoice or Contract format.")


st.markdown("---")
st.write("Note: This is a basic rule-based implementation.")
st.write("Rule-based extraction is highly dependent on consistent document layouts and phrasing.")
st.write("Accuracy may vary significantly across different document formats.")
st.write("For improved accuracy and flexibility, consider integrating LLMs or specialized document AI services.")
st.write("Manual installation of Tesseract OCR engine is required.")
