import streamlit as st
from PIL import Image
import pytesseract
import re
from pdf2image import convert_from_bytes

st.title('Simple IDP App')

uploaded_file = st.file_uploader('Upload an image or PDF', type=['png', 'jpg', 'jpeg', 'pdf'])

def extract_text(file):
    if file.type == 'application/pdf':
        images = convert_from_bytes(file.read())
        text = pytesseract.image_to_string(images[0])  # first page only
    else:
        image = Image.open(file)
        text = pytesseract.image_to_string(image)
    return text

if uploaded_file is not None:
    with st.spinner('Extracting text...'):
        text = extract_text(uploaded_file)
    st.header('Extracted Text')
    st.text_area('Text output', text, height=300)
    
    dates = re.findall(r'\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b', text)
    if dates:
        st.subheader('Extracted Dates')
        st.write(dates)
    else:
        st.write('No dates found')

    emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', text)
    if emails:
        st.subheader('Extracted Emails')
        st.write(emails)
    else:
        st.write('No emails found')
