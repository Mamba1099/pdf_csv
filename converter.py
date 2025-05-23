import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import re

# extracting 168 pages pdf to excel using OCR

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
PDF_PATH = r'C:\Users\mamba\Downloads\print_MembersOnlineD233B698B1A93788F3BBD0CBC1EA0A16.cfusion1039.pdf'
OUTPUT_EXCEL = r'C:\Users\mamba\Downloads\info_extracted.xlsx'
print("Converting PDF to images...")
pages = convert_from_path(PDF_PATH, dpi=300, poppler_path=r"C:\poppler-24.08.0\Library\bin")


data = []

def extract_contact_block(text):
    company = re.search(r'^(.+?)\nContact:', text)
    name = re.search(r'Contact:\s*(.+)', text)
    function = re.search(r'Function:\s*(.+)', text)
    address = re.search(r'Address:\s*(.+?)(?:Phone:|E-mail:)', text, re.DOTALL)
    phone_fax = re.findall(r'Phone:\s*([\+\d\s\-()]+)(?:\s*Fax:\s*([\+\d\s\-()]+))?', text)
    email = re.search(r'E-mail:\s*(.+)', text)
    website = re.search(r'More information:\s*(.+)', text)

    # Clean Contact
    contact_person = name.group(1).strip() if name else 'NA'
    if contact_person.lower().startswith("contact:"):
        contact_person = contact_person[8:].strip()

    # Clean Phone and Fax
    phone = phone_fax[0][0].strip() if phone_fax else 'NA'
    fax = phone_fax[0][1].strip() if phone_fax and phone_fax[0][1] else 'NA'

    # Clean address
    raw_address = address.group(1).replace('\n', ', ').strip() if address else 'NA'
    address_parts = [part.strip() for part in raw_address.split(',') if part.strip()]
    street_address = ', '.join(address_parts[:-2]) if len(address_parts) > 2 else raw_address
    city = address_parts[-2] if len(address_parts) > 1 else 'NA'
    postal_code = re.findall(r'\b\d{4,6}\b', raw_address)
    postal_code = postal_code[-1] if postal_code else 'NA'
    country = address_parts[-1] if address_parts else 'NA'

    return {
        'Company': company.group(1).strip() if company else 'NA',
        'Contact Person': contact_person,
        'Function': function.group(1).strip() if function else 'NA',
        'Street Address': street_address,
        'City': city,
        'Postal Code': postal_code,
        'Country': country,
        'Phone': phone,
        'Fax': fax,
        'Email': email.group(1).strip() if email else 'NA',
        'Website': website.group(1).strip() if website else 'NA'
    }

print("Running OCR on pages...")
for i, page in enumerate(pages):
    print(f"Processing page {i+1}/{len(pages)}...")
    text = pytesseract.image_to_string(page)
    blocks = re.split(r'\n(?=[A-Z][^\n]+\nContact:)', text)
    for block in blocks:
        if 'Contact:' in block:
            entry = extract_contact_block(block)
            data.append(entry)

df = pd.DataFrame(data)
df.to_excel(OUTPUT_EXCEL, index=False)
print(f"\n✅ Extraction complete! Data saved to '{OUTPUT_EXCEL}'")
