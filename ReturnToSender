import fitz  # PyMuPDF
import os
import pandas as pd
from PIL import Image
from pyzbar.pyzbar import decode
import io

# Function to extract and decode barcodes from PDF files
def extract_and_decode_barcodes(pdf_file_name):
    barcodes = []
    # Open the PDF file
    document = fitz.open(pdf_file_name)

    # Iterate over each page
    for page_num in range(len(document)):
        page = document[page_num]
        # Iterate through each image on the page
        for img in page.get_images(full=True):
            xref = img[0]
            base_image = document.extract_image(xref)
            image_bytes = base_image["image"]
            image = Image.open(io.BytesIO(image_bytes))
            decoded_objects = decode(image)
            for obj in decoded_objects:
                if obj.type == 'CODE39':
                    barcode_data = {
                        "File": pdf_file_name.replace('.pdf', ''),  # Remove ".pdf" from the file name
                        "Data": obj.data.decode("utf-8")
                    }
                    barcodes.append(barcode_data)

    return barcodes

# Function to extract address from XLSX files
def extract_address_from_xlsx(xlsx_file_name):
    address_data = []
    try:
        df = pd.read_excel(xlsx_file_name, engine='openpyxl')
        address_line_1 = df.iloc[13, 1]  # Cell B15
        city_state_zip = df.iloc[14, 1]  # Cell B16
        full_address = f"{address_line_1}, {city_state_zip}"
        file_name_without_extension = os.path.splitext(xlsx_file_name)[0]
        address_data.append({'File': file_name_without_extension, 'Address': full_address})
    except Exception as e:
        print(f"Error processing {xlsx_file_name}: {str(e)}")
    
    return address_data

# Get the list of all files in the directory
file_list = os.listdir('.')

# Initialize empty lists to store barcode and address data
all_barcodes = []
all_addresses = []

# Process each file
for file_name in file_list:
    if file_name.lower().endswith('.pdf'):
        print(f'Processing PDF: {file_name}...')
        barcodes = extract_and_decode_barcodes(file_name)
        all_barcodes.extend(barcodes)
    elif file_name.lower().endswith('.xlsx'):
        print(f'Processing XLSX: {file_name}...')
        addresses = extract_address_from_xlsx(file_name)
        all_addresses.extend(addresses)

# Create DataFrames from barcode and address data
barcode_df = pd.DataFrame(all_barcodes)
address_df = pd.DataFrame(all_addresses)

# Merge the DataFrames on the "File" column
combined_df = pd.merge(barcode_df, address_df, on="File")

# Save the merged DataFrame to a CSV file
combined_csv_file_name = 'Combined_Output.csv'
combined_df.to_csv(combined_csv_file_name, index=False)

print(f'Combined data saved to {combined_csv_file_name}')
