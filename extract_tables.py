import pdfplumber
import pandas as pd
import re
import os

# Folder containing PDF files
folder_path = "W:\\SHARE8 Physics\\RDC Invoices\\DoNotDeleteData\\ToImport"  # Update this to your folder path

# Regular expression pattern
pattern = re.compile(r'(\d+)\s([\d\.]+)\s(\d{2}/\d{2}/\d{4})\s(\d{2}/\d{2}/\d{4})\s(PIN \d+\s[,\-\w\s]+)\s(\d+)\s(\d+\.\d{2})\s(\d+\.\d{2})')

# List to accumulate all records
all_records = []

# Loop through each PDF file in the folder
for filename in os.listdir(folder_path):
    if filename.endswith(".pdf"):
        pdf_path = os.path.join(folder_path, filename)
        
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:  # Ensure there's text on the page
                    full_text += text + "\n"
            
            # Check for "Unreturned Dosimeters" before adding anything to the DataFrame
            if "Unreturned Dosimeters" in full_text:
                # Find all matches in the full text of the current PDF
                matches = pattern.findall(full_text)
                all_records.extend(matches)

# Convert all records to a DataFrame
df = pd.DataFrame(all_records, columns=["Group", "Order", "Shipped", "Due Date", "Identifier", "Quantity", "Price", "Amount"])

# Display the first few rows to verify the extraction
print(df.head())


print(df.head())  # Display the first few rows to verify the extraction
print(df.shape)  # Display the shape of the DataFrame to verify the extraction

import os
import pandas as pd

# Assuming 'df' is your DataFrame
excel_path = "W:\\SHARE8 Physics\\RDC Invoices\\DoNotDeleteData\\LateBadgesData.xlsx"  # Update this to your Excel file path

# Check if the Excel file already exists
if os.path.exists(excel_path):
    # Open the Excel file in append mode and write the DataFrame without headers
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)
else:
    # If the file does not exist, write the DataFrame with headers
    df.to_excel(excel_path, index=False)

print("Data saved to Excel.")
