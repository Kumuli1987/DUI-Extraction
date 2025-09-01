import pandas as pd
from pypdf import PdfReader
import os
import re

pdf_folder = r"C:\Users\Priyanka\OneDrive\Documents\FASTAPI\Dinesh\Input"
#Flixible pattern
#DO_pattern = r'doi[\s\-\:]*(\d+)'
DO_pattern = r'https?://doi\.org/(.+)'
# Store Result
DO_list = [ ]
#Loop through pdfs
print("start")
for filename in os.listdir(pdf_folder):
    if filename.lower().endswith('.pdf'):
        #print(filename)
        pdf_path = os.path.join(pdf_folder, filename)
        print(f'pdf_path: {pdf_path}')
        reader = PdfReader(pdf_path)
        found = False
        
        for page_num, page in enumerate(reader.pages):
            text = page.extract_text()
            matches = re.findall(DO_pattern, text)
            DOI_Match = re.search(DO_pattern,text)
            if DOI_Match:
                DOI_group = DOI_Match.group(0)
            
            if matches:
                DO_list.append({
                    
                    "File Name": filename,
                    "Page": page_num + 1,
                    "DOI Number": matches[0],
                    "DOI_Group": DOI_group
                })
                found = True
                break
            
        reader.close()
        
        if not found:
            DO_list.append({
                
                "File Name": filename,
                "Page":page_num,
                "DOI Number": "Not Found"
            })
            
df = pd.DataFrame(DO_list)
df.to_excel(r"C:\Users\Priyanka\OneDrive\Documents\FASTAPI\Dinesh\Input\Extracted_DO.xlsx", index=False)
                
print("End")
