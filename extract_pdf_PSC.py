import sys
import pandas as pd
import os
import PyPDF2
import re

def process_files(batch):
    search_keywords = ['Clust er ID', 'EPIC No.', 'SLNO in P art', 'RLN Name', 'AC / P art No.', 'Name']
    batch_results = []

    for filename in batch:
        if filename.lower().endswith(".pdf"):
            #print('Processing PDF: ' + filename)
            try:
                with open(filename, 'rb') as pdfFileObject:
                    pdfReader = PyPDF2.PdfReader(pdfFileObject)
                    text = ''.join([page.extract_text() for page in pdfReader.pages if page.extract_text()])
                    extracted_values = {'FileName': filename}
                    for keyword in search_keywords:
                        if keyword == 'Name':
                            pattern = r'^(?!.*RLN \s+Name).*\bName:\s*(.*)$'
                        elif keyword == 'AC / P art No.':
                            pattern = r'\b' + re.escape(keyword) + r'\s*[:.-]?\s*([\w\d\s\/]+)'
                        else:
                            pattern = r'\b' + re.escape(keyword) + r'\s*[:.-]?\s*([\w\d\s]+)'
                        match = re.search(pattern, text)
                        extracted_values[keyword] = match.group(1).strip() if match else None
                    batch_results.append(extracted_values)
            except Exception as e:
                print(f"An error occurred while processing {filename}: {e}")
    return batch_results

def main(argv):
    path = 'C:/Users/WIN 10/Desktop/New folder/AC 239 MISSING PSE 27.12.2023'
    filelist = [os.path.join(root, file) for root, dirs, files in os.walk(path) for file in files]
    
    results = []
    batch_size = 100  # Processing 100 files at a time

    for i in range(0, len(filelist), batch_size):
        batch = filelist[i:i+batch_size]
        results += process_files(batch)
        print(f"Processed files {i+1} to {min(i+batch_size, len(filelist))}") 

    df = pd.DataFrame(results)

    split_columns = df['AC / P art No.'].str.split('/', expand=True)
    df['AC_No.'] = split_columns[0].str.strip()
    df['Part_No.'] = split_columns[1].str.strip()

    print(f'renaming columns...')

    df.rename(columns={
        'Clust er ID': 'Cluster_ID',
        'EPIC No.': 'Epic No',
        'SLNO in P art': 'Sl.No_Part', 
        'RLN Name': 'Rel.Name', 
       
    }, inplace=True)

    print(f'cleaning data...')
    df.drop(columns=['AC / P art No.'], inplace=True)
    for column in df.columns:
        df[column] = df[column].str.split('\n').str[0]
    df.to_csv('C:/Users/WIN 10/Desktop/ExtractPdf/PSE_239.csv', index=False)  # Saving as CSV for large data
    return df

if __name__ == "__main__":
    df = main(sys.argv[1:])
    print(df)