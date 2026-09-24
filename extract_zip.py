import zipfile, os, sys
zip_path = 'CallCenterViolazioni_main.zip'
extract_path = 'repo'
with zipfile.ZipFile(zip_path, 'r') as z:
    z.extractall(extract_path)
print('Extraction completed')
