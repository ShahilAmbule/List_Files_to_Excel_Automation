import os
import xlsxwriter
import time

# Set the path to the desktop and the folder
desktop_path = r"Desktop PAth"
folder_path = r"folder path which is on desktop"

# Check if the folder exists
if not os.path.exists(folder_path):
    print(f"The folder {folder_path} does not exist.")
    exit(1)

print(f"Found the folder: {folder_path}")

# Create a list to store the file names
file_names = []

#if not os.path.exists(folder_path)
    #printf(f'The folder {folder_path} does not exists)
    #exit(1)    
# Iterate through the files in the folder
print("Starting to iterate through files in the folder...")
for file in os.listdir(folder_path):
    # Get the file name without the extension
    file_name = os.path.splitext(file)[0]
    file_names.append(file_name)
    print(f"Added file name: {file_name}")
    
    time.sleep(0.1)

print(f"Collected file names: {file_names}")

# Create an Excel file
output_path = os.path.join(desktop_path, 'File_List.xlsx')
workbook = xlsxwriter.Workbook(output_path)
worksheet = workbook.add_worksheet()

print("Writing file names to Excel sheet...")
# Write the file names to the Excel sheet
for i, file_name in enumerate(file_names):
    worksheet.write(i, 0, file_name)
    print(f"Wrote to Excel: {file_name}")

# Close the Excel file
workbook.close()

print(f"File list saved to {output_path}")
