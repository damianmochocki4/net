import os
from tkinter import Tk, filedialog
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Border, Side

# Function to select the directory
def select_directory():
    root = Tk()
    root.withdraw()  # Close the root window
    folder_selected = filedialog.askdirectory()
    return folder_selected

# Function to find the first folder starting with "Re" in the selected directory path
def find_re_folder(path):
    # Split the path into components using both '/' and '\'
    parts = path.replace("\\", "/").split("/")
    for part in parts:
        if part.startswith("Re"):
            return part  # Return the first part found starting with "Re"
    print("No folder starting with 'Re' found.")
    return None

# Function to scan a folder and create Excel summary if it contains files
def process_folder(folder_path, aggregated_data, deleted_files_data):
    folder_name = os.path.basename(folder_path)
    summary_data = []
    local_deleted_files_data = []  # Only for the current folder
    
    # List files and subfolders in the folder
    subdirs = next(os.walk(folder_path))[1]
    files = next(os.walk(folder_path))[2]
    
    num_subfolders = len(subdirs)
    num_files = len(files)
    num_pdf_files = len([file for file in files if file.lower().endswith('.pdf')])
    num_jpeg_files = len([file for file in files if file.lower().endswith(('.jpg', '.jpeg'))])
    num_other_files = num_files - num_pdf_files - num_jpeg_files
    num_deleted_files = 0

    # Process PDF files and delete them
    for file in files:
        if file.lower().endswith('.pdf'):
            file_path = os.path.join(folder_path, file)
            file_size = os.path.getsize(file_path)
            delete_time = datetime.now().strftime("%d.%m.%Y %H:%M")
            
            # Collect deleted file info for summary and local folder
            deleted_files_data.append({
                "Dateiname": file,
                "Ordnername": folder_name,
                "Dateityp": "PDF",
                "Löschdatum": delete_time,
                "Dateigröße": f"{file_size / 1024:.0f}kB",  # File size in KB
            })
            
            local_deleted_files_data.append({
                "Dateiname": file,
                "Ordnername": folder_name,
                "Dateityp": "PDF",
                "Löschdatum": delete_time,
                "Dateigröße": f"{file_size / 1024:.0f}kB",
            })
            
            os.remove(file_path)  # Delete the file
            num_deleted_files += 1

    # Append folder summary data if it contains files
    if num_files > 0:
        summary_row = {
            "Ordnername": folder_name,
            "Number of subfolders": num_subfolders,
            "Number of files": num_files,
            "Number of PDF files": num_pdf_files,
            "Number of JPEG files": num_jpeg_files,
            "Number of other files": num_other_files,
            "Number of deleted files": num_deleted_files
        }
        summary_data.append(summary_row)
        aggregated_data.append(summary_row)  # Append to aggregated data for final summary
    
    # Save summary to Excel if there are files in the folder
    if summary_data:
        save_to_excel(folder_path, folder_name, summary_data, local_deleted_files_data)


# Function to save data into Excel for each folder and adjust column width
def save_to_excel(folder_path, folder_name, summary_data, deleted_files_data):
    # Create a Pandas Excel writer
    excel_filename = f"Register_{folder_name}.xlsx"
    excel_path = os.path.join(folder_path, excel_filename)
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        # Create DataFrame for folder summary
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Folder Summary', index=False)
    
    # Load the workbook to add the deleted files below the summary
    wb = load_workbook(excel_path)
    sheet = wb['Folder Summary']

    # Add a title row for deleted files section
    start_row = sheet.max_row + 2  # Two rows after the last row of the summary
    sheet[f"A{start_row}"] = "Gelöschte Dateien / Deleted files"
    sheet[f"A{start_row}"].font = Font(bold=True)
    
    # Check if there are deleted files and add them to the sheet
    if deleted_files_data:
        deleted_files_df = pd.DataFrame(deleted_files_data, columns=[
            "Dateiname", "Ordnername", "Dateityp", "Löschdatum", "Dateigröße"
        ])
        
        # Append the deleted files data to the same sheet
        for r_idx, row in enumerate(dataframe_to_rows(deleted_files_df, index=False, header=True), start=start_row + 1):
            for c_idx, value in enumerate(row, start=1):
                cell = sheet.cell(row=r_idx, column=c_idx, value=value)
                
                # Apply bold font and thin border to the header row
                if r_idx == start_row + 1:  # First row of deleted files is the header
                    cell.font = Font(bold=True)
                    thin_border = Border(left=Side(style='thin'), 
                                         right=Side(style='thin'), 
                                         top=Side(style='thin'), 
                                         bottom=Side(style='thin'))
                    cell.border = thin_border

    # Adjust column width after adding data
    adjust_column_width(sheet)

    # Save the workbook
    wb.save(excel_path)

    print(f"Excel file '{excel_filename}' created in '{folder_path}'")

# Function to adjust column width in each sheet
def adjust_column_width(sheet):
    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column letter
        
        for cell in col:
            try:
                # Get length of the cell's value as a string
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        
        # Adjust width (add some padding to account for different character widths)
        adjusted_width = max_length + 2
        sheet.column_dimensions[column].width = adjusted_width

# Function to save the aggregated summary in a separate Excel file
def save_aggregated_summary(base_path, folder_name, aggregated_data, deleted_files_data, excel_subfolders_count):
    # Create a DataFrame for the aggregated data
    summary_df = pd.DataFrame(aggregated_data)
    
    # Create the summary Excel file above the selected directory
    summary_filename = f"Inventory_{folder_name}.xlsx"
    parent_path = os.path.dirname(base_path)  # Get parent directory
    summary_path = os.path.join(parent_path, summary_filename)
    
    # Save the aggregated summary to Excel
    with pd.ExcelWriter(summary_path, engine='openpyxl') as writer:
        # Create structured summary DataFrame
        structured_summary_df = pd.DataFrame({
            "Invoice name": [folder_name],
            "Number of subfolders": [excel_subfolders_count],
            "Number of files before deletion incl. Inventory file": [summary_df["Number of files"].sum() + 1],
            "Number of PDF files": [summary_df["Number of PDF files"].sum()],
            "Number of JPEG files": [summary_df["Number of JPEG files"].sum()],
            "Number of other files": [summary_df["Number of other files"].sum()],
            "Number of deleted files": [summary_df["Number of deleted files"].sum()],
        })

        structured_summary_df.to_excel(writer, sheet_name='Aggregated Summary', index=False)

        # Add a title row for deleted files section
        start_row = structured_summary_df.shape[0] + 3  # Two rows after the summary table
        writer.sheets['Aggregated Summary'].cell(row=start_row, column=1, value="Gelöschte Dateien / Deleted files").font = Font(bold=True)

        # Check if there are deleted files and add them to the summary
        if deleted_files_data:
            deleted_files_df = pd.DataFrame(deleted_files_data, columns=[
                "Dateiname", "Ordnername", "Dateityp", "Löschdatum", "Dateigröße"
            ])
            for r_idx, row in enumerate(dataframe_to_rows(deleted_files_df, index=False, header=True), start=start_row + 1):
                for c_idx, value in enumerate(row, start=1):
                    cell = writer.sheets['Aggregated Summary'].cell(row=r_idx, column=c_idx, value=value)
                    if r_idx == start_row + 1:  # Header row
                        cell.font = Font(bold=True)
                        # Apply thin borders to the header row
                        cell.border = Border(left=Side(style='thin'), 
                                             right=Side(style='thin'), 
                                             top=Side(style='thin'), 
                                             bottom=Side(style='thin'))
                    else:  # Apply borders to data rows
                        cell.border = Border(left=Side(style='thin'), 
                                             right=Side(style='thin'), 
                                             top=Side(style='thin'), 
                                             bottom=Side(style='thin'))

        # Apply formatting (borders and bold) to the structured summary
        for row in range(1, structured_summary_df.shape[0] + 1):
            for col in range(1, structured_summary_df.shape[1] + 1):
                cell = writer.sheets['Aggregated Summary'].cell(row=row, column=col)
                cell.border = Border(left=Side(style='thin'), 
                                     right=Side(style='thin'), 
                                     top=Side(style='thin'), 
                                     bottom=Side(style='thin'))

        adjust_column_width(writer.sheets['Aggregated Summary'])

    print(f"Aggregated summary saved as '{summary_filename}' in '{parent_path}'")

# Main execution: scan through each folder in the selected directory
def main():
    # Step 1: Select directory
    selected_directory = select_directory()
    if not selected_directory:
        print("No directory selected. Exiting.")
        return

    # Step 2: Identify the folder starting with "Re"
    base_folder = find_re_folder(selected_directory)
    if not base_folder:
        print(f"No folder starting with 'Re' found in '{selected_directory}'. Exiting.")
        return
    
    # Step 3: Prepare to collect aggregated data and count subfolders
    aggregated_data = []
    deleted_files_data = []
    excel_subfolders_count = 0  # Count the number of subfolders where Excel was created

    # Step 4: Walk through the directory and process each subfolder
    for root, subdirs, files in os.walk(selected_directory):
        if len(files) > 0:  # Only process folders with files
            process_folder(root, aggregated_data, deleted_files_data)
            excel_subfolders_count += 1  # Increment count for each processed folder

    # Step 5: Save the aggregated summary in the parent directory
    save_aggregated_summary(selected_directory, base_folder, aggregated_data, deleted_files_data, excel_subfolders_count)

if __name__ == "__main__":
    main()
