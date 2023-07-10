import pandas as pd
import openpyxl

# Load the existing Excel file
filename = 'apr.xlsx'
book = openpyxl.load_workbook(filename)

# Create a new sheet with the modified data
new_data = pd.read_excel('modified_file.xlsx')
sheet_name = 'DATA_INV_updated'
book.remove(book[sheet_name])  # Remove the existing sheet
book.create_sheet(sheet_name, index=0)  # Create a new sheet
writer = pd.ExcelWriter(filename, engine='openpyxl')
writer.book = book
new_data.to_excel(writer, sheet_name=sheet_name, index=False)
writer.save()
writer.close()
