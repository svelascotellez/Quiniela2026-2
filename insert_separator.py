import openpyxl

file_path = 'Quiniela_Mundial_2026_Final_vacia.xlsx'
wb = openpyxl.load_workbook(file_path)
ws = wb['Grupos']

# Insert a blank column at index 10 (J). This pushes the current column J to K.
ws.insert_cols(10)

# Make the new column J a narrow separator column
ws.column_dimensions['J'].width = 3

wb.save(file_path)
print("Separator column inserted successfully.")
