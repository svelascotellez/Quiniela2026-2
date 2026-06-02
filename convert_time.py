import openpyxl
import re

wb = openpyxl.load_workbook('Quiniela_Mundial_2026_Final_vacia.xlsx')
sheet = wb['Calendario de Juegos']

pattern = re.compile(r'^(\d{2}):(\d{2})\s*-\s*(.*)')

for row in sheet.iter_rows():
    for cell in row:
        if isinstance(cell.value, str):
            val = cell.value.strip()
            if "Todos los horarios son del Este de Estados Unidos" in val:
                cell.value = "Todos los horarios son de la Ciudad de México."
                continue
            
            match = pattern.match(val)
            if match:
                hour = int(match.group(1))
                minute = int(match.group(2))
                rest = match.group(3)
                
                new_hour = (hour - 2) % 24
                
                new_time_str = f"{new_hour:02d}:{minute:02d} - {rest}"
                cell.value = new_time_str

wb.save('Quiniela_Mundial_2026_Final_vacia.xlsx')
print("Times converted to CDMX successfully.")
