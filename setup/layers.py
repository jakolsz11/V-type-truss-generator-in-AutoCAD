import os
import openpyxl

def line_width_to_AutoCad_format(width):
    # if width == -1:
    #     return width
    # else:
    return width*100    
    
def load_linetype_if_needed(doc, linetype_name):
    """Sprawdza i ładuje styl linii, jeśli nie jest dostępny."""
    linetypes = doc.Linetypes
    try:
        _ = linetypes.Item(linetype_name)
    except:
        print(f"Ładowanie stylu linii: {linetype_name}")
        doc.Linetypes.Load(linetype_name, "acad.lin") 
        

def get_data(file_path, sheet_name):

    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb[sheet_name]
    
    data = []

    for row in range(7, 25): 
        row_data = []
        for col in range(24, 29):
            one_data = sheet.cell(row, col).value
            if col == 27:
                one_data = line_width_to_AutoCad_format(one_data)
            
            row_data.append(one_data) 
            
        
        data.append(row_data) 

    if None in data:
        raise ValueError("Nie wszystkie wartości zostały odczytane poprawnie z Excela.")

    return data


def create_layers(doc):

    script_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(script_dir)
    excel_path = os.path.join(parent_dir, "data.xlsm")
    sheet_name = "Parametry"
    
    data = get_data(excel_path, sheet_name)

    for row in data:
        name, color, style, width, scale = row
        load_linetype_if_needed(doc, style)
    
        layer = doc.Layers.Add(name)
        layer.Color = color 
        layer.Linetype = style
        
        if width != -1:
            layer.Lineweight = width
            
            

    
    