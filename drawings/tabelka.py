import os
import openpyxl
from utils.point_double_variant import APoint, ADouble, variants



def open_excel():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(script_dir)
    excel_path = os.path.join(parent_dir, "data.xlsm")
    sheet_name = "Obliczenia i dane"
    
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet = wb[sheet_name]
        return sheet
    except Exception as e:
        print(f"{e}")
        
        
def draw_tabela(model_space, sheet):
    
    ilość_wierszy = 24
    ilość_kolumn = 3
    wysokość_wierszy = 10
    szerokość_kolumn = 60
    wysokość_tekstu = 5
    start_x = -500
    start_y = -1000
    
    
    tabela = model_space.AddTable(APoint(start_x, start_y), ilość_wierszy, ilość_kolumn, wysokość_wierszy, szerokość_kolumn)
    
    tabela.Layer = "Tekst"
    # tabela.StyleName = "Elementy" żeby rysowało w tym stylu to najpierw sam musze dodac ten styl w AutoCadzie
    tabela.SetColumnWidth(0, 30)  # Kolumna 0 → szerokość 12
    tabela.SetColumnWidth(1, 80)
    tabela.SetColumnWidth(2, 50)
    
    tabela.SetText(0, 0, "Zestawienie elementów")
    # tabela.MergeCells(0, 0, 0, 2)  # Scal kolumny 0, 1 i 2 w wierszu 0
    # tabela.SetCellAlignment(0, 0, 5)
    
    tabela.SetText(1, 0, "Numer")  # Nagłówek kolumny "Nr"
    tabela.SetText(1, 1, "Nazwa elementu")  # Nagłówek kolumny "Nazwa"
    tabela.SetText(1, 2, "Długość [mm]")  # Nagłówek kolumny "Długość"
    # tabela.SetCellAlignment(1, 0, 5)  # Wyrównanie tekstu do środka w pierwszej kolumnie
    # tabela.SetCellAlignment(1, 1, 5)  # Wyrównanie tekstu do środka w drugiej kolumnie
    # tabela.SetCellAlignment(1, 2, 5)
    
    
    for i, row in enumerate(range(633, 655)):  # enumerate() doda indeks od 0
        row_data = []
        for col in range(8, 11):
            one_data = sheet.cell(row, col).value      
            row_data.append(one_data)

        # Pobranie danych z wiersza
        nr = row_data[0]  
        nazwa = row_data[1]  
        dlugosc = row_data[2]  

        # Wstawienie do tabeli AutoCAD
        tabela.SetText(i+2, 0, str(nr))  # Numer
        tabela.SetText(i+2, 1, str(nazwa))  # Nazwa
        tabela.SetText(i+2, 2, str(dlugosc))  # Długość

        # Wyrównanie tekstu do środka
        # tabela.SetCellAlignment(i+2, 0, 5)  
        # tabela.SetCellAlignment(i+2, 1, 5)  
        # tabela.SetCellAlignment(i+2, 2, 5)
        
        
    print("Tabelka narysowana")
    

    


def draw_tabelka(doc):
    
    model_space = doc.ModelSpace
    
    sheet = open_excel()
    
    draw_tabela(model_space, sheet)

