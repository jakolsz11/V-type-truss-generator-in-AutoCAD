import win32com.client
from setup.layers import create_layers
from drawings.axes import draw_axes
from drawings.przekroj_pasa_gornego import draw_przekroj_pasa_gornego
from drawings.słup import draw_słup
from drawings.przekroj_pasa_dolnego import draw_przekroj_pasa_dolnego
from drawings.pas_dolny import draw_pas_dolny
from drawings.pas_gorny import draw_pas_gorny
from drawings.krzyzulec_rozciagany import draw_krzyzulec_rozciagany
from drawings.przekroj_krzyzulca_rozciaganego import draw_przekroj_krzyzulca_rozciaganego
from drawings.krzyzulec_sciskany import draw_krzyzulec_sciskany
from drawings.przekroj_krzyzulca_sciskanego import draw_przekroj_krzyzulca_sciskanego
from drawings.krzyzulec_3 import draw_krzyzulec_3
from drawings.blachy import draw_blachy
from drawings.przewiazki import draw_przewiazki
from drawings.katownik_pasa_dolnego import draw_katownik_pasa_dolnego
from drawings.male_trojkaty import draw_male_trojkaty
from drawings.sruby import draw_sruby
from drawings.platwie import draw_platwie
from drawings.przekroj_A_A import draw_przekrojAA
from drawings.przekroj_B_B import draw_przekrojBB
from drawings.przekroj_C_C import draw_przekrojCC
from drawings.blachy_wyciagniete import draw_blachy_wyciagniete
from drawings.tabelka import draw_tabelka


if __name__ == "__main__":
    
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
    except:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        
    acad.Visible = True
    doc = acad.ActiveDocument
    
    try:
        create_layers(doc)
        draw_axes(doc)
        draw_przekroj_pasa_gornego(doc)
        draw_słup(doc)
        draw_przekroj_pasa_dolnego(doc)
        draw_pas_dolny(doc)
        draw_pas_gorny(doc)
        draw_krzyzulec_rozciagany(doc)
        draw_przekroj_krzyzulca_rozciaganego(doc)
        draw_krzyzulec_sciskany(doc)
        draw_przekroj_krzyzulca_sciskanego(doc)
        draw_krzyzulec_3(doc)
        draw_blachy(doc)
        draw_przewiazki(doc)
        draw_katownik_pasa_dolnego(doc)
        draw_male_trojkaty(doc)
        draw_sruby(doc)
        draw_platwie(doc)
        draw_przekrojAA(doc)
        draw_przekrojBB(doc)
        draw_przekrojCC(doc)
        draw_blachy_wyciagniete(doc)
        draw_tabelka(doc)

    except Exception as e:
        print(f"Błąd: {e}")
        
