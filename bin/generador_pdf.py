import os
import io
import xlrd
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

current_folder = os.path.dirname (__file__)
parent_folder = os.path.dirname (current_folder)
files_folder = os.path.join (parent_folder, "files")
data = os.path.join (files_folder, f"Data.xlsx")
original_pdf = os.path.join (current_folder, f"DIPLOMA_Basico.pdf")

def generatePDF(alumno, dni, fecha_inicio, fecha_fin, docente1, docente2):
    packet = io.BytesIO()
    # Fonts with epecific path
    pdfmetrics.registerFont(TTFont('arial','arial.ttf'))

    c = canvas.Canvas(packet, letter)

    #Página 1
    c.setFont('arial', 12)
    c.drawString(106, 340, alumno)
    c.drawString(176, 325, dni)


    c.setFont('arial', 10)
    c.drawString(324, 226, docente1)
    c.drawString(294, 214, docente2)

    c.setFont('arial', 11)
    c.drawString(106, 188, f"Realizado del {fecha_inicio} al {fecha_fin} en modalidad presencial")

    c.drawString(268, 149.5, fecha_fin)
    
    c.showPage()

    c.showPage()
    c.save()

    packet.seek(0)

    new_pdf = PdfFileReader(packet)
    
    existing_pdf = PdfFileReader(open(original_pdf, "rb"))
    output = PdfFileWriter()
    
    #Página sin editar
    page=existing_pdf.pages[0]
    output.add_page(page)
    
    #Página Editada
    page = existing_pdf.pages[1]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)

    new_pdf = os.path.join (files_folder, f"Diploma {alumno}.pdf")
    output_stream = open(new_pdf, "wb")
    output.write(output_stream)
    output_stream.close()
  
wb = xlrd.open_workbook(data) 

hoja = wb.sheet_by_index(0) 
for i in range (1, hoja.nrows):
    print(hoja.cell_value(i, 0))
    print(hoja.cell_value(i, 1))
    print(hoja.cell_value(i, 2))
    print(hoja.cell_value(i, 3))
    print(hoja.cell_value(i, 4))
    print(hoja.cell_value(i, 5))
    print(hoja.cell_value(i, 6))

    print("_______________________________")
    
    
    alumno = hoja.cell_value(i, 0)
    dni = hoja.cell_value(i, 1)
    fecha_inicio = hoja.cell_value(i, 2)
    fecha_fin = hoja.cell_value(i, 3)
    docente1 = hoja.cell_value(i, 4)
    docente2 = hoja.cell_value(i, 5)
    
    generatePDF(alumno, dni, fecha_inicio, fecha_fin, docente1, docente2)
print("Documentos generados correctamente")    
input()