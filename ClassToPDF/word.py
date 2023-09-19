import win32com.client
import os
import fitz
class FileToPDF():

    def __init__(self, origen, destino):
        try:
            self.origen = os.path.normpath(origen)
            self.destino = os.path.normpath(destino)
        except Exception as e:
            print('Ingrese una ruta valida')
        self.lista_docx = lista_docx(self.origen)
        self.lista_xlsx = lista_xlsx(self.origen)
        self.lista_pdf = lista_pdf(self.origen)

    def word_to_pdf(self):
        if len(self.lista_docx) == 0:
            print(f'No hay archivos .docx en el directorio {self.origen}')
            return 
        try:
            # Crear una instancia de Word
            word = win32com.client.Dispatch('Word.Application')
            docx_n = 0
            word.Visible = False
            for file in self.lista_docx:
                docs_name = file.split('.')[0].replace(' ', '_')
                
                # Abrir el archivo de Word
                file_path = os.path.join(self.origen, file)
                doc = word.Documents.Open(file_path)

                # Guardar el archivo como PDF
                pdf_path = os.path.join(self.destino, docs_name + '.pdf')
                doc.SaveAs(pdf_path, FileFormat=17)  # 17 es el formato para PDF

                # Cerrar el archivo de Word
                doc.Close()
                docx_n = docx_n + 1

            # Salir de Word
            word.Quit()
            print(f'Se realizo con exito la conversion de {docx_n} archivos .docx')
        except Exception as e:
            print(f"Se produjo un error: {str(e)}")

    def excel_to_pdf(self):
        if len(self.lista_xlsx) == 0:
            print(f'No hay archivos .xlsx en el directorio {self.origen}')
            return
        try:
            # Crear una instancia de Excel
            excel = win32com.client.Dispatch('Excel.Application')
            excel.Visible = False  # Evitar que Excel sea visible durante la conversión
            excel_n = 0

            for file in self.lista_xlsx:

                excel_file_name = file.split('.')[0].replace(' ', '_')

                # Abrir el archivo de Excel
                file_path = os.path.join(self.origen, file)
                workbook = excel.Workbooks.Open(file_path)

                # Guardar el archivo como PDF
                pdf_path = os.path.join(self.destino, excel_file_name + '.pdf')
                workbook.ExportAsFixedFormat(0, pdf_path)  # 0 es el formato para PDF

                # Cerrar el archivo de Excel
                workbook.Close()

                excel_n = excel_n + 1

            # Salir de Excel
            excel.Quit()
            print(f'Se realizo con exito la conversion de {excel_n} archivos .xlsx')
        except Exception as e:
            print(f"Se produjo un error: {str(e)}")

    def compile(self, lista_pdf, origen):
        if len(lista_pdf) == 0:
            print (f'No hay archivos .pdf en {origen}')
            return
        
        result = fitz.open()

        for pdf in lista_pdf:
            with fitz.open(os.path.join(origen, pdf)) as mfile:
                result.insert_pdf(mfile)
            
        result.save(os.path.join(self.destino, 'resultado.pdf'))

        print(f'Se unieron los siguientes pdf en {os.path.join(self.destino, "resultado.pdf")}:')
        for pdf in lista_pdf: print(pdf)



def es_docx(file):
    return file.split('.')[-1] == 'docx'

def lista_docx(dir):
    lista_docx = []
    for file in os.listdir(dir):
     if es_docx(file):
         lista_docx.append(file)
    return lista_docx

def es_xlsx(file):
    return file.split('.')[-1] == 'xlsx'

def lista_xlsx(dir):
    lista_xlsx = []
    for file in os.listdir(dir):
        if es_xlsx(file) and not file.startswith("~$"): # Saltea archivo temporal
            lista_xlsx.append(file)
    return lista_xlsx

def es_pdf(file):
    return file.split('.')[-1] == 'pdf'

def lista_pdf(dir):
    lista_pdf = []
    for file in os.listdir(dir):
        if es_pdf(file):
            lista_pdf.append(file)
    return lista_pdf

