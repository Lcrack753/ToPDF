import win32com.client
import os

class FileToPDF():

    def __init__(self, origen, destino):
        try:
            self.origen = os.path.normpath(origen)
            self.destino = os.path.normpath(destino)
        except Exception as e:
            print('Ingrese una ruta valida')

    def word_to_pdf(self):
        if len(lista_docx(self.origen)) == 0:
            print(f'No hay archivos .docx en el directorio {self.origen}')
            return 
        try:
            # Crear una instancia de Word
            word = win32com.client.Dispatch('Word.Application')
            docx_n = 0
            word.Visible = False
            for file in lista_docx(self.origen):
                # Saltea archivo temporal
                if file.startswith("~$"):
                    continue  

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
        if len(lista_xlsx(self.origen)) == 0:
            print(f'No hay archivos .xlsx en el directorio {self.origen}')
            return
        try:
            # Crear una instancia de Excel
            excel = win32com.client.Dispatch('Excel.Application')
            excel.Visible = False  # Evitar que Excel sea visible durante la conversión

            for file in lista_xlsx(self.origen):
                # Saltea archivo temporal
                if file.startswith("~$"):
                    continue

                excel_n = 0

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
