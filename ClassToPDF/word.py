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
        self.lista_docx = lista_files(self.origen, tipo='docx')
        self.lista_xlsx = lista_files(self.origen, tipo='xlsx')
        self.lista_pdf = lista_files(self.origen, tipo='pdf')

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

    #Combina Los PDF
    def compile(self, origen):
        lista_pdf = lista_files(origen, tipo='pdf')
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

    def embebido(self):
        pass

#Comprueba si el archivo es un tipo de archivo especifico
def es_file(file,tipo):
    return file.split('.')[-1] == tipo

#Genera una lista con los archivos de un directorio de un tipo especifico
def lista_files(dir, tipo):
    lista_files = []
    for file in os.listdir(dir):
     if es_file(file, tipo):
         lista_files.append(file)
    return lista_files
    lista_pdf = []
    for file in os.listdir(dir):
        if es_pdf(file):
            lista_pdf.append(file)
    return lista_pdf




