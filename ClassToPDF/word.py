import win32com.client
import re
import os

class FileWord():

    def __init__(self, origen,destino):
        self.origen= origen.replace('\\', '/')
        self.destino = destino.replace('\\', '/') + '/'
        self.nombre = self.origen.split('/')[-1].split('.')[0]

    def word_to_pdf(self):
        try:
            # Crear una instancia de Word
            word = win32com.client.Dispatch('Word.Application')

            # Abrir el archivo de Word
            doc = word.Documents.Open(self.origen)

            # Guardar el archivo como PDF
            doc.SaveAs(self.destino + self.nombre + '.pdf', FileFormat=17)  # 17 es el formato para PDF

            # Cerrar el archivo de Word
            doc.Close()

            # Salir de Word
            word.Quit()

            print(f"El archivo {self.nombre} se ha convertido a {self.nombre}.pdf exitosamente.")
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
        if es_xlsx(file):
            lista_xlsx.append(file)
    return lista_xlsx
