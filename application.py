import os
from ClassToPDF.word import *

dir_actual = input("Escriba el directorio contenedor de los archivos docx: ")

if dir_actual.strip()[-1] != '/':
    dir_actual = dir_actual + '/'

if not os.path.exists(dir_actual + '/pdfs'):
    os.mkdir(dir_actual + '/pdfs')

dir_destino = dir_actual + '/pdfs'

lista_de_docs = lista_docx(os.listdir(dir_actual))

for archivo in lista_de_docs:
    archivo = FileWord(dir_actual + archivo, dir_destino)
    archivo.word_to_pdf()
