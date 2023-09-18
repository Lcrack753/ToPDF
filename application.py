from ClassToPDF.word import *

dir_actual = input("Escriba el directorio: ")

if not os.path.exists(dir_actual + '/pdfs'):
    os.mkdir(dir_actual + '/pdfs')


dir_destino = dir_actual + '/pdfs'
direc_origen= FileToPDF(dir_actual, dir_destino)

n = int(input('escriba 1 para convertir documentos .docx\nEscriba 2 para convertir documentos .xlxs: '))
if n == 1:
    direc_origen.word_to_pdf()
elif n == 2:
    direc_origen.excel_to_pdf()
else:
    print('ERROR inesperado')
