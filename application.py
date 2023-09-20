from ClassToPDF.word import *

dir_actual = os.path.normpath(input("Escriba el directorio: "))

if not os.path.exists(dir_actual + '/pdfs'):
    os.mkdir(dir_actual + '/pdfs')
        
dir_destino = dir_actual + '/pdfs'
dir = FileToPDF(dir_actual, dir_destino)

while True:
    print('Convertir archivos [1]')
    print('Combinar PDFs [2]')
    print('Lista de archivos [3]')
    accion = int(input('>>> '))
    

    if accion == 1:
        print('Archivos Word [1]')
        print('Archivos Excel [2]')
        tipo = int(input('>>> '))
        
        if tipo == 1:
            dir.word_to_pdf()
            dir.compile(dir.destino)

        if tipo == 2:
            dir.excel_to_pdf()
            dir.compile(dir.destino)
    
    if accion == 2:
        dir.compile(dir.origen)
    
    if accion == 3:
        print('archivos word:')
        for file in dir.lista_docx: print(r'- ',file)
        print('\narchivos excel:')
        for file in dir.lista_xlsx: print(r'- ',file)
        print('\narchivos pdfs:')
        for file in dir.lista_pdf: print(r'- ',file)
    
    print('¿desea realizar otra accion?')
    print('Si [1]')
    print('No [2]')
    x = int(input('>>> '))
    if x == 1:
        continue
    if x == 2:
        break
