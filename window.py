import curses
import os
from ClassToPDF.word import *
# Definir la clase FileToPDF aquí o importarla adecuadamente

menu_principal = ['CONVERTIR', 'COMBINAR PDFs', 'LISTA', 'CAMBIAR DIRECTORIO', 'EXIT']
menu_convertir = ['Word ----> PDF', 'Excel ----> PDF', 'Todo ----> PDF']
menu_lista = []

def print_menu(stdscr, selected_row_idx, lista):
    stdscr.clear()
    h, w = stdscr.getmaxyx()

    for idx, row in enumerate(lista):
        x = w // 2 - len(row) // 2
        y = h // 2 - len(lista) // 2 + idx
        if idx == selected_row_idx:
            stdscr.attron(curses.color_pair(1))
            stdscr.addstr(y, x, row)
            stdscr.attroff(curses.color_pair(1))
        else:
            stdscr.addstr(y, x, row)

    stdscr.refresh()

def sec_combinar(stdscr,dir_origen,dir_destino):
    stdscr.clear()
    print_menu(stdscr,0,['CARGANDO...'])
    stdscr.refresh()
    dir = FileToPDF(dir_origen,dir_destino)
    dir.compile(dir_origen)
    stdscr.clear()
    print_menu(stdscr,0,[f'Se unieron {len(dir.lista_pdf)} archivos',''] + dir.lista_pdf)
    stdscr.refresh()
    stdscr.getch()

def sec_convertir(stdscr, dir_origen, dir_destino):
    dir = FileToPDF(dir_origen, dir_destino)
    current_row_idx = 0
    print_menu(stdscr, current_row_idx, menu_convertir)

    while True:
        key = stdscr.getch()
        stdscr.clear()
        if key == 450 and current_row_idx > 0:  # 450 Flecha para arriba
            current_row_idx -= 1
        elif key == 456 and current_row_idx < len(menu_convertir) - 1:  # 456 Flecha para abajo
            current_row_idx += 1
        elif key in [10, 13]:
            if current_row_idx == 0:
                print_menu(stdscr,1,['CARGANDO...'])
                dir.word_to_pdf()
                print_menu(stdscr,2,[f'Se han convertido {len(dir.lista_docx)} archivos correctamente','','VOLVER'])
                stdscr.getch()
            if current_row_idx == 1:
                print_menu(stdscr,1,['CARGANDO...'])
                dir.excel_to_pdf()
                print_menu(stdscr,2,[f'Se han convertido {len(dir.lista_xlsx)} archivos correctamente','','VOLVER'])
                stdscr.getch()
            if current_row_idx == 2:
                print_menu(stdscr,1,['CARGANDO...'])
                dir.word_to_pdf()
                dir.excel_to_pdf()
                print_menu(stdscr,2,[f'Se han convertido {len(dir.lista_xlsx) + len(dir.lista_docx)} archivos correctamente','','VOLVER'])
                stdscr.getch()
        elif key == 27:  # 27 escape
            break
        print_menu(stdscr, current_row_idx, menu_convertir)
        stdscr.refresh()

def sec_cambiar_dir(stdscr):
    curses.curs_set(1)
    curses.echo()
    n=0
    while True:
        stdscr.clear()
        stdscr.refresh()
        if n == 1:
            stdscr.attron(curses.color_pair(2))
            stdscr.addstr(0,0,'El directorio NO EXISTE')
            stdscr.attroff(curses.color_pair(2))
        
        stdscr.addstr(1, 0, 'Escriba el Directorio ORIGEN:')
        stdscr.refresh()
        dir_origen = stdscr.getstr(2, 0).decode("utf-8")
        
        if os.path.exists(os.path.normpath(dir_origen)) == True and dir_origen != '':
            break
        n=1

    stdscr.clear()
    n=0

    while True:
        stdscr.clear()
        stdscr.refresh()
        if n == 1:
            stdscr.attron(curses.color_pair(2))
            stdscr.addstr(0,0,'El directorio NO EXISTE')
            stdscr.attroff(curses.color_pair(2))
        
        stdscr.addstr(1, 0, 'Escriba el Directorio DESTINO:')
        stdscr.refresh()
        dir_destino = stdscr.getstr(2, 0).decode("utf-8")
        
        if os.path.exists(os.path.normpath(dir_destino)) == True or dir_destino == "":
            break
        n=1

    if dir_destino == '':
        dir_destino = dir_origen

    if not os.path.exists(dir_destino + '/pdfs'):
        os.mkdir(dir_destino + '/pdfs')

    dir_destino = dir_origen + '/pdfs'
    curses.curs_set(0)
    curses.noecho()
    return dir_origen, dir_destino

def sec_lista(stdscr,dir_origen,dir_destino):
    stdscr.clear()
    dir = FileToPDF(dir_origen,dir_destino)
    files = os.listdir(dir.origen)
    for idx, file in enumerate(files):
        if file in dir.lista_docx:
            stdscr.attron(curses.color_pair(3))
            stdscr.addstr(idx,0,file)
            stdscr.attroff(curses.color_pair(3))
        elif file in dir.lista_xlsx:
            stdscr.attron(curses.color_pair(4))
            stdscr.addstr(idx,0,file)
            stdscr.attroff(curses.color_pair(4))
        elif file in dir.lista_pdf:
            stdscr.attron(curses.color_pair(2))
            stdscr.addstr(idx,0,file)
            stdscr.attroff(curses.color_pair(2))
        elif file.find('.') == -1:
            stdscr.attron(curses.color_pair(1))
            stdscr.addstr(idx,0,file)
            stdscr.attroff(curses.color_pair(1))
        else:
            stdscr.addstr(idx,0,file)
    stdscr.refresh()
    stdscr.getch()

def main(stdscr):
    curses.init_pair(1, curses.COLOR_BLACK, curses.COLOR_WHITE)
    curses.init_pair(2, curses.COLOR_WHITE, curses.COLOR_RED)
    curses.init_pair(3, curses.COLOR_WHITE, curses.COLOR_BLUE)
    curses.init_pair(4, curses.COLOR_WHITE, curses.COLOR_GREEN)
    curses.cbreak()
    stdscr.keypad(True)
    current_row_idx = 0
    
    dir_origen, dir_destino = sec_cambiar_dir(stdscr) 

    while True:
        print_menu(stdscr, current_row_idx, menu_principal)
        stdscr.addstr(0,(stdscr.getmaxyx()[1] - len(dir_origen)) //2,dir_origen)
        key = stdscr.getch()
        stdscr.clear()
        if key == 450 and current_row_idx > 0:  # 450 Flecha para arriba
            current_row_idx -= 1
        elif key == 456 and current_row_idx < len(menu_principal) - 1:  # 456 Flecha para abajo
            current_row_idx += 1
        elif key in [10, 13]:
            if current_row_idx == 0:
                sec_convertir(stdscr, dir_origen, dir_destino)
            elif current_row_idx == 1:
                sec_combinar(stdscr,dir_origen,dir_destino)
            elif current_row_idx == 2:
                sec_lista(stdscr,dir_origen,dir_destino)
            elif current_row_idx == 3:
                dir_origen, dir_destino = sec_cambiar_dir(stdscr)
            else:
                break
        elif key == 27:  # 27 Escape
            break
        print_menu(stdscr, current_row_idx, menu_principal)
        stdscr.refresh()
    
    


curses.wrapper(main)
