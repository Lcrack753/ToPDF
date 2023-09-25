import curses
import os
from ClassToPDF.word import *

import curses
import os

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

def sec_convertir(stdscr):
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
                stdscr.clear()
                stdscr.addstr(2, 0, 'Escriba el Directorio ORIGEN: ')
                stdscr.refresh()
                dir_origen = stdscr.getstr(3, 0).decode("utf-8")
                # ... realiza acciones con dir_origen ...
        elif key == 27:  # 27 escape
            break
        print_menu(stdscr, current_row_idx, menu_convertir)
        stdscr.refresh()

def main(stdscr):
    curses.init_pair(1, curses.COLOR_BLACK, curses.COLOR_WHITE)
    curses.curs_set(0)
    curses.noecho()
    curses.cbreak()
    stdscr.keypad(True)
    current_row_idx = 0
    
    stdscr.clear()
    stdscr.addstr(0, 0, 'Escriba el Directorio ORIGEN: ')
    stdscr.refresh()
    dir_origen = stdscr.getstr(1, 0).decode("utf-8")

    stdscr.clear()
    stdscr.addstr(0, 0, 'Escriba el Directorio de DESTINO (presione Enter para usar el mismo directorio): ')
    stdscr.refresh()
    dir_destino = stdscr.getstr(1, 0).decode("utf-8")

    if dir_destino == '':
        dir_destino = dir_origen

    if not os.path.exists(dir_destino + '/pdfs'):
        os.mkdir(dir_destino + '/pdfs')

    dir_destino = dir_origen + '/pdfs'

    # Crear una instancia de FileToPDF con los directorios obtenidos
    dir = FileToPDF(dir_origen, dir_destino)

    print_menu(stdscr, current_row_idx, menu_principal)

    while True:
        key = stdscr.getch()
        stdscr.clear()
        if key == 450 and current_row_idx > 0:  # 450 Flecha para arriba
            current_row_idx -= 1
        elif key == 456 and current_row_idx < len(menu_principal) - 1:  # 456 Flecha para abajo
            current_row_idx += 1
        elif key in [10, 13]:
            if current_row_idx == 0:
                sec_convertir(stdscr)
            elif current_row_idx == 1:
                pass  # Combinar PDF
        elif key == 27:  # 27 Escape
            break
        print_menu(stdscr, current_row_idx, menu_principal)
        stdscr.refresh()

    curses.endwin()

curses.wrapper(main)
