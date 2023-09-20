import curses
from ClassToPDF.word import *

menu_principal = ['CONVERTIR', 'COMBINAR PDFs', 'LISTA', 'EXIT']
menu_convertir = ['Word ----> PDF', 'Excel ----> PDF', 'Todo ----> PDF']


def print_menu(stdscr, selected_row_idx, list):
    stdscr.clear()
    h, w = stdscr.getmaxyx()

    for idx, row in enumerate(list):
        x = w//2 - len(row) //2
        y = h//2 - len(list)//2 + idx
        if idx == selected_row_idx:
            stdscr.attron(curses.color_pair(1))
            stdscr.addstr(y, x, row)
            stdscr.attroff(curses.color_pair(1))
        else:
            stdscr.addstr(y, x, row)

    stdscr.refresh()

def print_convertir(stdfscr):
    pass

def sec_convertir(stdscr):
    current_row_idx = 0
    print_menu(stdscr, current_row_idx, menu_convertir)

    while True:
        key = stdscr.getch()
        stdscr.clear()
        if key == 450 and current_row_idx > 0: #450 Flecha para arriba
            current_row_idx -= 1
        elif key == 456 and current_row_idx < len(menu_convertir) - 1: #456 Flecha para abajo
            current_row_idx += 1
        elif key in [10,13]:
            if current_row_idx == 0:
                pass
        elif key == 27: #27 escape
            break
        print_menu(stdscr, current_row_idx, menu_convertir)
        stdscr.refresh()



def main(stdscr):
    curses.curs_set(0)
    curses.noecho()
    curses.cbreak()
    stdscr.keypad(True)
    curses.init_pair(1, curses.COLOR_BLACK,   curses.COLOR_WHITE)
    current_row_idx = 0
    print_menu(stdscr, current_row_idx, menu_principal)

    while True:
        key = stdscr.getch()
        stdscr.clear()
        if key == 450 and current_row_idx > 0: #450 Flecha para arriba
            current_row_idx -= 1
        elif key == 456 and current_row_idx < len(menu_principal) - 1: #456 Flecha para abajo
            current_row_idx += 1
        elif key in [10,13]: #Enter
            if current_row_idx == 0: #Convertir
                sec_convertir(stdscr)
            elif current_row_idx == 1: #Combinar PDF
                pass
        elif key == 27: #27 Escape
            break
        print_menu(stdscr, current_row_idx, menu_principal)
        stdscr.refresh()
    

curses.wrapper(main)