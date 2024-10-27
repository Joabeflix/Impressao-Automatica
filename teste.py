import sys
import win32print

# Lista de impressoras
impressoras_lista = []

def listar_impressoras():
    # Obtém todas as impressoras disponíveis
    impressoras = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    imp = win32print.Enum
    print(imp)
    for impressora in impressoras:
        x, y, nome, z = impressora[:4]
        print()
        print(y)
        impressoras_lista.append(nome)

listar_impressoras()