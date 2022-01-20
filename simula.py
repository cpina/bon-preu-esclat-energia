#!/usr/bin/env python3

import argparse
from datetime import datetime

import openpyxl

acumulat_energia = {'baix': 0,
                    'mitja': 0,
                    'car': 0
                    }


def calcular_preu_dia_i_nit(data, hora, activa_entrant):
    global acumulat_energia
    hora = int(hora)

    baix = 0.079990
    mitja = 0.129990
    car = 0.229990

    preus = {1: baix, 2: baix, 3: baix, 4: baix, 5: baix, 6: baix, 7: baix, 8: mitja, 9: mitja, 10: car, 11: car,
             12: car, 13: car, 14: mitja, 15: mitja, 16: mitja, 17: mitja, 18: car, 19: car, 20: car, 21: car, 22: baix,
             23: baix, 24: baix
             }

    dt = datetime.strptime(data, "%d/%m/%Y")
    dia_de_la_setmana = dt.weekday()

    if data == "x06/12/2021" or data == "x08/12/2021":
        preu = baix
    elif dia_de_la_setmana == 5 or dia_de_la_setmana == 6:
        preu = baix
    else:
        preu = preus[hora]

    return preu * activa_entrant / 1000


def simula_dia_i_nit(fitxer):
    workbook = openpyxl.load_workbook(fitxer)
    sheet = workbook.get_active_sheet()

    row = 2

    preu_acumulat = 0

    while True:
        cups = sheet.cell(row=row, column=1).value
        if cups is None:
            break

        data = sheet.cell(row=row, column=2).value
        hora = sheet.cell(row=row, column=3).value
        activa_entrant = sheet.cell(row=row, column=4).value

        preu_acumulat += calcular_preu_dia_i_nit(data, hora, activa_entrant)

        row += 1

    return preu_acumulat


def simula(fitxer):
    preu_dia_i_nit = simula_dia_i_nit(fitxer)
    print(preu_dia_i_nit)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()

    parser.add_argument('fitxer', help='Fitxer .xlsx de Bon Preu Esclat en franges horaries')
    args = parser.parse_args()

    simula(args.fitxer)
